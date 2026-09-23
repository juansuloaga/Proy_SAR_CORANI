/**
 * ============================================================================
 *  SAR CORANI - DASHBOARD API (Google Apps Script)
 * ============================================================================
 *  Este script va DENTRO del Google Sheets (Extensiones > Apps Script),
 *  como script "bound" (enlazado). No necesita ID de spreadsheet porque
 *  usa SpreadsheetApp.getActiveSpreadsheet().
 *
 *  ENDPOINTS (todos vía GET, mismo /exec ):
 *    ?action=base    -> JSON con registros de participación (anonimizado)
 *    ?action=metas   -> JSON con metas mensuales (Meta/Ejecutado/%)
 *    ?action=casos   -> JSON con casos atendidos (mapeo automático de columnas)
 *    ?action=export  -> JSON { filename, mimeType, base64 } del Excel
 *
 *  IMPORTANTE - PRIVACIDAD:
 *  Este script NUNCA envía DNI ni nombres/apellidos al cliente. La
 *  deduplicación (nombres+apellidos concatenados) se calcula aquí, en el
 *  servidor, y solo se expone un identificador anónimo (hash corto) más
 *  la bandera "esDuplicado". Así, si en el futuro compartes el dashboard
 *  fuera del equipo, lo único que circula son datos agregados/anónimos.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// 1. CONFIGURACIÓN — AJUSTA ESTO A TUS HOJAS REALES
// ---------------------------------------------------------------------------

const CONFIG = {
  SHEET_BASE: 'SHEET_PARTICIPANTES', // nombre exacto de la hoja de registros
  SHEET_METAS: 'SHEET_METAS',     // nombre exacto de la hoja de metas
  SHEET_CASOS: 'SHEET_CASOS',     // nombre exacto de la hoja de casos (AJUSTA si es distinto)

  // Nombres de columna EXACTOS tal como aparecen en la fila 1 de SHEET_BASE
  COLS_BASE: {
    FICHA: 'N°_FICHA',
    NOM_ACCION: 'NOM_ACCION',
    COD_ACCION: 'COD_ACCION',
    TIPO_ACCION: 'TIPO_ACCION',
    LOCALIDAD: 'NOMBRE_LOCALIDAD',
    FECHA: 'FECHA',
    PROFESIONAL: 'PROF_REALIZA_ACCION',
    TEMATICA: 'NOM_TEMATICA',
    DNI: 'DNI',
    A_PATERNO: 'PARTICIPANTES_A_PATERNO',
    A_MATERNO: 'PARTICIPANTES_A_MATERNO',
    NOMBRES: 'PARTICIPANTES_NOMBRES',
    SEXO: 'SEXO',
    EDAD: 'EDAD',
    COD_PARTICIPANTE: 'COD_PARTICIPANTE'
  },

  // Fila donde empiezan los datos reales en METAS (después de las 2 filas de encabezado)
  METAS_HEADER_ROWS: 2,

  // Control de acceso opcional por correo (además del control de "Implementar").
  // Déjalo como [] para no restringir aquí (recomendado si el Deploy ya está
  // limitado a "Cualquier usuario de tu organización").
  ALLOWED_EMAILS: [],

  // Antigüedad máxima (segundos) que puede tener la caché en Drive antes de
  // recalcularse. Si configuras el disparador refrescarCacheProgramado cada
  // 10 min, este valor evita servir datos más viejos que eso.
  CACHE_SECONDS: 600
};

// Grupos etarios — ajusta los rangos si tu institución usa otros cortes
const GRUPOS_ETARIOS = [
  { min: 0, max: 11, label: 'Niño/a (0-11)' },
  { min: 12, max: 17, label: 'Adolescente (12-17)' },
  { min: 18, max: 29, label: 'Joven (18-29)' },
  { min: 30, max: 59, label: 'Adulto (30-59)' },
  { min: 60, max: 200, label: 'Adulto mayor (60+)' }
];

// Grupos etarios para víctimas de Casos, según lo solicitado (3 tramos)
const GRUPOS_ETARIOS_CASOS = [
  { min: 0, max: 17, label: 'Menor de 18 años' },
  { min: 18, max: 59, label: '18 a 59 años' },
  { min: 60, max: 200, label: '60 años a más' }
];

// ---------------------------------------------------------------------------
// 2. ROUTER PRINCIPAL
// ---------------------------------------------------------------------------

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';
    const forzar = e && e.parameter && e.parameter.force === '1';

    if (CONFIG.ALLOWED_EMAILS.length > 0 && !usuarioAutorizado_()) {
      return jsonOutput_({ error: 'No autorizado' }, 403);
    }

    let payload;
    switch (action) {
      case 'base':
        payload = forzar ? refrescarYObtener_('base', getBaseData_) : getBaseDataCacheada_();
        break;
      case 'metas':
        payload = forzar ? refrescarYObtener_('metas', getMetasData_) : getMetasDataCacheada_();
        break;
      case 'casos':
        payload = forzar ? refrescarYObtener_('casos', getCasosData_) : getCasosDataCacheada_();
        break;
      case 'export':
        payload = generarExportExcel_(); // no se cachea: siempre fresco
        break;
      default:
        payload = {
          error: 'Acción no reconocida. Usa ?action=base | metas | export'
        };
    }
    return jsonOutput_(payload);
  } catch (err) {
    return jsonOutput_({ error: String(err) }, 500);
  }
}

function usuarioAutorizado_() {
  try {
    const email = Session.getActiveUser().getEmail();
    return CONFIG.ALLOWED_EMAILS.indexOf(email) !== -1;
  } catch (e) {
    // Si el deploy es "Ejecutar como: Yo", Session.getActiveUser() puede
    // venir vacío para el visitante; en ese caso, la seguridad depende
    // de la configuración de acceso del propio Deploy (ver notas).
    return true;
  }
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// 2.1 CACHÉ RÁPIDA (CacheService, fraccionada en trozos de <100KB)
// ---------------------------------------------------------------------------
// CacheService es mucho más rápido que Drive (no hace llamadas a la API de
// Drive, que puede tardar segundos y a veces fallar). Su única limitación es
// 100KB por valor, así que si el JSON es más grande, lo partimos en varios
// trozos bajo esa misma clave y los volvemos a unir al leer.

function guardarEnCache_(key, data) {
  const cache = CacheService.getScriptCache();
  const json = JSON.stringify(data);
  const TAM_TROZO = 90000; // caracteres por trozo, con margen bajo el límite de 100KB
  const n = Math.max(1, Math.ceil(json.length / TAM_TROZO));
  try {
    for (let i = 0; i < n; i++) {
      cache.put(key + '_' + i, json.substring(i * TAM_TROZO, (i + 1) * TAM_TROZO), CONFIG.CACHE_SECONDS);
    }
    cache.put(key + '_n', String(n), CONFIG.CACHE_SECONDS);
  } catch (e) { /* si falla el guardado, igual se devuelven los datos frescos al usuario */ }
}

function leerDeCache_(key) {
  const cache = CacheService.getScriptCache();
  const nStr = cache.get(key + '_n');
  if (!nStr) return null;
  const n = parseInt(nStr, 10);
  const partes = [];
  for (let i = 0; i < n; i++) {
    const p = cache.get(key + '_' + i);
    if (p === null) return null; // algún trozo expiró o no existe: se recalcula todo
    partes.push(p);
  }
  try { return JSON.parse(partes.join('')); } catch (e) { return null; }
}

function getBaseDataCacheada_() {
  const cache = leerDeCache_('base');
  if (cache) return cache;
  const data = getBaseData_();
  guardarEnCache_('base', data);
  return data;
}

function getMetasDataCacheada_() {
  const cache = leerDeCache_('metas');
  if (cache) return cache;
  const data = getMetasData_();
  guardarEnCache_('metas', data);
  return data;
}

function getCasosDataCacheada_() {
  const cache = leerDeCache_('casos');
  if (cache) return cache;
  const data = getCasosData_();
  guardarEnCache_('casos', data);
  return data;
}

/**
 * Ignora la caché por completo: relee la hoja tal cual está AHORA MISMO y
 * renueva la caché con ese resultado fresco. La usan los botones
 * "Actualizar" del dashboard (?action=X&force=1) para que agregar filas en
 * Google Sheets se refleje de inmediato, sin esperar a que la caché expire.
 */
function refrescarYObtener_(key, fnCalcular) {
  const data = fnCalcular();
  guardarEnCache_(key, data);
  return data;
}

/**
 * Recalcula y guarda la caché de las 3 hojas SIN esperar a que alguien abra
 * el dashboard. Configúrala como disparador de tiempo (recomendado: cada 10
 * minutos) para que el dashboard casi nunca tenga que esperar:
 * Apps Script → ícono de reloj (Activadores) → Añadir activador →
 * función: refrescarCacheProgramado → Basado en tiempo → cada 10 minutos.
 */
function refrescarCacheProgramado() {
  guardarEnCache_('base', getBaseData_());
  guardarEnCache_('metas', getMetasData_());
  try { guardarEnCache_('casos', getCasosData_()); } catch (e) { /* aún no configurado, se omite sin romper el resto */ }
}

// ---------------------------------------------------------------------------
// 3. LECTURA Y TRANSFORMACIÓN — BASE DE REGISTROS
// ---------------------------------------------------------------------------

function getBaseData_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_BASE);
  if (!sheet) throw new Error('No se encontró la hoja ' + CONFIG.SHEET_BASE);

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1);

  const idx = {};
  Object.keys(CONFIG.COLS_BASE).forEach(function (key) {
    idx[key] = headers.indexOf(CONFIG.COLS_BASE[key]);
  });

  const vistos = {}; // clave única -> true una vez vista (para marcar duplicados)
  const registros = [];

  rows.forEach(function (row) {
    // saltar filas sin N° de ficha (vacías, de relleno, o con solo espacios en alguna celda)
    const fichaVal = row[idx.FICHA];
    if (fichaVal === '' || fichaVal === null || fichaVal === undefined) return;

    const paterno = String(row[idx.A_PATERNO] || '').trim();
    const materno = String(row[idx.A_MATERNO] || '').trim();
    const nombres = String(row[idx.NOMBRES] || '').trim();

    const claveUnica = normalizarTexto_(paterno + '|' + materno + '|' + nombres);
    const esDuplicado = !!vistos[claveUnica];
    vistos[claveUnica] = true;

    const edad = Number(row[idx.EDAD]) || null;
    const fecha = row[idx.FECHA] instanceof Date
      ? row[idx.FECHA]
      : new Date(row[idx.FECHA]);

    registros.push({
      ficha: row[idx.FICHA],
      nomAccion: row[idx.NOM_ACCION],
      codAccion: row[idx.COD_ACCION],
      tipoAccion: row[idx.TIPO_ACCION],
      localidad: row[idx.LOCALIDAD],
      fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      anio: fecha.getFullYear(),
      mes: fecha.getMonth() + 1,
      trimestre: Math.ceil((fecha.getMonth() + 1) / 3),
      profesional: row[idx.PROFESIONAL],
      tematica: row[idx.TEMATICA],
      sexo: row[idx.SEXO],
      edad: edad,
      grupoEtario: calcularGrupoEtario_(edad),
      codParticipante: row[idx.COD_PARTICIPANTE],
      // ID anónimo estable (hash corto) en vez de nombre/DNI real
      idParticipante: hashCorto_(claveUnica),
      esDuplicado: esDuplicado
      // NOTA: A PROPÓSITO no se incluyen DNI, apellidos ni nombres aquí.
    });
  });

  return {
    actualizado: new Date().toISOString(),
    totalFilas: registros.length,
    registros: registros
  };
}

function normalizarTexto_(str) {
  return String(str)
    .toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // quita tildes
    .replace(/\s+/g, ' ')
    .trim();
}

function hashCorto_(str) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  return digest.map(function (b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('').substring(0, 12);
}

function calcularGrupoEtario_(edad) {
  if (edad === null || isNaN(edad)) return 'No especificado';
  const grupo = GRUPOS_ETARIOS.find(function (g) { return edad >= g.min && edad <= g.max; });
  return grupo ? grupo.label : 'No especificado';
}

// ---------------------------------------------------------------------------
// 4. LECTURA Y TRANSFORMACIÓN — METAS (Prog. / Ejec. / % por mes)
// ---------------------------------------------------------------------------

function getMetasData_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_METAS);
  if (!sheet) throw new Error('No se encontró la hoja ' + CONFIG.SHEET_METAS);

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function (h) { return normalizarTexto_(h); });

  const idx = {
    anio: headers.indexOf('ANO'),                 // "AÑO" sin tilde tras normalizar
    mes: headers.indexOf('MES'),
    metaP: headers.indexOf('META_PART'),
    ejecP: headers.indexOf('EJECUTADO_PART'),
    pctP: headers.findIndex(function (h) { return h.indexOf('CUMPLIMIENTO') !== -1 && h.indexOf('PART') !== -1; }),
    metaC: headers.indexOf('META_CASOS'),
    ejecC: headers.indexOf('EJECUTADO_CASOS'),
    pctC: headers.findIndex(function (h) { return h.indexOf('CUMPLIMIENTO') !== -1 && h.indexOf('CASOS') !== -1; })
  };

  // La columna "BRECHA" se repite dos veces con el mismo texto (una para
  // Participantes y otra para Casos) — se distinguen por su POSICIÓN:
  // la primera (antes de META_CASOS) es de Participantes, la segunda de Casos.
  const columnasBrecha = [];
  headers.forEach(function (h, i) { if (h === 'BRECHA') columnasBrecha.push(i); });
  idx.brechaP = columnasBrecha.filter(function (i) { return idx.metaC === -1 || i < idx.metaC; })[0];
  idx.brechaC = columnasBrecha.filter(function (i) { return idx.metaC !== -1 && i > idx.metaC; })[0];
  if (idx.brechaP === undefined) idx.brechaP = -1;
  if (idx.brechaC === undefined) idx.brechaC = -1;

  const faltantesCriticos = ['anio', 'mes', 'metaP', 'ejecP', 'pctP'].filter(function (k) { return idx[k] === -1; });
  if (faltantesCriticos.length > 0) {
    throw new Error('Faltan columnas de Participantes en ' + CONFIG.SHEET_METAS + ': ' + faltantesCriticos.join(', '));
  }
  const tieneCasos = idx.metaC !== -1 && idx.ejecC !== -1; // si no existen, esas métricas se omiten sin romper nada

  const NOMBRES_MES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO',
    'AGOSTO','SEPTIEMBRE','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

  const metas = [];
  let anual = null;

  values.slice(1).forEach(function (row) {
    const mesTexto = String(row[idx.mes] || '').trim();
    if (mesTexto === '') return;
    const mesNum = NOMBRES_MES.indexOf(normalizarTexto_(mesTexto)) + 1; // 0 si es "ANUAL" u otro texto

    const metaP = Number(row[idx.metaP]) || 0;
    const ejecP = Number(row[idx.ejecP]) || 0;
    const pctP = parsearPorcentaje_(row[idx.pctP]);
    const brechaP = idx.brechaP !== -1 ? (Number(row[idx.brechaP]) || 0) : (ejecP - metaP);

    const registro = {
      anio: row[idx.anio],
      mes: mesTexto,
      mesNum: mesNum,
      meta: metaP,
      ejecutado: ejecP,
      porcentaje: pctP,
      brecha: brechaP,
      estado: calcularEstado_(pctP)
    };

    if (tieneCasos) {
      const metaC = Number(row[idx.metaC]) || 0;
      const ejecC = Number(row[idx.ejecC]) || 0;
      const pctC = idx.pctC !== -1 ? parsearPorcentaje_(row[idx.pctC]) : (metaC > 0 ? (ejecC / metaC * 100) : 0);
      const brechaC = idx.brechaC !== -1 ? (Number(row[idx.brechaC]) || 0) : (ejecC - metaC);
      registro.metaCasos = metaC;
      registro.ejecutadoCasos = ejecC;
      registro.porcentajeCasos = pctC;
      registro.brechaCasos = brechaC;
      registro.estadoCasos = calcularEstado_(pctC);
    }

    if (normalizarTexto_(mesTexto) === 'ANUAL' || mesNum === 0) {
      anual = registro; // fila resumen, no va en el arreglo mensual
    } else {
      metas.push(registro);
    }
  });

  return {
    actualizado: new Date().toISOString(),
    tieneCasos: tieneCasos,
    metas: metas,
    anual: anual
  };
}

function calcularEstado_(pct) {
  if (pct === null || pct === undefined || isNaN(pct)) return null;
  return pct >= 95 ? 'BUENO' : (pct >= 75 ? 'MODERADO' : 'BAJO');
}

function parsearPorcentaje_(valor) {
  if (typeof valor === 'number') {
    // Una celda con formato de porcentaje en Sheets SIEMPRE guarda la fracción
    // internamente (97.44% -> 0.9744, e incluso 159.31% -> 1.5931), sin
    // importar que el valor visible supere el 100%. Por eso aquí SIEMPRE se
    // multiplica por 100 — la condición anterior (solo si valor<=1) rompía
    // justo los meses con cumplimiento mayor a 100%.
    return valor * 100;
  }
  const limpio = String(valor || '0').replace('%', '').replace(',', '.').trim();
  return Number(limpio) || 0;
}

// ---------------------------------------------------------------------------
// 4.5 LECTURA Y TRANSFORMACIÓN — CASOS (mapeo automático de columnas)
// ---------------------------------------------------------------------------
// Como aún no se confirmó el esquema exacto de la hoja de Casos, en vez de
// asumir nombres de columna fijos (que romperían todo si no coinciden),
// buscamos cada campo esperado por PALABRAS CLAVE dentro de tus encabezados
// reales. Si un campo no se encuentra, simplemente no se incluye — el
// dashboard lo mostrará como "no disponible" en vez de fallar por completo.
// Ajusta CAMPOS_CASOS si tus encabezados usan otros términos.

// Cada campo puede tener varias listas de palabras clave candidatas (en
// orden de prioridad); se usa la primera que encuentre coincidencia.
const CAMPOS_CASOS = {
  ficha:          [['FICHA']],
  fecha:          [['FECHA', 'INGRESO'], ['FECHA']],           // evita confundir con FECHA_NAC
  centroPoblado:  [['CCPP'], ['CENTRO', 'POBLADO'], ['RESIDENCIA'], ['LOCALIDAD']],
  vinculo:        [['VINCULO']],                                 // VINCULO/SIN_VINCULO
  planAtencion:   [['PLAN', 'ATENCION']],
  victimaSexo:    [['VICTIMA', 'SEXO'], ['SEXO']],                // "SEXO" (no AGRESOR_SEXO, que va después en la hoja)
  victimaEdad:    [['VICTIMA', 'EDAD'], ['EDAD']],                // "EDAD" (no AGRESOR_EDAD)
  victimaTrabaja: [['VICTIMA', 'TRABAJA'], ['TRABAJO'], ['TRABAJA']],
  nivelEducativo: [['NIVEL', 'EDUCATIVO']],
  tipoViolencia:  [['TIPO', 'VIOLENCIA']]
};

function buscarColumna_(headersNorm, listaCandidatos) {
  for (let i = 0; i < listaCandidatos.length; i++) {
    const palabrasClave = listaCandidatos[i];
    for (let c = 0; c < headersNorm.length; c++) {
      const coincideTodo = palabrasClave.every(function (palabra) {
        return headersNorm[c].indexOf(palabra) !== -1;
      });
      if (coincideTodo) return c;
    }
  }
  return -1;
}

function getCasosData_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CASOS);
  if (!sheet) throw new Error('No se encontró la hoja ' + CONFIG.SHEET_CASOS + ' (revisa CONFIG.SHEET_CASOS)');

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { actualizado: new Date().toISOString(), registros: [], camposDetectados: [], camposFaltantes: Object.keys(CAMPOS_CASOS) };

  const headersNorm = values[0].map(normalizarTexto_);

  const idx = {};
  const camposDetectados = [];
  const camposFaltantes = [];
  Object.keys(CAMPOS_CASOS).forEach(function (campo) {
    const c = buscarColumna_(headersNorm, CAMPOS_CASOS[campo]);
    idx[campo] = c;
    (c === -1 ? camposFaltantes : camposDetectados).push(campo);
  });

  // sin columna de ficha no hay forma confiable de identificar filas válidas
  if (idx.ficha === -1) {
    throw new Error('No se encontró una columna de FICHA en ' + CONFIG.SHEET_CASOS + '. Columnas detectadas: ' + headersNorm.join(' | '));
  }

  const registros = [];
  values.slice(1).forEach(function (row) {
    const fichaVal = row[idx.ficha];
    if (fichaVal === '' || fichaVal === null || fichaVal === undefined) return;

    const reg = { ficha: fichaVal };

    if (idx.fecha !== -1) {
      const f = row[idx.fecha] instanceof Date ? row[idx.fecha] : new Date(row[idx.fecha]);
      if (!isNaN(f)) {
        reg.fecha = Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        reg.anio = f.getFullYear();
        reg.mes = f.getMonth() + 1;
        reg.trimestre = Math.ceil((f.getMonth() + 1) / 3);
      }
    }

    if (idx.centroPoblado !== -1) reg.centroPoblado = row[idx.centroPoblado];
    if (idx.vinculo !== -1) reg.vinculo = row[idx.vinculo];
    if (idx.planAtencion !== -1) reg.planAtencion = row[idx.planAtencion];
    if (idx.victimaSexo !== -1) reg.victimaSexo = row[idx.victimaSexo];
    if (idx.victimaTrabaja !== -1) reg.victimaTrabaja = row[idx.victimaTrabaja];
    if (idx.nivelEducativo !== -1) reg.nivelEducativo = row[idx.nivelEducativo];
    if (idx.tipoViolencia !== -1) reg.tipoViolencia = row[idx.tipoViolencia];
    if (idx.victimaEdad !== -1) {
      const edad = Number(row[idx.victimaEdad]) || null;
      reg.victimaEdad = edad;
      reg.victimaGrupoEtario = calcularGrupoEtarioCaso_(edad);
    }

    registros.push(reg);
  });

  return {
    actualizado: new Date().toISOString(),
    totalFilas: registros.length,
    camposDetectados: camposDetectados,   // para que puedas verificar qué sí se mapeó bien
    camposFaltantes: camposFaltantes,     // y qué falta (dime los nombres reales y los agrego)
    registros: registros
  };
}

function calcularGrupoEtarioCaso_(edad) {
  if (edad === null || isNaN(edad)) return 'No especificado';
  const grupo = GRUPOS_ETARIOS_CASOS.find(function (g) { return edad >= g.min && edad <= g.max; });
  return grupo ? grupo.label : 'No especificado';
}

// ---------------------------------------------------------------------------
// 5. EXPORTACIÓN A EXCEL (descarga directa desde Apps Script)
// ---------------------------------------------------------------------------

function generarExportExcel_() {
  const base = getBaseData_();
  const metas = getMetasData_();
  let casos = { registros: [] };
  try { casos = getCasosData_(); } catch (e) { /* aún no configurado */ }

  const temp = SpreadsheetApp.create('export_temp_' + new Date().getTime());

  // --- Hoja Participantes (agregada/anonimizada, sin nombres ni DNI) ---
  const hojaP = temp.getSheets()[0];
  hojaP.setName('Participantes');
  const encabezadosP = [
    'Ficha', 'Accion', 'CodAccion', 'TipoAccion', 'Localidad', 'Sesion',
    'Fecha', 'Anio', 'Mes', 'Trimestre', 'Profesional', 'Tematica',
    'Sexo', 'Edad', 'GrupoEtario', 'Lengua', 'CodParticipante',
    'IdParticipanteAnonimo', 'EsDuplicado'
  ];
  const filasP = base.registros.map(function (r) {
    return [
      r.ficha, r.nomAccion, r.codAccion, r.tipoAccion, r.localidad, r.sesion,
      r.fecha, r.anio, r.mes, r.trimestre, r.profesional, r.tematica,
      r.sexo, r.edad, r.grupoEtario, r.lengua, r.codParticipante,
      r.idParticipante, r.esDuplicado
    ];
  });
  if (filasP.length > 0) {
    hojaP.getRange(1, 1, 1, encabezadosP.length).setValues([encabezadosP]);
    hojaP.getRange(2, 1, filasP.length, encabezadosP.length).setValues(filasP);
  }

  // --- Hoja Metas (Participantes + Casos) ---
  const hojaM = temp.insertSheet('Metas');
  const encabezadosM = ['Año','Mes','Meta_Part','Ejecutado_Part','% Cumpl_Part','Brecha_Part','Estado_Part',
    'Meta_Casos','Ejecutado_Casos','% Cumpl_Casos','Brecha_Casos','Estado_Casos'];
  const filaMetaCasos = function (m) {
    return metas.tieneCasos
      ? [m.metaCasos, m.ejecutadoCasos, m.porcentajeCasos, m.brechaCasos, m.estadoCasos]
      : ['', '', '', '', ''];
  };
  const filasM = metas.metas.map(function (m) {
    return [m.anio, m.mes, m.meta, m.ejecutado, m.porcentaje, m.brecha, m.estado].concat(filaMetaCasos(m));
  });
  if (metas.anual) {
    filasM.push([metas.anual.anio, 'ANUAL', metas.anual.meta, metas.anual.ejecutado, metas.anual.porcentaje, metas.anual.brecha, metas.anual.estado].concat(filaMetaCasos(metas.anual)));
  }
  if (filasM.length > 0) {
    hojaM.getRange(1, 1, 1, encabezadosM.length).setValues([encabezadosM]);
    hojaM.getRange(2, 1, filasM.length, encabezadosM.length).setValues(filasM);
  }

  // --- Hoja Casos (si ya está configurada) ---
  if (casos.registros.length > 0) {
    const hojaC = temp.insertSheet('Casos');
    const encabezadosC = ['Ficha','Fecha','CentroPoblado','Vinculo','PlanAtencion','VictimaSexo','VictimaEdad','VictimaGrupoEtario','VictimaTrabaja','NivelEducativo','TipoViolencia'];
    const filasC = casos.registros.map(function (r) {
      return [r.ficha, r.fecha||'', r.centroPoblado||'', r.vinculo||'', r.planAtencion||'', r.victimaSexo||'', r.victimaEdad||'', r.victimaGrupoEtario||'', r.victimaTrabaja||'', r.nivelEducativo||'', r.tipoViolencia||''];
    });
    hojaC.getRange(1, 1, 1, encabezadosC.length).setValues([encabezadosC]);
    hojaC.getRange(2, 1, filasC.length, encabezadosC.length).setValues(filasC);
  }

  SpreadsheetApp.flush();

  // Convertir el spreadsheet temporal a bytes .xlsx y codificar en base64
  const file = DriveApp.getFileById(temp.getId());
  const blob = file.getAs(MimeType.MICROSOFT_EXCEL);
  const base64 = Utilities.base64Encode(blob.getBytes());

  // Limpieza: eliminar el archivo temporal de Drive
  file.setTrashed(true);

  return {
    filename: 'SAR_Corani_Dashboard_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    base64: base64
  };
}
