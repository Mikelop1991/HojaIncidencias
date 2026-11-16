/**
 * REGISTRO DE INCIDENCIAS · ESCUELA 26-27
 * Backend (Google Apps Script) de la app móvil de incidencias.
 *
 * Trabaja directamente sobre la hoja de cálculo de Google donde está el
 * registro mes a mes. Hojas que usa:
 *   - "Octubre", "Noviembre", ... "Junio"  -> una fila por responsable+monitor (las mismas de siempre)
 *   - "Responsables"                      -> datos propios de cada responsable por mes
 *   - "_Accesos"    (se crea)  -> código de acceso de cada responsable y qué grupos ve
 *   - "_Monitores"  (se crea)  -> lista maestra de monitores por grupo y sus días fijos (L M X J V)
 *   - "_Festivos"   (se crea)  -> días sin actividad (no se premarcan en el calendario)
 *   - "_Ajustes"    (se crea)  -> temporada y meses cerrados
 *   - "_Dias"       (se crea)  -> qué días concretos marcó cada responsable (para poder reabrir)
 *   - "_Registro"   (se crea)  -> historial de quién guardó qué y cuándo
 *
 * Primer uso: ejecutar la función configurarInicial() desde el editor.
 */

/* ================================================================
 * URL DE LA EXCEL DEL REGISTRO
 * - Si has abierto este código desde la Excel (Extensiones → Apps Script),
 *   déjalo vacío: ''.
 * - Si el código está en un proyecto aparte, pega aquí la dirección de la
 *   Excel, por ejemplo:
 *   'https://docs.google.com/spreadsheets/d/XXXXXXXX/edit'
 * ================================================================ */
var URL_HOJA_REGISTRO = '';

var MESES = ['Octubre', 'Noviembre', 'Diciembre', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio'];
var MES_NUM = { Octubre: 10, Noviembre: 11, Diciembre: 12, Enero: 1, Febrero: 2, Marzo: 3, Abril: 4, Mayo: 5, Junio: 6 };
var DIAS = ['L', 'M', 'X', 'J', 'V'];

var H_ACCESOS = ['Código', 'Persona', 'Grupos (separados por ; o * = todos)', 'Nombre en hoja Responsables', 'Admin', 'Teléfono (WhatsApp)'];
var H_MONITORES = ['Grupo', 'Nombre', 'Primer Apellido', 'Segundo Apellido', 'L', 'M', 'X', 'J', 'V', 'Activo'];
var H_FESTIVOS = ['Fecha', 'Descripción'];
var H_AJUSTES = ['Clave', 'Valor', 'Nota'];
var H_DIAS = ['Mes', 'Grupo', 'Monitor', 'Días', 'Actualizado', 'Por'];
var H_REGISTRO = ['Fecha', 'Persona', 'Acción', 'Detalle'];

var CABECERA_MES = ['Responsable', 'Nombre', 'Primer Apellido', 'Segundo Apellido',
  'LUNES\n(nº viajes)', 'MARTES\n(nº viajes)', 'MIÉRCOLES\n(nº viajes)', 'JUEVES\n(nº viajes)', 'VIERNES\n(nº viajes)',
  'Nº PARTIDOS\nLIGA ', 'AUSENCIAS\n(Nº)', 'SUSTITUCIONES \n(Nº)',
  'KILÓMETROS ESPECIALES\nQUE LLEVES COCHE \n(formacion, liga, partidos, etc.)', 'PEAJES', 'OBSERVACIONES\nSUSTITUCIONES'];

// ---------------------------------------------------------------- utilidades

var _registro = null;
/** Hoja compartida del registro (meses + Responsables). */
function ss_() {
  if (!_registro) _registro = URL_HOJA_REGISTRO ? SpreadsheetApp.openByUrl(URL_HOJA_REGISTRO) : SpreadsheetApp.getActiveSpreadsheet();
  if (!_registro) throw new Error('Este proyecto de Apps Script no está dentro de la Excel. Ábrelo desde la Excel (Extensiones → Apps Script) o pega la dirección de la Excel en URL_HOJA_REGISTRO, arriba del todo.');
  return _registro;
}
/** Hoja donde están _Accesos, _Monitores... (la propia Excel, salvo que el script viva en otra hoja). */
function cfg_() { return SpreadsheetApp.getActiveSpreadsheet() || ss_(); }

function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

function claveMonitor_(nombre, ap1, ap2) { return norm_(nombre) + '|' + norm_(ap1) + '|' + norm_(ap2); }

function num_(v) {
  if (v === '' || v == null) return null;
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? null : n;
}

function esVerdadero_(v) {
  if (v === true) return true;
  var s = norm_(v);
  return s === 'TRUE' || s === 'VERDADERO' || s === 'SI' || s === 'X' || s === '1';
}

function hoja_(nombre, cabecera) {
  var sh = cfg_().getSheetByName(nombre);
  if (!sh && cabecera) {
    sh = cfg_().insertSheet(nombre);
    sh.getRange(1, 1, 1, cabecera.length).setValues([cabecera]).setFontWeight('bold').setBackground('#1f2a44').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function filas_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

/** Lecturas de configuración en caché 60 s (acelera mucho cada llamada). */
function cacheado_(clave, fn) {
  var c = CacheService.getScriptCache(), v = c.get(clave);
  if (v) return JSON.parse(v);
  var r = fn();
  try { c.put(clave, JSON.stringify(r), 60); } catch (e) { }
  return r;
}
function filasCache_(nombre) {
  return cacheado_('f_' + nombre, function () {
    return filas_(hoja_(nombre)).map(function (r) { return r.map(function (v) { return v instanceof Date ? { d: v.getTime() } : v; }); });
  });
}
function limpiarCache_() { CacheService.getScriptCache().removeAll(['f__Accesos', 'f__Monitores', 'f__Festivos', 'f__Ajustes']); }

function ajustes_() {
  var out = { TEMPORADA_INICIO: 2026, MESES_CERRADOS: '' };
  filasCache_('_Ajustes').forEach(function (r) { if (r[0]) out[String(r[0]).trim()] = r[1]; });
  out.TEMPORADA_INICIO = parseInt(out.TEMPORADA_INICIO, 10) || 2026;
  out.cerrados = String(out.MESES_CERRADOS || '').split(/[;,]/).map(function (s) { return norm_(s); }).filter(String);
  return out;
}

function anioDeMes_(mes, aj) { return MES_NUM[mes] >= 9 ? aj.TEMPORADA_INICIO : aj.TEMPORADA_INICIO + 1; }

function log_(persona, accion, detalle) {
  try { hoja_('_Registro', H_REGISTRO).appendRow([new Date(), persona, accion, detalle]); } catch (e) { }
}

/** Localiza columnas de una hoja de mes por el texto de su cabecera (con valores por defecto). */
function columnasMes_(sh) {
  var cab = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 15)).getValues()[0].map(norm_);
  function busca(txt, def) {
    for (var i = 0; i < cab.length; i++) if (cab[i].indexOf(txt) >= 0) return i;
    return def;
  }
  return {
    grupo: 0, nombre: 1, ap1: 2, ap2: 3,
    dias: [busca('LUNES', 4), busca('MARTES', 5), busca('MIERCOLES', 6), busca('JUEVES', 7), busca('VIERNES', 8)],
    partidos: busca('PARTIDOS', 9), ausencias: busca('AUSENCIAS', 10), sust: busca('SUSTITUCIONES (', 11),
    km: busca('KILOMETROS', 12), peajes: busca('PEAJES', 13), obs: busca('OBSERVACIONES', 14)
  };
}

// ---------------------------------------------------------------- configuración inicial

/**
 * Crea las hojas de configuración a partir de los datos que ya hay.
 * Se puede volver a ejecutar: solo añade lo que falte, nunca borra.
 */
var HOJAS_PRIVADAS = ['_Accesos', '_Monitores', '_Festivos', '_Ajustes', '_Dias', '_Registro'];

/** Mueve las hojas de configuración de la hoja compartida a la privada (si estaban allí). */
function moverConfiguracionAPrivada_() {
  var reg = ss_(), cfg = cfg_(), movidas = [];
  if (reg.getId() === cfg.getId()) return movidas;
  HOJAS_PRIVADAS.forEach(function (n) {
    var vieja = reg.getSheetByName(n);
    if (!vieja) return;
    if (!cfg.getSheetByName(n)) { vieja.copyTo(cfg).setName(n); movidas.push(n); }
    reg.deleteSheet(vieja);
  });
  var s1 = cfg.getSheetByName('Hoja 1') || cfg.getSheetByName('Sheet1');
  if (s1 && cfg.getSheets().length > 1 && s1.getLastRow() === 0) cfg.deleteSheet(s1);
  return movidas;
}

/** Festivos y días sin escuela del calendario oficial 26-27 (solo lunes a viernes). */
var FESTIVOS_2627 = [
  ['2026-10-01', 'Antes del inicio de la EdF'], ['2026-10-12', 'Festivo (Pilar)'],
  ['2026-11-02', 'Sin actividad'],
  ['2026-12-01', 'Sin actividad'], ['2026-12-02', 'Sin actividad'], ['2026-12-03', 'San Francisco Javier'], ['2026-12-04', 'Sin actividad'],
  ['2026-12-07', 'Puente'], ['2026-12-08', 'Inmaculada'],
  ['2026-12-23', 'Vacaciones de Navidad'], ['2026-12-24', 'Vacaciones de Navidad'], ['2026-12-25', 'Navidad'],
  ['2026-12-28', 'Vacaciones de Navidad'], ['2026-12-29', 'Vacaciones de Navidad'], ['2026-12-30', 'Vacaciones de Navidad'], ['2026-12-31', 'Vacaciones de Navidad'],
  ['2027-01-01', 'Año Nuevo'], ['2027-01-04', 'Vacaciones de Navidad'], ['2027-01-05', 'Vacaciones de Navidad'], ['2027-01-06', 'Reyes'],
  ['2027-01-07', 'Vacaciones de Navidad'], ['2027-01-08', 'Vacaciones de Navidad'],
  ['2027-02-08', 'Sin actividad (semana blanca)'], ['2027-02-09', 'Sin actividad (semana blanca)'], ['2027-02-10', 'Sin actividad (semana blanca)'],
  ['2027-02-11', 'Sin actividad (semana blanca)'], ['2027-02-12', 'Sin actividad (semana blanca)'],
  ['2027-03-19', 'San José'], ['2027-03-23', 'Semana Santa'], ['2027-03-24', 'Semana Santa'], ['2027-03-25', 'Jueves Santo'], ['2027-03-26', 'Viernes Santo'],
  ['2027-03-29', 'Lunes de Pascua'], ['2027-03-30', 'Semana Santa'], ['2027-03-31', 'Semana Santa'], ['2027-04-01', 'Semana Santa'], ['2027-04-02', 'Semana Santa'],
  ['2027-05-17', 'Fin de la EdF'], ['2027-05-18', 'Fin de la EdF'], ['2027-05-19', 'Fin de la EdF'], ['2027-05-20', 'Fin de la EdF'], ['2027-05-21', 'Fin de la EdF'],
  ['2027-05-24', 'Fin de la EdF'], ['2027-05-25', 'Fin de la EdF'], ['2027-05-26', 'Fin de la EdF'], ['2027-05-27', 'Fin de la EdF'], ['2027-05-28', 'Fin de la EdF'],
  ['2027-05-31', 'Fin de la EdF']
];

/** Sustituye la hoja _Festivos por los días del calendario oficial 26-27. */
function cargarFestivosCalendario() {
  var sh = hoja_('_Festivos', H_FESTIVOS);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  var f = FESTIVOS_2627.map(function (x) { var p = x[0].split('-'); return [new Date(+p[0], +p[1] - 1, +p[2]), x[1]]; });
  sh.getRange(2, 1, f.length, 2).setValues(f);
  sh.getRange(2, 1, f.length, 1).setNumberFormat('dd/mm/yyyy');
  limpiarCache_();
  return f.length + ' días sin actividad cargados.';
}

function configurarInicial() {
  var movidas = moverConfiguracionAPrivada_();
  var ss = ss_();
  var mesesExistentes = MESES.filter(function (m) { return ss.getSheetByName(m); });
  if (!mesesExistentes.length) throw new Error('No encuentro ninguna hoja de mes (Octubre, Noviembre...).');

  // --- Monitores y días fijos (unión de todos los meses existentes)
  var shMon = hoja_('_Monitores', H_MONITORES);
  var yaMon = {};
  filas_(shMon).forEach(function (r) { yaMon[norm_(r[0]) + '#' + claveMonitor_(r[1], r[2], r[3])] = true; });
  var nuevos = {}, orden = [];
  mesesExistentes.forEach(function (m) {
    var sh = ss.getSheetByName(m), c = columnasMes_(sh);
    filas_(sh).forEach(function (r) {
      if (!r[c.grupo] || !r[c.nombre]) return;
      var k = norm_(r[c.grupo]) + '#' + claveMonitor_(r[c.nombre], r[c.ap1], r[c.ap2]);
      if (yaMon[k]) return;
      if (!nuevos[k]) { nuevos[k] = [String(r[c.grupo]).trim(), r[c.nombre], r[c.ap1], r[c.ap2], false, false, false, false, false, true]; orden.push(k); }
      c.dias.forEach(function (col, i) {
        var v = r[col];
        if (v !== '' && v != null && v !== 0) nuevos[k][4 + i] = true;
      });
    });
  });
  if (orden.length) {
    var start = shMon.getLastRow() + 1;
    shMon.getRange(start, 1, orden.length, H_MONITORES.length).setValues(orden.map(function (k) { return nuevos[k]; }));
    shMon.getRange(start, 5, orden.length, 6).insertCheckboxes();
  }

  // --- Accesos: un código por responsable (los grupos con el mismo nombre de persona se juntan)
  var shAcc = hoja_('_Accesos', H_ACCESOS);
  if (!String(shAcc.getRange(1, 6).getValue() || '').trim()) shAcc.getRange(1, 6).setValue(H_ACCESOS[5]).setFontWeight('bold').setBackground('#1f2a44').setFontColor('#ffffff');
  var accExist = filas_(shAcc);
  var gruposCubiertos = {};
  accExist.forEach(function (r) { String(r[2]).split(';').forEach(function (g) { gruposCubiertos[norm_(g)] = true; }); });
  var nombresResp = {};
  filas_(ss.getSheetByName('Responsables')).forEach(function (r) { if (r[1]) nombresResp[norm_(r[1])] = String(r[1]).trim(); });
  var porPersona = {}, ordenP = [];
  filas_(shMon).forEach(function (r) {
    var g = String(r[0]).trim();
    if (!g || gruposCubiertos[norm_(g)]) return;
    var palabras = norm_(g).split(' ');
    var persona = palabras.slice(0, 2).join(' ');
    if (!porPersona[persona]) { porPersona[persona] = []; ordenP.push(persona); }
    if (porPersona[persona].indexOf(g) < 0) porPersona[persona].push(g);
  });
  var filasAcc = ordenP.map(function (p) {
    var nombreResp = nombresResp[p] || '';
    var bonito = nombreResp || p.toLowerCase().replace(/(^|\s)\S/g, function (t) { return t.toUpperCase(); });
    return [nuevoCodigo_(), bonito, porPersona[p].join('; '), nombreResp, false, ''];
  });
  if (!accExist.some(function (r) { return esVerdadero_(r[4]); })) {
    filasAcc.unshift([nuevoCodigo_(), 'Administración', '*', '', true, '']);
  }
  if (filasAcc.length) {
    var st = shAcc.getLastRow() + 1;
    shAcc.getRange(st, 1, filasAcc.length, H_ACCESOS.length).setValues(filasAcc);
    shAcc.getRange(st, 5, filasAcc.length, 1).insertCheckboxes();
  }

  // --- Festivos, ajustes y hojas auxiliares
  var shFes = hoja_('_Festivos', H_FESTIVOS);
  if (shFes.getLastRow() < 2) cargarFestivosCalendario();
  var shAj = hoja_('_Ajustes', H_AJUSTES);
  if (shAj.getLastRow() < 2) {
    shAj.getRange(2, 1, 2, 3).setValues([
      ['TEMPORADA_INICIO', 2026, 'Año en que empieza la temporada (octubre)'],
      ['MESES_CERRADOS', '', 'Meses que ya no se pueden editar desde la app, separados por ; (ej: Octubre; Noviembre)']
    ]);
  }
  hoja_('_Dias', H_DIAS);
  hoja_('_Registro', H_REGISTRO);
  limpiarCache_();
  return 'Configuración lista: ' + orden.length + ' monitores y ' + filasAcc.length + ' accesos nuevos.' +
    (movidas.length ? ' Movidas a esta hoja privada: ' + movidas.join(', ') + '.' : '');
}

function nuevoCodigo_() {
  var abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 6; i++) s += abc.charAt(Math.floor(Math.random() * abc.length));
  return s;
}

// ---------------------------------------------------------------- acceso

function usuario_(codigo) {
  var c = norm_(codigo);
  if (!c) throw new Error('Introduce tu código.');
  var filas = filasCache_('_Accesos');
  for (var i = 0; i < filas.length; i++) {
    var r = filas[i];
    if (norm_(r[0]) === c) {
      var admin = esVerdadero_(r[4]) || String(r[2]).trim() === '*';
      var grupos;
      if (admin) {
        var set = {};
        MESES.forEach(function (mm) { var sh = ss_().getSheetByName(mm); if (sh) filas_(sh).forEach(function (m) { if (m[0]) set[String(m[0]).trim()] = 1; }); });
        filasCache_('_Monitores').forEach(function (m) { if (m[0]) set[String(m[0]).trim()] = 1; });
        grupos = Object.keys(set).sort();
      } else {
        grupos = String(r[2]).split(';').map(function (g) { return g.trim(); }).filter(String);
      }
      // nombres de hoja Responsables a los que puede acceder en "Mis datos"
      var propios = admin
        ? uniq_(filas_(ss_().getSheetByName('Responsables')).map(function (x) { return String(x[1] || '').trim(); }).filter(String))
        : (String(r[3]).trim() ? [String(r[3]).trim()] : []);
      var contactos = admin ? filas.filter(function (x) { return x[0] && !(esVerdadero_(x[4]) || String(x[2]).trim() === '*'); }).map(function (x) {
        return { persona: String(x[1]).trim(), codigo: String(x[0]).trim(), telefono: String(x[5] || '').replace(/[^0-9+]/g, ''),
          grupos: String(x[2]).split(';').map(function (g) { return g.trim(); }).filter(String) };
      }) : [];
      return { persona: String(r[1]).trim(), grupos: grupos, admin: admin, responsables: propios, contactos: contactos };
    }
  }
  throw new Error('Código no válido.');
}

function uniq_(a) { var s = {}; return a.filter(function (x) { return s[x] ? false : (s[x] = true); }); }

function comprobarGrupo_(u, grupo) {
  if (!u.grupos.some(function (g) { return norm_(g) === norm_(grupo); })) throw new Error('No tienes acceso a ese grupo.');
}

function comprobarMes_(mes, aj, escribir) {
  if (MESES.indexOf(mes) < 0) throw new Error('Mes no válido.');
  if (escribir && aj.cerrados.indexOf(norm_(mes)) >= 0) throw new Error('El mes de ' + mes + ' está cerrado. Habla con Mikel si necesitas cambiar algo.');
}

function festivosDelMes_(mes, aj) {
  var y = anioDeMes_(mes, aj), m = MES_NUM[mes], out = {};
  filasCache_('_Festivos').forEach(function (r) {
    var d = r[0];
    if (d && d.d) d = new Date(d.d);
    if (!(d instanceof Date)) {
      var p = String(d).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (!p) return;
      d = new Date(+p[3], +p[2] - 1, +p[1]);
    }
    if (d.getFullYear() === y && d.getMonth() + 1 === m) out[d.getDate()] = String(r[1] || 'Festivo');
  });
  return out;
}

// ---------------------------------------------------------------- hoja de mes

/** Meses que salen en la app: los que ya tienen hoja + el mes en curso. */
function mesesDisponibles_() {
  var ss = ss_(), hoy = new Date().getMonth() + 1;
  return MESES.filter(function (m) { return ss.getSheetByName(m) || MES_NUM[m] === hoy; });
}

/** Hoja de mes que sirve de plantilla: la última que exista antes de "mes" (o la última que haya). */
function plantillaPara_(mes) {
  var ss = ss_(), i = MESES.indexOf(mes), tpl = null;
  for (var j = i - 1; j >= 0 && !tpl; j--) tpl = ss.getSheetByName(MESES[j]);
  for (var k = MESES.length - 1; k >= 0 && !tpl; k--) tpl = ss.getSheetByName(MESES[k]);
  return tpl;
}

/** Devuelve la hoja del mes; si no existe y crear=true, la crea DUPLICANDO la del mes anterior
 *  (mismo formato, anchos y lista de monitores) y vacía las columnas de datos. */
function hojaMes_(mes, crear) {
  var ss = ss_(), sh = ss.getSheetByName(mes);
  if (sh || !crear) return sh;
  var tpl = plantillaPara_(mes);
  if (tpl) {
    sh = tpl.copyTo(ss).setName(mes);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(tpl.getIndex() + 1);
    var c = columnasMes_(sh), cols = colsDatos_(c);
    if (sh.getLastRow() > 1) sh.getRange(2, cols.desde + 1, sh.getLastRow() - 1, cols.hasta - cols.desde + 1).clearContent();
  } else {
    sh = ss.insertSheet(mes);
    sh.getRange(1, 1, 1, CABECERA_MES.length).setValues([CABECERA_MES]).setFontWeight('bold').setWrap(true);
    sh.setFrozenRows(1);
  }
  log_('App', 'Crear hoja', mes + (tpl ? ' (copia de ' + tpl.getName() + ')' : ''));
  return sh;
}

/** Rango de columnas de datos (E..O normalmente), 0-based. */
function colsDatos_(c) {
  var todas = c.dias.concat([c.partidos, c.ausencias, c.sust, c.km, c.peajes, c.obs]);
  return { desde: Math.min.apply(null, todas), hasta: Math.max.apply(null, todas) };
}

/** Busca (o añade) la fila de un monitor dentro de la hoja de mes. Devuelve nº de fila (1-based). */
function filaMonitor_(sh, c, grupo, m, crear) {
  var datos = filas_(sh), k = claveMonitor_(m.nombre, m.ap1, m.ap2), g = norm_(grupo);
  var ultimaDelGrupo = -1;
  for (var i = 0; i < datos.length; i++) {
    if (norm_(datos[i][c.grupo]) === g) {
      ultimaDelGrupo = i;
      if (claveMonitor_(datos[i][c.nombre], datos[i][c.ap1], datos[i][c.ap2]) === k) return i + 2;
    }
  }
  if (!crear) return -1;
  // insertar justo debajo del último monitor del grupo (o al final)
  var fila;
  if (ultimaDelGrupo >= 0) { fila = ultimaDelGrupo + 3; sh.insertRowAfter(ultimaDelGrupo + 2); }
  else { fila = Math.max(sh.getLastRow(), 1) + 1; }
  sh.getRange(fila, 1, 1, 4).setValues([[grupo, m.nombre, m.ap1, m.ap2]]);
  return fila;
}

/** Días fijos por grupo+monitor desde _Monitores. */
function fijosMap_() {
  var out = {};
  filasCache_('_Monitores').forEach(function (r) {
    if (!r[1]) return;
    var f = [];
    for (var i = 0; i < 5; i++) if (esVerdadero_(r[4 + i])) f.push(i);
    out[norm_(r[0]) + '#' + claveMonitor_(r[1], r[2], r[3])] = f;
  });
  return out;
}

/** Filas del mes (o de la plantilla si el mes aún no tiene hoja). Devuelve {c, filas:[{fila, r}], existe}. */
function datosHojaMes_(mes) {
  var sh = hojaMes_(mes, false), existe = !!sh;
  if (!sh) sh = plantillaPara_(mes);
  if (!sh) return { c: null, filas: [], existe: false };
  var c = columnasMes_(sh);
  var filas = filas_(sh).map(function (r, i) { return { fila: i + 2, r: r }; })
    .filter(function (x) { return x.r[c.grupo] && x.r[c.nombre]; });
  return { c: c, filas: filas, existe: existe };
}

/** Monitores de un grupo según la hoja del mes (la plantilla manda). */
function monitoresDelGrupo_(mes, grupo, hm, fijos) {
  hm = hm || datosHojaMes_(mes); fijos = fijos || fijosMap_();
  var g = norm_(grupo), c = hm.c, vistos = {};
  if (!c) return [];
  return hm.filas.filter(function (x) { return norm_(x.r[c.grupo]) === g; }).map(function (x) {
    var r = x.r, key = claveMonitor_(r[c.nombre], r[c.ap1], r[c.ap2]);
    if (vistos[key]) return null; vistos[key] = 1;
    return { key: key, nombre: String(r[c.nombre]).trim(), ap1: String(r[c.ap1] || '').trim(), ap2: String(r[c.ap2] || '').trim(),
      fijos: fijos[g + '#' + key] || [], _r: hm.existe ? r : null };
  }).filter(Boolean);
}

function diasGuardados_(mes, grupo) {
  var out = {}, g = grupo == null ? null : norm_(grupo);
  filas_(hoja_('_Dias', H_DIAS)).forEach(function (r) {
    if (r[0] !== mes) return;
    var rg = norm_(r[1]);
    if (g !== null && rg !== g) return;
    var reg = {
      dias: String(r[3]).split(',').map(function (x) { return parseInt(x, 10); }).filter(function (x) { return !isNaN(x); }),
      actualizado: r[4] instanceof Date ? r[4].toISOString() : String(r[4] || ''), por: String(r[5] || '')
    };
    if (g === null) { (out[rg] = out[rg] || {})[r[2]] = reg; } else out[r[2]] = reg;
  });
  return out;
}

/** Datos de un mes para TODOS los grupos (se guarda en caché 60 s y se borra al guardar). */
function datosMesTodos_(mes, aj) {
  var hm = datosHojaMes_(mes), fijos = fijosMap_(), c = hm.c, y = anioDeMes_(mes, aj), mn = MES_NUM[mes];
  if (!c) return { anio: y, mesNum: mn, grupos: {} };
  var todosDias = diasGuardados_(mes, null), grupos = [], porGrupo = {};
  hm.filas.forEach(function (x) { var g = String(x.r[c.grupo]).trim(); if (grupos.indexOf(g) < 0) grupos.push(g); });
  grupos.forEach(function (g) {
    var dias = todosDias[norm_(g)] || {};
    porGrupo[g] = monitoresDelGrupo_(mes, g, hm, fijos).map(function (m) {
      var r = m._r, d = dias[m.key];
      delete m._r;
      m.datos = { conteo: [null, null, null, null, null], partidos: null, ausencias: null, sust: null, km: null, peajes: null, obs: '' };
      if (r) {
        m.datos.conteo = c.dias.map(function (col) { var v = r[col]; return v === '' || v == null ? null : (num_(v) == null ? String(v) : num_(v)); });
        m.datos.partidos = num_(r[c.partidos]); m.datos.ausencias = num_(r[c.ausencias]); m.datos.sust = num_(r[c.sust]);
        m.datos.km = num_(r[c.km]); m.datos.peajes = num_(r[c.peajes]); m.datos.obs = String(r[c.obs] || '');
      }
      // ¿alguien cambió los viajes directamente en la hoja? entonces manda la hoja
      if (d) {
        var cnt = [0, 0, 0, 0, 0];
        d.dias.forEach(function (dd) { var w = new Date(y, mn - 1, dd).getDay(); if (w >= 1 && w <= 5) cnt[w - 1]++; });
        var igual = m.datos.conteo.every(function (v, i) { return (typeof v === 'number' ? v : 0) === cnt[i] || (v !== null && typeof v !== 'number'); });
        if (!igual) d = null;
      }
      m.dias = d ? d.dias : null;
      m.actualizado = d ? d.actualizado : '';
      m.por = d ? d.por : '';
      var conDatos = r && [c.partidos, c.ausencias, c.sust, c.km, c.peajes, c.obs].concat(c.dias)
        .some(function (col) { var v = r[col]; return v !== '' && v != null && String(v).toLowerCase() !== 'x'; });
      m.estado = d ? 'guardado' : (conDatos ? 'enhoja' : 'pendiente');
      return m;
    });
  });
  return { anio: y, mesNum: mn, grupos: porGrupo };
}

// ---------------------------------------------------------------- bloqueo (reentrante dentro de una misma llamada)
var _bloqueado = false;
function conBloqueo_(fn) {
  if (_bloqueado) return fn();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  _bloqueado = true;
  try { return fn(); } finally { _bloqueado = false; lock.releaseLock(); }
}
function invalidarMes_(mes) {
  try { CacheService.getScriptCache().removeAll(['resumen', 'mes_' + mes]); } catch (e) { }
}

// ---------------------------------------------------------------- acciones de la API

var ACCIONES = {

  login: function (p) {
    var u = usuario_(p.codigo), aj = ajustes_();
    var hoy = new Date(), mesHoy = null;
    MESES.forEach(function (m) { if (MES_NUM[m] === hoy.getMonth() + 1) mesHoy = m; });
    return {
      persona: u.persona, grupos: u.grupos, admin: u.admin, responsables: u.responsables, contactos: u.contactos,
      meses: mesesDisponibles_(),
      mesActual: (function () { var d = mesesDisponibles_(); return d.indexOf(mesHoy) >= 0 ? mesHoy : d[d.length - 1]; })(),
      cerrados: MESES.filter(function (m) { return aj.cerrados.indexOf(norm_(m)) >= 0; })
    };
  },

  mes: function (p) {
    var u = usuario_(p.codigo), aj = ajustes_();
    comprobarMes_(p.mes, aj, false);
    var grupos = p.grupo ? [p.grupo] : u.grupos;
    grupos.forEach(function (g) { comprobarGrupo_(u, g); });
    var todo = cacheado_('mes_' + p.mes, function () { return datosMesTodos_(p.mes, aj); });
    var porGrupo = {};
    grupos.forEach(function (g) {
      var k = Object.keys(todo.grupos).filter(function (x) { return norm_(x) === norm_(g); })[0];
      porGrupo[g] = k ? todo.grupos[k] : [];
    });
    return {
      mes: p.mes, anio: todo.anio, mesNum: todo.mesNum, festivos: festivosDelMes_(p.mes, aj),
      cerrado: aj.cerrados.indexOf(norm_(p.mes)) >= 0, grupos: porGrupo,
      grupo: p.grupo || null, monitores: p.grupo ? porGrupo[p.grupo] : undefined
    };
  },

  /** Resumen de toda la temporada (solo administración). */
  resumen: function (p) {
    var u = usuario_(p.codigo);
    if (!u.admin) throw new Error('Solo para administración.');
    var c0 = CacheService.getScriptCache(), v = c0.get('resumen');
    if (v) return JSON.parse(v);
    var ss = ss_(), out = { meses: [], filas: {}, resp: [], generado: new Date().toISOString() };
    MESES.forEach(function (mes) {
      var sh = ss.getSheetByName(mes); if (!sh) return;
      var c = columnasMes_(sh);
      out.meses.push(mes);
      out.filas[mes] = filas_(sh).filter(function (r) { return r[c.grupo] && r[c.nombre]; }).map(function (r) {
        var viajes = c.dias.reduce(function (a, col) { return a + (num_(r[col]) || 0); }, 0);
        return [String(r[c.grupo]).trim(), (String(r[c.nombre]).trim() + ' ' + String(r[c.ap1] || '').trim()).trim(),
          viajes, num_(r[c.ausencias]) || 0, num_(r[c.sust]) || 0, num_(r[c.partidos]) || 0, num_(r[c.km]) || 0, num_(r[c.peajes]) || 0];
      });
    });
    var shR = ss.getSheetByName('Responsables');
    if (shR) filas_(shR).forEach(function (r) {
      if (!r[0] || !r[1]) return;
      var mes = MESES.filter(function (m) { return norm_(m) === norm_(r[0]); })[0];
      if (mes) out.resp.push([mes, String(r[1]).trim(), num_(r[2]) || 0, num_(r[3]) || 0, num_(r[4]) || 0, num_(r[6]) || 0]);
    });
    try { c0.put('resumen', JSON.stringify(out), 120); } catch (e) { }
    return out;
  },

  /** Varios guardados en una sola llamada (más rápido desde el móvil). */
  lote: function (p) {
    var res = [];
    conBloqueo_(function () {
      (p.items || []).forEach(function (it) {
        if (it.action !== 'guardarMonitor' && it.action !== 'guardarMisDatos') { res.push({ ok: false, error: 'Acción no permitida' }); return; }
        var q = it.params || {}; q.codigo = p.codigo;
        try { res.push({ ok: true, data: ACCIONES[it.action](q) }); }
        catch (e) { res.push({ ok: false, error: e && e.message ? e.message : String(e) }); }
      });
    });
    return res;
  },

  /** login + datos del mes en una sola llamada (arranque rápido) */
  inicio: function (p) {
    var l = ACCIONES.login(p);
    var mes = p.mes && l.meses.indexOf(p.mes) >= 0 ? p.mes : l.mesActual;
    l.datos = ACCIONES.mes({ codigo: p.codigo, mes: mes });
    return l;
  },

  guardarMonitor: function (p) {
    var u = usuario_(p.codigo), aj = ajustes_();
    comprobarMes_(p.mes, aj, true);
    comprobarGrupo_(u, p.grupo);
    var m = monitoresDelGrupo_(p.mes, p.grupo).filter(function (x) { return x.key === p.key; })[0];
    if (!m) throw new Error('Ese monitor no está en tu grupo este mes.');
    return conBloqueo_(function () {
      var y = anioDeMes_(p.mes, aj), mn = MES_NUM[p.mes];
      var diasMes = new Date(y, mn, 0).getDate();
      var dias = uniq_((p.dias || []).map(Number).filter(function (d) { return d >= 1 && d <= diasMes; })).sort(function (a, b) { return a - b; });
      var conteo = [0, 0, 0, 0, 0];
      dias.forEach(function (d) { var wd = new Date(y, mn - 1, d).getDay(); if (wd >= 1 && wd <= 5) conteo[wd - 1]++; });

      var sh = hojaMes_(p.mes, true), c = columnasMes_(sh);
      var fila = filaMonitor_(sh, c, p.grupo, m, true);
      // solo se escriben las columnas de datos (E..O); el resto de la plantilla no se toca
      var cd = colsDatos_(c), ancho = cd.hasta - cd.desde + 1;
      var rango = sh.getRange(fila, cd.desde + 1, 1, ancho), v = rango.getValues()[0];
      function put(col, x) { v[col - cd.desde] = (x === null || x === undefined || x === '') ? '' : x; }
      c.dias.forEach(function (col, i) { put(col, conteo[i] > 0 ? conteo[i] : (m.fijos.indexOf(i) >= 0 ? 0 : '')); });
      put(c.partidos, num_(p.partidos)); put(c.ausencias, num_(p.ausencias)); put(c.sust, num_(p.sust));
      put(c.km, num_(p.km)); put(c.peajes, num_(p.peajes)); put(c.obs, String(p.obs || '').trim());
      rango.setValues([v]);

      // guardar el detalle de días
      var shD = hoja_('_Dias', H_DIAS), fd = filas_(shD), ahora = new Date(), hecho = false;
      for (var i = 0; i < fd.length; i++) {
        if (fd[i][0] === p.mes && norm_(fd[i][1]) === norm_(p.grupo) && fd[i][2] === p.key) {
          shD.getRange(i + 2, 4, 1, 3).setValues([[dias.join(','), ahora, u.persona]]); hecho = true; break;
        }
      }
      if (!hecho) shD.appendRow([p.mes, p.grupo, p.key, dias.join(','), ahora, u.persona]);
      invalidarMes_(p.mes);
      log_(u.persona, 'Guardar monitor', p.mes + ' · ' + p.grupo + ' · ' + m.nombre + ' ' + m.ap1 + ' · días ' + dias.join(','));
      return { ok: true, conteo: conteo, actualizado: ahora.toISOString() };
    });
  },

  misDatos: function (p) {
    var u = usuario_(p.codigo);
    var nombre = p.responsable || u.responsables[0];
    if (!nombre || u.responsables.indexOf(nombre) < 0) throw new Error('No tienes datos propios configurados.');
    var sh = ss_().getSheetByName('Responsables'), M = norm_(p.mes);
    var r = filas_(sh).filter(function (x) { return norm_(x[0]) === M && norm_(x[1]) === norm_(nombre); })[0];
    return {
      responsable: nombre,
      km: r ? num_(r[2]) : null, partidos: r ? num_(r[3]) : null, sust: r ? num_(r[4]) : null,
      obs: r ? String(r[5] || '') : '', peajes: r ? num_(r[6]) : null,
      cerrado: ajustes_().cerrados.indexOf(M) >= 0
    };
  },

  guardarMisDatos: function (p) {
    var u = usuario_(p.codigo), aj = ajustes_();
    comprobarMes_(p.mes, aj, true);
    var nombre = p.responsable || u.responsables[0];
    if (!nombre || u.responsables.indexOf(nombre) < 0) throw new Error('No tienes datos propios configurados.');
    return conBloqueo_(function () {
      var sh = ss_().getSheetByName('Responsables'), M = norm_(p.mes), datos = filas_(sh), fila = -1;
      for (var i = 0; i < datos.length; i++) if (norm_(datos[i][0]) === M && norm_(datos[i][1]) === norm_(nombre)) { fila = i + 2; break; }
      var valores = [[num_(p.km) == null ? '' : num_(p.km), num_(p.partidos) == null ? '' : num_(p.partidos),
        num_(p.sust) == null ? '' : num_(p.sust), String(p.obs || '').trim(), num_(p.peajes) == null ? '' : num_(p.peajes)]];
      if (fila < 0) {
        // añadir al final del bloque de ese mes
        var ultimo = -1;
        for (var j = 0; j < datos.length; j++) if (norm_(datos[j][0]) === M) ultimo = j;
        if (ultimo >= 0) { fila = ultimo + 3; sh.insertRowAfter(ultimo + 2); } else fila = sh.getLastRow() + 1;
        sh.getRange(fila, 1, 1, 2).setValues([[p.mes.toUpperCase(), nombre]]);
      }
      sh.getRange(fila, 3, 1, 5).setValues(valores);
      invalidarMes_(p.mes);
      log_(u.persona, 'Guardar mis datos', p.mes + ' · ' + nombre);
      return { ok: true };
    });
  }
};

function ejecutar_(accion, params) {
  try {
    if (!ACCIONES[accion]) throw new Error('Acción desconocida: ' + accion);
    return { ok: true, data: ACCIONES[accion](params || {}) };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

// ---------------------------------------------------------------- puntos de entrada web

/** Llamadas desde la app alojada fuera (GitHub Pages, etc.). */
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { }
  var res = ejecutar_(body.action, body);
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

/** Abrir la URL del script en el navegador sirve la propia app. */
function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Incidencias Escuela 26-27')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Llamadas desde la app servida por el propio script (google.script.run). */
function apiRun(accion, paramsJson) {
  return JSON.stringify(ejecutar_(accion, JSON.parse(paramsJson || '{}')));
}

/** Menú en la hoja para Mikel. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('App incidencias')
    .addItem('Configurar / actualizar hojas de la app', 'menuConfigurar_')
    .addItem('Crear hoja del mes siguiente', 'menuMesSiguiente_')
    .addItem('Cargar festivos del calendario 26-27', 'menuFestivos_')
    .addItem('Aplicar cambios de configuración ya', 'limpiarCache_')
    .addToUi();
}
function menuFestivos_() { SpreadsheetApp.getUi().alert(cargarFestivosCalendario()); }
function menuMesSiguiente_() {
  var ss = ss_(), ultimo = -1;
  MESES.forEach(function (m, i) { if (ss.getSheetByName(m)) ultimo = i; });
  if (ultimo >= MESES.length - 1) return SpreadsheetApp.getUi().alert('Ya están creados todos los meses.');
  var mes = MESES[ultimo + 1];
  hojaMes_(mes, true);
  SpreadsheetApp.getUi().alert('Creada la hoja "' + mes + '" copiando la anterior (sin datos).');
}
function menuConfigurar_() { SpreadsheetApp.getUi().alert(configurarInicial()); }
