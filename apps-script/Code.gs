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

var MESES = ['Octubre', 'Noviembre', 'Diciembre', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio'];
var MES_NUM = { Octubre: 10, Noviembre: 11, Diciembre: 12, Enero: 1, Febrero: 2, Marzo: 3, Abril: 4, Mayo: 5, Junio: 6 };
var DIAS = ['L', 'M', 'X', 'J', 'V'];

var H_ACCESOS = ['Código', 'Persona', 'Grupos (separados por ; o * = todos)', 'Nombre en hoja Responsables', 'Admin'];
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

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

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
  var sh = ss_().getSheetByName(nombre);
  if (!sh && cabecera) {
    sh = ss_().insertSheet(nombre);
    sh.getRange(1, 1, 1, cabecera.length).setValues([cabecera]).setFontWeight('bold').setBackground('#1f2a44').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function filas_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

function ajustes_() {
  var out = { TEMPORADA_INICIO: 2026, MESES_CERRADOS: '' };
  filas_(hoja_('_Ajustes')).forEach(function (r) { if (r[0]) out[String(r[0]).trim()] = r[1]; });
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
function configurarInicial() {
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
    return [nuevoCodigo_(), bonito, porPersona[p].join('; '), nombreResp, false];
  });
  if (!accExist.some(function (r) { return esVerdadero_(r[4]); })) {
    filasAcc.unshift([nuevoCodigo_(), 'Administración', '*', '', true]);
  }
  if (filasAcc.length) {
    var st = shAcc.getLastRow() + 1;
    shAcc.getRange(st, 1, filasAcc.length, H_ACCESOS.length).setValues(filasAcc);
    shAcc.getRange(st, 5, filasAcc.length, 1).insertCheckboxes();
  }

  // --- Festivos, ajustes y hojas auxiliares
  var shFes = hoja_('_Festivos', H_FESTIVOS);
  if (shFes.getLastRow() < 2) {
    var f = [
      [new Date(2026, 11, 3), 'San Francisco Javier (Navarra)'],
      [new Date(2026, 11, 8), 'Inmaculada'],
      [new Date(2026, 11, 25), 'Navidad'],
      [new Date(2027, 0, 1), 'Año Nuevo'],
      [new Date(2027, 0, 6), 'Reyes'],
      [new Date(2027, 2, 19), 'San José'],
      [new Date(2027, 2, 25), 'Jueves Santo'],
      [new Date(2027, 2, 26), 'Viernes Santo'],
      [new Date(2027, 2, 29), 'Lunes de Pascua']
    ];
    shFes.getRange(2, 1, f.length, 2).setValues(f);
    shFes.getRange(2, 1, f.length, 1).setNumberFormat('dd/mm/yyyy');
  }
  var shAj = hoja_('_Ajustes', H_AJUSTES);
  if (shAj.getLastRow() < 2) {
    shAj.getRange(2, 1, 2, 3).setValues([
      ['TEMPORADA_INICIO', 2026, 'Año en que empieza la temporada (octubre)'],
      ['MESES_CERRADOS', '', 'Meses que ya no se pueden editar desde la app, separados por ; (ej: Octubre; Noviembre)']
    ]);
  }
  hoja_('_Dias', H_DIAS);
  hoja_('_Registro', H_REGISTRO);
  return 'Configuración lista: ' + orden.length + ' monitores y ' + filasAcc.length + ' accesos nuevos.';
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
  var filas = filas_(hoja_('_Accesos'));
  for (var i = 0; i < filas.length; i++) {
    var r = filas[i];
    if (norm_(r[0]) === c) {
      var admin = esVerdadero_(r[4]) || String(r[2]).trim() === '*';
      var grupos;
      if (admin) {
        var set = {};
        filas_(hoja_('_Monitores')).forEach(function (m) { if (m[0]) set[String(m[0]).trim()] = 1; });
        grupos = Object.keys(set).sort();
      } else {
        grupos = String(r[2]).split(';').map(function (g) { return g.trim(); }).filter(String);
      }
      // nombres de hoja Responsables a los que puede acceder en "Mis datos"
      var propios = admin
        ? uniq_(filas_(ss_().getSheetByName('Responsables')).map(function (x) { return String(x[1] || '').trim(); }).filter(String))
        : (String(r[3]).trim() ? [String(r[3]).trim()] : []);
      return { persona: String(r[1]).trim(), grupos: grupos, admin: admin, responsables: propios };
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
  filas_(hoja_('_Festivos')).forEach(function (r) {
    var d = r[0];
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

/** Devuelve la hoja del mes; si no existe la crea con todos los monitores activos. */
function hojaMes_(mes, crear) {
  var ss = ss_(), sh = ss.getSheetByName(mes);
  if (sh || !crear) return sh;
  sh = ss.insertSheet(mes);
  sh.getRange(1, 1, 1, CABECERA_MES.length).setValues([CABECERA_MES]).setFontWeight('bold').setWrap(true);
  sh.setFrozenRows(1);
  var mons = filas_(hoja_('_Monitores')).filter(function (r) { return r[0] && r[1] && esVerdadero_(r[9]); });
  if (mons.length) {
    sh.getRange(2, 1, mons.length, 4).setValues(mons.map(function (r) { return [r[0], r[1], r[2], r[3]]; }));
  }
  return sh;
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

function monitoresDelGrupo_(grupo) {
  var g = norm_(grupo);
  return filas_(hoja_('_Monitores')).filter(function (r) {
    return norm_(r[0]) === g && r[1] && esVerdadero_(r[9]);
  }).map(function (r) {
    var fijos = [];
    for (var i = 0; i < 5; i++) if (esVerdadero_(r[4 + i])) fijos.push(i);
    return { key: claveMonitor_(r[1], r[2], r[3]), nombre: String(r[1]).trim(), ap1: String(r[2] || '').trim(), ap2: String(r[3] || '').trim(), fijos: fijos };
  });
}

function diasGuardados_(mes, grupo) {
  var out = {}, g = norm_(grupo);
  filas_(hoja_('_Dias', H_DIAS)).forEach(function (r) {
    if (r[0] === mes && norm_(r[1]) === g) {
      out[r[2]] = {
        dias: String(r[3]).split(',').map(function (x) { return parseInt(x, 10); }).filter(function (x) { return !isNaN(x); }),
        actualizado: r[4] instanceof Date ? r[4].toISOString() : String(r[4] || ''), por: String(r[5] || '')
      };
    }
  });
  return out;
}

// ---------------------------------------------------------------- acciones de la API

var ACCIONES = {

  login: function (p) {
    var u = usuario_(p.codigo), aj = ajustes_();
    var hoy = new Date(), mesHoy = null;
    MESES.forEach(function (m) { if (MES_NUM[m] === hoy.getMonth() + 1) mesHoy = m; });
    return {
      persona: u.persona, grupos: u.grupos, admin: u.admin, responsables: u.responsables,
      meses: MESES, mesActual: mesHoy || (hoy.getMonth() + 1 >= 7 && hoy.getMonth() + 1 < 10 ? 'Octubre' : 'Junio'),
      cerrados: MESES.filter(function (m) { return aj.cerrados.indexOf(norm_(m)) >= 0; })
    };
  },

  mes: function (p) {
    var u = usuario_(p.codigo), aj = ajustes_();
    comprobarMes_(p.mes, aj, false);
    comprobarGrupo_(u, p.grupo);
    var mons = monitoresDelGrupo_(p.grupo);
    var sh = hojaMes_(p.mes, false), filas = {}, c = null;
    if (sh) {
      c = columnasMes_(sh);
      filas_(sh).forEach(function (r) {
        if (norm_(r[c.grupo]) === norm_(p.grupo)) filas[claveMonitor_(r[c.nombre], r[c.ap1], r[c.ap2])] = r;
      });
    }
    var dias = diasGuardados_(p.mes, p.grupo);
    mons.forEach(function (m) {
      var r = filas[m.key], d = dias[m.key];
      m.datos = { conteo: [null, null, null, null, null], partidos: null, ausencias: null, sust: null, km: null, peajes: null, obs: '' };
      if (r) {
        m.datos.conteo = c.dias.map(function (col) { var v = r[col]; return v === '' || v == null ? null : (num_(v) == null ? String(v) : num_(v)); });
        m.datos.partidos = num_(r[c.partidos]); m.datos.ausencias = num_(r[c.ausencias]); m.datos.sust = num_(r[c.sust]);
        m.datos.km = num_(r[c.km]); m.datos.peajes = num_(r[c.peajes]); m.datos.obs = String(r[c.obs] || '');
      }
      m.dias = d ? d.dias : null;
      m.actualizado = d ? d.actualizado : '';
      m.por = d ? d.por : '';
      var conDatos = r && [c.partidos, c.ausencias, c.sust, c.km, c.peajes, c.obs].concat(c.dias)
        .some(function (col) { var v = r[col]; return v !== '' && v != null && String(v).toLowerCase() !== 'x'; });
      m.estado = d ? 'guardado' : (conDatos ? 'enhoja' : 'pendiente');
    });
    return {
      mes: p.mes, grupo: p.grupo, anio: anioDeMes_(p.mes, aj), mesNum: MES_NUM[p.mes],
      festivos: festivosDelMes_(p.mes, aj), cerrado: aj.cerrados.indexOf(norm_(p.mes)) >= 0, monitores: mons
    };
  },

  guardarMonitor: function (p) {
    var u = usuario_(p.codigo), aj = ajustes_();
    comprobarMes_(p.mes, aj, true);
    comprobarGrupo_(u, p.grupo);
    var m = monitoresDelGrupo_(p.grupo).filter(function (x) { return x.key === p.key; })[0];
    if (!m) throw new Error('Ese monitor no está en tu grupo.');
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var y = anioDeMes_(p.mes, aj), mn = MES_NUM[p.mes];
      var diasMes = new Date(y, mn, 0).getDate();
      var dias = uniq_((p.dias || []).map(Number).filter(function (d) { return d >= 1 && d <= diasMes; })).sort(function (a, b) { return a - b; });
      var conteo = [0, 0, 0, 0, 0];
      dias.forEach(function (d) { var wd = new Date(y, mn - 1, d).getDay(); if (wd >= 1 && wd <= 5) conteo[wd - 1]++; });

      var sh = hojaMes_(p.mes, true), c = columnasMes_(sh);
      var fila = filaMonitor_(sh, c, p.grupo, m, true);
      var ancho = Math.max(sh.getLastColumn(), c.obs + 1);
      var r = sh.getRange(fila, 1, 1, ancho).getValues()[0];
      c.dias.forEach(function (col, i) { r[col] = conteo[i] > 0 ? conteo[i] : (m.fijos.indexOf(i) >= 0 ? 0 : ''); });
      function put(col, v) { r[col] = (v === null || v === undefined || v === '' ) ? '' : v; }
      put(c.partidos, num_(p.partidos)); put(c.ausencias, num_(p.ausencias)); put(c.sust, num_(p.sust));
      put(c.km, num_(p.km)); put(c.peajes, num_(p.peajes)); put(c.obs, String(p.obs || '').trim());
      sh.getRange(fila, 1, 1, ancho).setValues([r]);

      // guardar el detalle de días
      var shD = hoja_('_Dias', H_DIAS), fd = filas_(shD), ahora = new Date(), hecho = false;
      for (var i = 0; i < fd.length; i++) {
        if (fd[i][0] === p.mes && norm_(fd[i][1]) === norm_(p.grupo) && fd[i][2] === p.key) {
          shD.getRange(i + 2, 4, 1, 3).setValues([[dias.join(','), ahora, u.persona]]); hecho = true; break;
        }
      }
      if (!hecho) shD.appendRow([p.mes, p.grupo, p.key, dias.join(','), ahora, u.persona]);
      log_(u.persona, 'Guardar monitor', p.mes + ' · ' + p.grupo + ' · ' + m.nombre + ' ' + m.ap1 + ' · días ' + dias.join(','));
      return { ok: true, conteo: conteo, actualizado: ahora.toISOString() };
    } finally { lock.releaseLock(); }
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
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
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
      log_(u.persona, 'Guardar mis datos', p.mes + ' · ' + nombre);
      return { ok: true };
    } finally { lock.releaseLock(); }
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
    .addToUi();
}
function menuConfigurar_() { SpreadsheetApp.getUi().alert(configurarInicial()); }
