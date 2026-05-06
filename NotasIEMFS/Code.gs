// ===============================
// UTILIDAD GLOBAL
// ===============================
function limpiarTexto(txt) {
  return String(txt)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Curso/grupo: alinea guiones tipográficos (Sheets/Word) con ASCII para que
 * «8-9 B» en DocenteAsignaciones coincida con «8–9 B» (raya en) en Estudiantes.
 * Debe usarse en toda comparación grupo↔grupo (portal estudiante, materias del curso).
 */
function normalizarGrupoCursoIEMFS_(txt) {
  return limpiarTexto(String(txt == null ? "" : txt).replace(/[\u2013\u2014\u2212]/g, "-"));
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Compara fechas de hoja / JSON / string de forma tolerante. */
function fechaMs_(v) {
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) return v.getTime();
  const d = new Date(v);
  if (isNaN(d.getTime())) return null;
  return d.getTime();
}

/** Mismo día civil entre ISO del cliente y Date/serial de la hoja (evita fallar la búsqueda por ms distintos). */
function ymdCalendarioDesdeValorNota_(v) {
  if (v === null || v === undefined || v === "") return "";
  var d = fechaCeldaNotaColumnaH_(v) || fechaCeldaInasistencia_(v);
  if (d && !isNaN(d.getTime())) {
    var tz = Session.getScriptTimeZone();
    return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  }
  var m = /^(\d{4}-\d{2}-\d{2})/.exec(String(v).trim());
  return m ? m[1] : "";
}

function notasIguales_(a, b) {
  const x = parseFloat(a);
  const y = parseFloat(b);
  if (isNaN(x) || isNaN(y)) return String(a) === String(b);
  return Math.abs(x - y) < 1e-6;
}

// ===============================
// NOTAS / INASISTENCIAS POR PERÍODO (NotasP1–P3, InasistenciasP1–P3)
// Activar desde el editor: EJECUTAR_HABILITAR_MULTIHJA_NOTAS_E_INASISTENCIAS()
// ===============================
var NOTAS_HOJAS_P_ = ["NotasP1", "NotasP2", "NotasP3"];
var INAS_HOJAS_P_ = ["InasistenciasP1", "InasistenciasP2", "InasistenciasP3"];
var PROP_NOTAS_MULTIH_ = "NOTAS_IEMFS_NOTAS_MULTIHJA";
var PROP_INAS_MULTIH_ = "NOTAS_IEMFS_INASIST_MULTIHJA";
var NOTAS_HOJA_LEGACY_ = "Notas";
var INAS_HOJA_LEGACY_ = "Inasistencias";
var CACHE_NOTAS_GRUPO_SEC_ = 90;
/**
 * Debe coincidir con NOTAS_IEMFS_BUILD_ESPERADO en src/config.js (el cliente lo comprueba vía ping).
 * Si cambia la lógica de guardado masivo, actualice ambos y vuelva a desplegar el script web.
 */
var NOTAS_IEMFS_BUILD_ = "mosaico-grupo-20260506a";

/** Acepta nombres canónicos y alias por errores al crear propiedades en la consola de Google. */
function scriptPropAlgunaEsUno_(nombres) {
  var p = PropertiesService.getScriptProperties();
  var i;
  for (i = 0; i < nombres.length; i++) {
    if (p.getProperty(nombres[i]) === "1") return true;
  }
  return false;
}

function notasMultihojaActivas_() {
  return scriptPropAlgunaEsUno_([
    PROP_NOTAS_MULTIH_,
    "NOTAS_IEMFS_NOTAS_MULTI_HJA",
    "NOTAS_IEMFS_NOTAS_MULTIHA"
  ]);
}

function inasistenciasMultihojaActivas_() {
  return scriptPropAlgunaEsUno_([
    PROP_INAS_MULTIH_,
    "NOTAS_IEMFS_INASIST_MULTI_HJA",
    "NOTAS_IEMFS_INASIST_MULTIHA"
  ]);
}

/** Si las notas están en hojas por periodo, no hace falta filtrar otra vez por fecha de periodo. */
function limitePeriodoFechaLecturaNotas_(periodoId) {
  if (notasMultihojaActivas_()) return null;
  return limitePeriodoNotasOpcional_(periodoId);
}

function periodoNumDesdeFechaObj_(fd) {
  if (!(fd instanceof Date) || isNaN(fd.getTime())) return 1;
  var lista = listarPeriodosAcademicos_();
  var i;
  for (i = 0; i < lista.length; i++) {
    try {
      var desde = fechaDiaDesdeIso_(lista[i].desde);
      var hasta = fechaDiaDesdeIso_(lista[i].hasta);
      if (fechaEnRangoOrd_(fd, desde, hasta)) return lista[i].id;
    } catch (e) {
      /* ignore */
    }
  }
  return 1;
}

function periodoNumDesdeFechaIsoInas_(fechaIso) {
  try {
    return periodoNumDesdeFechaObj_(fechaDiaDesdeIso_(fechaIso));
  } catch (e) {
    return 1;
  }
}

/** Periodo 1–3 enviado por la app docente; si no viene o es inválido, se infiere por la fecha de guardado. */
function periodoNumParaNuevaNota_(notaObj, fGuardado) {
  var p = parseInt(String(notaObj.periodo != null ? notaObj.periodo : "").trim(), 10);
  if (!isNaN(p) && p >= 1 && p <= 3) return p;
  return periodoNumDesdeFechaObj_(fGuardado);
}

function obtenerHojaNotasPeriodoNum_(n) {
  var idx = parseInt(String(n), 10);
  if (isNaN(idx) || idx < 1 || idx > 3) idx = 1;
  var ss = SpreadsheetApp.getActive();
  var name = NOTAS_HOJAS_P_[idx - 1];
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(["ts", "estudiante", "materia", "grupo", "tipo", "titulo", "nota", "fecha"]);
  }
  return sh;
}

function obtenerHojaNotasParaLecturaPorPeriodoId_(periodoId) {
  if (!notasMultihojaActivas_()) {
    return SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
  }
  var n = parseInt(String(periodoId || "1").trim(), 10);
  if (isNaN(n) || n < 1 || n > 3) n = 1;
  return obtenerHojaNotasPeriodoNum_(n);
}

/** Filas de datos (desde fila 2) columnas A:H — sin getDataRange de toda la hoja. */
function valoresNotasHojaDesdeFila2_(hoja) {
  if (!hoja) return [];
  var lr = hoja.getLastRow();
  if (lr < 2) return [];
  return hoja.getRange(2, 1, lr - 1, 8).getValues();
}

function valoresInasistenciasHojaDesdeFila2_(hoja) {
  if (!hoja) return [];
  var lr = hoja.getLastRow();
  if (lr < 2) return [];
  return hoja.getRange(2, 1, lr - 1, 6).getValues();
}

function parseNotaIdRef_(notaId) {
  var s = String(notaId == null ? "" : notaId).trim();
  var m = /^([123]):(\d+)$/.exec(s);
  if (m) {
    return { periodo: parseInt(m[1], 10), fila: parseInt(m[2], 10), soloLegacy: false };
  }
  var r = parseInt(s, 10);
  if (!isNaN(r) && r >= 2) {
    return { periodo: 0, fila: r, soloLegacy: true };
  }
  return null;
}

function hojaNotaDesdeRef_(ref) {
  if (!notasMultihojaActivas_() || ref.soloLegacy) {
    return SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
  }
  return obtenerHojaNotasPeriodoNum_(ref.periodo);
}

function notaRefAuditoria_(periodoNum, filaHoja) {
  if (!notasMultihojaActivas_()) return String(filaHoja);
  return String(periodoNum) + ":" + String(filaHoja);
}

function cacheNotasGrupoKey_(grupo, materia, periodoId) {
  return (
    "ng2:" +
    normalizarGrupoCursoIEMFS_(String(grupo || "")) +
    ":" +
    limpiarTexto(String(materia || "")) +
    ":" +
    String(periodoId || "")
  );
}

function cacheGetJson_(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function cachePutJson_(key, obj, ttlSec) {
  try {
    var s = JSON.stringify(obj);
    if (s.length > 95000) return;
    CacheService.getScriptCache().put(key, s, ttlSec);
  } catch (e) {
    /* ignore */
  }
}

function invalidarCacheNotasGrupo_(grupo, materia) {
  var g = normalizarGrupoCursoIEMFS_(String(grupo || ""));
  var m = limpiarTexto(String(materia || ""));
  var p;
  for (p = 1; p <= 3; p++) {
    try {
      CacheService.getScriptCache().remove(cacheNotasGrupoKey_(g, m, String(p)));
    } catch (e) {
      /* ignore */
    }
  }
  try {
    CacheService.getScriptCache().remove(cacheNotasGrupoKey_(g, m, ""));
  } catch (e2) {
    /* ignore */
  }
}

function obtenerHojaInasistenciasPeriodoNum_(n) {
  var idx = parseInt(String(n), 10);
  if (isNaN(idx) || idx < 1 || idx > 3) idx = 1;
  var ss = SpreadsheetApp.getActive();
  var name = INAS_HOJAS_P_[idx - 1];
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(["timestamp", "estudiante", "grupo", "materia", "fecha", "falta"]);
  }
  return sh;
}

function hojaInasistenciasParaFechaIso_(fechaIso) {
  if (!inasistenciasMultihojaActivas_()) {
    return obtenerHojaInasistencias_();
  }
  var n = periodoNumDesdeFechaIsoInas_(fechaIso);
  return obtenerHojaInasistenciasPeriodoNum_(n);
}

/**
 * Crea NotasP1–P3 e InasistenciasP1–P3 y activa el modo multihoja.
 * Ejecutar una sola vez desde el editor de Apps Script (con el spreadsheet abierto).
 * Opcional: migrar filas desde Notas / Inasistencias según la fecha de cada fila.
 */
function EJECUTAR_HABILITAR_MULTIHJA_NOTAS_E_INASISTENCIAS() {
  var ss = SpreadsheetApp.getActive();
  var i;
  for (i = 0; i < NOTAS_HOJAS_P_.length; i++) {
    obtenerHojaNotasPeriodoNum_(i + 1);
  }
  for (i = 0; i < INAS_HOJAS_P_.length; i++) {
    obtenerHojaInasistenciasPeriodoNum_(i + 1);
  }
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_NOTAS_MULTIH_, "1");
  props.setProperty(PROP_INAS_MULTIH_, "1");
  return (
    "Multihoja activada. Hojas " +
    NOTAS_HOJAS_P_.join(", ") +
    " e " +
    INAS_HOJAS_P_.join(", ") +
    " listas. Puede ejecutar MIGRAR_NOTAS_LEGACY_A_HOJAS_POR_PERIODO() si tenía datos en «Notas»."
  );
}

/** Copia filas de la hoja «Notas» a NotasP1/P2/P3 según la columna fecha (H). No borra el origen. */
function MIGRAR_NOTAS_LEGACY_A_HOJAS_POR_PERIODO() {
  if (!notasMultihojaActivas_()) {
    throw new Error("Primero ejecute EJECUTAR_HABILITAR_MULTIHJA_NOTAS_E_INASISTENCIAS().");
  }
  var leg = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
  if (!leg) return "No hay hoja Notas.";
  var datos = leg.getDataRange().getValues();
  var mov = 0;
  var j;
  for (j = 1; j < datos.length; j++) {
    var fd = fechaCeldaInasistencia_(datos[j][7]);
    if (!fd) continue;
    var p = periodoNumDesdeFechaObj_(fd);
    if (p < 1 || p > 3) p = 1;
    var dest = obtenerHojaNotasPeriodoNum_(p);
    dest.appendRow([
      datos[j][0],
      datos[j][1],
      datos[j][2],
      datos[j][3],
      datos[j][4],
      datos[j][5],
      datos[j][6],
      datos[j][7]
    ]);
    mov++;
  }
  return "Filas copiadas a NotasP* (según fecha): " + mov + ". Revise y, si está conforme, archive o vacíe la hoja Notas manualmente.";
}

/** Copia filas de «Inasistencias» a InasistenciasP1/P2/P3 según la fecha (columna E). */
function MIGRAR_INASISTENCIAS_LEGACY_A_HOJAS_POR_PERIODO() {
  if (!inasistenciasMultihojaActivas_()) {
    throw new Error("Primero ejecute EJECUTAR_HABILITAR_MULTIHJA_NOTAS_E_INASISTENCIAS().");
  }
  var leg = SpreadsheetApp.getActive().getSheetByName(INAS_HOJA_LEGACY_);
  if (!leg) return "No hay hoja Inasistencias.";
  var datos = leg.getDataRange().getValues();
  var mov = 0;
  var j;
  for (j = 1; j < datos.length; j++) {
    var fd = fechaCeldaInasistencia_(datos[j][4]);
    if (!fd) continue;
    var p = periodoNumDesdeFechaObj_(fd);
    if (p < 1 || p > 3) p = 1;
    var dest = obtenerHojaInasistenciasPeriodoNum_(p);
    dest.appendRow([
      datos[j][0],
      datos[j][1],
      datos[j][2],
      datos[j][3],
      datos[j][4],
      datos[j][5]
    ]);
    mov++;
  }
  return "Filas copiadas a InasistenciasP* (según fecha): " + mov + ".";
}

// ===============================
// CLAVE DE ACCESO (Propiedades del script)
// Propiedad: NOTAS_IEMFS_CLAVE
// Configura una vez ejecutando en el editor: EJECUTAR_PARA_GUARDAR_MI_CLAVE (no pulses Ejecutar en establecerClaveAcceso)
// ===============================
var PROP_CLAVE_NOTAS_ = "NOTAS_IEMFS_CLAVE";

function obtenerClaveEsperada_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_CLAVE_NOTAS_) || "";
}

/** Código de usuario para coordinación (misma clave que NOTAS_IEMFS_CLAVE). Opcional: propiedad NOTAS_IEMFS_CODIGO_COORD. */
var PROP_CODIGO_COORD_ = "NOTAS_IEMFS_CODIGO_COORD";
var CODIGO_COORDINADOR_DEFAULT_ = "C001";

function obtenerCodigoCoordinador_() {
  var p = PropertiesService.getScriptProperties().getProperty(PROP_CODIGO_COORD_);
  var c = String(p || "").trim();
  return c || CODIGO_COORDINADOR_DEFAULT_;
}

/**
 * Opcional: ejecutar una vez si no usa el valor por defecto (C001).
 * Si cambia el código aquí, actualice también `CODIGO_COORDINADOR` en `src/config.js` y regenere `app.js`.
 */
function EJECUTAR_PARA_GUARDAR_CODIGO_COORDINADOR() {
  var miCodigo = "C001";

  var c = String(miCodigo || "").trim();
  if (c.length < 2) {
    throw new Error(
      'Edita EJECUTAR_PARA_GUARDAR_CODIGO_COORDINADOR: pon var miCodigo = "C001";'
    );
  }
  PropertiesService.getScriptProperties().setProperty(PROP_CODIGO_COORD_, c);
  return "Código de coordinación guardado en propiedades del script.";
}

/** @return {string} mensaje de error o "" si OK */
function validarClaveEnGet_(e) {
  var recibida = String(e && e.parameter && e.parameter.clave ? e.parameter.clave : "").trim();
  var esp = obtenerClaveEsperada_();
  if (!esp) {
    return "Falta configurar la clave en Apps Script. En el editor ejecuta EJECUTAR_PARA_GUARDAR_MI_CLAVE (ver Código.gs).";
  }
  if (recibida !== esp) {
    return "Clave incorrecta.";
  }
  return "";
}

/** @return {string} mensaje de error o "" si OK — solo modo coordinador (clave maestra). */
function validarClaveMaestraPost_(data) {
  var recibida = String(data && data.clave ? data.clave : "").trim();
  var esp = obtenerClaveEsperada_();
  if (!esp) {
    return "Falta configurar la clave en Apps Script.";
  }
  if (recibida !== esp) {
    return "Clave incorrecta.";
  }
  return "";
}

/** @deprecated Usar resolverAuthPost_; se mantiene nombre para búsquedas en código antiguo. */
function validarClaveEnPost_(data) {
  var r = resolverAuthPost_(data);
  return r.error || "";
}

// ===============================
// CLAVE ADMIN (traslados y acciones sensibles) — Propiedad NOTAS_IEMFS_CLAVE_ADMIN
// ===============================
var PROP_CLAVE_ADMIN_ = "NOTAS_IEMFS_CLAVE_ADMIN";

function obtenerClaveAdmin_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_CLAVE_ADMIN_) || "";
}

function establecerClaveAdminAcceso(clave) {
  var c = String(clave || "").trim();
  if (c.length < 8) {
    throw new Error("La clave de administración debe tener al menos 8 caracteres.");
  }
  PropertiesService.getScriptProperties().setProperty(PROP_CLAVE_ADMIN_, c);
  return "Clave de administración guardada en propiedades del script.";
}

/**
 * Igual que EJECUTAR_PARA_GUARDAR_MI_CLAVE pero para NOTAS_IEMFS_CLAVE_ADMIN.
 * Pon la clave en miClave y Ejecutar (solo coordinador con acceso al editor).
 */
function EJECUTAR_PARA_GUARDAR_CLAVE_ADMIN() {
  var miClave = "";

  if (!miClave || String(miClave).trim().length < 8) {
    throw new Error(
      'Edita EJECUTAR_PARA_GUARDAR_CLAVE_ADMIN: pon la clave en var miClave = "...";'
    );
  }
  return establecerClaveAdminAcceso(miClave);
}

/** @return {string} error o "" */
function validarClaveAdminEnData_(data) {
  var rec = String(data && data.claveAdmin ? data.claveAdmin : "").trim();
  var esp = obtenerClaveAdmin_();
  if (!esp) {
    return "Falta configurar la clave de administración (ejecute EJECUTAR_PARA_GUARDAR_CLAVE_ADMIN en Apps Script).";
  }
  if (rec !== esp) {
    return "Clave de administración incorrecta.";
  }
  return "";
}

// ===============================
// DOCENTES Y ASIGNACIONES (hojas Docentes + DocenteAsignaciones)
// ===============================
function obtenerHojaDocentes_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName("Docentes");
  if (!sh) {
    sh = ss.insertSheet("Docentes");
    sh.appendRow(["codigo", "nombre", "clave"]);
  }
  return sh;
}

function obtenerHojaDocenteAsignaciones_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName("DocenteAsignaciones");
  if (!sh) {
    sh = ss.insertSheet("DocenteAsignaciones");
    sh.appendRow(["docente_codigo", "grupo", "materia"]);
  }
  return sh;
}

function obtenerDocentePorCodigoYClave_(codigo, clave) {
  var c = String(codigo || "").trim();
  var cl = String(clave || "").trim();
  if (!c || !cl) return null;
  var sh = obtenerHojaDocentes_();
  var datos = sh.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var cod = String(datos[i][0] || "").trim();
    if (cod !== c) continue;
    var pass = String(datos[i][2] == null ? "" : datos[i][2]).trim();
    if (pass !== cl) return null;
    return {
      codigo: cod,
      nombre: String(datos[i][1] || "").trim() || cod
    };
  }
  return null;
}

function obtenerAsignacionesDocente_(codigo) {
  var c = String(codigo || "").trim();
  if (!c) return [];
  var sh = obtenerHojaDocenteAsignaciones_();
  var datos = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < datos.length; i++) {
    var dc = String(datos[i][0] || "").trim();
    if (dc !== c) continue;
    var g = String(datos[i][1] == null ? "" : datos[i][1]).trim();
    var m = String(datos[i][2] == null ? "" : datos[i][2]).trim();
    if (!g || !m) continue;
    out.push({ grupo: g, materia: m });
  }
  return out;
}

/** Materias únicas asignadas a un curso (DocenteAsignaciones: B=grupo, C=materia). */
function listarMateriasUnicasGrupo_(grupo) {
  var gW = String(grupo || "").trim();
  if (!gW) return [];
  var sh = obtenerHojaDocenteAsignaciones_();
  var datos = sh.getDataRange().getValues();
  var seen = {};
  var out = [];
  var gNorm = normalizarGrupoCursoIEMFS_(gW);
  var i;
  for (i = 1; i < datos.length; i++) {
    var gRow = String(datos[i][1] == null ? "" : datos[i][1]).trim();
    if (normalizarGrupoCursoIEMFS_(gRow) !== gNorm) continue;
    var m = String(datos[i][2] == null ? "" : datos[i][2]).trim();
    if (!m) continue;
    var mk = limpiarTexto(m);
    if (seen[mk]) continue;
    seen[mk] = true;
    out.push(m);
  }
  out.sort(function (a, b) {
    return limpiarTexto(a).localeCompare(limpiarTexto(b), "es", { sensitivity: "base", numeric: true });
  });
  return out;
}

function listarMateriasDistintasNotasEstudiante_(estudianteId) {
  var idB = String(estudianteId || "").trim();
  if (!idB) return [];
  var seen = {};
  var out = [];
  function acumularDesdeDatos_(datos) {
    var j;
    for (j = 0; j < datos.length; j++) {
      if (String(datos[j][1] || "").trim() !== idB) continue;
      var m = String(datos[j][2] == null ? "" : datos[j][2]).trim();
      if (!m) continue;
      var mk = limpiarTexto(m);
      if (seen[mk]) continue;
      seen[mk] = true;
      out.push(m);
    }
  }
  if (notasMultihojaActivas_()) {
    var p;
    for (p = 1; p <= 3; p++) {
      var sh = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJAS_P_[p - 1]);
      if (!sh) continue;
      acumularDesdeDatos_(valoresNotasHojaDesdeFila2_(sh));
    }
    var leg = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
    if (leg) {
      acumularDesdeDatos_(valoresNotasHojaDesdeFila2_(leg));
    }
  } else {
    var hoja = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
    if (!hoja) return [];
    acumularDesdeDatos_(valoresNotasHojaDesdeFila2_(hoja));
  }
  out.sort(function (a, b) {
    return limpiarTexto(a).localeCompare(limpiarTexto(b), "es", { sensitivity: "base", numeric: true });
  });
  return out;
}

/**
 * Catálogo institucional: hoja **Materias**, columna B = nombre (fila 1 = encabezados).
 * El orden del portal estudiante sigue el orden de filas del catálogo; luego se añaden
 * materias de DocenteAsignaciones/Notas que no estén en el catálogo (p. ej. nombre distinto).
 */
function listarNombresMateriasCatalogoHoja_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName("Materias");
  if (!hoja) return [];
  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return [];
  var out = [];
  var seen = {};
  var r;
  var nombre;
  var nk;
  for (r = 1; r < datos.length; r++) {
    nombre = String(datos[r][1] == null ? "" : datos[r][1]).trim();
    if (!nombre) continue;
    nk = limpiarTexto(nombre);
    if (seen[nk]) continue;
    seen[nk] = true;
    out.push(nombre);
  }
  return out;
}

/**
 * Boletín y portal estudiante: (1) hoja **Materias** si existe; si no, unión de DocenteAsignaciones
 * (B=grupo, C=materia) y Notas del estudiante. (2) Se añaden transversales en materiasTransversalesIEMFS_
 * si no vinieron en el catálogo (p. ej. alias «Competencia lectora»).
 */
function materiasTransversalesIEMFS_() {
  return ["Competencia lectora"];
}

function fusionarMateriasConCatalogoHoja_(baseLista, catalogo) {
  var cat = Array.isArray(catalogo) ? catalogo : [];
  var seen = {};
  var out = [];
  var i;
  var m;
  var mk;
  var base = Array.isArray(baseLista) ? baseLista : [];
  for (i = 0; i < cat.length; i++) {
    m = cat[i];
    mk = limpiarTexto(m);
    if (seen[mk]) continue;
    seen[mk] = true;
    out.push(m);
  }
  var appendix = [];
  for (i = 0; i < base.length; i++) {
    m = base[i];
    mk = limpiarTexto(m);
    if (seen[mk]) continue;
    seen[mk] = true;
    appendix.push(m);
  }
  appendix.sort(function (a, b) {
    return limpiarTexto(a).localeCompare(limpiarTexto(b), "es", { sensitivity: "base", numeric: true });
  });
  for (i = 0; i < appendix.length; i++) {
    out.push(appendix[i]);
  }
  var extras = materiasTransversalesIEMFS_();
  for (i = 0; i < extras.length; i++) {
    m = extras[i];
    mk = limpiarTexto(m);
    if (seen[mk]) continue;
    seen[mk] = true;
    out.push(m);
  }
  return out;
}

function fusionarMateriasExtrasCurriculoIEMFS_(lista) {
  var base = Array.isArray(lista) ? lista : [];
  var seen = {};
  var out = [];
  var i;
  var m;
  var mk;
  for (i = 0; i < base.length; i++) {
    m = base[i];
    mk = limpiarTexto(m);
    if (seen[mk]) continue;
    seen[mk] = true;
    out.push(m);
  }
  var extras = materiasTransversalesIEMFS_();
  for (i = 0; i < extras.length; i++) {
    m = extras[i];
    mk = limpiarTexto(m);
    if (seen[mk]) continue;
    seen[mk] = true;
    out.push(m);
  }
  out.sort(function (a, b) {
    return limpiarTexto(a).localeCompare(limpiarTexto(b), "es", { sensitivity: "base", numeric: true });
  });
  return out;
}

function listarMateriasParaEstudiante_(grupo, estudianteId) {
  var fromAsig = listarMateriasUnicasGrupo_(grupo);
  var fromNotas = listarMateriasDistintasNotasEstudiante_(estudianteId);
  var base;
  var seen = {};
  var i;
  var m;
  var mk;
  if (!fromAsig.length) {
    base = fromNotas.slice();
  } else if (!fromNotas.length) {
    base = fromAsig.slice();
  } else {
    base = [];
    for (i = 0; i < fromAsig.length; i++) {
      m = fromAsig[i];
      mk = limpiarTexto(m);
      if (seen[mk]) continue;
      seen[mk] = true;
      base.push(m);
    }
    for (i = 0; i < fromNotas.length; i++) {
      m = fromNotas[i];
      mk = limpiarTexto(m);
      if (seen[mk]) continue;
      seen[mk] = true;
      base.push(m);
    }
  }
  var catalogo = listarNombresMateriasCatalogoHoja_();
  if (catalogo.length > 0) {
    return fusionarMateriasConCatalogoHoja_(base, catalogo);
  }
  base.sort(function (a, b) {
    return limpiarTexto(a).localeCompare(limpiarTexto(b), "es", { sensitivity: "base", numeric: true });
  });
  return fusionarMateriasExtrasCurriculoIEMFS_(base);
}

function asignacionPermite_(codigoDocente, grupo, materia) {
  var g = limpiarTexto(grupo);
  var m = limpiarTexto(materia);
  var arr = obtenerAsignacionesDocente_(codigoDocente);
  for (var i = 0; i < arr.length; i++) {
    if (limpiarTexto(arr[i].grupo) === g && limpiarTexto(arr[i].materia) === m) return true;
  }
  return false;
}

function gruposUnicosDesdeAsignacionesDocente_(codigo) {
  var arr = obtenerAsignacionesDocente_(codigo);
  var seen = {};
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var g = String(arr[i].grupo || "").trim();
    if (!g || seen[g]) continue;
    seen[g] = true;
    out.push(g);
  }
  out.sort(function (a, b) {
    return limpiarTexto(a).localeCompare(limpiarTexto(b), "es", { sensitivity: "base", numeric: true });
  });
  return out;
}

/**
 * Ejecutar desde el editor (▶): lista docentes y asignaciones en el Registro de ejecución (Ctrl+Enter / Ver → Registros).
 * Si la hoja no existía, la crea vacía con encabezados (codigo, nombre, clave).
 *
 * Por seguridad, por defecto NO imprime la contraseña en texto plano; solo longitud.
 * Para ver la clave mientras configuras cuentas: ponga VER_CLAVES_DOCENTES_EN_PLANO a true abajo y borre después.
 */
function EJECUTAR_VER_DOCENTES_Y_ASIGNACIONES() {
  var VER_CLAVES_DOCENTES_EN_PLANO = false;

  var ss = SpreadsheetApp.getActive();
  var sh = obtenerHojaDocentes_();
  var datos = sh.getDataRange().getValues();

  Logger.log("=== Libro: " + ss.getName() + " ===");
  Logger.log(
    "=== Hoja «Docentes» (filas de datos: " + Math.max(0, datos.length - 1) + ") ==="
  );
  Logger.log("Columnas esperadas: A=codigo | B=nombre | C=clave (la misma que escribe el docente en el login junto al código)");
  if (datos.length < 2) {
    Logger.log("(Vacío: solo encabezado.) Agregue filas debajo del encabezado.");
  }
  for (var i = 1; i < datos.length; i++) {
    var cod = String(datos[i][0] || "").trim();
    var nom = String(datos[i][1] || "").trim();
    var cl = String(datos[i][2] == null ? "" : datos[i][2]).trim();
    var infoClave;
    if (!cl.length) {
      infoClave = "clave (vacía — no podrá entrar)";
    } else if (VER_CLAVES_DOCENTES_EN_PLANO) {
      infoClave = "clave=" + cl;
    } else {
      infoClave = "clave (" + cl.length + " caracteres, oculta)";
    }
    Logger.log("Fila " + (i + 1) + " | codigo=[" + cod + "] nombre=[" + nom + "] " + infoClave);
  }

  var sha = obtenerHojaDocenteAsignaciones_();
  var da = sha.getDataRange().getValues();
  Logger.log(
    "=== Hoja «DocenteAsignaciones» (filas de datos: " + Math.max(0, da.length - 1) + ") ==="
  );
  Logger.log("Columnas: A=docente_codigo | B=grupo | C=materia (debe coincidir con grupos/materias del sistema)");
  if (da.length < 2) {
    Logger.log("(Vacío.) Sin filas aquí el docente verá error: «no tiene asignaciones».");
  }
  for (var j = 1; j < da.length; j++) {
    Logger.log(
      "Fila " +
        (j + 1) +
        " | " +
        String(da[j][0] || "").trim() +
        " | " +
        String(da[j][1] || "").trim() +
        " | " +
        String(da[j][2] || "").trim()
    );
  }

  try {
    ss.setActiveSheet(sh);
  } catch (e0) {
    /* ignore */
  }
  return "Listo. Abra «Ver» → «Registros» o el panel de registro de ejecución para leer la salida.";
}

/**
 * Ejecutar una sola vez (▶): inserta un docente y una asignación de ejemplo si no existe el código D001.
 * Edite abajo CODIGO, CLAVE, grupo o materia antes de ejecutar si lo desea.
 * After testing, cambie la clave en la hoja y borre esta fila de ejemplo si no la necesita.
 */
function EJECUTAR_INSERTAR_DOCENTE_EJEMPLO() {
  var CODIGO = "D001";
  var NOMBRE = "Docente de prueba (editar)";
  /** Texto que debe escribir el docente en «Contraseña» del login (columna C de «Docentes»). */
  var CLAVE_DOCENTE = "DocenteDemo2024";
  var MATERIA = "Matemáticas";
  var GRUPO_SI_NO_HAY_ESTUDIANTES = "8-9 B";

  var shD = obtenerHojaDocentes_();
  var datosD = shD.getDataRange().getValues();
  for (var i = 1; i < datosD.length; i++) {
    if (String(datosD[i][0] || "").trim() === CODIGO) {
      Logger.log(
        'Ya existe una fila con codigo="' +
          CODIGO +
          '". No se insertó nada. Cambie CODIGO arriba o borre esa fila en «Docentes».'
      );
      return "Sin cambios: código ya existe.";
    }
  }

  var grupo = GRUPO_SI_NO_HAY_ESTUDIANTES;
  var hEst = SpreadsheetApp.getActive().getSheetByName("Estudiantes");
  if (hEst) {
    var r = hEst.getDataRange().getValues();
    if (r.length > 1) {
      var gTomado = String(r[1][2] == null ? "" : r[1][2]).trim();
      if (gTomado) grupo = gTomado;
    }
  }

  shD.appendRow([CODIGO, NOMBRE, CLAVE_DOCENTE]);

  var shA = obtenerHojaDocenteAsignaciones_();
  shA.appendRow([CODIGO, grupo, MATERIA]);

  Logger.log("Insertado en Docentes: codigo=" + CODIGO + " | clave (longitud)=" + String(CLAVE_DOCENTE).length);
  Logger.log(
    "Insertado en DocenteAsignaciones: " + CODIGO + " | grupo=" + grupo + " | materia=" + MATERIA
  );
  Logger.log(
    "Login web: campo «Código de docente» = " +
      CODIGO +
      "; campo «Contraseña» = exactamente lo guardado en columna C (clave propia del docente), NO la clave de coordinación."
  );
  Logger.log("Tras probar, cambie la contraseña en la columna C de «Docentes» y ejecute de nuevo «Ver» docentes si lo necesita.");

  try {
    SpreadsheetApp.getActive().setActiveSheet(shD);
  } catch (e1) {
    /* ignore */
  }
  return "Filas de ejemplo añadidas. Revise el registro y la hoja Docentes.";
}

function obtenerGrupoEstudiantePorId_(estudianteId) {
  var idB = String(estudianteId || "").trim();
  if (!idB) return "";
  var hoja = SpreadsheetApp.getActive().getSheetByName("Estudiantes");
  if (!hoja) return "";
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0] || "").trim() !== idB) continue;
    return String(datos[i][2] == null ? "" : datos[i][2]).trim();
  }
  return "";
}

function docentePuedeVerEstudiantePorGrupo_(codigoDocente, grupoEstudiante) {
  var g = limpiarTexto(grupoEstudiante);
  var arr = obtenerAsignacionesDocente_(codigoDocente);
  for (var i = 0; i < arr.length; i++) {
    if (limpiarTexto(arr[i].grupo) === g) return true;
  }
  return false;
}

function leerMetaNotaPorFila_(filaNota) {
  var ref = parseNotaIdRef_(filaNota);
  if (!ref || ref.fila < 2) return null;
  var hoja = hojaNotaDesdeRef_(ref);
  if (!hoja) return null;
  var lr = hoja.getLastRow();
  if (ref.fila > lr) return null;
  var row = hoja.getRange(ref.fila, 1, 1, 8).getValues()[0];
  return {
    estudiante: String(row[1] || "").trim(),
    materia: limpiarTexto(row[2]),
    grupo: String(row[3] == null ? "" : row[3]).trim()
  };
}

/** Normaliza número de documento para comparar (sin puntos ni espacios). */
function normalizarDocumentoEstudiante_(txt) {
  return String(txt || "")
    .replace(/\./g, "")
    .replace(/\s/g, "")
    .trim();
}

/**
 * PIN numérico de 1 a 4 dígitos: unifica 394, 0394 y "394" a "0394" para el login.
 * Si el PIN no es solo dígitos, se compara tal cual.
 */
function normalizarPinEstudianteParaLogin_(raw) {
  var s = String(raw == null ? "" : raw).trim();
  if (s === "") return "";
  if (!/^\d{1,4}$/.test(s)) return s;
  var n = parseInt(s, 10);
  if (isNaN(n) || n < 0 || n > 9999) return s;
  return ("0000" + String(n)).slice(-4);
}

/**
 * Hoja Estudiantes: columnas E=documento, F=PIN (portal estudiante solo lectura).
 * @return {{ id: string, nombre: string, grupo: string } | null}
 */
function obtenerEstudiantePorDocumentoYPin_(documentoNorm, pin) {
  var p = String(pin || "").trim();
  if (!documentoNorm || !p) return null;
  var pinNorm = normalizarPinEstudianteParaLogin_(p);
  var hoja = SpreadsheetApp.getActive().getSheetByName("Estudiantes");
  if (!hoja) return null;
  var datos = hoja.getDataRange().getValues();
  var colFoto = indiceColumnaFotoEstudiantes_(datos[0] || []);
  var colSheetFoto = colFoto + 1;
  for (var i = 1; i < datos.length; i++) {
    var docCell = normalizarDocumentoEstudiante_(String(datos[i][4] == null ? "" : datos[i][4]));
    var pinCell = String(datos[i][5] == null ? "" : datos[i][5]).trim();
    if (docCell !== documentoNorm) continue;
    if (normalizarPinEstudianteParaLogin_(pinCell) !== pinNorm) continue;
    return {
      id: String(datos[i][0] || "").trim(),
      nombre: String(datos[i][1] == null ? "" : datos[i][1]).trim(),
      grupo: String(datos[i][2] == null ? "" : datos[i][2]).trim(),
      fotoUrl: fotoEstudianteDesdeFilaYFormulaIdx_(
        datos[i],
        hoja.getRange(i + 1, colSheetFoto).getFormula(),
        colFoto
      )
    };
  }
  return null;
}

/** PIN aleatorio de exactamente 4 dígitos (0000–9999). */
function generarPinCuatroDigitos_() {
  var n = Math.floor(Math.random() * 10000);
  return ("0000" + String(n)).slice(-4);
}

/**
 * Ejecutar una vez desde el editor (▶): escribe en columna **F** un PIN de 4 dígitos
 * solo en filas que aún no tienen PIN. No modifica PIN ya existentes.
 * Hoja **Estudiantes**: columnas E=documento, F=PIN.
 */
function EJECUTAR_ASIGNAR_PIN_FALTANTES_ESTUDIANTES() {
  var hoja = SpreadsheetApp.getActive().getSheetByName("Estudiantes");
  if (!hoja) {
    throw new Error("No existe la hoja Estudiantes.");
  }
  var datos = hoja.getDataRange().getValues();
  if (datos.length < 2) {
    return "No hay filas de datos bajo el encabezado.";
  }

  var pinsUsados = {};
  for (var r = 1; r < datos.length; r++) {
    var pinEx = String(datos[r][5] == null ? "" : datos[r][5]).trim();
    if (pinEx) {
      pinsUsados[pinEx] = true;
    }
  }

  var actualizados = 0;
  for (var i = 1; i < datos.length; i++) {
    var pinActual = String(datos[i][5] == null ? "" : datos[i][5]).trim();
    if (pinActual !== "") {
      continue;
    }
    var nuevo;
    var intentos = 0;
    do {
      nuevo = generarPinCuatroDigitos_();
      intentos++;
    } while (pinsUsados[nuevo] && intentos < 200);
    if (pinsUsados[nuevo]) {
      throw new Error(
        "No se pudo generar un PIN único tras varios intentos. Ejecute de nuevo o revise PIN duplicados en columna F."
      );
    }
    pinsUsados[nuevo] = true;
    hoja.getRange(i + 1, 6).setValue(nuevo);
    actualizados++;
  }

  return (
    "Listo: PIN asignado en " +
    actualizados +
    " fila(s) (columna F). Las filas que ya tenían PIN no se modificaron."
  );
}

/**
 * @typedef {{ tipo: string, docenteId: string, nombre: string, error: string, estudianteId: string, grupoEstudiante: string }}
 */
function resolverAuthGet_(e) {
  var p = (e && e.parameter) || {};
  var loginTipo = String(p.loginTipo || p.logintipo || "").trim().toLowerCase();
  if (loginTipo === "estudiante") {
    var docRaw = String(p.documento || "").trim();
    var pinEs = String(p.pin || "").trim();
    var docN = normalizarDocumentoEstudiante_(docRaw);
    if (!docN || !pinEs) {
      return {
        tipo: "",
        docenteId: "",
        nombre: "",
        error: "Ingrese documento y PIN.",
        estudianteId: "",
        grupoEstudiante: ""
      };
    }
    var estG = obtenerEstudiantePorDocumentoYPin_(docN, pinEs);
    if (!estG) {
      return {
        tipo: "",
        docenteId: "",
        nombre: "",
        error: "Documento o PIN incorrectos.",
        estudianteId: "",
        grupoEstudiante: ""
      };
    }
    return {
      tipo: "estudiante",
      docenteId: "",
      nombre: estG.nombre,
      error: "",
      estudianteId: estG.id,
      grupoEstudiante: estG.grupo
    };
  }

  var docenteId = String(p.docenteId || "").trim();
  var clave = String(p.clave || "").trim();
  if (!clave) {
    return { tipo: "", docenteId: "", nombre: "", error: "Falta clave de acceso." };
  }
  if (!docenteId) {
    return {
      tipo: "",
      docenteId: "",
      nombre: "",
      error: "Ingrese el código de usuario (docente o coordinación, p. ej. " + obtenerCodigoCoordinador_() + ")."
    };
  }
  var codCoordG = obtenerCodigoCoordinador_();
  if (limpiarTexto(docenteId) === limpiarTexto(codCoordG)) {
    var espG = obtenerClaveEsperada_();
    if (!espG) {
      return {
        tipo: "",
        docenteId: "",
        nombre: "",
        error: "Falta configurar la clave en Apps Script."
      };
    }
    if (clave !== espG) {
      return { tipo: "", docenteId: "", nombre: "", error: "Clave incorrecta." };
    }
    return { tipo: "coordinador", docenteId: "", nombre: "", error: "" };
  }
  var doc = obtenerDocentePorCodigoYClave_(docenteId, clave);
  if (!doc) {
    return {
      tipo: "",
      docenteId: "",
      nombre: "",
      error: "Código de coordinación o de docente, o contraseña incorrecta."
    };
  }
  var asig = obtenerAsignacionesDocente_(doc.codigo);
  if (!asig.length) {
    return {
      tipo: "",
      docenteId: "",
      nombre: "",
      error: "Este docente no tiene filas en la hoja DocenteAsignaciones (grupo + materia)."
    };
  }
  return { tipo: "docente", docenteId: doc.codigo, nombre: doc.nombre, error: "" };
}

function resolverAuthPost_(data) {
  data = data || {};
  var loginTipoP = String(data.loginTipo || data.logintipo || "").trim().toLowerCase();
  if (loginTipoP === "estudiante") {
    var docRawP = String(data.documento || "").trim();
    var pinP = String(data.pin || "").trim();
    var docNP = normalizarDocumentoEstudiante_(docRawP);
    if (!docNP || !pinP) {
      return {
        tipo: "",
        docenteId: "",
        nombre: "",
        error: "Ingrese documento y PIN.",
        estudianteId: "",
        grupoEstudiante: ""
      };
    }
    var estP = obtenerEstudiantePorDocumentoYPin_(docNP, pinP);
    if (!estP) {
      return {
        tipo: "",
        docenteId: "",
        nombre: "",
        error: "Documento o PIN incorrectos.",
        estudianteId: "",
        grupoEstudiante: ""
      };
    }
    return {
      tipo: "estudiante",
      docenteId: "",
      nombre: estP.nombre,
      error: "",
      estudianteId: estP.id,
      grupoEstudiante: estP.grupo
    };
  }

  var docenteId = String(data.docenteId || "").trim();
  var clave = String(data.clave || "").trim();
  if (!clave) {
    return { tipo: "", docenteId: "", nombre: "", error: "Falta clave de acceso." };
  }
  if (!docenteId) {
    return {
      tipo: "",
      docenteId: "",
      nombre: "",
      error: "Ingrese el código de usuario (docente o coordinación, p. ej. " + obtenerCodigoCoordinador_() + ")."
    };
  }
  var codCoordP = obtenerCodigoCoordinador_();
  if (limpiarTexto(docenteId) === limpiarTexto(codCoordP)) {
    var espP = obtenerClaveEsperada_();
    if (!espP) {
      return {
        tipo: "",
        docenteId: "",
        nombre: "",
        error: "Falta configurar la clave en Apps Script."
      };
    }
    if (clave !== espP) {
      return { tipo: "", docenteId: "", nombre: "", error: "Clave incorrecta." };
    }
    return { tipo: "coordinador", docenteId: "", nombre: "", error: "" };
  }
  var doc = obtenerDocentePorCodigoYClave_(docenteId, clave);
  if (!doc) {
    return {
      tipo: "",
      docenteId: "",
      nombre: "",
      error: "Código de coordinación o de docente, o contraseña incorrecta."
    };
  }
  var asig = obtenerAsignacionesDocente_(doc.codigo);
  if (!asig.length) {
    return {
      tipo: "",
      docenteId: "",
      nombre: "",
      error: "Este docente no tiene asignaciones en DocenteAsignaciones."
    };
  }
  return { tipo: "docente", docenteId: doc.codigo, nombre: doc.nombre, error: "" };
}

function errorPermisoGrupoMateria_() {
  return "No tiene permiso para esta combinación de grupo y materia.";
}

function errorPermisoEstudiante_() {
  return "No tiene permiso para ver o modificar este estudiante.";
}

function errorCoordinadorSoloLectura_() {
  return "El acceso de coordinación es solo consulta. Para crear o modificar datos debe iniciar sesión como docente.";
}

function asegurarGrupoMateriaDocente_(auth, grupo, materia) {
  if (auth.tipo === "coordinador") return "";
  if (auth.tipo === "estudiante") {
    var ge = String(auth.grupoEstudiante || "").trim();
    if (!ge || normalizarGrupoCursoIEMFS_(grupo) !== normalizarGrupoCursoIEMFS_(ge)) {
      return "No tiene permiso para ver este curso.";
    }
    return "";
  }
  if (auth.tipo === "docente") {
    if (!asignacionPermite_(auth.docenteId, grupo, materia)) {
      return errorPermisoGrupoMateria_();
    }
  }
  return "";
}

/**
 * Permiso por grupo (sin materia): docente debe tener al menos una asignación en ese grupo;
 * coordinación pasa siempre; estudiante solo si coincide con su grupo.
 * Útil para vistas globales del grupo (p. ej. mosaico, ficha del grupo).
 */
function asegurarGrupoSoloDocente_(auth, grupo) {
  if (auth.tipo === "coordinador") return "";
  if (auth.tipo === "estudiante") {
    var ge = String(auth.grupoEstudiante || "").trim();
    if (!ge || normalizarGrupoCursoIEMFS_(grupo) !== normalizarGrupoCursoIEMFS_(ge)) {
      return "No tiene permiso para ver este curso.";
    }
    return "";
  }
  if (auth.tipo === "docente") {
    if (!docentePuedeVerEstudiantePorGrupo_(auth.docenteId, grupo)) {
      return errorPermisoGrupoMateria_();
    }
  }
  return "";
}

function asegurarEstudianteMateriaDocente_(auth, estudianteId, materia) {
  if (auth.tipo === "coordinador") return "";
  if (auth.tipo === "estudiante") {
    if (String(auth.estudianteId || "").trim() !== String(estudianteId || "").trim()) {
      return errorPermisoEstudiante_();
    }
    return "";
  }
  var gr = obtenerGrupoEstudiantePorId_(estudianteId);
  if (!gr) return "Estudiante no encontrado.";
  if (!asignacionPermite_(auth.docenteId, gr, materia)) {
    return errorPermisoGrupoMateria_();
  }
  return "";
}

function asegurarEstudianteVistaDocente_(auth, estudianteId) {
  if (auth.tipo === "coordinador") return "";
  if (auth.tipo === "estudiante") {
    if (String(auth.estudianteId || "").trim() !== String(estudianteId || "").trim()) {
      return errorPermisoEstudiante_();
    }
    return "";
  }
  var gr = obtenerGrupoEstudiantePorId_(estudianteId);
  if (!gr) return "Estudiante no encontrado.";
  if (!docentePuedeVerEstudiantePorGrupo_(auth.docenteId, gr)) {
    return errorPermisoEstudiante_();
  }
  return "";
}

/**
 * Guarda la clave en Propiedades del script (NO la escribes en la web, solo aquí una vez).
 *
 * IMPORTANTE: Si eliges "establecerClaveAcceso" en el menú y pulsas Ejecutar ▶, Google NO puede
 * pasar el parámetro "clave" → queda vacío → error "al menos 6 caracteres".
 * Siempre usa EJECUTAR_PARA_GUARDAR_MI_CLAVE (abajo).
 */
function establecerClaveAcceso(clave) {
  var c = String(clave || "").trim();
  if (c.length < 6) {
    throw new Error("Usa una clave de al menos 6 caracteres.");
  }
  PropertiesService.getScriptProperties().setProperty(PROP_CLAVE_NOTAS_, c);
  return "Clave guardada en propiedades del proyecto.";
}

/**
 * ✅ PRIMERA VEZ — Haz solo esto:
 * 1. En la línea "var miClave = ..." pon TU clave entre comillas (mínimo 6 caracteres).
 * 2. Guarda el archivo (Ctrl+S).
 * 3. En el desplegable de funciones elige: EJECUTAR_PARA_GUARDAR_MI_CLAVE (NO establecerClaveAcceso).
 * 4. Pulsa Ejecutar ▶. Acepta permisos.
 * 5. Borra la clave del código o deja miClave = "" para no dejarla escrita.
 */
function EJECUTAR_PARA_GUARDAR_MI_CLAVE() {
  var miClave = ""; // ← ESCRIBE TU CLAVE ENTRE COMILLAS, ejemplo: "IEMFS_Notas_2026"

  if (!miClave || String(miClave).trim().length < 6) {
    throw new Error(
      'Edita EJECUTAR_PARA_GUARDAR_MI_CLAVE: pon tu clave en var miClave = "tu_clave_aqui";'
    );
  }
  return establecerClaveAcceso(miClave);
}

/** @deprecated Usa EJECUTAR_PARA_GUARDAR_MI_CLAVE (mismo propósito). */
function configurarClaveUnaSolaVez() {
  return EJECUTAR_PARA_GUARDAR_MI_CLAVE();
}

// ===============================
// HOJA AUDITORÍA
// ===============================
function obtenerHojaAuditoria_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Auditoria");
  if (!sh) {
    sh = ss.insertSheet("Auditoria");
    sh.appendRow([
      "timestamp",
      "estudiante",
      "materia",
      "grupo",
      "filaNota",
      "detalle",
      "motivo",
      "fechaRegistro"
    ]);
  }
  return sh;
}

function registrarAuditoria_(filaNota, estudianteId, materiaOriginal, grupo, detalle, motivo) {
  obtenerHojaAuditoria_().appendRow([
    new Date().getTime(),
    String(estudianteId).trim(),
    String(materiaOriginal),
    String(grupo || ""),
    String(filaNota),
    String(detalle || ""),
    String(motivo || "").trim(),
    new Date()
  ]);
}

function obtenerHistorialPorFilaNota_(filaNota) {
  const idBus = String(filaNota).trim();
  if (!idBus) return [];

  const sh = SpreadsheetApp.getActive().getSheetByName("Auditoria");
  if (!sh) return [];

  const datos = sh.getDataRange().getValues();
  const items = [];

  for (let r = 1; r < datos.length; r++) {
    const filaReg = String(datos[r][4] != null ? datos[r][4] : "").trim();
    if (filaReg !== idBus) continue;

    const fechaCell = datos[r][7];
    const ts = datos[r][0];
    let fechaIso;
    if (fechaCell instanceof Date) {
      fechaIso = fechaCell.toISOString();
    } else if (ts instanceof Date) {
      fechaIso = ts.toISOString();
    } else {
      fechaIso = new Date(Number(ts) || Date.now()).toISOString();
    }

    const detalle = String(datos[r][5] || "");
    const motivo = String(datos[r][6] || "");
    const texto = motivo ? detalle + " · Motivo: " + motivo : detalle;

    items.push({
      fecha: fechaIso,
      detalle: texto,
      motivo: motivo
    });
  }

  items.sort(function (a, b) {
    return new Date(a.fecha).getTime() - new Date(b.fecha).getTime();
  });

  return items;
}

// ===============================
// doGet
// ===============================
function doGet(e) {
  const accion = String(e && e.parameter && e.parameter.accion ? e.parameter.accion : "").trim();

  var auth = resolverAuthGet_(e);
  if (auth.error) {
    return json_({ ok: false, error: auth.error });
  }

  if (!accion) {
    return json_({
      ok: true,
      message: "API de NotasIEMFS activa",
      endpoints: [
        "ping",
        "miContexto",
        "grupos",
        "estudiantes",
        "notas",
        "resumen",
        "historial",
        "historialEstudiante",
        "notasGrupo",
        "inasistenciasDia",
        "inasistenciasFechasGrupo",
        "inasistenciasEstudiante",
        "periodosAcademicos",
        "inasistenciasResumenPeriodo",
        "boletinResumenEstudiante",
        "notasEstudianteAgregado",
        "materiasEstudiante",
        "mosaicoGrupo",
        "migrarGrupoNotas (POST + claveAdmin)",
        "inasistenciasGuardar (POST)",
        "guardarMasivo (POST)"
      ]
    });
  }

  if (accion === "ping") {
    var pingOut = {
      ok: true,
      auth: true,
      tipo: auth.tipo,
      nombre: auth.nombre || "",
      notasMultihoja: notasMultihojaActivas_(),
      inasistenciasMultihoja: inasistenciasMultihojaActivas_(),
      notasIemfsBuild: NOTAS_IEMFS_BUILD_
    };
    if (auth.tipo === "estudiante") {
      pingOut.estudianteId = auth.estudianteId;
      pingOut.grupoEstudiante = auth.grupoEstudiante;
    }
    return json_(pingOut);
  }

  if (accion === "miContexto") {
    if (auth.tipo === "coordinador") {
      return json_({ tipo: "coordinador", nombre: "", asignaciones: null });
    }
    if (auth.tipo === "estudiante") {
      return json_({
        tipo: "estudiante",
        nombre: auth.nombre,
        estudianteId: auth.estudianteId,
        grupo: auth.grupoEstudiante,
        asignaciones: null
      });
    }
    return json_({
      tipo: "docente",
      nombre: auth.nombre,
      asignaciones: obtenerAsignacionesDocente_(auth.docenteId)
    });
  }

  if (accion === "grupos") {
    if (auth.tipo === "coordinador") {
      return json_(obtenerGruposLista_());
    }
    if (auth.tipo === "estudiante") {
      var gUn = String(auth.grupoEstudiante || "").trim();
      return json_(gUn ? [gUn] : []);
    }
    return json_(gruposUnicosDesdeAsignacionesDocente_(auth.docenteId));
  }

  if (accion === "estudiantes") {
    const grupo = e.parameter.grupo || "";
    if (auth.tipo === "estudiante") {
      if (normalizarGrupoCursoIEMFS_(grupo) !== normalizarGrupoCursoIEMFS_(auth.grupoEstudiante || "")) {
        return json_({ ok: false, error: "No tiene permiso para ver este grupo." });
      }
      var listaFull = obtenerEstudiantesPorGrupo(grupo);
      var soloYo = [];
      for (var si = 0; si < listaFull.length; si++) {
        if (String(listaFull[si].id).trim() === String(auth.estudianteId).trim()) {
          soloYo.push(listaFull[si]);
          break;
        }
      }
      return json_(soloYo);
    }
    if (auth.tipo === "docente") {
      var puedeEst = false;
      var asigEst = obtenerAsignacionesDocente_(auth.docenteId);
      for (var gi = 0; gi < asigEst.length; gi++) {
        if (limpiarTexto(asigEst[gi].grupo) === limpiarTexto(grupo)) {
          puedeEst = true;
          break;
        }
      }
      if (!puedeEst) {
        return json_({ ok: false, error: "No tiene permiso para ver este grupo." });
      }
    }
    return json_(obtenerEstudiantesPorGrupo(grupo));
  }

  if (accion === "materiasEstudiante") {
    if (auth.tipo !== "estudiante") {
      return json_({
        ok: false,
        error: "Solo el acceso estudiante puede obtener la lista de materias de su curso."
      });
    }
    var grMe = String(auth.grupoEstudiante || "").trim();
    var idMe = String(auth.estudianteId || "").trim();
    /** Objeto con clave `materias` para localizar la petición en DevTools (antes era solo un array JSON). */
    return json_({
      ok: true,
      materias: listarMateriasParaEstudiante_(grMe, idMe)
    });
  }

  if (accion === "notas") {
    const id = e.parameter.id || "";
    const materia = e.parameter.materia || "";
    var errN = asegurarEstudianteMateriaDocente_(auth, id, materia);
    if (errN) {
      return json_({ ok: false, error: errN });
    }
    var periodoN = e.parameter.periodo || "";
    return json_(obtenerDetalleEstudiante(id, materia, periodoN));
  }

  if (accion === "notasGrupo") {
    if (auth.tipo === "estudiante") {
      return json_({
        ok: false,
        error:
          "Los estudiantes ven las calificaciones en «Ver mi expediente» por materia, no la tabla del grupo completo."
      });
    }
    const grupo = e.parameter.grupo || "";
    const materia = e.parameter.materia || "";
    var errNg = asegurarGrupoMateriaDocente_(auth, grupo, materia);
    if (errNg) {
      return json_({ ok: false, error: errNg });
    }
    var periodoNg = e.parameter.periodo || "";
    return json_(obtenerNotasGrupoMateria_(grupo, materia, periodoNg));
  }

  if (accion === "resumen") {
    const id = e.parameter.id || "";
    const materia = e.parameter.materia || "";
    var errR = asegurarEstudianteMateriaDocente_(auth, id, materia);
    if (errR) {
      return json_({ ok: false, error: errR });
    }
    var periodoR = e.parameter.periodo || "";
    return json_(calcularResumen(id, materia, periodoR));
  }

  if (accion === "boletinResumenEstudiante") {
    if (auth.tipo !== "estudiante") {
      return json_({
        ok: false,
        error: "Solo el acceso estudiante puede ver el boletín resumen por materias."
      });
    }
    var idBol = String(auth.estudianteId || "").trim();
    if (!idBol) {
      return json_({ ok: false, error: "Sesión de estudiante no válida." });
    }
    var periodoBol = e.parameter.periodo || "";
    var grBol = String(auth.grupoEstudiante || "").trim();
    var itemsBol = boletinResumenEstudianteUnaPasada_(idBol, periodoBol, grBol);
    return json_({
      items: itemsBol,
      aviso:
        "La columna «Definitiva orientativa» usa la misma fórmula que el sistema (70% seguimiento, 10% actitudinal, 20% prueba). Si faltan registros de seguimiento, actitudinal o prueba, la cifra no es la calificación final oficial: es solo una referencia de cómo vas en ese momento."
    });
  }

  /**
   * Una sola lectura de la hoja Notas; notas del estudiante agrupadas por materia (fallback cliente rápido).
   */
  if (accion === "notasEstudianteAgregado") {
    if (auth.tipo !== "estudiante") {
      return json_({
        ok: false,
        error: "Solo estudiantes pueden usar esta consulta."
      });
    }
    var idAgg = String(auth.estudianteId || "").trim();
    if (!idAgg) {
      return json_({ ok: false, error: "Sesión de estudiante no válida." });
    }
    var periodoAgg = e.parameter.periodo || "";
    var grAgg = String(auth.grupoEstudiante || "").trim();
    return json_(notasEstudianteAgregadoPorMateria_(idAgg, periodoAgg, grAgg));
  }

  if (accion === "historial") {
    const notaId = e.parameter.notaId || e.parameter.id || "";
    var meta = leerMetaNotaPorFila_(notaId);
    if (!meta) {
      return json_({ ok: false, error: "Nota no encontrada." });
    }
    var errH = asegurarGrupoMateriaDocente_(auth, meta.grupo, meta.materia);
    if (errH) {
      return json_({ ok: false, error: errH });
    }
    return json_(obtenerHistorialPorFilaNota_(notaId));
  }

  if (accion === "historialEstudiante") {
    const estId = e.parameter.estudiante || e.parameter.id || "";
    var errHe = asegurarEstudianteVistaDocente_(auth, estId);
    if (errHe) {
      return json_({ ok: false, error: errHe });
    }
    return json_(obtenerHistorialEstadoMatricula_(estId));
  }

  if (accion === "inasistenciasDia") {
    if (auth.tipo === "estudiante") {
      return json_({
        ok: false,
        error: "El registro diario de lista es solo para docentes. Consulte inasistencias en su expediente."
      });
    }
    const grupo = e.parameter.grupo || "";
    const materia = e.parameter.materia || "";
    const fecha = e.parameter.fecha || "";
    var errI = asegurarGrupoMateriaDocente_(auth, grupo, materia);
    if (errI) {
      return json_({ ok: false, error: errI });
    }
    return json_(obtenerInasistenciasDia_(grupo, materia, fecha));
  }

  if (accion === "inasistenciasFechasGrupo") {
    if (auth.tipo === "estudiante") {
      return json_({
        ok: false,
        error: "Use el expediente por materia para ver sus inasistencias."
      });
    }
    const grupo = e.parameter.grupo || "";
    const materia = e.parameter.materia || "";
    var errF = asegurarGrupoMateriaDocente_(auth, grupo, materia);
    if (errF) {
      return json_({ ok: false, error: errF });
    }
    return json_(listarFechasConListaGrupoMateria_(grupo, materia));
  }

  if (accion === "inasistenciasEstudiante") {
    const estId = e.parameter.estudiante || e.parameter.id || "";
    const materia = e.parameter.materia || "";
    var errIe = asegurarEstudianteMateriaDocente_(auth, estId, materia);
    if (errIe) {
      return json_({ ok: false, error: errIe });
    }
    const limite = parseInt(String(e.parameter.limite || "60"), 10);
    var periodoIe = e.parameter.periodo || "";
    return json_(
      listarInasistenciasEstudiante_(estId, materia, isNaN(limite) ? 60 : limite, periodoIe)
    );
  }

  if (accion === "periodosAcademicos") {
    return json_(listarPeriodosAcademicos_());
  }

  if (accion === "inasistenciasResumenPeriodo") {
    if (auth.tipo === "estudiante") {
      return json_({
        ok: false,
        error:
          "El resumen de faltas por grupo es para coordinación y docentes. Consulte en su expediente por materia."
      });
    }
    const grupo = e.parameter.grupo || "";
    const materia = e.parameter.materia || "";
    const periodo = e.parameter.periodo || "1";
    var errRf = asegurarGrupoMateriaDocente_(auth, grupo, materia);
    if (errRf) {
      return json_({ ok: false, error: errRf });
    }
    return json_(obtenerResumenInasistenciasPeriodo_(grupo, materia, periodo));
  }

  if (accion === "mosaicoGrupo") {
    if (auth.tipo === "estudiante") {
      return json_({
        ok: false,
        error: "El mosaico de grupo es solo para coordinación y docentes."
      });
    }
    var grMos = String(e.parameter.grupo || "").trim();
    if (!grMos) {
      return json_({ ok: false, error: "Falta el grupo." });
    }
    var errMos = asegurarGrupoSoloDocente_(auth, grMos);
    if (errMos) {
      return json_({ ok: false, error: errMos });
    }
    return json_(obtenerMosaicoGrupo_(grMos));
  }

  return json_({
    ok: false,
    error: "Acción no válida"
  });
}

// ===============================
// doPost
// ===============================
function doPost(e) {
  try {
    const data = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};

    /**
     * GET al URL .../exec redirige (302) a script.googleusercontent.com y la query suele perderse:
     * el handler recibe acción sin loginTipo/documento/pin → «Falta clave de acceso».
     * Las lecturas van por POST (cuerpo intacto) y se delegan en doGet con parámetros planos.
     */
    if (data.__readViaPost) {
      var param = {};
      Object.keys(data).forEach(function (k) {
        if (k === "__readViaPost") return;
        var v = data[k];
        if (v === undefined || v === null) return;
        if (typeof v === "string" && v === "") return;
        param[k] = String(v);
      });
      return doGet({ parameter: param });
    }

    const accion = String(data.accion || "").trim();

    if (accion === "migrarGrupoNotas") {
      var errAdmMg = validarClaveAdminEnData_(data);
      if (errAdmMg) {
        return json_({ ok: false, error: errAdmMg });
      }
      var authMg = resolverAuthPost_(data);
      if (authMg.error) {
        return json_({ ok: false, error: authMg.error });
      }
      if (authMg.tipo === "estudiante") {
        return json_({ ok: false, error: "Acceso no autorizado." });
      }
      if (authMg.tipo === "docente") {
        return json_({ ok: false, error: "Solo coordinación puede realizar traslados de curso." });
      }
      const rM = migrarGrupoNotasEstudiante_(data);
      return json_({
        ok: true,
        message: "Traslado de curso aplicado",
        actualizadas: rM.actualizadas,
        estudianteActualizado: rM.estudianteActualizado
      });
    }

    if (accion === "estadoEstudiante") {
      var errAdmEst = validarClaveAdminEnData_(data);
      if (errAdmEst) {
        return json_({ ok: false, error: errAdmEst });
      }
      var authEst = resolverAuthPost_(data);
      if (authEst.error) {
        return json_({ ok: false, error: authEst.error });
      }
      if (authEst.tipo === "estudiante") {
        return json_({ ok: false, error: "Acceso no autorizado." });
      }
      if (authEst.tipo === "docente") {
        return json_({ ok: false, error: "Solo coordinación puede modificar el estado de matrícula." });
      }
      const rEstAdm = actualizarEstadoMatriculaEstudiante_(data);
      return json_({
        ok: true,
        message: "Estado actualizado",
        estadoMatricula: rEstAdm.estadoMatricula
      });
    }

    var authP = resolverAuthPost_(data);
    if (authP.error) {
      return json_({ ok: false, error: authP.error });
    }

    if (authP.tipo === "coordinador") {
      return json_({ ok: false, error: errorCoordinadorSoloLectura_() });
    }
    if (authP.tipo === "estudiante") {
      return json_({ ok: false, error: "El acceso de estudiante es solo consulta." });
    }

    if (accion === "guardarMasivo") {
      var errGm = asegurarGrupoMateriaDocente_(authP, String(data.grupo || "").trim(), String(data.materia || ""));
      if (errGm) {
        return json_({ ok: false, error: errGm });
      }
      var rGm = guardarNotasMasivo_(data, authP);
      return json_({
        ok: true,
        message: "Notas guardadas",
        guardadas: rGm.guardadas
      });
    }

    if (accion === "guardar") {
      var idGu = String(data.estudiante || "").trim();
      var grRealGu = String(obtenerGrupoEstudiantePorId_(idGu) || "").trim();
      if (!grRealGu) {
        return json_({
          ok: false,
          error:
            "Estudiante no encontrado en la hoja Estudiantes. Recargue el grupo en pantalla tras cambios manuales."
        });
      }
      var errGu = asegurarGrupoMateriaDocente_(authP, grRealGu, String(data.materia || "").trim());
      if (errGu) {
        return json_({ ok: false, error: errGu });
      }
      guardarNota(data);
      return json_({ ok: true, message: "Nota guardada" });
    }

    if (accion === "actualizar") {
      var errAc = asegurarGrupoMateriaDocente_(authP, String(data.grupo || "").trim(), String(data.materia || ""));
      if (errAc) {
        return json_({ ok: false, error: errAc });
      }
      actualizarNota_(data);
      return json_({ ok: true, message: "Nota actualizada" });
    }

    if (accion === "mejorar") {
      var grMe = obtenerGrupoEstudiantePorId_(data.estudiante);
      var errMe = asegurarGrupoMateriaDocente_(authP, grMe, String(data.materia || ""));
      if (errMe) {
        return json_({ ok: false, error: errMe });
      }
      mejorarNota(
        data.estudiante,
        data.materia,
        data.titulo,
        data.nuevaNota,
        data.motivo
      );
      return json_({ ok: true, message: "Mejora registrada" });
    }

    if (accion === "inasistenciasGuardar") {
      var errIn = asegurarGrupoMateriaDocente_(authP, String(data.grupo || "").trim(), String(data.materia || ""));
      if (errIn) {
        return json_({ ok: false, error: errIn });
      }
      const rIn = guardarInasistenciasMasivo_(data);
      return json_({
        ok: true,
        message: "Inasistencias guardadas",
        actualizadas: rIn.actualizadas,
        creadas: rIn.creadas
      });
    }

    return json_({ ok: false, error: "Acción no válida" });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// ===============================
// ESTADOS DE MATRÍCULA (columna D en hoja Estudiantes, índice 3)
// Valores guardados en hoja: texto canónico (ej. Activo, Inactivo…)
// ===============================
var MAP_ESTADO_MATRICULA_ = {
  activo: "Activo",
  inactivo: "Inactivo",
  desertor: "Desertor",
  "matricula cancelada": "Matrícula cancelada"
};

function etiquetaEstadoMatriculaDesdeCelda_(celda) {
  var norm = limpiarTexto(celda);
  if (!norm) return "Activo";
  if (MAP_ESTADO_MATRICULA_[norm]) return MAP_ESTADO_MATRICULA_[norm];
  var t = String(celda || "").trim();
  return t || "Activo";
}

function normalizarEstadoMatriculaInput_(s) {
  var norm = limpiarTexto(s);
  if (MAP_ESTADO_MATRICULA_[norm]) return MAP_ESTADO_MATRICULA_[norm];
  throw new Error(
    "Estado no válido. Use: Activo, Inactivo, Desertor o Matrícula cancelada."
  );
}

function obtenerHistorialEstadoMatricula_(estudianteId) {
  var idB = String(estudianteId || "").trim();
  if (!idB) return [];

  var sh = SpreadsheetApp.getActive().getSheetByName("Auditoria");
  if (!sh) return [];

  var datos = sh.getDataRange().getValues();
  var items = [];

  for (var r = 1; r < datos.length; r++) {
    var idF = String(datos[r][1] || "").trim();
    if (idF !== idB) continue;

    var filaN = parseInt(String(datos[r][4]), 10);
    if (filaN !== 0) continue;

    var fechaCell = datos[r][7];
    var ts = datos[r][0];
    var fechaIso;
    if (fechaCell instanceof Date) {
      fechaIso = fechaCell.toISOString();
    } else if (ts instanceof Date) {
      fechaIso = ts.toISOString();
    } else {
      fechaIso = new Date(Number(ts) || Date.now()).toISOString();
    }

    var det = String(datos[r][5] || "");
    var motivo = String(datos[r][6] || "");
    var texto = motivo ? det + " · Motivo: " + motivo : det;

    items.push({
      fecha: fechaIso,
      detalle: texto,
      motivo: motivo
    });
  }

  items.sort(function (a, b) {
    return new Date(a.fecha).getTime() - new Date(b.fecha).getTime();
  });

  return items;
}

function actualizarEstadoMatriculaEstudiante_(data) {
  var idEst = String(data.estudiante || "").trim();
  var grupoReq = String(data.grupo || "").trim();
  var nuevo = normalizarEstadoMatriculaInput_(data.estadoMatricula || data.estado || "");
  var motivo = String(data.motivo || "").trim();

  if (!idEst || !grupoReq) {
    throw new Error("Faltan estudiante o grupo.");
  }

  var hoja = SpreadsheetApp.getActive().getSheetByName("Estudiantes");
  var datos = hoja.getDataRange().getValues();
  var grupoBuscado = limpiarTexto(grupoReq);
  var filaHoja = null;
  var anteriorEtiqueta = "Activo";

  for (var i = 1; i < datos.length; i++) {
    var idF = String(datos[i][0] || "").trim();
    var grF = limpiarTexto(datos[i][2]);
    if (idF === idEst && grF === grupoBuscado) {
      filaHoja = i + 1;
      anteriorEtiqueta = etiquetaEstadoMatriculaDesdeCelda_(datos[i][3]);
      break;
    }
  }

  if (!filaHoja) {
    throw new Error("No se encontró el estudiante en ese grupo.");
  }

  if (anteriorEtiqueta === nuevo) {
    return { estadoMatricula: nuevo };
  }

  hoja.getRange(filaHoja, 4).setValue(nuevo);

  var detalle = "Estado matrícula: " + anteriorEtiqueta + " → " + nuevo;
  registrarAuditoria_(0, idEst, "", grupoReq, detalle, motivo);

  return { estadoMatricula: nuevo };
}

// ===============================
// ESTUDIANTES
// Hoja Estudiantes: A=id, B=nombre, C=grupo, D=estadoMatricula, E=documento, F=PIN, G=fotoUrl (opcional: URL o ruta, p. ej. WebP)
// La columna de foto también se detecta por la fila 1 si el encabezado contiene «foto», «fotoUrl», «imagen», etc.
// ===============================

/** Índice 0-based de la columna de foto (fila 1 = encabezados). Por defecto G (6). */
function indiceColumnaFotoEstudiantes_(filaEncabezado) {
  var def = 6;
  if (!filaEncabezado || !filaEncabezado.length) return def;
  var c;
  for (c = 0; c < filaEncabezado.length; c++) {
    var h = limpiarTexto(String(filaEncabezado[c] == null ? "" : filaEncabezado[c]));
    if (!h) continue;
    if (h === "fotourl" || h === "foto url" || h.indexOf("fotourl") >= 0) return c;
    if (h === "url foto" || h === "ruta foto" || h === "ruta imagen") return c;
  }
  for (c = 0; c < filaEncabezado.length; c++) {
    var h2 = limpiarTexto(String(filaEncabezado[c] == null ? "" : filaEncabezado[c]));
    if (h2 === "foto" || h2.indexOf("foto ") === 0 || h2 === "imagen") return c;
  }
  return def;
}

/** Normaliza texto de celda G (comillas, barra invertida). */
function normalizarFotoUrlCeldaGs_(raw) {
  var s = String(raw == null ? "" : raw).trim();
  if (!s) return "";
  var q0 = s.charAt(0);
  var q1 = s.charAt(s.length - 1);
  if ((q0 === '"' && q1 === '"') || (q0 === "'" && q1 === "'")) {
    s = s.substring(1, s.length - 1).trim();
  }
  s = s.replace(/\\/g, "/");
  return s;
}

/**
 * Si la celda tiene =HYPERLINK("url",...), getValues() solo devuelve el texto visible, no la URL.
 * Extrae el primer argumento (ruta o enlace) de la fórmula.
 */
function extraerUrlPrimeraHYPERLINKGs_(formula) {
  var f = String(formula || "").trim();
  if (!f || f.indexOf("HYPERLINK") === -1) return "";
  var m = f.match(/HYPERLINK\s*\(\s*"([^"]*)"/i);
  if (m) return m[1].replace(/""/g, '"');
  m = f.match(/HYPERLINK\s*\(\s*'([^']*)'/i);
  if (m) return m[1];
  return "";
}

/** Valor de celda de foto (índice de columna 0-based). */
function fotoEstudianteDesdeFilaIdx_(filaArr, colIdx) {
  if (!filaArr || colIdx < 0 || filaArr.length <= colIdx) return "";
  return normalizarFotoUrlCeldaGs_(filaArr[colIdx]);
}

/** Valor de celda o URL extraída de =HYPERLINK en esa columna. */
function fotoEstudianteDesdeFilaYFormulaIdx_(filaArr, formulaCelda, colIdx) {
  var desdeFormula = extraerUrlPrimeraHYPERLINKGs_(formulaCelda);
  if (desdeFormula) return normalizarFotoUrlCeldaGs_(desdeFormula);
  return fotoEstudianteDesdeFilaIdx_(filaArr, colIdx);
}

/**
 * Hoja "Grupos": columnas típicas A=ID_Grupo, B=Nombre_Grupo (fila 1 encabezados).
 * El desplegable usa los nombres de la columna B.
 */
function obtenerGruposDesdeHojaGrupos_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName("Grupos");
  if (!hoja) return [];

  var datos = hoja.getDataRange().getValues();
  if (!datos.length) return [];

  var start = 0;
  var h0 = String(datos[0][0] == null ? "" : datos[0][0]).toLowerCase();
  var h1 = String(datos[0][1] == null ? "" : datos[0][1]).toLowerCase();
  if (h0.indexOf("id") >= 0 && (h1.indexOf("nombre") >= 0 || h1.indexOf("grupo") >= 0)) {
    start = 1;
  }

  var seen = {};
  var out = [];

  for (var i = start; i < datos.length; i++) {
    var nombre = String(datos[i][1] == null ? "" : datos[i][1]).trim();
    if (!nombre) continue;
    var key = limpiarTexto(nombre);
    if (seen[key]) continue;
    seen[key] = true;
    out.push(nombre);
  }

  out.sort(function (a, b) {
    return limpiarTexto(a).localeCompare(limpiarTexto(b), "es", { sensitivity: "base", numeric: true });
  });

  return out;
}

/** Lista única de grupos según columna C de Estudiantes (respaldo si no hay hoja Grupos). */
function obtenerGruposUnicosDesdeEstudiantes_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName("Estudiantes");
  if (!hoja) return [];

  var datos = hoja.getDataRange().getValues();
  var seen = {};
  var out = [];

  for (var i = 1; i < datos.length; i++) {
    var g = String(datos[i][2] == null ? "" : datos[i][2]).trim();
    if (!g) continue;
    var key = limpiarTexto(g);
    if (seen[key]) continue;
    seen[key] = true;
    out.push(g);
  }

  out.sort(function (a, b) {
    return limpiarTexto(a).localeCompare(limpiarTexto(b), "es", { sensitivity: "base", numeric: true });
  });

  return out;
}

function obtenerGruposLista_() {
  var desdeCat = obtenerGruposDesdeHojaGrupos_();
  if (desdeCat.length) return desdeCat;
  return obtenerGruposUnicosDesdeEstudiantes_();
}

/**
 * Lee la hoja `MosaicosGrupo` (A: grupo, B: imagen URL o data URI, C opcional: actualizado_en).
 * No depende de Drive; el cliente recibe el texto y lo aplica directamente al `<img src>`.
 */
var MOSAICO_HOJA_NOMBRE_ = "MosaicosGrupo";

function obtenerMosaicoGrupo_(grupo) {
  var grupoNorm = normalizarGrupoCursoIEMFS_(grupo);
  var hoja = SpreadsheetApp.getActive().getSheetByName(MOSAICO_HOJA_NOMBRE_);
  if (!hoja) {
    return {
      ok: true,
      grupo: String(grupo || "").trim(),
      imagen: "",
      actualizado: "",
      mensaje:
        "Aún no existe la hoja «" +
        MOSAICO_HOJA_NOMBRE_ +
        "». Cree una con columnas A=grupo, B=imagen (URL o data URI) y C=actualizado_en (opcional)."
    };
  }
  var lr = hoja.getLastRow();
  if (lr < 2) {
    return {
      ok: true,
      grupo: String(grupo || "").trim(),
      imagen: "",
      actualizado: ""
    };
  }
  var lc = Math.max(2, hoja.getLastColumn());
  var datos = hoja.getRange(2, 1, lr - 1, lc).getValues();
  for (var i = 0; i < datos.length; i++) {
    var gFila = normalizarGrupoCursoIEMFS_(datos[i][0]);
    if (gFila !== grupoNorm) continue;
    var img = String(datos[i][1] == null ? "" : datos[i][1]).trim();
    var act = lc >= 3 ? datos[i][2] : "";
    var actStr = "";
    if (act instanceof Date && !isNaN(act.getTime())) {
      actStr = Utilities.formatDate(act, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (act != null) {
      actStr = String(act).trim();
    }
    return {
      ok: true,
      grupo: String(grupo || "").trim(),
      imagen: img,
      actualizado: actStr
    };
  }
  return {
    ok: true,
    grupo: String(grupo || "").trim(),
    imagen: "",
    actualizado: ""
  };
}

function obtenerEstudiantesPorGrupo(grupo) {
  const hoja = SpreadsheetApp.getActive().getSheetByName("Estudiantes");
  if (!hoja) return [];
  const datos = hoja.getDataRange().getValues();

  const grupoBuscado = normalizarGrupoCursoIEMFS_(grupo);
  const resultado = [];
  var colFoto = indiceColumnaFotoEstudiantes_(datos[0] || []);
  var colSheetFoto = colFoto + 1;
  var formulasFoto = [];
  if (datos.length > 1) {
    formulasFoto = hoja.getRange(2, colSheetFoto, datos.length - 1, 1).getFormulas();
  }

  for (let i = 1; i < datos.length; i++) {
    const grupoFila = normalizarGrupoCursoIEMFS_(datos[i][2]);
    if (grupoFila !== grupoBuscado) continue;

    var filaFormula = formulasFoto[i - 1];
    var formulaCelda = filaFormula && filaFormula[0] ? filaFormula[0] : "";

    resultado.push({
      id: String(datos[i][0]).trim(),
      nombre: datos[i][1],
      estadoMatricula: etiquetaEstadoMatriculaDesdeCelda_(datos[i][3]),
      grupo: String(datos[i][2] == null ? "" : datos[i][2]).trim(),
      fotoUrl: fotoEstudianteDesdeFilaYFormulaIdx_(datos[i], formulaCelda, colFoto)
    });
  }

  resultado.sort(function (a, b) {
    return String(a.nombre || "").localeCompare(String(b.nombre || ""), "es");
  });

  return resultado;
}

/**
 * Traslado de curso (automático): actualiza columna C (grupo) en Estudiantes y columna D en todas las filas de Notas del id.
 * Opcional data.grupoAnterior: si viene, debe coincidir con el grupo actual en hoja (evita traslados con pantalla desactualizada).
 * El grupo destino debe existir en la hoja Grupos (o en el listado de grupos del sistema).
 */
function migrarGrupoNotasEstudiante_(data) {
  var idEst = String(data.estudiante || data.estudianteId || "").trim();
  var grupoNuevoRaw = String(data.grupoNuevo || "").trim();
  var motivo = String(data.motivo || "").trim();
  var grupoAnteriorCliente = String(data.grupoAnterior || "").trim();

  if (!idEst || !grupoNuevoRaw) {
    throw new Error("Faltan datos del estudiante o el curso de destino.");
  }

  var gruposOk = obtenerGruposLista_();
  var normNuevo = limpiarTexto(grupoNuevoRaw);
  var grupoNuevo = grupoNuevoRaw;
  var permitido = gruposOk.length === 0;

  for (var gi = 0; gi < gruposOk.length; gi++) {
    if (limpiarTexto(gruposOk[gi]) === normNuevo) {
      permitido = true;
      grupoNuevo = String(gruposOk[gi]).trim();
      normNuevo = limpiarTexto(grupoNuevo);
      break;
    }
  }

  if (!permitido) {
    throw new Error(
      "El curso de destino no está en el listado institucional. Si es un grupo nuevo, regístrelo primero en la hoja Grupos."
    );
  }

  var hojaEst = SpreadsheetApp.getActive().getSheetByName("Estudiantes");
  if (!hojaEst) {
    throw new Error("No existe la hoja Estudiantes.");
  }

  var datosEst = hojaEst.getDataRange().getValues();
  var filasEstudiante = [];
  var grupoEnHoja = "";

  for (var i = 1; i < datosEst.length; i++) {
    var idF = String(datosEst[i][0] || "").trim();
    if (idF !== idEst) continue;
    filasEstudiante.push(i + 1);
    if (!grupoEnHoja) {
      grupoEnHoja = String(datosEst[i][2] == null ? "" : datosEst[i][2]).trim();
    }
  }

  if (filasEstudiante.length === 0) {
    throw new Error("No se encontró al estudiante en la base de datos.");
  }

  var grupoActualNorm = limpiarTexto(grupoEnHoja);
  if (grupoAnteriorCliente && limpiarTexto(grupoAnteriorCliente) !== grupoActualNorm) {
    throw new Error(
      "El curso del estudiante cambió en el sistema respecto a lo que muestra la pantalla. Vuelva atrás, pulse «Cargar grupo» otra vez y repita el traslado."
    );
  }

  if (grupoActualNorm === normNuevo) {
    throw new Error("El estudiante ya está en ese curso. Elija otro destino.");
  }

  for (var fi = 0; fi < filasEstudiante.length; fi++) {
    hojaEst.getRange(filasEstudiante[fi], 3).setValue(grupoNuevo);
  }

  var actualizadas = 0;

  function actualizarGrupoEnHojaNotas_(sh) {
    if (!sh) return;
    var datos = valoresNotasHojaDesdeFila2_(sh);
    var j;
    for (j = 0; j < datos.length; j++) {
      var idN = String(datos[j][1] || "").trim();
      if (idN !== idEst) continue;
      var grCelda = String(datos[j][3] == null ? "" : datos[j][3]).trim();
      if (limpiarTexto(grCelda) === normNuevo) continue;
      sh.getRange(j + 2, 4).setValue(grupoNuevo);
      actualizadas++;
    }
  }

  if (notasMultihojaActivas_()) {
    var pi;
    for (pi = 1; pi <= 3; pi++) {
      actualizarGrupoEnHojaNotas_(SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJAS_P_[pi - 1]));
    }
  }
  actualizarGrupoEnHojaNotas_(SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_));

  function actualizarGrupoEnHojaInas_(sh) {
    if (!sh) return;
    var dIna = valoresInasistenciasHojaDesdeFila2_(sh);
    var ki;
    for (ki = 0; ki < dIna.length; ki++) {
      if (String(dIna[ki][1] || "").trim() !== idEst) continue;
      sh.getRange(ki + 2, 3).setValue(grupoNuevo);
    }
  }

  if (inasistenciasMultihojaActivas_()) {
    var qi;
    for (qi = 1; qi <= 3; qi++) {
      actualizarGrupoEnHojaInas_(SpreadsheetApp.getActive().getSheetByName(INAS_HOJAS_P_[qi - 1]));
    }
  }
  actualizarGrupoEnHojaInas_(SpreadsheetApp.getActive().getSheetByName(INAS_HOJA_LEGACY_));

  var detalle =
    "Traslado de curso: «" +
    (grupoEnHoja || "—") +
    "» → «" +
    grupoNuevo +
    "». Calificaciones (filas Notas) actualizadas: " +
    actualizadas +
    ". Ficha del estudiante (curso) actualizada.";

  registrarAuditoria_(0, idEst, "", grupoNuevo, detalle, motivo);

  return { actualizadas: actualizadas, estudianteActualizado: true };
}

// ===============================
// INASISTENCIAS
// Hoja Inasistencias: A=ts, B=estudiante id, C=grupo, D=materia (texto), E=fecha (día), F=falta (1=faltó, 0=asistió)
// ===============================
function obtenerHojaInasistencias_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName("Inasistencias");
  if (!sh) {
    sh = ss.insertSheet("Inasistencias");
    sh.appendRow(["timestamp", "estudiante", "grupo", "materia", "fecha", "falta"]);
  }
  return sh;
}

function fechaDiaDesdeIso_(fechaStr) {
  var p = String(fechaStr || "").trim().split("-");
  if (p.length !== 3) {
    throw new Error("Fecha no válida (use AAAA-MM-DD).");
  }
  var y = parseInt(p[0], 10);
  var m = parseInt(p[1], 10) - 1;
  var d = parseInt(p[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) {
    throw new Error("Fecha no válida.");
  }
  return new Date(y, m, d, 12, 0, 0);
}

function mismoDiaInasistencia_(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return false;
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function celdaFaltaANumero_(v) {
  if (v === 1 || v === true) return 1;
  if (v === 0 || v === false) return 0;
  var s = String(v == null ? "" : v)
    .trim()
    .toLowerCase();
  if (s === "1" || s === "sí" || s === "si" || s === "falta") return 1;
  return 0;
}

/**
 * Fechas (ISO yyyy-MM-dd) con al menos un registro de lista para grupo + materia.
 * Orden: la más reciente primero.
 */
function listarFechasConListaGrupoMateria_(grupo, materia) {
  var g = String(grupo || "").trim();
  var m = String(materia || "").trim();
  if (!g || !m) {
    throw new Error("Indique grupo y materia.");
  }
  var grupoB = limpiarTexto(g);
  var materiaNorm = limpiarTexto(m);
  var tz = Session.getScriptTimeZone();
  var seen = {};

  function acumular_(datos, startIndex) {
    var i;
    for (i = startIndex; i < datos.length; i++) {
      var gF = limpiarTexto(datos[i][2]);
      var mF = limpiarTexto(datos[i][3]);
      if (gF !== grupoB || mF !== materiaNorm) continue;
      var fCell = datos[i][4];
      var fd = fechaCeldaInasistencia_(fCell);
      if (!fd) continue;
      var fk = Utilities.formatDate(fd, tz, "yyyy-MM-dd");
      seen[fk] = true;
    }
  }

  if (inasistenciasMultihojaActivas_()) {
    var p;
    for (p = 1; p <= 3; p++) {
      var sh = SpreadsheetApp.getActive().getSheetByName(INAS_HOJAS_P_[p - 1]);
      if (!sh) continue;
      acumular_(valoresInasistenciasHojaDesdeFila2_(sh), 0);
    }
  } else {
    var hLeg = SpreadsheetApp.getActive().getSheetByName(INAS_HOJA_LEGACY_);
    if (hLeg) {
      acumular_(hLeg.getDataRange().getValues(), 1);
    }
  }

  var out = Object.keys(seen);
  out.sort();
  out.reverse();
  return out;
}

function obtenerMapaInasistenciasDia_(grupo, materiaNorm, fechaObj) {
  var tz = Session.getScriptTimeZone();
  var fechaIso = Utilities.formatDate(fechaObj, tz, "yyyy-MM-dd");
  var hoja = hojaInasistenciasParaFechaIso_(fechaIso);
  var datos = inasistenciasMultihojaActivas_()
    ? valoresInasistenciasHojaDesdeFila2_(hoja)
    : hoja.getDataRange().getValues();
  var grupoB = limpiarTexto(grupo);
  var map = {};
  var i0 = inasistenciasMultihojaActivas_() ? 0 : 1;

  for (var i = i0; i < datos.length; i++) {
    var idF = String(datos[i][1] || "").trim();
    var gF = limpiarTexto(datos[i][2]);
    var mF = limpiarTexto(datos[i][3]);
    var fCell = datos[i][4];
    if (gF !== grupoB || mF !== materiaNorm || !idF) continue;

    var fd = fechaCeldaInasistencia_(fCell);
    if (!fd || !mismoDiaInasistencia_(fd, fechaObj)) continue;

    var filaReal = inasistenciasMultihojaActivas_() ? i + 2 : i + 1;
    map[idF] = {
      falta: celdaFaltaANumero_(datos[i][5]),
      fila: filaReal
    };
  }

  return { hoja: hoja, map: map };
}

function obtenerInasistenciasDia_(grupo, materia, fechaIso) {
  var g = String(grupo || "").trim();
  var m = String(materia || "").trim();
  if (!g || !m) {
    throw new Error("Indique grupo y materia.");
  }
  var fechaObj = fechaDiaDesdeIso_(fechaIso);
  var materiaNorm = limpiarTexto(m);
  var ctx = obtenerMapaInasistenciasDia_(g, materiaNorm, fechaObj);
  var estudiantes = obtenerEstudiantesPorGrupo(g);

  return estudiantes.map(function (e) {
    var id = String(e.id).trim();
    var rec = ctx.map[id];
    return {
      id: id,
      nombre: e.nombre,
      estadoMatricula: e.estadoMatricula,
      grupo: e.grupo,
      falta: rec ? rec.falta : 0,
      filaInasistencia: rec ? rec.fila : null
    };
  });
}

function guardarInasistenciasMasivo_(data) {
  var grupo = String(data.grupo || "").trim();
  var materia = String(data.materia || "").trim();
  var fechaIso = String(data.fecha || "").trim();
  var items = data.items;

  if (!grupo || !materia || !fechaIso) {
    throw new Error("Faltan grupo, materia o fecha.");
  }
  if (!Array.isArray(items) || !items.length) {
    throw new Error("No hay registros para guardar.");
  }

  var fechaObj = fechaDiaDesdeIso_(fechaIso);
  var materiaNorm = limpiarTexto(materia);
  var ctx = obtenerMapaInasistenciasDia_(grupo, materiaNorm, fechaObj);
  var hoja = ctx.hoja;
  var map = ctx.map;

  var actualizadas = 0;
  var creadas = 0;

  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var id = String(it.id || it.estudiante || "").trim();
    if (!id) continue;

    var falta = celdaFaltaANumero_(it.falta);
    var rec = map[id];

    if (rec) {
      hoja.getRange(rec.fila, 6).setValue(falta);
      actualizadas++;
    } else {
      hoja.appendRow([new Date().getTime(), id, grupo, materia, fechaObj, falta]);
      var newRow = hoja.getLastRow();
      map[id] = { falta: falta, fila: newRow };
      creadas++;
    }
  }

  return { actualizadas: actualizadas, creadas: creadas };
}

function listarInasistenciasEstudiante_(estudianteId, materia, limite, periodoId) {
  var idB = String(estudianteId || "").trim();
  var mB = limpiarTexto(materia || "");
  if (!idB) return [];

  var limPer = limitePeriodoNotasOpcional_(periodoId);

  var tz = Session.getScriptTimeZone();
  var out = [];

  function acumularDesde_(datos, startIndex) {
    var i;
    for (i = startIndex; i < datos.length; i++) {
      if (String(datos[i][1] || "").trim() !== idB) continue;
      if (limpiarTexto(datos[i][3]) !== mB) continue;

      var fCell = datos[i][4];
      if (limPer && !notaFechaDentroDePeriodo_(fCell, limPer)) continue;
      var fd = fechaCeldaInasistencia_(fCell);
      var fechaIso = fd
        ? Utilities.formatDate(fd, tz, "yyyy-MM-dd")
        : String(fCell || "").slice(0, 10);

      var filaH = startIndex === 0 ? i + 2 : i + 1;
      out.push({
        fecha: fechaIso,
        falta: celdaFaltaANumero_(datos[i][5]),
        fila: filaH
      });
    }
  }

  if (inasistenciasMultihojaActivas_()) {
    if (limPer) {
      var nPer = parseInt(String(periodoId || "1").trim(), 10);
      if (isNaN(nPer) || nPer < 1 || nPer > 3) nPer = 1;
      var sh = SpreadsheetApp.getActive().getSheetByName(INAS_HOJAS_P_[nPer - 1]);
      if (sh) {
        acumularDesde_(valoresInasistenciasHojaDesdeFila2_(sh), 0);
      }
    } else {
      var p;
      for (p = 1; p <= 3; p++) {
        var sh2 = SpreadsheetApp.getActive().getSheetByName(INAS_HOJAS_P_[p - 1]);
        if (sh2) {
          acumularDesde_(valoresInasistenciasHojaDesdeFila2_(sh2), 0);
        }
      }
    }
  } else {
    var hoja = obtenerHojaInasistencias_();
    acumularDesde_(hoja.getDataRange().getValues(), 1);
  }

  out.sort(function (a, b) {
    return String(b.fecha).localeCompare(String(a.fecha));
  });

  if (limite > 0 && out.length > limite) {
    out = out.slice(0, limite);
  }

  return out;
}

/** Matrícula cancelada / retirados: no acumulan control de faltas (institucional). Desertor e Inactivo sí. */
function estadoSinAcumuloFaltas_(etiquetaEstado) {
  var s = limpiarTexto(String(etiquetaEstado || ""));
  if (s.indexOf("cancel") >= 0) return true;
  if (s.indexOf("retir") >= 0) return true;
  return false;
}

/**
 * En el resumen por periodo, Desertor e Inactivo: cada fecha en que hubo lista (grupo+materia)
 * cuenta como falta aunque en el día no se haya marcado F (celda en blanco o asistió).
 * Activos: solo suman filas con falta explícita.
 */
function imputaFaltaPorCadaDiaConLista_(etiquetaEstado) {
  var s = limpiarTexto(String(etiquetaEstado || ""));
  if (s.indexOf("desert") >= 0) return true;
  if (s.indexOf("inactiv") >= 0) return true;
  return false;
}

function obtenerHojaPeriodosAcademicos_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName("PeriodosAcademicos");
  if (!sh) {
    sh = ss.insertSheet("PeriodosAcademicos");
    sh.appendRow(["periodo", "fecha_inicio", "fecha_fin", "etiqueta"]);
    var y = new Date().getFullYear();
    sh.appendRow([1, new Date(y, 1, 3), new Date(y, 4, 15), "Periodo 1 (ajuste fechas en esta hoja)"]);
    sh.appendRow([2, new Date(y, 4, 16), new Date(y, 7, 15), "Periodo 2"]);
    sh.appendRow([3, new Date(y, 7, 16), new Date(y, 10, 30), "Periodo 3"]);
    sh.getRange(2, 2, 3, 2).setNumberFormat("dd/mm/yyyy");
  }
  return sh;
}

function listarPeriodosAcademicos_() {
  var sh = obtenerHojaPeriodosAcademicos_();
  var datos = sh.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var out = [];

  for (var i = 1; i < datos.length; i++) {
    var pid = parseInt(String(datos[i][0]), 10);
    if (isNaN(pid)) continue;
    var di = datos[i][1];
    var df = datos[i][2];
    if (!(di instanceof Date)) di = new Date(di);
    if (!(df instanceof Date)) df = new Date(df);
    if (isNaN(di.getTime()) || isNaN(df.getTime())) continue;

    var etiqueta = String(datos[i][3] != null && datos[i][3] !== "" ? datos[i][3] : "Periodo " + pid);
    out.push({
      id: pid,
      desde: Utilities.formatDate(di, tz, "yyyy-MM-dd"),
      hasta: Utilities.formatDate(df, tz, "yyyy-MM-dd"),
      etiqueta: etiqueta
    });
  }

  out.sort(function (a, b) {
    return a.id - b.id;
  });
  return out;
}

function obtenerLimitePeriodo_(periodoNum) {
  var n = parseInt(String(periodoNum), 10);
  var lista = listarPeriodosAcademicos_();
  for (var i = 0; i < lista.length; i++) {
    if (lista[i].id === n) return lista[i];
  }
  return null;
}

/**
 * Fecha en columna Inasistencias (Date de hoja, ISO, dd/mm/aaaa, etc.).
 * Sin esto, texto tipo "15/05/2026" o celdas mal interpretadas quedan fuera del resumen (0 días con lista).
 */
function fechaCeldaInasistencia_(fCell) {
  if (fCell instanceof Date) {
    return isNaN(fCell.getTime()) ? null : fCell;
  }
  if (fCell === null || fCell === undefined || fCell === "") return null;

  var s = String(fCell).trim();
  if (!s) return null;

  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (iso) {
    var y = parseInt(iso[1], 10);
    var mo = parseInt(iso[2], 10) - 1;
    var d = parseInt(iso[3], 10);
    var dt = new Date(y, mo, d, 12, 0, 0);
    if (!isNaN(dt.getTime())) return dt;
  }

  var dm = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\s*$/.exec(s);
  if (!dm) {
    dm = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})\s*$/.exec(s);
  }
  if (dm) {
    var d2 = parseInt(dm[1], 10);
    var m2 = parseInt(dm[2], 10) - 1;
    var y2 = parseInt(dm[3], 10);
    if (y2 < 100) {
      y2 += y2 >= 70 ? 1900 : 2000;
    }
    var dt2 = new Date(y2, m2, d2, 12, 0, 0);
    if (!isNaN(dt2.getTime())) return dt2;
  }

  var d3 = new Date(s);
  if (!isNaN(d3.getTime())) return d3;
  return null;
}

/**
 * Columna H (Notas): Date de hoja, número serial de Sheets, ISO o texto (dd/mm/aaaa).
 * Si getValues() devuelve serial en vez de Date, el fallback «sin texto en payload» no debe usar `new Date()` (hoy).
 */
function fechaCeldaNotaColumnaH_(fCell) {
  if (fCell instanceof Date) {
    return isNaN(fCell.getTime()) ? null : fCell;
  }
  if (fCell === null || fCell === undefined || fCell === "") return null;
  if (typeof fCell === "number" && isFinite(fCell)) {
    var ent = Math.floor(Math.abs(fCell));
    if (ent > 20000 && ent < 120000) {
      var dS = new Date((ent - 25569) * 86400 * 1000);
      if (!isNaN(dS.getTime())) return dS;
    }
  }
  return fechaCeldaInasistencia_(fCell);
}

function fechaEnRangoOrd_(fd, desde, hasta) {
  var ordF = fd.getFullYear() * 10000 + (fd.getMonth() + 1) * 100 + fd.getDate();
  var ord0 = desde.getFullYear() * 10000 + (desde.getMonth() + 1) * 100 + desde.getDate();
  var ord1 = hasta.getFullYear() * 10000 + (hasta.getMonth() + 1) * 100 + hasta.getDate();
  var mn = Math.min(ord0, ord1);
  var mx = Math.max(ord0, ord1);
  return ordF >= mn && ordF <= mx;
}

function obtenerResumenInasistenciasPeriodo_(grupo, materia, periodoId) {
  var g = String(grupo || "").trim();
  var m = String(materia || "").trim();
  if (!g || !m) {
    throw new Error("Indique grupo y materia.");
  }

  var lim = obtenerLimitePeriodo_(periodoId);
  if (!lim) {
    throw new Error(
      "Periodo no encontrado. Revise la hoja PeriodosAcademicos (columna periodo: 1, 2 o 3)."
    );
  }

  var desde = fechaDiaDesdeIso_(lim.desde);
  var hasta = fechaDiaDesdeIso_(lim.hasta);
  var grupoB = limpiarTexto(g);
  var materiaNorm = limpiarTexto(m);
  var tz = Session.getScriptTimeZone();

  var datos;
  if (inasistenciasMultihojaActivas_()) {
    var nPer = parseInt(String(periodoId || "1").trim(), 10);
    if (isNaN(nPer) || nPer < 1 || nPer > 3) nPer = 1;
    var shR = SpreadsheetApp.getActive().getSheetByName(INAS_HOJAS_P_[nPer - 1]);
    datos = shR ? [["__h__"]].concat(valoresInasistenciasHojaDesdeFila2_(shR)) : [];
  } else {
    datos = obtenerHojaInasistencias_().getDataRange().getValues();
  }

  var fechasListaGrupo = {};
  var porEstudiante = {};

  for (var i = 1; i < datos.length; i++) {
    var idEst = String(datos[i][1] || "").trim();
    var gF = limpiarTexto(datos[i][2]);
    var mF = limpiarTexto(datos[i][3]);
    if (gF !== grupoB || mF !== materiaNorm || !idEst) continue;

    var fd = fechaCeldaInasistencia_(datos[i][4]);
    if (!fd) continue;
    if (!fechaEnRangoOrd_(fd, desde, hasta)) continue;

    var fk = Utilities.formatDate(fd, tz, "yyyy-MM-dd");
    fechasListaGrupo[fk] = true;

    if (!porEstudiante[idEst]) {
      porEstudiante[idEst] = { faltas: 0, fechas: {}, faltaDia: {} };
    }
    porEstudiante[idEst].fechas[fk] = true;
    if (celdaFaltaANumero_(datos[i][5]) === 1) {
      porEstudiante[idEst].faltas++;
      porEstudiante[idEst].faltaDia[fk] = true;
    }
  }

  var diasListaGrupo = Object.keys(fechasListaGrupo).length;
  var fechasColumnas = Object.keys(fechasListaGrupo).sort();

  var estudiantes = obtenerEstudiantesPorGrupo(g);

  var lista = estudiantes.map(function (e) {
    var id = String(e.id).trim();
    var sinAc = estadoSinAcumuloFaltas_(e.estadoMatricula);
    var st = porEstudiante[id];
    var faltas = st ? st.faltas : 0;
    var faltaDia = !sinAc && st && st.faltaDia ? st.faltaDia : {};

    if (
      !sinAc &&
      imputaFaltaPorCadaDiaConLista_(e.estadoMatricula) &&
      fechasColumnas.length
    ) {
      faltas = fechasColumnas.length;
      faltaDia = {};
      for (var ic = 0; ic < fechasColumnas.length; ic++) {
        faltaDia[fechasColumnas[ic]] = true;
      }
    }

    return {
      id: id,
      nombre: e.nombre,
      estadoMatricula: e.estadoMatricula,
      sinAcumuloFaltas: sinAc,
      totalFaltas: sinAc ? null : faltas,
      faltaDia: sinAc ? null : faltaDia
    };
  });

  return {
    periodo: lim,
    diasListaGrupoEnPeriodo: diasListaGrupo,
    fechasColumnas: fechasColumnas,
    estudiantes: lista
  };
}

function fechaNotaPermitidaEnHojaPeriodo_(fd, periodoSheetNum) {
  if (!(fd instanceof Date) || isNaN(fd.getTime())) return false;
  if (!notasMultihojaActivas_()) return true;
  var n = parseInt(String(periodoSheetNum || "0"), 10);
  if (isNaN(n) || n < 1 || n > 3) return true;
  var lim = obtenerLimitePeriodo_(n);
  if (!lim) return true;
  var desde = fechaDiaDesdeIso_(lim.desde);
  var hasta = fechaDiaDesdeIso_(lim.hasta);
  return fechaEnRangoOrd_(fd, desde, hasta);
}

function fechaColumnaHDesdePayloadActualizar_(s, fechaCeldaAnterior) {
  var t0 = String(s != null ? s : "").trim();
  if (!t0) {
    var kept = fechaCeldaNotaColumnaH_(fechaCeldaAnterior);
    if (kept && !isNaN(kept.getTime())) return kept;
    return new Date();
  }
  var ymd = /^(\d{4}-\d{2}-\d{2})/.exec(t0);
  if (ymd) {
    return fechaDiaDesdeIso_(ymd[1]);
  }
  var d = new Date(t0);
  if (isNaN(d.getTime())) return null;
  return d;
}

/** Lee fecha desde POST (alias por compatibilidad y proxies que pierdan un nombre de campo). */
function fechaTextoDesdePayloadGenerico_(data) {
  if (!data || typeof data !== "object") return "";
  var keys = ["fecha", "fechaCalificacion", "fechaStr", "fechaNota", "fecha_nota"];
  var i;
  for (i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (data[k] != null && String(data[k]).trim() !== "") {
      return String(data[k]).trim();
    }
  }
  return "";
}

// ===============================
// GUARDAR NOTA
// Columnas Notas: A=ts, B=estudiante, C=materia, D=grupo, E=tipo, F=título, G=nota, H=fecha
// ===============================
function guardarNota(notaObj) {
  const nota = parseFloat(notaObj.nota);
  if (isNaN(nota) || nota < 0 || nota > 5) {
    throw new Error("La nota debe estar entre 0 y 5.");
  }

  var fGuardado = new Date();
  var fechaTxt = fechaTextoDesdePayloadGenerico_(notaObj);
  if (!fechaTxt) {
    throw new Error(
      "Falta la fecha de la calificación (campo fecha, formato AAAA-MM-DD). Si ya eligió la fecha en pantalla, recargue la aplicación (Ctrl+F5) y vuelva a intentar."
    );
  }
  var colFecha = fechaColumnaHDesdePayloadActualizar_(fechaTxt, null);
  if (!colFecha) {
    throw new Error("Fecha no válida. Use AAAA-MM-DD.");
  }

  var hoja;
  var perN = 0;
  if (notasMultihojaActivas_()) {
    perN = periodoNumParaNuevaNota_(notaObj, fGuardado);
    if (!fechaNotaPermitidaEnHojaPeriodo_(colFecha, perN)) {
      throw new Error(
        "La fecha cae fuera del periodo académico seleccionado (periodo " +
          String(perN) +
          "). Ajuste la fecha o el periodo en la barra superior."
      );
    }
    hoja = obtenerHojaNotasPeriodoNum_(perN);
  } else {
    hoja = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
    if (!hoja) {
      throw new Error("No existe la hoja Notas.");
    }
  }

  var idN = String(notaObj.estudiante || "").trim();
  var grupoFila = String(obtenerGrupoEstudiantePorId_(idN) || "").trim();
  if (!grupoFila) {
    throw new Error(
      "El estudiante no figura en la hoja Estudiantes. Recargue el curso en pantalla si hubo cambios manuales."
    );
  }

  hoja.appendRow([
    new Date().getTime(),
    idN,
    limpiarTexto(notaObj.materia),
    grupoFila,
    String(notaObj.tipo).trim(),
    String(notaObj.titulo).trim(),
    nota,
    colFecha
  ]);

  invalidarCacheNotasGrupo_(grupoFila, String(notaObj.materia || ""));

  return true;
}

/**
 * Anexa un lote (N×8) al final de la hoja bajo bloqueo de documento.
 * Usa solo appendRow por fila (igual que guardarNota): evita fallos de setValues por
 * dimensiones, celdas combinadas o reglas de validación en el bloque destino.
 */
function escribirLoteNotasAlFinalHoja_(hoja, filas) {
  if (!hoja || !filas || !filas.length) {
    return;
  }
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(45000);
  } catch (eLock) {
    throw new Error(
      "La hoja está ocupada por otro guardado. Espere unos segundos e intente de nuevo."
    );
  }
  try {
    var n = filas.length;
    var r;
    for (r = 0; r < n; r++) {
      if (!filas[r] || filas[r].length !== 8) {
        throw new Error("Lote inconsistente: cada fila debe tener 8 columnas (indice " + String(r) + ").");
      }
    }
    var lr = hoja.getLastRow();
    if (lr < 1) {
      throw new Error("La hoja de notas no tiene encabezado valido.");
    }
    var a;
    for (a = 0; a < n; a++) {
      hoja.appendRow(filas[a]);
    }
  } finally {
    try {
      lock.releaseLock();
    } catch (eRel) {
      /* ignore */
    }
  }
}

/**
 * Registro masivo docente: una sola petición HTTP en lugar de N × guardar.
 * items: [{ estudiante, tipo, nota, titulo? }, ...]; titulo por ítem sustituye al título común en esa fila.
 * authP: permisos por cada curso real de la hoja Estudiantes (tras traslados o edición manual).
 */
function guardarNotasMasivo_(data, authP) {
  var grupo = String(data.grupo || "").trim();
  var materia = String(data.materia || "").trim();
  var tituloComun = String(data.titulo || "").trim();
  var items = data.items;
  if (!grupo || !materia) {
    throw new Error("Faltan grupo o materia.");
  }
  if (!Array.isArray(items) || !items.length) {
    throw new Error("No hay notas para guardar.");
  }
  if (items.length > 800) {
    throw new Error("Máximo 800 filas por lote (p. ej. modo definitiva con muchos estudiantes).");
  }

  var fGuardado = new Date();

  var fechaComun = fechaTextoDesdePayloadGenerico_(data);
  if (!fechaComun) {
    throw new Error(
      "Indique la fecha de la evaluación (campo fecha, AAAA-MM-DD). Recargue la página (Ctrl+F5) si el campo fecha no se envía."
    );
  }
  var colFecha = fechaColumnaHDesdePayloadActualizar_(fechaComun, null);
  if (!colFecha) {
    throw new Error("Fecha no válida. Use AAAA-MM-DD.");
  }

  var baseNota = {
    grupo: grupo,
    materia: materia,
    fecha: fechaComun,
    periodo: data.periodo
  };

  var hoja;
  var perN = 0;
  if (notasMultihojaActivas_()) {
    perN = periodoNumParaNuevaNota_(baseNota, fGuardado);
    if (!fechaNotaPermitidaEnHojaPeriodo_(colFecha, perN)) {
      throw new Error(
        "La fecha cae fuera del periodo académico seleccionado (periodo " +
          String(perN) +
          "). Ajuste la fecha o el periodo en la barra superior."
      );
    }
    hoja = obtenerHojaNotasPeriodoNum_(perN);
  } else {
    hoja = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
    if (!hoja) {
      throw new Error("No existe la hoja Notas.");
    }
  }

  var materiaNorm = limpiarTexto(materia);
  var filas = [];
  var t0 = new Date().getTime();
  var gruposRealesPorNorma = {};
  var i;
  for (i = 0; i < items.length; i++) {
    var it = items[i];
    var nota = parseFloat(it.nota);
    if (isNaN(nota) || nota < 0 || nota > 5) {
      throw new Error(
        "Nota inválida para el estudiante " + String(it.estudiante || "").trim() + " (use 0 a 5)."
      );
    }
    var idEst = String(it.estudiante || "").trim();
    if (!idEst) {
      throw new Error("Falta id de estudiante en un ítem del lote.");
    }
    var grupoReal = String(obtenerGrupoEstudiantePorId_(idEst) || "").trim();
    if (!grupoReal) {
      throw new Error(
        "Estudiante «" +
          idEst +
          "» no aparece en la hoja Estudiantes. Pulse «Cargar grupo» otra vez tras cambios en la hoja."
      );
    }
    var normG = normalizarGrupoCursoIEMFS_(grupoReal);
    if (!gruposRealesPorNorma[normG]) {
      gruposRealesPorNorma[normG] = grupoReal;
    }
    var tituloFila = String(it.titulo != null ? it.titulo : "").trim();
    if (!tituloFila) tituloFila = tituloComun;
    if (!tituloFila) {
      throw new Error(
        "Indique el título común de la evaluación o envíe «titulo» en cada ítem del lote."
      );
    }
    filas.push([
      t0 + i,
      idEst,
      materiaNorm,
      grupoReal,
      String(it.tipo || "Seguimiento").trim(),
      tituloFila,
      nota,
      colFecha
    ]);
  }

  if (authP) {
    var gk;
    for (gk in gruposRealesPorNorma) {
      if (!Object.prototype.hasOwnProperty.call(gruposRealesPorNorma, gk)) continue;
      var gDisplay = gruposRealesPorNorma[gk];
      var errPerm = asegurarGrupoMateriaDocente_(authP, gDisplay, materia);
      if (errPerm) {
        throw new Error(
          errPerm +
            " El lote incluye estudiantes del curso «" +
            gDisplay +
            "» (según la hoja Estudiantes). Si editó cursos a mano, cargue ese curso o revise DocenteAsignaciones."
        );
      }
    }
  }

  escribirLoteNotasAlFinalHoja_(hoja, filas);

  for (gk in gruposRealesPorNorma) {
    if (!Object.prototype.hasOwnProperty.call(gruposRealesPorNorma, gk)) continue;
    invalidarCacheNotasGrupo_(gruposRealesPorNorma[gk], materia);
  }
  invalidarCacheNotasGrupo_(grupo, materia);
  return { guardadas: filas.length };
}

// ===============================
// LISTAR NOTAS (incluye id = fila en hoja)
// ===============================
/** @param periodoId {string} id numérico según hoja PeriodosAcademicos; vacío = todas las notas. */
function limitePeriodoNotasOpcional_(periodoId) {
  var s = String(periodoId == null ? "" : periodoId).trim();
  if (!s) return null;
  return obtenerLimitePeriodo_(s);
}

function notaFechaDentroDePeriodo_(fechaVal, lim) {
  if (!lim) return true;
  var fd = fechaCeldaNotaColumnaH_(fechaVal) || fechaCeldaInasistencia_(fechaVal);
  if (!fd) return false;
  try {
    var desde = fechaDiaDesdeIso_(lim.desde);
    var hasta = fechaDiaDesdeIso_(lim.hasta);
    return fechaEnRangoOrd_(fd, desde, hasta);
  } catch (e) {
    return true;
  }
}

function obtenerDetalleEstudiante(estudianteId, materia, periodoId) {
  const hoja = obtenerHojaNotasParaLecturaPorPeriodoId_(periodoId);
  if (!hoja) return [];
  const datos = valoresNotasHojaDesdeFila2_(hoja);

  const idBuscado = String(estudianteId).trim();
  const materiaBuscada = limpiarTexto(materia);
  var lim = limitePeriodoFechaLecturaNotas_(periodoId);
  var perNum = parseInt(String(periodoId || "1").trim(), 10);
  if (isNaN(perNum) || perNum < 1 || perNum > 3) perNum = 1;

  const notas = [];

  for (let i = 0; i < datos.length; i++) {
    const idFila = String(datos[i][1]).trim();
    const materiaFila = limpiarTexto(datos[i][2]);

    if (idFila === idBuscado && materiaFila === materiaBuscada) {
      const fechaVal = datos[i][7];
      if (!notaFechaDentroDePeriodo_(fechaVal, lim)) continue;
      const filaHoja = i + 2;
      var idNota = notasMultihojaActivas_() ? notaRefAuditoria_(perNum, filaHoja) : String(filaHoja);
      var fNorm = fechaCeldaNotaColumnaH_(fechaVal) || fechaCeldaInasistencia_(fechaVal);
      notas.push({
        id: idNota,
        notaId: idNota,
        fila: filaHoja,
        tipo: datos[i][4],
        titulo: datos[i][5],
        nota: datos[i][6],
        fecha: fNorm instanceof Date && !isNaN(fNorm.getTime()) ? fNorm.toISOString() : fechaVal
      });
    }
  }

  return notas;
}

/**
 * La hoja Notas usa B = id de estudiante (columna A de Estudiantes). Si alguien editó
 * manualmente y puso el nombre en B, el visor del docente no enlazaba filas (mapa por id).
 * Resolvemos por id exacto o por nombre normalizado (primer coincidencia en el curso).
 */
function construirResolucionEstudianteNotasGrupo_(grupo) {
  var alumnos = obtenerEstudiantesPorGrupo(grupo);
  var porId = {};
  var porNombre = {};
  var i;
  var id;
  var nom;
  for (i = 0; i < alumnos.length; i++) {
    id = String(alumnos[i].id || "").trim();
    if (!id) continue;
    porId[id] = id;
    nom = limpiarTexto(String(alumnos[i].nombre || ""));
    if (nom && porNombre[nom] === undefined) {
      porNombre[nom] = id;
    }
  }
  return { porId: porId, porNombre: porNombre };
}

function resolverEstudianteIdParaNotaGrupo_(rawCelda, resol) {
  if (!resol) return String(rawCelda == null ? "" : rawCelda).trim();
  var s = String(rawCelda == null ? "" : rawCelda).trim();
  if (!s) return "";
  if (resol.porId[s]) return s;
  var kn = limpiarTexto(s);
  if (kn && resol.porNombre[kn] !== undefined) return resol.porNombre[kn];
  return s;
}

function obtenerNotasGrupoMateria_(grupo, materia, periodoId) {
  const grupoBuscadoNorm = normalizarGrupoCursoIEMFS_(grupo);
  const materiaBuscada = limpiarTexto(materia);
  var perKey = notasMultihojaActivas_() ? String(periodoId || "1").trim() : "";
  var ck = cacheNotasGrupoKey_(grupo, materia, perKey);
  var cached = cacheGetJson_(ck);
  if (cached) return cached;

  var resolEst = construirResolucionEstudianteNotasGrupo_(grupo);

  var hoja = notasMultihojaActivas_()
    ? obtenerHojaNotasParaLecturaPorPeriodoId_(periodoId)
    : SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
  if (!hoja) return [];

  var datos = notasMultihojaActivas_()
    ? valoresNotasHojaDesdeFila2_(hoja)
    : hoja.getDataRange().getValues();

  const out = [];
  var i0 = notasMultihojaActivas_() ? 0 : 1;
  var perNumG = parseInt(String(periodoId || "1").trim(), 10);
  if (isNaN(perNumG) || perNumG < 1 || perNumG > 3) perNumG = 1;
  var i;
  for (i = i0; i < datos.length; i++) {
    const grupoFilaNorm = normalizarGrupoCursoIEMFS_(datos[i][3]);
    const materiaFila = limpiarTexto(datos[i][2]);
    if (grupoFilaNorm !== grupoBuscadoNorm || materiaFila !== materiaBuscada) continue;

    const fechaVal = datos[i][7];
    var fNormG = fechaCeldaNotaColumnaH_(fechaVal) || fechaCeldaInasistencia_(fechaVal);
    var filaHojaG = notasMultihojaActivas_() ? i + 2 : i + 1;
    var idNotaG = notasMultihojaActivas_() ? notaRefAuditoria_(perNumG, filaHojaG) : String(filaHojaG);
    out.push({
      estudiante: resolverEstudianteIdParaNotaGrupo_(datos[i][1], resolEst),
      notaId: idNotaG,
      tipo: String(datos[i][4] || "").trim(),
      titulo: String(datos[i][5] || "").trim(),
      nota: datos[i][6],
      fecha: fNormG instanceof Date && !isNaN(fNormG.getTime()) ? fNormG.toISOString() : fechaVal
    });
  }

  cachePutJson_(ck, out, CACHE_NOTAS_GRUPO_SEC_);
  return out;
}

// ===============================
// RESUMEN
// ===============================
function construirItemsBoletinPorMateriasOrdenadas_(porMatNorm, etiquetaPorNorm, ordenDisplay) {
  var items = [];
  var seen = {};
  var i;
  var m;
  var kn;
  for (i = 0; i < ordenDisplay.length; i++) {
    m = ordenDisplay[i];
    kn = limpiarTexto(m);
    items.push(calcularResumenBoletinMetaDesdeFilas_(m, porMatNorm[kn] || []));
    seen[kn] = true;
  }
  var extras = Object.keys(porMatNorm).filter(function (k) {
    return !seen[k];
  });
  extras.sort(function (a, b) {
    return a.localeCompare(b, "es");
  });
  for (i = 0; i < extras.length; i++) {
    kn = extras[i];
    m = etiquetaPorNorm[kn] || kn;
    items.push(calcularResumenBoletinMetaDesdeFilas_(m, porMatNorm[kn]));
  }
  return items;
}

function construirMateriasAgregadoOrdenadas_(porMatNorm, etiquetaPorNorm, ordenDisplay) {
  var materias = [];
  var seen = {};
  var i;
  var m;
  var kn;
  for (i = 0; i < ordenDisplay.length; i++) {
    m = ordenDisplay[i];
    kn = limpiarTexto(m);
    materias.push({ materia: m, notas: porMatNorm[kn] || [] });
    seen[kn] = true;
  }
  var extras = Object.keys(porMatNorm).filter(function (k) {
    return !seen[k];
  });
  extras.sort(function (a, b) {
    return a.localeCompare(b, "es");
  });
  for (i = 0; i < extras.length; i++) {
    kn = extras[i];
    m = etiquetaPorNorm[kn] || kn;
    materias.push({ materia: m, notas: porMatNorm[kn] });
  }
  return materias;
}

/**
 * Boletín: una sola lectura de la hoja Notas (evita N lecturas completas).
 * Materias: DocenteAsignaciones del curso del estudiante; si no hay, las que figuren en Notas.
 */
function boletinResumenEstudianteUnaPasada_(estudianteId, periodoId, grupoEstudiante) {
  var idB = String(estudianteId).trim();
  var lim = limitePeriodoFechaLecturaNotas_(periodoId);
  var hoja = obtenerHojaNotasParaLecturaPorPeriodoId_(periodoId);
  if (!hoja) {
    return construirItemsBoletinPorMateriasOrdenadas_({}, {}, listarMateriasParaEstudiante_(String(grupoEstudiante || "").trim(), idB));
  }
  var datos = valoresNotasHojaDesdeFila2_(hoja);
  var porMatNorm = {};
  var etiquetaPorNorm = {};
  var i;
  var k;
  var rawM;
  for (i = 0; i < datos.length; i++) {
    if (String(datos[i][1]).trim() !== idB) continue;
    if (!notaFechaDentroDePeriodo_(datos[i][7], lim)) continue;
    rawM = String(datos[i][2] == null ? "" : datos[i][2]).trim();
    k = limpiarTexto(rawM);
    if (!k) continue;
    if (!porMatNorm[k]) porMatNorm[k] = [];
    if (!etiquetaPorNorm[k]) etiquetaPorNorm[k] = rawM;
    porMatNorm[k].push({
      tipo: String(datos[i][4] || "").trim(),
      nota: datos[i][6]
    });
  }
  var orden = listarMateriasParaEstudiante_(String(grupoEstudiante || "").trim(), idB);
  return construirItemsBoletinPorMateriasOrdenadas_(porMatNorm, etiquetaPorNorm, orden);
}

/**
 * Notas del estudiante agrupadas por materia (una lectura de hoja).
 */
function notasEstudianteAgregadoPorMateria_(estudianteId, periodoId, grupoEstudiante) {
  var idB = String(estudianteId).trim();
  var lim = limitePeriodoFechaLecturaNotas_(periodoId);
  var hoja = obtenerHojaNotasParaLecturaPorPeriodoId_(periodoId);
  if (!hoja) {
    return { materias: [] };
  }
  var datos = valoresNotasHojaDesdeFila2_(hoja);
  var porMatNorm = {};
  var etiquetaPorNorm = {};
  var i;
  var k;
  var rawM;
  for (i = 0; i < datos.length; i++) {
    if (String(datos[i][1]).trim() !== idB) continue;
    if (!notaFechaDentroDePeriodo_(datos[i][7], lim)) continue;
    rawM = String(datos[i][2] == null ? "" : datos[i][2]).trim();
    k = limpiarTexto(rawM);
    if (!k) continue;
    if (!porMatNorm[k]) porMatNorm[k] = [];
    if (!etiquetaPorNorm[k]) etiquetaPorNorm[k] = rawM;
    porMatNorm[k].push({
      tipo: String(datos[i][4] || "").trim(),
      titulo: String(datos[i][5] || "").trim(),
      nota: datos[i][6],
      fecha: datos[i][7]
    });
  }
  var orden = listarMateriasParaEstudiante_(String(grupoEstudiante || "").trim(), idB);
  var materias = construirMateriasAgregadoOrdenadas_(porMatNorm, etiquetaPorNorm, orden);
  return { materias: materias };
}

/**
 * @param filas {Array<{ tipo: string, nota: * }>}
 */
function calcularResumenBoletinMetaDesdeFilas_(materia, filas) {
  const seguimiento = [];
  let actitudinal = 0;
  let prueba = 0;
  let tieneActitudinal = false;
  let tienePrueba = false;

  for (let i = 0; i < filas.length; i++) {
    const tipo = String(filas[i].tipo || "").trim();
    const nota = parseFloat(filas[i].nota);

    if (tipo === "Seguimiento" && !isNaN(nota)) seguimiento.push(nota);
    if (tipo === "Actitudinal") {
      tieneActitudinal = true;
      actitudinal = isNaN(nota) ? 0 : nota;
    }
    if (tipo === "Prueba") {
      tienePrueba = true;
      prueba = isNaN(nota) ? 0 : nota;
    }
  }

  const promSeg =
    seguimiento.length > 0
      ? seguimiento.reduce(function (a, b) {
          return a + b;
        }, 0) / seguimiento.length
      : 0;

  const final = promSeg * 0.7 + actitudinal * 0.1 + prueba * 0.2;

  const incompleto = seguimiento.length === 0 || !tieneActitudinal || !tienePrueba;

  const avisos = [];
  if (seguimiento.length === 0) avisos.push("Falta seguimiento");
  if (!tieneActitudinal) avisos.push("Falta Actitudinal");
  if (!tienePrueba) avisos.push("Falta Prueba");

  return {
    materia: materia,
    definitiva: final.toFixed(2),
    seguimiento: promSeg.toFixed(2),
    incompleto: incompleto,
    avisos: avisos
  };
}

function calcularResumen(estudianteId, materia, periodoId) {
  const hoja = obtenerHojaNotasParaLecturaPorPeriodoId_(periodoId);
  if (!hoja) {
    return {
      seguimiento: "0.00",
      actitudinal: 0,
      prueba: 0,
      final: "0.00"
    };
  }
  const datos = valoresNotasHojaDesdeFila2_(hoja);

  const idBuscado = String(estudianteId).trim();
  const materiaBuscada = limpiarTexto(materia);
  var lim = limitePeriodoFechaLecturaNotas_(periodoId);

  const seguimiento = [];
  let actitudinal = 0;
  let prueba = 0;

  for (let i = 0; i < datos.length; i++) {
    const idFila = String(datos[i][1]).trim();
    const materiaFila = limpiarTexto(datos[i][2]);

    if (idFila === idBuscado && materiaFila === materiaBuscada) {
      if (!notaFechaDentroDePeriodo_(datos[i][7], lim)) continue;
      const tipo = String(datos[i][4]).trim();
      const nota = parseFloat(datos[i][6]);

      if (tipo === "Seguimiento") seguimiento.push(nota);
      if (tipo === "Actitudinal") actitudinal = nota;
      if (tipo === "Prueba") prueba = nota;
    }
  }

  const promSeg = seguimiento.length > 0
    ? seguimiento.reduce(function (a, b) {
      return a + b;
    }, 0) / seguimiento.length
    : 0;

  const final = (promSeg * 0.7) + (actitudinal * 0.1) + (prueba * 0.2);

  return {
    seguimiento: promSeg.toFixed(2),
    actitudinal: isNaN(actitudinal) ? 0 : actitudinal,
    prueba: isNaN(prueba) ? 0 : prueba,
    final: final.toFixed(2)
  };
}

// ===============================
// BUSCAR FILA POR CRITERIOS (sin id)
// ===============================
function buscarFilaNotaPorAntes_(datos, estudianteId, materiaNorm, grupo, antes) {
  const idB = String(estudianteId).trim();
  const grupoBNorm = normalizarGrupoCursoIEMFS_(grupo);
  const titAnt = limpiarTexto(antes.titulo || "");
  const tipoAnt = String(antes.tipo || "").trim();
  const notaAnt = antes.nota;
  const reqFecha = antes.fecha !== undefined && antes.fecha !== null && String(antes.fecha) !== "";
  const ymdAnt = reqFecha ? ymdCalendarioDesdeValorNota_(antes.fecha) : "";

  for (let i = 1; i < datos.length; i++) {
    const idF = String(datos[i][1]).trim();
    const matF = limpiarTexto(datos[i][2]);
    const grF = normalizarGrupoCursoIEMFS_(datos[i][3]);

    if (idF !== idB || matF !== materiaNorm || grF !== grupoBNorm) continue;

    const tipoF = String(datos[i][4] == null ? "" : datos[i][4]).trim();
    const titF = limpiarTexto(datos[i][5] == null ? "" : datos[i][5]);
    const notaF = datos[i][6];

    if (titF !== titAnt) continue;
    if (tipoF !== tipoAnt) continue;
    if (!notasIguales_(notaF, notaAnt)) continue;

    if (reqFecha) {
      var ymdF = ymdCalendarioDesdeValorNota_(datos[i][7]);
      if (!ymdAnt || !ymdF || ymdAnt !== ymdF) continue;
    }

    return i + 1;
  }

  return null;
}

function buscarFilaNotaPorAntesMultihoja_(estudiante, materiaNorm, grupo, antes) {
  if (!notasMultihojaActivas_()) {
    var h0 = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
    if (!h0) return null;
    var d0 = h0.getDataRange().getValues();
    var fh0 = buscarFilaNotaPorAntes_(d0, estudiante, materiaNorm, grupo, antes);
    if (fh0) return { hoja: h0, fila: fh0, periodo: 0 };
    return null;
  }
  var p;
  for (p = 1; p <= 3; p++) {
    var sh = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJAS_P_[p - 1]);
    if (!sh) continue;
    var raw = valoresNotasHojaDesdeFila2_(sh);
    var datosPrep = [["__hdr__"]].concat(raw);
    var fh = buscarFilaNotaPorAntes_(datosPrep, estudiante, materiaNorm, grupo, antes);
    if (fh) return { hoja: sh, fila: fh, periodo: p };
  }
  var leg = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
  if (leg) {
    var d2 = leg.getDataRange().getValues();
    var fh2 = buscarFilaNotaPorAntes_(d2, estudiante, materiaNorm, grupo, antes);
    if (fh2) return { hoja: leg, fila: fh2, periodo: 0 };
  }
  return null;
}

// ===============================
// ACTUALIZAR NOTA (frontend: actualizar)
// ===============================
function actualizarNota_(data) {
  const estudiante = String(data.estudiante).trim();
  const materiaNorm = limpiarTexto(data.materia);
  const grupo = String(data.grupo || "").trim();

  const nuevaNota = parseFloat(data.nota);
  if (isNaN(nuevaNota) || nuevaNota < 0 || nuevaNota > 5) {
    throw new Error("La nota debe estar entre 0 y 5.");
  }

  const nuevoTitulo = String(data.titulo || "").trim();
  const nuevoTipo = String(data.tipo || "").trim();
  const motivo = String(data.motivo || "").trim();

  if (!nuevoTitulo) {
    throw new Error("El título no puede estar vacío.");
  }

  var hoja;
  var filaHoja = null;
  var periodoAud = 0;

  if (data.notaId !== undefined && data.notaId !== null && String(data.notaId).trim() !== "") {
    var ref = parseNotaIdRef_(String(data.notaId));
    if (!ref) {
      throw new Error("notaId no válido.");
    }
    hoja = hojaNotaDesdeRef_(ref);
    if (!hoja) {
      throw new Error("Hoja de notas no disponible.");
    }
    var lr = hoja.getLastRow();
    if (ref.fila < 2 || ref.fila > lr) {
      throw new Error("notaId no válido (fila fuera de rango).");
    }
    var idF = String(hoja.getRange(ref.fila, 2).getValue() || "").trim();
    var matF = limpiarTexto(hoja.getRange(ref.fila, 3).getValue());
    if (idF !== estudiante || matF !== materiaNorm) {
      throw new Error("notaId no coincide con estudiante/materia.");
    }
    filaHoja = ref.fila;
    periodoAud = ref.soloLegacy || !notasMultihojaActivas_() ? 0 : ref.periodo;
  } else if (data.antes && typeof data.antes === "object") {
    var ubic = buscarFilaNotaPorAntesMultihoja_(estudiante, materiaNorm, grupo, data.antes);
    if (!ubic) {
      throw new Error("No se encontró la fila con los datos «antes». Comprueba título, tipo, nota, grupo y fecha.");
    }
    hoja = ubic.hoja;
    filaHoja = ubic.fila;
    periodoAud = ubic.periodo;
  } else {
    throw new Error("Falta notaId o el objeto antes para ubicar la nota.");
  }

  const row = hoja.getRange(filaHoja, 1, 1, 8).getValues()[0];
  const tipoOld = String(row[4] == null ? "" : row[4]).trim();
  const titOld = String(row[5] == null ? "" : row[5]).trim();
  const notaOld = row[6];
  const matOriginal = String(row[2] == null ? "" : row[2]);
  const fechaOldCell = row[7];

  var fechaTxtPayload = fechaTextoDesdePayloadGenerico_(data);
  if (!fechaTxtPayload) {
    throw new Error(
      "El servidor no recibió la fecha de la nota (campos fecha / fechaNota). " +
        "Actualice con Ctrl+F5, confirme que guarda desde http://127.0.0.1 (no abrir index.html como archivo) " +
        "y vuelva a publicar Code.gs si el mensaje persiste."
    );
  }
  var fechaNueva = fechaColumnaHDesdePayloadActualizar_(fechaTxtPayload, fechaOldCell);
  if (!fechaNueva) {
    throw new Error("Fecha no válida. Use el formato AAAA-MM-DD.");
  }
  if (!fechaNotaPermitidaEnHojaPeriodo_(fechaNueva, periodoAud)) {
    throw new Error(
      "La fecha debe caer dentro del periodo académico de esta nota (periodo " +
        String(periodoAud || "—") +
        "). Si la evaluación pertenece a otro periodo, registre una nota nueva en esa franja."
    );
  }

  var tz = Session.getScriptTimeZone();
  function ymd_(d) {
    return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  }
  var dOld =
    fechaOldCell instanceof Date && !isNaN(fechaOldCell.getTime())
      ? fechaOldCell
      : fechaCeldaInasistencia_(fechaOldCell);
  var labOld = dOld && !isNaN(dOld.getTime()) ? ymd_(dOld) : String(fechaOldCell || "—");
  var labNew = ymd_(fechaNueva);

  const partes = [];
  if (!notasIguales_(notaOld, nuevaNota)) {
    partes.push("Nota: " + notaOld + " → " + nuevaNota);
  }
  if (tipoOld !== nuevoTipo) {
    partes.push("Tipo: " + tipoOld + " → " + nuevoTipo);
  }
  if (titOld !== nuevoTitulo) {
    partes.push("Título: «" + titOld + "» → «" + nuevoTitulo + "»");
  }
  if (labOld !== labNew) {
    partes.push("Fecha: " + labOld + " → " + labNew);
  }
  const detalle = partes.length ? partes.join(" · ") : "Sin cambios de contenido.";

  hoja.getRange(filaHoja, 5).setValue(nuevoTipo);
  hoja.getRange(filaHoja, 6).setValue(nuevoTitulo);
  hoja.getRange(filaHoja, 7).setValue(nuevaNota);
  hoja.getRange(filaHoja, 8).setValue(fechaNueva);

  var refAud =
    notasMultihojaActivas_() && periodoAud >= 1 ? notaRefAuditoria_(periodoAud, filaHoja) : String(filaHoja);
  registrarAuditoria_(refAud, estudiante, matOriginal, grupo, detalle, motivo);

  invalidarCacheNotasGrupo_(grupo, String(data.materia || ""));

  return true;
}

// ===============================
// MEJORAR NOTA (compatibilidad: accion mejorar)
// ===============================
function mejorarNota(estudianteId, materia, titulo, nuevaNota, motivo) {
  const idBuscado = String(estudianteId).trim();
  const materiaBuscada = limpiarTexto(materia);
  const tituloBuscado = limpiarTexto(titulo);

  let hoja = null;
  let filaHoja = null;
  let notaAnterior = null;
  let grupo = "";
  let matOriginal = "";
  let periodoAud = 0;

  function buscarEnDatos_(datos, filaBase, sh, perN) {
    var i;
    for (i = 0; i < datos.length; i++) {
      const idFila = String(datos[i][1]).trim();
      const materiaFila = limpiarTexto(datos[i][2]);
      const tituloFila = limpiarTexto(datos[i][5]);
      if (idFila !== idBuscado || materiaFila !== materiaBuscada || tituloFila !== tituloBuscado) continue;
      notaAnterior = datos[i][6];
      filaHoja = filaBase + i;
      grupo = String(datos[i][3] == null ? "" : datos[i][3]).trim();
      matOriginal = String(datos[i][2] == null ? "" : datos[i][2]);
      hoja = sh;
      periodoAud = perN;
      return true;
    }
    return false;
  }

  if (notasMultihojaActivas_()) {
    var p;
    for (p = 1; p <= 3; p++) {
      var sh = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJAS_P_[p - 1]);
      if (!sh) continue;
      var raw = valoresNotasHojaDesdeFila2_(sh);
      if (buscarEnDatos_(raw, 2, sh, p)) break;
    }
  }
  if (!filaHoja) {
    var leg = SpreadsheetApp.getActive().getSheetByName(NOTAS_HOJA_LEGACY_);
    if (leg) {
      var datos = leg.getDataRange().getValues();
      var j;
      for (j = 1; j < datos.length; j++) {
        const idFila = String(datos[j][1]).trim();
        const materiaFila = limpiarTexto(datos[j][2]);
        const tituloFila = limpiarTexto(datos[j][5]);
        if (idFila !== idBuscado || materiaFila !== materiaBuscada || tituloFila !== tituloBuscado) continue;
        notaAnterior = datos[j][6];
        filaHoja = j + 1;
        grupo = String(datos[j][3] == null ? "" : datos[j][3]).trim();
        matOriginal = String(datos[j][2] == null ? "" : datos[j][2]);
        hoja = leg;
        periodoAud = 0;
        break;
      }
    }
  }

  if (!filaHoja || !hoja) {
    throw new Error("No se encontró la nota para mejorar.");
  }

  const nn = parseFloat(nuevaNota);
  if (isNaN(nn) || nn < 0 || nn > 5) {
    throw new Error("La nota debe estar entre 0 y 5.");
  }

  hoja.getRange(filaHoja, 7).setValue(nn);
  hoja.getRange(filaHoja, 8).setValue(new Date());

  const detalle = "Nota: " + notaAnterior + " → " + nn + " (mejorar)";
  var refAud =
    notasMultihojaActivas_() && periodoAud >= 1 ? notaRefAuditoria_(periodoAud, filaHoja) : String(filaHoja);
  registrarAuditoria_(refAud, idBuscado, matOriginal, grupo, detalle, String(motivo || "").trim());

  invalidarCacheNotasGrupo_(grupo, materia);

  return true;
}
