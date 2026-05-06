/**
 * Recordatorio: el navegador carga `app.js` (bundle), no este archivo suelto.
 * Tras cambiar la URL de la API u otras constantes usadas en el cliente, ejecute en la raíz del proyecto:
 *   npm run build
 * y vuelva a servir / subir el `app.js` generado; si no, seguirá usando valores viejos empaquetados.
 */
/** URL del despliegue web de tu proyecto Apps Script (cámbiala si generas uno nuevo). */
export const API =
  "https://script.google.com/macros/s/AKfycbzt2kH5PkDUyypHlZuhFogqO97wvkuIGmEYRNb2N9h5Kg_3v2o4cLRcHiu1CqRcHaRT1w/exec";

/** Igual que NOTAS_IEMFS_BUILD_ en Code.gs; el ping del servidor debe traer este valor. */
export const NOTAS_IEMFS_BUILD_ESPERADO = "mosaico-grupo-20260506a";

export const STORAGE = {
  grupo: "notas_iemfs_grupo",
  materia: "notas_iemfs_materia",
  modo: "notas_iemfs_modo",
  theme: "notas_iemfs_theme",
  authSession: "notas_iemfs_auth_sess",
  authPersist: "notas_iemfs_auth_persist",
  docenteId: "notas_iemfs_docente_id",
  docentePersist: "notas_iemfs_docente_persist",
  docenteNombre: "notas_iemfs_docente_nombre",
  docenteNombrePersist: "notas_iemfs_docente_nombre_persist",
  authTipo: "notas_iemfs_auth_tipo",
  estudianteDoc: "notas_iemfs_est_doc",
  estudianteDocPersist: "notas_iemfs_est_doc_p",
  estudiantePin: "notas_iemfs_est_pin",
  estudiantePinPersist: "notas_iemfs_est_pin_p",
  /** Sesión estudiante: última materia elegida para expediente (no forzar Matemáticas al entrar). */
  estudianteMateriaExp: "notas_iemfs_est_mat_exp",
  /** Periodo académico elegido para notas del portal estudiante (calificaciones / boletín / expediente). */
  estudiantePeriodoCal: "notas_iemfs_est_periodo_cal",
  /** Periodo para la tabla de calificaciones del grupo (docente / coordinación). */
  docentePeriodoCal: "notas_iemfs_doc_periodo_cal"
};

/**
 * Respaldo si la API no devuelve grupos (sin red, error de despliegue).
 * Debe estar vacío: los cursos salen solo de la hoja Grupos / columna C de Estudiantes (Code.gs).
 * Antes aquí había nombres de ejemplo (p. ej. «8-9 A») y parecían cursos reales.
 */
export const GRUPOS_FALLBACK = [];
/**
 * Respaldo del navegador (coordinación o API caída). El listado oficial del estudiante sale de la hoja
 * **Materias** en Apps Script + DocenteAsignaciones/Notas. Aquí **Sociales**, **Ética** y **Religión**
 * coinciden con el catálogo típico IEMFS para que no falten si solo corre el legado cliente.
 */
export const MATERIAS = [
  "Matemáticas",
  "Geometría",
  "Estadística",
  "Ciencias Naturales",
  "Lengua Castellana",
  "Competencia lectora",
  "Inglés",
  "Sociales",
  "Educación Física",
  "Ética",
  "Religión"
];

/**
 * Áreas curriculares IEMFS: promedio orientativo a partir de las definitivas por asignatura del boletín.
 * `etiquetas`: nombres posibles en hoja Notas / DocenteAsignaciones (misma lógica que limpiarTexto en Code.gs).
 */
export const AREAS_BOLETIN_ESTUDIANTE = [
  {
    nombre: "MEG",
    formulaCorta: "70% Matemáticas · 15% Geometría · 15% Estadística",
    componentes: [
      { etiquetas: ["Matemáticas"], peso: 0.7 },
      { etiquetas: ["Geometría"], peso: 0.15 },
      { etiquetas: ["Estadística"], peso: 0.15 }
    ]
  },
  {
    nombre: "Humanidades",
    formulaCorta: "80% Lengua Castellana · 20% Competencia lectora",
    componentes: [
      { etiquetas: ["Lengua Castellana", "Lengua castellana", "Castellano"], peso: 0.8 },
      { etiquetas: ["Competencia lectora", "Competencia Lectora"], peso: 0.2 }
    ]
  }
];

/** Estados de matrícula (coinciden con Code.gs / hoja Estudiantes columna D). */
export const ESTADOS_MATRICULA = [
  { value: "Activo", label: "Activo — asiste con regularidad" },
  { value: "Inactivo", label: "Inactivo — ausencias prolongadas o seguimiento" },
  { value: "Desertor", label: "Desertor — abandonó el proceso escolar" },
  { value: "Matrícula cancelada", label: "Matrícula cancelada — retiro o baja administrativa" }
];

export const NOTA_MIN = 0;
export const NOTA_MAX = 5;

/**
 * Código de acceso de coordinación (misma clave institucional que en Apps Script).
 * Debe coincidir con Code.gs: valor por defecto o propiedad del script NOTAS_IEMFS_CODIGO_COORD.
 */
export const CODIGO_COORDINADOR = "C001";
