/** Estado mutable compartido entre módulos (live bindings vía import). */
export const state = {
  estudianteActual: null,
  cacheEstudiantes: [],
  cacheNotasGrupo: [],
  cacheTitulosGrupo: [],
  filtroBusqueda: "",
  /** "calificaciones" | "inasistencias" | "resumen-faltas" | "resumen-area" | "mosaico-grupo" */
  appModo: "calificaciones",
  cacheInasistenciasDia: null,
  /** Evita que una respuesta lenta de un día anterior pise la tabla del día que eligió después. */
  inasistenciasDiaFetchSeq: 0,
  cacheResumenFaltas: null,
  /** Docente/coordinación: filas del resumen MEG / Humanidades (vista grupo). */
  cacheResumenArea: null,
  /** Mosaico del grupo (URL o data URI desde hoja MosaicosGrupo). null = no cargado. */
  cacheMosaicoGrupo: null,
  cachePeriodosAcademicos: [],
  /** Docente: pares { grupo, materia } desde el servidor (misión miContexto). */
  cacheDocenteAsignaciones: null,
  /** Nombre del docente (miContexto); también puede leerse de almacenamiento. */
  cacheDocenteNombre: "",
  /** Nombre del estudiante (ping) para la insignia de sesión. */
  estudianteNombreSesion: "",
  /**
   * Estudiante: materias del curso desde servidor (DocenteAsignaciones / Notas).
   * null = aún no cargado (usar lista fija MATERIAS en config); [] = cargado vacío.
   */
  cacheMateriasEstudiante: null,
  /** Se asigna en DOMContentLoaded; cancela el debounce del campo Buscar al cambiar de grupo. */
  cancelDebouncedFiltroBusqueda() {},
  /** Última lista `accion:notas` del expediente abierto (selector de evaluación en ficha). */
  fichaListaNotas: null
};
