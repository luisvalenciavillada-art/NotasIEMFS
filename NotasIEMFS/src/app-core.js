import {
  API,
  NOTAS_IEMFS_BUILD_ESPERADO,
  STORAGE,
  GRUPOS_FALLBACK,
  MATERIAS,
  AREAS_BOLETIN_ESTUDIANTE,
  ESTADOS_MATRICULA,
  NOTA_MIN,
  NOTA_MAX,
  CODIGO_COORDINADOR
} from "./config.js";
import { state } from "./state.js";

/**
 * En http://127.0.0.1 u http://localhost el POST directo a script.google.com suele seguir un 302
 * y el navegador reenvía sin el cuerpo → el servidor ve la petición sin auth («Falta clave de acceso»).
 * `servidor_local.py` expone /notas-gas-proxy (curl -L) para mantener el body intacto.
 */
function esHostLocalDesarrollo() {
  const h = String(window.location.hostname || "").toLowerCase();
  return h === "127.0.0.1" || h === "localhost" || h === "::1" || h === "[::1]";
}

function urlApiNotas() {
  try {
    const pr = String(window.location.protocol || "");
    if (esHostLocalDesarrollo() && pr === "http:") {
      return `${window.location.origin}/notas-gas-proxy`;
    }
  } catch (_) {
    /* ignore */
  }
  return API;
}

/**
 * Falla en cliente antes de llamar a GAS si faltan credenciales (evita el error genérico del servidor).
 * `ping` se omite: el login envía documento/PIN o acaba de guardar clave en almacenamiento.
 */
function asegurarCredencialesParaApiSalvoPing_(merged) {
  const m = merged && typeof merged === "object" && !Array.isArray(merged) ? merged : {};
  const accion = String(m.accion || "").trim();
  if (accion === "ping") return;
  if (esEstudiante()) {
    const doc = String(m.documento || "").trim();
    const pin = String(m.pin || "").trim();
    if (!doc || !pin) {
      throw new Error("Falta documento o PIN. Inicie sesión de nuevo como estudiante.");
    }
    return;
  }
  const clave = String(m.clave || "").trim();
  if (!clave) {
    throw new Error(
      "Falta la clave de acceso (sesión no disponible o caducada en esta pestaña). " +
        "Use «Cerrar sesión», entre de nuevo con su contraseña y marque «Recordarme en este equipo» si no quiere repetir al cerrar el navegador. " +
        "En esta PC, abra la app con iniciar-servidor.bat (servidor con proxy); con otro servidor local sin /notas-gas-proxy el guardado puede fallar igual."
    );
  }
}

// =====================
// UTILIDADES
// =====================
/** Misma idea que Code.gs `limpiarTexto` para emparejar nombres de materia en el cliente. */
function normalizarMateriaClaveIEMFS(txt) {
  return String(txt || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

/** Ejecuta `fn` tras `waitMs` sin actividad; incluye `cancel()` para reseteos. */
function debounce(fn, waitMs) {
  let t = null;
  function wrapped() {
    if (t !== null) clearTimeout(t);
    t = setTimeout(() => {
      t = null;
      fn();
    }, waitMs);
  }
  wrapped.cancel = function cancelDebounced() {
    if (t !== null) {
      clearTimeout(t);
      t = null;
    }
  };
  return wrapped;
}

function getGrupo() {
  const el = document.getElementById("grupo");
  return el ? el.value.trim() : "";
}

function getMateria() {
  const el = document.getElementById("materia");
  if (!el) return esEstudiante() ? "" : "Matemáticas";
  const v = el.value.trim();
  return v;
}

/**
 * Materias del selector estudiante: vienen de la API `materiasEstudiante` (DocenteAsignaciones + Notas).
 * Si la petición falla, el script no está desplegado o devuelve algo inesperado, se usa la lista fija
 * MATERIAS de config.js — conviene desplegar Code.gs y mantener DocenteAsignaciones para el curso completo.
 */
function materiasListaEstudianteActual() {
  const x = state.cacheMateriasEstudiante;
  if (Array.isArray(x) && x.length > 0) return x;
  return [...MATERIAS];
}

async function refrescarMateriasEstudianteCache() {
  if (!esEstudiante()) return;
  try {
    const data = await apiGet({ accion: "materiasEstudiante" });
    let list = null;
    if (Array.isArray(data)) list = data;
    else if (data && typeof data === "object" && Array.isArray(data.materias)) {
      list = data.materias;
    }
    state.cacheMateriasEstudiante = Array.isArray(list) ? list : null;
  } catch (_) {
    state.cacheMateriasEstudiante = null;
  }
}

function leerMateriaExpedienteEstudiante() {
  try {
    const v = sessionStorage.getItem(STORAGE.estudianteMateriaExp);
    const lista = materiasListaEstudianteActual();
    return v && lista.includes(v) ? v : "";
  } catch (_) {
    return "";
  }
}

function guardarMateriaExpedienteEstudiante(materia) {
  try {
    const m = String(materia || "").trim();
    const lista = materiasListaEstudianteActual();
    if (m && lista.includes(m)) sessionStorage.setItem(STORAGE.estudianteMateriaExp, m);
    else sessionStorage.removeItem(STORAGE.estudianteMateriaExp);
  } catch (_) {
    /* ignore */
  }
}

/** Lista de materias con opción vacía inicial (solo estudiante; expediente requiere elegir). */
function llenarSelectMateriaEstudiante(valorGuardado) {
  const sel = document.getElementById("materia");
  if (!sel) return;
  const mats = materiasListaEstudianteActual();
  sel.innerHTML = "";
  const ph = document.createElement("option");
  ph.value = "";
  ph.textContent = "— Elija materia —";
  sel.appendChild(ph);
  mats.forEach((m) => {
    const o = document.createElement("option");
    o.value = m;
    o.textContent = m;
    sel.appendChild(o);
  });
  const ok = valorGuardado && mats.includes(valorGuardado);
  sel.value = ok ? valorGuardado : "";
}

function notaIdDe(n) {
  if (!n || typeof n !== "object") return null;
  /** `notaId` primero: evita confundir la fila de la nota con otro `id` (p. ej. documento) si el objeto se mezcla. */
  const claves = [
    "notaId",
    "notaRef",
    "id",
    "rowId",
    "fila",
    "row",
    "Row",
    "ID",
    "Id",
    "_id",
    "sheetRow",
    "linea",
    "nroFila",
    "numeroFila"
  ];
  for (const k of claves) {
    const v = n[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") {
      return v;
    }
  }
  return null;
}

/**
 * Extrae AAAA-MM-DD de ISO/fecha del servidor sin aplicar desfase UTC → día calendario local estable.
 * Evita que «1 may» en hoja se vea como otro día al usar solo `new Date(iso).getDate()`.
 */
function extraerYmdDesdeFechaApi_(val) {
  if (val == null || val === "") return "";
  const m = /^(\d{4}-\d{2}-\d{2})/.exec(String(val).trim());
  return m ? m[1] : "";
}

/** Texto corto para tarjetas de nota (misma lógica de día calendario que el input date). */
function fechaTextoMostrarNota_(val) {
  const ymd = extraerYmdDesdeFechaApi_(val);
  if (ymd) {
    const p = ymd.split("-");
    const d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    return d.toLocaleDateString("es-CO", { day: "numeric", month: "short", year: "numeric" });
  }
  if (!val) return "";
  const d = new Date(val);
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleDateString("es-CO", { day: "numeric", month: "short", year: "numeric" });
}

/** Valor `yyyy-MM-dd` para `<input type="date">` a partir de ISO o fecha del detalle. */
function fechaIsoParaInputDate(val) {
  if (val == null || val === "") {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }
  const desdeIso = extraerYmdDesdeFechaApi_(val);
  if (desdeIso) return desdeIso;
  try {
    const d = new Date(val);
    if (Number.isNaN(d.getTime())) {
      const d2 = new Date();
      const y2 = d2.getFullYear();
      const m2 = String(d2.getMonth() + 1).padStart(2, "0");
      const day2 = String(d2.getDate()).padStart(2, "0");
      return `${y2}-${m2}-${day2}`;
    }
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  } catch {
    const d3 = new Date();
    return `${d3.getFullYear()}-${String(d3.getMonth() + 1).padStart(2, "0")}-${String(d3.getDate()).padStart(2, "0")}`;
  }
}

/** Valores originales para localizar la fila cuando el servidor no envía id (debe soportarlo en Apps Script). */
function snapshotAntesParaBuscar(n) {
  if (!n || typeof n !== "object") return null;
  const titulo = String(n.titulo ?? "").trim();
  const tipo = String(n.tipo ?? "").trim();
  const fecha = n.fecha != null && n.fecha !== "" ? n.fecha : "";
  const notaRaw = n.nota;
  if (notaRaw === undefined || notaRaw === null || String(notaRaw).trim() === "") {
    return null;
  }
  if (!titulo && !fecha) {
    return null;
  }
  return {
    titulo,
    tipo,
    nota: notaRaw,
    fecha
  };
}

function claseTipo(tipo) {
  const t = String(tipo || "").toLowerCase();
  if (t.includes("seguimiento")) return "note-card--tipo-seguimiento";
  if (t.includes("actitudinal")) return "note-card--tipo-actitudinal";
  if (t.includes("prueba")) return "note-card--tipo-prueba";
  return "";
}

function clasesEstadoMatriculaChip(estado) {
  const s = String(estado || "").toLowerCase();
  let mod = "estado-chip--activo";
  if (s.includes("cancel")) mod = "estado-chip--cancelado";
  else if (s.includes("desertor")) mod = "estado-chip--desertor";
  else if (s.includes("inactivo")) mod = "estado-chip--inactivo";
  return `chip estado-chip ${mod}`;
}

function llenarSelectEstadoMatricula() {
  const sel = document.getElementById("estado-matricula");
  if (!sel) return;
  sel.innerHTML = "";
  ESTADOS_MATRICULA.forEach((opt) => {
    const o = document.createElement("option");
    o.value = opt.value;
    o.textContent = opt.label;
    sel.appendChild(o);
  });
}

function actualizarChipEstadoEnPanel(estado) {
  const chip = document.getElementById("estudiante-estado-chip");
  if (!chip) return;
  const txt = estado || "Activo";
  chip.textContent = txt;
  chip.className = clasesEstadoMatriculaChip(txt);
}

function estadoEsActivo(estado) {
  const s = String(estado || "").toLowerCase();
  return s.includes("activo") && !s.includes("inactivo");
}

function normalizarTextoSimple(txt) {
  return String(txt || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

/** Grupo/cursos: colapsa espacios para coincidir con la hoja aunque haya dobles espacios. */
function normalizarGrupoCurso(txt) {
  return normalizarTextoSimple(String(txt || "").replace(/\s+/g, " ").trim());
}

/** Valor canónico en `opciones` que coincide con el curso `valor` (texto exacto o normalizado). */
function coincidenciaGrupoEnOpciones(valor, opciones) {
  if (!valor || !Array.isArray(opciones) || !opciones.length) return "";
  const v0 = String(valor).trim();
  if (!v0) return "";
  for (let i = 0; i < opciones.length; i++) {
    const g = String(opciones[i] == null ? "" : opciones[i]).trim();
    if (g === v0) return g;
  }
  const nv = normalizarGrupoCurso(v0);
  for (let i = 0; i < opciones.length; i++) {
    const g = String(opciones[i] == null ? "" : opciones[i]).trim();
    if (g && normalizarGrupoCurso(g) === nv) return g;
  }
  return "";
}

function parseNotaDesdeCampo(str) {
  const t = String(str ?? "").trim().replace(",", ".");
  if (t === "") return null;
  const n = Number(t);
  return Number.isNaN(n) ? null : n;
}

function notaEnRango(n) {
  return n != null && !Number.isNaN(n) && n >= NOTA_MIN && n <= NOTA_MAX;
}

function redondearNotaPasoDecima(n) {
  return Math.round(Math.min(NOTA_MAX, Math.max(NOTA_MIN, n)) * 10) / 10;
}

/** Al salir del campo: recorta a 0–5 y redondea a una decimal. */
function vincularValidacionNotaInput(el) {
  if (!el || el.dataset.notaValidBound) return;
  el.dataset.notaValidBound = "1";
  el.addEventListener("blur", () => {
    const raw = el.value;
    if (String(raw).trim() === "") return;
    const n = parseNotaDesdeCampo(raw);
    if (n === null) {
      el.value = "";
      toast("Usa un número entre 0 y 5.", "err");
      return;
    }
    let ajuste = false;
    let v = n;
    if (n < NOTA_MIN) {
      v = NOTA_MIN;
      ajuste = true;
    } else if (n > NOTA_MAX) {
      v = NOTA_MAX;
      ajuste = true;
    }
    const red = redondearNotaPasoDecima(v);
    el.value = String(red);
    if (ajuste) {
      toast("Las notas solo van de 0,0 a 5,0. Se ajustó al límite permitido.", "err");
    }
  });
}

async function cargarNotasGrupoMateria() {
  try {
    const data = await apiGet({
      accion: "notasGrupo",
      grupo: getGrupo(),
      materia: getMateria(),
      ...periodoParamParaApiDocente()
    });
    state.cacheNotasGrupo = Array.isArray(data) ? data : [];
  } catch (_) {
    state.cacheNotasGrupo = [];
  }
}

function actualizarTituloCargaRapida() {
  const out = document.getElementById("cr-sticky-title");
  if (!out) return;
  const triple = Boolean(document.getElementById("cr-definitiva-triple")?.checked);
  if (triple) {
    const m = String(getMateria() || "").trim() || "—";
    out.textContent = `Modo definitiva: misma nota en Seguimiento, Actitudinal y Prueba · ${m}`;
    return;
  }
  const t = document.getElementById("cr-titulo");
  const txt = t ? t.value.trim() : "";
  out.textContent = `Título: ${txt || "—"}`;
}

/** Modo «nota única → tres componentes»: aviso y desactiva título/tipo por defecto. */
function aplicarEstadoUiDefinitivaTriple() {
  const triple = Boolean(document.getElementById("cr-definitiva-triple")?.checked);
  const aviso = document.getElementById("cr-definitiva-aviso");
  const tituloEl = document.getElementById("cr-titulo");
  const tipoEl = document.getElementById("cr-tipo");
  if (aviso) {
    aviso.classList.toggle("is-hidden", !triple);
    aviso.hidden = !triple;
  }
  if (tituloEl) {
    tituloEl.disabled = triple;
    tituloEl.title = triple
      ? "En modo definitiva única los títulos se forman con la materia (Seguimiento, Actitudinal, Prueba)."
      : "";
  }
  if (tipoEl) {
    tipoEl.disabled = triple;
    tipoEl.title = triple
      ? "En modo definitiva única se registran automáticamente los tres tipos."
      : "";
  }
  actualizarTituloCargaRapida();
}

function aplicarVisibilidadColumnasCargaRapida() {
  const wrap = document.getElementById("cr-wrap");
  if (!wrap) return;
  wrap.dataset.colEstado = document.getElementById("cr-col-estado")?.checked ? "on" : "off";
  wrap.dataset.colTipo = document.getElementById("cr-col-tipo")?.checked ? "on" : "off";
  wrap.dataset.colNota = document.getElementById("cr-col-nota")?.checked ? "on" : "off";
}

/** Si el campo está vacío, deja la fecha del registro masivo en hoy (AAAA-MM-DD). */
function asegurarFechaCargaRapidaPorDefecto() {
  const el = document.getElementById("cr-fecha");
  if (!el || el.value) return;
  el.value = fechaIsoParaInputDate("");
}

function renderCargaRapida() {
  const body = document.getElementById("cr-body");
  if (!body) return;

  asegurarFechaCargaRapidaPorDefecto();

  const soloActivos = Boolean(document.getElementById("cr-solo-activos")?.checked);
  const tipoDefault = document.getElementById("cr-tipo")?.value || "Seguimiento";
  const estudiantes = (state.cacheEstudiantes || []).filter((e) => {
    if (!soloActivos) return true;
    return estadoEsActivo(e.estadoMatricula || "Activo");
  });

  if (!estudiantes.length) {
    body.innerHTML = `<tr><td colspan="5"><div class="state-msg">No hay estudiantes para el registro masivo. Carga el grupo arriba.</div></td></tr>`;
    return;
  }

  body.innerHTML = estudiantes.map((e, i) => {
    const estado = e.estadoMatricula || "Activo";
    return `
      <tr data-estudiante-id="${escapeHtml(e.id)}">
        <td>${i + 1}</td>
        <td><strong>${escapeHtml(e.nombre || "")}</strong></td>
        <td class="col-estado"><span class="${clasesEstadoMatriculaChip(estado)}">${escapeHtml(estado)}</span></td>
        <td class="col-tipo">
          <select class="input cr-tipo-select">
            <option value="Seguimiento" ${tipoDefault === "Seguimiento" ? "selected" : ""}>Seguimiento</option>
            <option value="Actitudinal" ${tipoDefault === "Actitudinal" ? "selected" : ""}>Actitudinal</option>
            <option value="Prueba" ${tipoDefault === "Prueba" ? "selected" : ""}>Prueba</option>
          </select>
        </td>
        <td class="col-nota">
          <input type="number" class="input cr-nota-input" min="0" max="5" step="0.1" placeholder="0.0" inputmode="decimal">
        </td>
      </tr>
    `;
  }).join("");

  document.querySelectorAll(".cr-nota-input").forEach((inp) => vincularValidacionNotaInput(inp));
  aplicarEstadoUiDefinitivaTriple();
}

function limpiarCargaRapida() {
  document.querySelectorAll(".cr-nota-input").forEach((inp) => {
    inp.value = "";
  });
  const crFecha = document.getElementById("cr-fecha");
  if (crFecha) crFecha.value = fechaIsoParaInputDate("");
}

/** Tras sincronizar Estudiantes, descarta ítems cuyo id ya no está en el curso mostrado (p. ej. traslado manual en la hoja). */
function alinearItemsConCacheEstudiantesCargaRapida_(itemsIn) {
  const idsEnCurso = new Set(
    (state.cacheEstudiantes || []).map((e) => String(e.id || "").trim()).filter(Boolean)
  );
  const antes = itemsIn.length;
  const out = itemsIn.filter((it) => idsEnCurso.has(String(it.estudiante || "").trim()));
  const omitidos = antes - out.length;
  const alumnosConNotaLocal = new Set(
    out.map((it) => String(it.estudiante || "").trim()).filter(Boolean)
  ).size;
  return { items: out, alumnosConNotaLocal, omitidos };
}

function moverFocoNotaSiguiente(inputActual) {
  const inputs = Array.from(document.querySelectorAll(".cr-nota-input"));
  const idx = inputs.indexOf(inputActual);
  if (idx >= 0 && idx < inputs.length - 1) {
    inputs[idx + 1].focus();
    inputs[idx + 1].select();
  }
}

/**
 * Si falla guardarMasivo y solo queda guardar nota por nota (lento), el docente elige explícitamente.
 * @returns {Promise<"continuar" | "declinar">}
 */
async function preguntarModoSeguroCargaRapida_(opts) {
  const numRegistros = opts.numRegistros ?? 0;
  const modoTriple = Boolean(opts.modoDefinitivaTriple);
  const mensajeCorto = String(opts.mensajeServidor || "")
    .trim()
    .slice(0, 320);

  ocultarBloqueoGuardandoCalificaciones();

  const btnCerrar = document.getElementById("modal-cerrar");
  const prevCerrarHidden = btnCerrar ? btnCerrar.hidden : false;
  if (btnCerrar) btnCerrar.hidden = true;

  const detalleHtml = mensajeCorto
    ? `<p class="empty-hint" style="word-break:break-word;"><strong>Servidor:</strong> ${escapeHtml(mensajeCorto)}</p>`
    : "";

  const hintRangoVsDatos =
    /n[uú]mero de filas|intervalo|rango/i.test(mensajeCorto) &&
    (/datos tienen|no coincide/i.test(mensajeCorto) || /los datos/i.test(mensajeCorto))
      ? `<p class="state-msg state-msg--warn" style="margin-top:0.75rem;">Este mensaje casi siempre indica que <strong>Google Apps Script sigue con un Code.gs antiguo</strong> (usa <code>setValues</code> mal). Abra el proyecto vinculado a la hoja, pegue <strong>todo</strong> el archivo Code.gs de su carpeta NotasIEMFS, guarde y use <strong>Implementar → Nueva implementación</strong> del despliegue web. Compruebe que la URL en <code>config.js</code> sea la de esa nueva implementación.</p>`
      : "";

  return await new Promise((resolve) => {
    let settled = false;
    const backdrop = document.getElementById("modal-backdrop");

    const finish = (valor) => {
      if (settled) return;
      settled = true;
      document.removeEventListener("keydown", onEscCapture, true);
      if (backdrop) backdrop.removeEventListener("click", onBackdropCapture, true);
      if (btnCerrar) btnCerrar.hidden = prevCerrarHidden;
      cerrarModal();
      resolve(valor);
    };

    const onBackdropCapture = (ev) => {
      if (ev.target === backdrop) {
        ev.preventDefault();
        ev.stopImmediatePropagation();
        finish("declinar");
      }
    };

    const onEscCapture = (ev) => {
      if (ev.key !== "Escape") return;
      const bd = document.getElementById("modal-backdrop");
      if (!bd || bd.hidden || bd.classList.contains("is-hidden")) return;
      ev.preventDefault();
      ev.stopImmediatePropagation();
      finish("declinar");
    };

    const intro = modoTriple
      ? `<p>El guardado <strong>en bloque</strong> falló. Hay <strong>${numRegistros}</strong> registros pendientes (tres por estudiante en modo definitiva). El <strong>modo seguro</strong> los envía uno a uno y puede tardar varios minutos.</p>`
      : `<p>El guardado <strong>en bloque</strong> falló. Hay <strong>${numRegistros}</strong> notas pendientes. El <strong>modo seguro</strong> las envía una a una y puede tardar varios minutos.</p>`;

    abrirModal(
      "¿Continuar en modo seguro?",
      `${intro}
      ${detalleHtml}
      ${hintRangoVsDatos}
      <p class="card__hint">Si prefiere revisar conexión, proxy local o que <strong>Code.gs</strong> esté actualizado en Apps Script, pulse <strong>Declinar</strong>; no se guardará nada hasta que vuelva a pulsar Guardar.</p>
      <div style="display:flex;gap:0.75rem;flex-wrap:wrap;margin-top:1rem;">
        <button type="button" class="btn btn--primary" id="cr-modoseguro-continue">Continuar (modo seguro)</button>
        <button type="button" class="btn" id="cr-modoseguro-decline">Declinar y revisar</button>
      </div>`
    );

    document.addEventListener("keydown", onEscCapture, true);
    if (backdrop) backdrop.addEventListener("click", onBackdropCapture, true);

    document.getElementById("cr-modoseguro-continue")?.addEventListener("click", () => finish("continuar"));
    document.getElementById("cr-modoseguro-decline")?.addEventListener("click", () => finish("declinar"));
  });
}

async function guardarCargaRapida() {
  if (esModoSoloConsulta()) {
    toast(
      esEstudiante()
        ? "El acceso de estudiante es solo consulta."
        : "El acceso de coordinación es solo consulta.",
      "err"
    );
    return;
  }
  const modoDefinitivaTriple = Boolean(document.getElementById("cr-definitiva-triple")?.checked);
  const titulo = document.getElementById("cr-titulo")?.value.trim() || "";
  const tipoDefault = document.getElementById("cr-tipo")?.value || "Seguimiento";
  asegurarFechaCargaRapidaPorDefecto();
  const fechaEval = document.getElementById("cr-fecha")?.value?.trim() || "";
  const btn = document.getElementById("cr-guardar");
  if (!modoDefinitivaTriple && !titulo) {
    toast("Escribe el título de la evaluación.", "err");
    return;
  }
  if (!fechaEval) {
    toast("Indique la fecha de la evaluación (registro masivo).", "err");
    return;
  }

  function capturarBorradorCargaRapida_() {
    const out = new Map();
    const rows = Array.from(document.querySelectorAll("#cr-body tr[data-estudiante-id]"));
    rows.forEach((row) => {
      const id = String(row.dataset.estudianteId || "").trim();
      if (!id) return;
      const notaEl = row.querySelector(".cr-nota-input");
      const tipoEl = row.querySelector(".cr-tipo-select");
      const nota = notaEl ? String(notaEl.value || "").trim() : "";
      const tipo = tipoEl ? String(tipoEl.value || "").trim() : tipoDefault;
      if (!nota) return;
      out.set(id, { nota, tipo });
    });
    return out;
  }

  function restaurarBorradorCargaRapida_(borrador) {
    if (!(borrador instanceof Map) || !borrador.size) return;
    const rows = Array.from(document.querySelectorAll("#cr-body tr[data-estudiante-id]"));
    rows.forEach((row) => {
      const id = String(row.dataset.estudianteId || "").trim();
      if (!id || !borrador.has(id)) return;
      const data = borrador.get(id);
      const notaEl = row.querySelector(".cr-nota-input");
      const tipoEl = row.querySelector(".cr-tipo-select");
      if (notaEl) notaEl.value = data.nota;
      if (tipoEl && data.tipo) tipoEl.value = data.tipo;
    });
  }

  function construirItemsDesdeTabla_() {
    const rows = Array.from(document.querySelectorAll("#cr-body tr[data-estudiante-id]"));
    let fueraDeRango = 0;
    let alumnosConNotaLocal = 0;
    const itemsLocal = [];
    for (const row of rows) {
      const notaEl = row.querySelector(".cr-nota-input");
      const tipoEl = row.querySelector(".cr-tipo-select");
      const notaTxt = notaEl ? notaEl.value.trim() : "";
      if (!notaTxt) continue;
      const n = parseNotaDesdeCampo(notaTxt);
      if (n === null || !notaEnRango(n)) {
        fueraDeRango += 1;
        continue;
      }
      alumnosConNotaLocal += 1;
      const num = redondearNotaPasoDecima(n);
      const idEst = row.dataset.estudianteId;
      if (modoDefinitivaTriple) {
        itemsLocal.push(
          {
            estudiante: idEst,
            tipo: "Seguimiento",
            titulo: `Seguimiento ${materiaLabel}`,
            nota: num
          },
          {
            estudiante: idEst,
            tipo: "Actitudinal",
            titulo: `Actitudinal ${materiaLabel}`,
            nota: num
          },
          {
            estudiante: idEst,
            tipo: "Prueba",
            titulo: `Prueba ${materiaLabel}`,
            nota: num
          }
        );
      } else {
        itemsLocal.push({
          estudiante: idEst,
          tipo: tipoEl ? tipoEl.value : tipoDefault,
          nota: num
        });
      }
    }
    return { rows, fueraDeRango, alumnosConNotaLocal, itemsLocal };
  }

  const periodoPayload = periodoParamParaApiDocente();
  const grupo = getGrupo();
  const materia = getMateria();
  const materiaLabel = String(materia || "").trim();

  if (modoDefinitivaTriple && !materiaLabel) {
    toast("Elija la materia arriba antes de usar la nota única definitiva.", "err");
    return;
  }

  const tituloComunApi = modoDefinitivaTriple
    ? `${materiaLabel} — definitiva periodo`
    : titulo;

  let { fueraDeRango, alumnosConNotaLocal: alumnosConNota, itemsLocal: items } = construirItemsDesdeTabla_();
  if (fueraDeRango) {
    toast("Hay notas fuera del rango 0 a 5. Corrígelas o sal del campo para ajustarlas automáticamente.", "err");
    return;
  }

  if (!items.length) {
    toast("Ingresa al menos una nota entre 0 y 5.", "err");
    return;
  }

  async function guardarCargaRapidaNotaPorNota_() {
    let ok = 0;
    let fail = 0;
    for (const it of items) {
      try {
        await apiPost({
          accion: "guardar",
          estudiante: it.estudiante,
          materia,
          grupo,
          tipo: it.tipo,
          titulo: it.titulo != null && String(it.titulo).trim() !== "" ? String(it.titulo).trim() : tituloComunApi,
          nota: it.nota,
          ...periodoPayload,
          fecha: fechaEval,
          fechaCalificacion: fechaEval,
          fechaNota: fechaEval
        });
        ok += 1;
      } catch {
        fail += 1;
      }
    }
    return { ok, fail };
  }

  if (btn) btn.disabled = true;
  mostrarBloqueoGuardandoCalificaciones();
  try {
    const borradorSync = capturarBorradorCargaRapida_();
    try {
      const dataSync = await apiGet({ accion: "estudiantes", grupo });
      state.cacheEstudiantes = ordenarEstudiantesPorNombreAsc(
        (Array.isArray(dataSync) ? dataSync : []).map((s) => {
          const row = { ...s };
          if (!row.grupo) row.grupo = grupo;
          return row;
        })
      );
      await cargarNotasGrupoMateria();
      renderCargaRapida();
      restaurarBorradorCargaRapida_(borradorSync);
    } catch {
      toast(
        "No se pudo actualizar la lista del curso antes de guardar. Revise la conexión, vuelva a cargar el grupo y reintente.",
        "err"
      );
      return;
    }

    ({ fueraDeRango, alumnosConNotaLocal: alumnosConNota, itemsLocal: items } = construirItemsDesdeTabla_());
    if (fueraDeRango) {
      toast("Hay notas fuera del rango 0 a 5. Corrígelas o sal del campo para ajustarlas automáticamente.", "err");
      return;
    }
    if (!items.length) {
      toast("Ingresa al menos una nota entre 0 y 5.", "err");
      return;
    }

    const alineados = alinearItemsConCacheEstudiantesCargaRapida_(items);
    items = alineados.items;
    alumnosConNota = alineados.alumnosConNotaLocal;
    if (alineados.omitidos > 0) {
      toast(
        `Lista del curso actualizada: se quitaron ${alineados.omitidos} registro(s) de estudiantes que ya no están en este grupo (revise la hoja Estudiantes o traslados).`,
        "ok"
      );
    }
    if (!items.length) {
      toast("Tras alinear con la hoja no quedan notas para guardar en este curso.", "err");
      return;
    }

    await apiPost(
      {
        accion: "guardarMasivo",
        grupo,
        materia,
        titulo: tituloComunApi,
        items,
        ...periodoPayload,
        fecha: fechaEval,
        fechaCalificacion: fechaEval,
        fechaNota: fechaEval
      },
      { timeoutMs: 240000 }
    );
    const msgOk = modoDefinitivaTriple
      ? `Información guardada: ${alumnosConNota} estudiante(s), ${items.length} registros de calificaciones (Seguimiento, Actitudinal y Prueba).`
      : `Información guardada: ${items.length} nota(s) de calificaciones.`;
    toast(msgOk, "ok");
    limpiarCargaRapida();
    if (state.estudianteActual) await abrirEstudiante(state.estudianteActual);
  } catch (e) {
    const msg = String(e.message || "");
    const servidorSinLote =
      msg.includes("Acción no válida") ||
      msg.includes("accion no valida") ||
      /acción no válida/i.test(msg);
    const servidorConDesfaseRango =
      (/n[uú]mero de filas/i.test(msg) && /no coincide con el n[uú]mero de filas del rango/i.test(msg)) ||
      (/n[uú]mero de filas de los datos/i.test(msg) && /intervalo/i.test(msg)) ||
      (/n[uú]mero de columnas/i.test(msg) &&
        (/no coincide con el n[uú]mero de columnas/i.test(msg) ||
          /no coincide con el n[uú]mero de columnas del rango/i.test(msg))) ||
      /does not match the number of rows in the range/i.test(msg) ||
      /does not match the number of columns/i.test(msg);
    if (servidorSinLote || servidorConDesfaseRango) {
      if (servidorConDesfaseRango) {
        const borrador = capturarBorradorCargaRapida_();
        try {
          const [data] = await Promise.all([
            apiGet({ accion: "estudiantes", grupo }),
            cargarNotasGrupoMateria()
          ]);
          state.cacheEstudiantes = ordenarEstudiantesPorNombreAsc(
            (Array.isArray(data) ? data : []).map((s) => {
              const row = { ...s };
              if (!row.grupo) row.grupo = grupo;
              return row;
            })
          );
          renderCargaRapida();
          restaurarBorradorCargaRapida_(borrador);
          ({ fueraDeRango, alumnosConNotaLocal: alumnosConNota, itemsLocal: items } = construirItemsDesdeTabla_());
          if (!fueraDeRango && items.length) {
            const alinRetry = alinearItemsConCacheEstudiantesCargaRapida_(items);
            items = alinRetry.items;
            alumnosConNota = alinRetry.alumnosConNotaLocal;
            if (!items.length) {
              throw new Error("Sin ítems tras alinear con la lista del curso.");
            }
            await apiPost(
              {
                accion: "guardarMasivo",
                grupo,
                materia,
                titulo: tituloComunApi,
                items,
                ...periodoPayload,
                fecha: fechaEval,
                fechaCalificacion: fechaEval,
                fechaNota: fechaEval
              },
              { timeoutMs: 240000 }
            );
            toast("Información guardada: reajuste automático aplicado y guardado masivo exitoso.", "ok");
            limpiarCargaRapida();
            if (state.estudianteActual) await abrirEstudiante(state.estudianteActual);
            return;
          }
        } catch {
          // Si no logra reajustar o reintentar masivo, continua al modo seguro.
        }
      }
      const decisionModoSeguro = await preguntarModoSeguroCargaRapida_({
        mensajeServidor: msg,
        numRegistros: items.length,
        modoDefinitivaTriple
      });
      if (decisionModoSeguro === "declinar") {
        toast(
          "No se guardó nada. Revise el mensaje del servidor, actualice Code.gs en Apps Script y la conexión; luego pulse Guardar de nuevo.",
          "err"
        );
        return;
      }
      mostrarBloqueoGuardandoCalificaciones(
        "Guardando calificaciones… Modo seguro (nota por nota); puede tardar varios minutos."
      );
      const { ok, fail } = await guardarCargaRapidaNotaPorNota_();
      if (ok) {
        toast(
          (modoDefinitivaTriple
            ? `Modo seguro: información guardada, ${ok} registro(s) de calificaciones.`
            : `Modo seguro: información guardada, ${ok} nota(s) de calificaciones.`) +
            (fail ? ` Fallidas: ${fail}.` : ""),
          fail ? "err" : "ok"
        );
        limpiarCargaRapida();
        if (state.estudianteActual) await abrirEstudiante(state.estudianteActual);
      } else {
        toast(msg || "No se pudieron guardar las calificaciones.", "err");
      }
    } else {
      toast(msg || "No se pudieron guardar las calificaciones.", "err");
    }
  } finally {
    ocultarBloqueoGuardandoCalificaciones();
    if (btn) btn.disabled = false;
  }
}

/** Overlay a pantalla completa (calificaciones masivas, asistencia del día, etc.). */
function mostrarOverlayEsperaApp_(mensaje) {
  const overlay = document.getElementById("app-busy-overlay");
  const msgEl = document.getElementById("app-busy-msg");
  const texto =
    mensaje ||
    "Procesando… Espere un momento; no cierre la ventana.";
  if (msgEl) msgEl.textContent = texto;
  if (overlay) {
    overlay.classList.remove("is-hidden");
    overlay.hidden = false;
    overlay.setAttribute("aria-hidden", "false");
  }
}

function ocultarOverlayEsperaApp_() {
  const overlay = document.getElementById("app-busy-overlay");
  if (overlay) {
    overlay.classList.add("is-hidden");
    overlay.hidden = true;
    overlay.setAttribute("aria-hidden", "true");
  }
}

/** Pantalla completa: bloquea clics hasta terminar el guardado masivo de notas. */
function mostrarBloqueoGuardandoCalificaciones(mensaje) {
  mostrarOverlayEsperaApp_(
    mensaje ||
      "Guardando calificaciones… Espere un momento; no cierre la ventana."
  );
}

function ocultarBloqueoGuardandoCalificaciones() {
  ocultarOverlayEsperaApp_();
}

function toast(message, tipo = "ok") {
  const host = document.getElementById("toast-host");
  if (!host) return;
  const div = document.createElement("div");
  div.className = `toast toast--${tipo === "err" ? "err" : "ok"}`;
  div.textContent = message;
  host.appendChild(div);
  const ms = tipo === "err" ? 5200 : 3400;
  setTimeout(() => {
    div.style.opacity = "0";
    div.style.transition = "opacity 0.25s ease";
    setTimeout(() => div.remove(), 280);
  }, ms);
}

function actualizarVisibilidadCargaRapida() {
  const el = document.getElementById("bulk-area");
  const inaBlock = document.getElementById("block-inasistencias");
  const rfBlock = document.getElementById("block-resumen-faltas");
  const raBlock = document.getElementById("block-resumen-area");
  const tiene = Array.isArray(state.cacheEstudiantes) && state.cacheEstudiantes.length > 0;

  if (el) {
    const showBulk =
      tiene && state.appModo === "calificaciones" && !esModoSoloConsulta();
    el.classList.toggle("is-hidden", !showBulk);
    el.setAttribute("aria-hidden", showBulk ? "false" : "true");
  }

  if (inaBlock) {
    const showIna = tiene && state.appModo === "inasistencias";
    inaBlock.classList.toggle("is-hidden", !showIna);
    inaBlock.hidden = !showIna;
    if (showIna) {
      void rellenarSelectFechasConListaInasistencias();
      void refrescarTablaInasistencias();
    }
  }

  if (rfBlock) {
    const showRf = state.appModo === "resumen-faltas";
    rfBlock.classList.toggle("is-hidden", !showRf);
    rfBlock.hidden = !showRf;
    if (showRf) {
      void cargarResumenFaltasPeriodo();
    }
  }

  if (raBlock) {
    const showRa = tiene && state.appModo === "resumen-area" && !esEstudiante();
    raBlock.classList.toggle("is-hidden", !showRa);
    raBlock.hidden = !showRa;
    if (showRa) {
      void cargarResumenAreaGrupoDocente_();
    }
  }

  const mgBlock = document.getElementById("block-mosaico-grupo");
  if (mgBlock) {
    const showMg = state.appModo === "mosaico-grupo" && !esEstudiante();
    mgBlock.classList.toggle("is-hidden", !showMg);
    mgBlock.hidden = !showMg;
    if (showMg) {
      void cargarMosaicoGrupoDocente_();
    }
  }

  actualizarBannerContextoInasistencias_();
}

/** Al cambiar de grupo en la barra: limpiar listado hasta el próximo «Cargar grupo» (o recarga automática del estudiante). */
function resetVistaGrupoSinCargar() {
  state.cacheEstudiantes = [];
  state.cacheNotasGrupo = [];
  state.cacheTitulosGrupo = [];
  const lista = document.getElementById("lista");
  const wrap = document.getElementById("grupo-board-wrap");
  if (lista) {
    lista.classList.remove("is-hidden");
    lista.innerHTML = esEstudiante()
      ? `<span class="state-msg state-msg--loading"><span class="spinner" aria-hidden="true"></span> Actualizando tu información…</span>`
      : `Selecciona grupo y pulsa “Cargar grupo”.`;
  }
  if (wrap) wrap.classList.add("is-hidden");
  state.filtroBusqueda = "";
  state.cancelDebouncedFiltroBusqueda();
  const buscar = document.getElementById("buscar");
  if (buscar) buscar.value = "";
  state.cacheInasistenciasDia = null;
  state.cacheResumenFaltas = null;
  state.cacheResumenArea = null;
  state.cacheMosaicoGrupo = null;
  const inaBody = document.getElementById("ina-body");
  if (inaBody) {
    inaBody.innerHTML = "";
  }
  const rfBody = document.getElementById("rf-body");
  if (rfBody) {
    rfBody.innerHTML = "";
  }
  const rfMeta = document.getElementById("rf-meta-lista");
  if (rfMeta) rfMeta.textContent = "";
  const raBody = document.getElementById("ra-body");
  if (raBody) raBody.innerHTML = "";
  const raFormula = document.getElementById("ra-formula");
  if (raFormula) raFormula.textContent = "";
  const raTitulo = document.getElementById("ra-titulo-area");
  if (raTitulo) raTitulo.textContent = "—";
  const raHead = document.getElementById("ra-head");
  if (raHead) raHead.innerHTML = "";
  const inaPick = document.getElementById("ina-fechas-con-lista");
  if (inaPick) {
    inaPick.innerHTML = '<option value="">Días con lista — cargue el grupo primero</option>';
  }
  const mgHost = document.getElementById("mg-host");
  if (mgHost) {
    mgHost.innerHTML = `<div class="state-msg">Elija el grupo arriba y abra la pestaña «Mosaico».</div>`;
  }
  const mgTitulo = document.getElementById("mg-titulo-grupo");
  if (mgTitulo) mgTitulo.textContent = "—";
  const mgMeta = document.getElementById("mg-meta");
  if (mgMeta) mgMeta.textContent = "";
  renderCargaRapida();
  actualizarVisibilidadCargaRapida();
}

function llenarSelect(id, opciones, valorGuardado) {
  const sel = document.getElementById(id);
  if (!sel) return;
  sel.innerHTML = "";
  opciones.forEach((op) => {
    const o = document.createElement("option");
    o.value = op;
    o.textContent = op;
    sel.appendChild(o);
  });
  if (valorGuardado && opciones.includes(valorGuardado)) {
    sel.value = valorGuardado;
  }
}

/** Un único ítem deshabilitado (no es un curso real del Sheet). */
function llenarSelectGrupoPlaceholder(mensaje) {
  const sel = document.getElementById("grupo");
  if (!sel) return;
  sel.innerHTML = "";
  const o = document.createElement("option");
  o.value = "";
  o.textContent = mensaje;
  o.disabled = true;
  sel.appendChild(o);
}

async function cargarAsignacionesDocenteCache() {
  if (getAuthTipo() !== "docente") {
    state.cacheDocenteAsignaciones = null;
    state.cacheDocenteNombre = "";
    return;
  }
  try {
    const d = await apiGet({ accion: "miContexto" });
    state.cacheDocenteAsignaciones = Array.isArray(d.asignaciones) ? d.asignaciones : [];
    state.cacheDocenteNombre = d && d.nombre != null ? String(d.nombre).trim() : "";
    let persistNom = false;
    try {
      persistNom = !!localStorage.getItem(STORAGE.docentePersist);
    } catch (_) {
      persistNom = false;
    }
    setStoredDocenteNombre(state.cacheDocenteNombre, persistNom);
  } catch (_) {
    state.cacheDocenteAsignaciones = [];
    state.cacheDocenteNombre = "";
  }
}

/** Docente: materias permitidas para el grupo elegido. Coordinador: lista fija MATERIAS. */
function aplicarMateriasSegunGrupoDesdeCache() {
  const sel = document.getElementById("materia");
  if (!sel) return;

  const actual = getMateria();
  const g = getGrupo();

  if (esEstudiante()) {
    llenarSelectMateriaEstudiante(leerMateriaExpedienteEstudiante());
    guardarContexto();
    return;
  }

  if (getAuthTipo() !== "docente") {
    const preferido = MATERIAS.includes(actual) ? actual : MATERIAS[0] || "";
    llenarSelect("materia", [...MATERIAS], preferido);
    guardarContexto();
    return;
  }

  if (!Array.isArray(state.cacheDocenteAsignaciones)) {
    llenarSelect("materia", ["… cargando materias …"], actual || "");
    return;
  }

  const gNorm = normalizarGrupoCurso(g);
  const mats = [];
  const seen = new Set();
  state.cacheDocenteAsignaciones.forEach((row) => {
    const rg = normalizarGrupoCurso(row.grupo || "");
    if (rg !== gNorm) return;
    const m = String(row.materia || "").trim();
    if (!m) return;
    const mk = normalizarTextoSimple(m);
    if (seen.has(mk)) return;
    seen.add(mk);
    mats.push(m);
  });
  mats.sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base", numeric: true }));

  if (!mats.length) {
    llenarSelect("materia", ["— sin materia para este grupo —"], "");
    return;
  }
  const actualNorm = normalizarTextoSimple(actual);
  const coincidencia = mats.find((x) => normalizarTextoSimple(x) === actualNorm);
  const preferido = coincidencia || mats[0];
  llenarSelect("materia", mats, preferido);
  guardarContexto();
}

function actualizarBadgeSesion() {
  const el = document.getElementById("auth-session-badge");
  if (!el) return;
  if (getAuthTipo() === "docente") {
    const cod = getStoredDocenteId();
    const nom = (state.cacheDocenteNombre || getStoredDocenteNombre() || "").trim();
    let label = "Docente";
    if (nom && cod) label = `${nom} · ${cod}`;
    else if (nom) label = nom;
    else if (cod) label = `Docente: ${cod}`;
    el.textContent = label;
    el.title = nom && cod ? `${nom} (${cod})` : cod ? String(cod) : "";
    el.hidden = false;
  } else if (getAuthTipo() === "estudiante") {
    const nom = String(state.estudianteNombreSesion || "").trim();
    el.textContent = nom ? `Estudiante · ${nom}` : "Estudiante";
    el.title = nom ? nom : "";
    el.hidden = false;
  } else {
    el.textContent = "Coordinación";
    el.title = "";
    el.hidden = false;
  }
}

function aplicarVisibilidadTrasladoCurso() {
  const card = document.querySelector(".card--reubicacion");
  if (!card) return;
  /** Solo coordinación: traslado no aplica a estudiante ni a docente en pantalla. */
  const soloCoord = getAuthTipo() === "coordinador";
  card.classList.toggle("is-hidden", !soloCoord);
  card.hidden = !soloCoord;
}

const LABEL_FICHA_TAB_ESTADO_STAFF = "Estado y traslado";
const LABEL_FICHA_TAB_ESTADO_ESTUDIANTE = "Estado de matrícula";

function actualizarEtiquetaFichaTabEstado() {
  const tab = document.getElementById("ficha-tab-estado");
  if (!tab) return;
  tab.textContent = esEstudiante() ? LABEL_FICHA_TAB_ESTADO_ESTUDIANTE : LABEL_FICHA_TAB_ESTADO_STAFF;
}

/**
 * Docente: solo lectura del estado + historial.
 * Coordinación: formulario + clave de administración (cambio único en hoja Estudiantes).
 */
function aplicarVisibilidadEstadoMatriculaPanel() {
  const card = document.getElementById("card-estado-matricula");
  if (!card) return;

  card.classList.remove("is-hidden");
  card.hidden = false;

  const esDoc = getAuthTipo() === "docente";
  const esEst = esEstudiante();
  const soloLecturaEst = esEst;
  const formEd = document.getElementById("estado-form-editable");
  const docVista = document.getElementById("estado-docente-solo-vista");
  const hintC = document.getElementById("estado-hint-coord");
  const hintD = document.getElementById("estado-hint-docente");
  const histRow = document.querySelector(".estado-actions--historial");
  const tx = document.getElementById("estado-valor-texto");
  if (tx && state.estudianteActual) {
    tx.textContent = state.estudianteActual.estadoMatricula || "—";
  }
  if (formEd) {
    formEd.classList.toggle("is-hidden", esDoc || soloLecturaEst);
    formEd.hidden = esDoc || soloLecturaEst;
  }
  if (docVista) {
    docVista.classList.toggle("is-hidden", !(esDoc || soloLecturaEst));
    docVista.hidden = !(esDoc || soloLecturaEst);
  }
  if (hintC) {
    hintC.classList.toggle("is-hidden", esDoc || soloLecturaEst);
    hintC.hidden = esDoc || soloLecturaEst;
  }
  if (hintD) {
    hintD.classList.toggle("is-hidden", !esDoc || soloLecturaEst);
    hintD.hidden = !esDoc || soloLecturaEst;
  }
  if (histRow) {
    histRow.classList.toggle("is-hidden", false);
    histRow.hidden = false;
  }
}

/** Grupos únicos según columna C de la hoja Estudiantes (GET accion=grupos). */
async function refrescarGruposSelect() {
  const sel = document.getElementById("grupo");
  if (!sel) return;

  const actual = getGrupo();
  let lista = [];

  try {
    const data = await apiGet({ accion: "grupos" });
    if (Array.isArray(data)) {
      lista = data.map((g) => String(g == null ? "" : g).trim()).filter(Boolean);
    }
  } catch (_) {
    lista = [];
  }

  if (!lista.length) {
    lista = [...GRUPOS_FALLBACK];
  }

  if (!lista.length) {
    llenarSelectGrupoPlaceholder(
      "Sin cursos disponibles (revise la hoja Grupos/Estudiantes o la conexión)"
    );
    aplicarMateriasSegunGrupoDesdeCache();
    guardarContexto();
    actualizarBadgeSesion();
    aplicarVisibilidadTrasladoCurso();
    if (state.estudianteActual) {
      poblarSelectTrasladoDestino();
    }
    return;
  }

  lista.sort((a, b) =>
    a.localeCompare(b, "es", { sensitivity: "base", numeric: true })
  );

  const preferido = coincidenciaGrupoEnOpciones(actual, lista) || lista[0] || "";
  llenarSelect("grupo", lista, preferido);
  aplicarMateriasSegunGrupoDesdeCache();
  guardarContexto();
  actualizarBadgeSesion();
  aplicarVisibilidadTrasladoCurso();
  if (state.estudianteActual) {
    poblarSelectTrasladoDestino();
  }
}

/** Opciones de curso de destino (excluye el curso actual del estudiante). */
function poblarSelectTrasladoDestino() {
  const sel = document.getElementById("traslado-grupo-destino");
  const ref = document.getElementById("grupo");
  if (!sel || !ref) return;

  const cursoAlumno = state.estudianteActual && state.estudianteActual.grupo ? state.estudianteActual.grupo : getGrupo();
  const actualNorm = normalizarTextoSimple(cursoAlumno);

  sel.innerHTML = "";
  let n = 0;
  Array.from(ref.options).forEach((opt) => {
    const v = String(opt.value || "").trim();
    if (!v) return;
    if (normalizarTextoSimple(v) === actualNorm) return;
    const o = document.createElement("option");
    o.value = v;
    o.textContent = opt.textContent || v;
    sel.appendChild(o);
    n++;
  });

  if (n === 0) {
    const o = document.createElement("option");
    o.value = "";
    o.textContent = "(No hay otros cursos en el listado)";
    sel.appendChild(o);
  }
}

// =====================
// INASISTENCIAS (vista grupo + detalle)
// =====================
function isoFechaLocalHoy() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function formatoFechaLargaEs(iso) {
  if (!iso) return "";
  try {
    const d = new Date(`${iso}T12:00:00`);
    return d.toLocaleDateString("es-CO", {
      weekday: "long",
      day: "numeric",
      month: "long",
      year: "numeric"
    });
  } catch (_) {
    return "";
  }
}

/** Encabezado corto para columnas del resumen de faltas (día / mes). */
function formatoFechaCortaRf(iso) {
  if (!iso) return "";
  try {
    const d = new Date(`${iso}T12:00:00`);
    if (Number.isNaN(d.getTime())) return iso;
    return d.toLocaleDateString("es-CO", { day: "numeric", month: "short" });
  } catch (_) {
    return "";
  }
}

function rfColspanTabla() {
  const d = state.cacheResumenFaltas;
  const n = d && Array.isArray(d.fechasColumnas) ? d.fechasColumnas.length : 0;
  return 4 + n;
}

function aplicarAppModoUI() {
  /** Estudiante: solo módulo calificaciones (expediente), sin lista diaria ni resumen por curso. */
  if (esEstudiante() && state.appModo !== "calificaciones") {
    state.appModo = "calificaciones";
  }
  /** Coordinación solo consulta: no usa el registro diario de lista (eso es docente). */
  if (esSoloLecturaCoordinacion() && state.appModo === "inasistencias") {
    state.appModo = "resumen-faltas";
  }

  const cal = document.getElementById("block-calificaciones");
  const ina = document.getElementById("block-inasistencias");
  const res = document.getElementById("block-resumen-faltas");
  const ra = document.getElementById("block-resumen-area");
  const mg = document.getElementById("block-mosaico-grupo");
  const btnCal = document.getElementById("modo-calificaciones");
  const btnIna = document.getElementById("modo-inasistencias");
  const btnRf = document.getElementById("modo-resumen-faltas");
  const btnRa = document.getElementById("modo-resumen-area");
  const btnMg = document.getElementById("modo-mosaico-grupo");

  const showCal = state.appModo === "calificaciones";
  const showIna = state.appModo === "inasistencias";
  const showRf = state.appModo === "resumen-faltas";
  const showRa = state.appModo === "resumen-area";
  const showMg = state.appModo === "mosaico-grupo";

  if (cal) {
    cal.hidden = !showCal;
    cal.classList.toggle("is-hidden", !showCal);
  }
  if (ina) {
    ina.hidden = !showIna;
    ina.classList.toggle("is-hidden", !showIna);
  }
  if (res) {
    res.hidden = !showRf;
    res.classList.toggle("is-hidden", !showRf);
  }
  if (ra) {
    ra.hidden = !showRa;
    ra.classList.toggle("is-hidden", !showRa);
  }
  if (mg) {
    mg.hidden = !showMg;
    mg.classList.toggle("is-hidden", !showMg);
  }

  if (btnCal) {
    btnCal.classList.toggle("is-active", showCal);
    btnCal.setAttribute("aria-selected", showCal ? "true" : "false");
  }
  if (btnIna) {
    const ocultarTabIna = esSoloLecturaCoordinacion() || esEstudiante();
    btnIna.classList.toggle("is-hidden", ocultarTabIna);
    btnIna.hidden = ocultarTabIna;
    btnIna.setAttribute("aria-hidden", ocultarTabIna ? "true" : "false");
    if (!ocultarTabIna) {
      btnIna.classList.toggle("is-active", showIna);
      btnIna.setAttribute("aria-selected", showIna ? "true" : "false");
    }
  }
  if (btnRf) {
    const ocultarTabRf = esEstudiante();
    btnRf.classList.toggle("is-hidden", ocultarTabRf);
    btnRf.hidden = ocultarTabRf;
    if (!ocultarTabRf) {
      btnRf.classList.toggle("is-active", showRf);
      btnRf.setAttribute("aria-selected", showRf ? "true" : "false");
    }
  }
  if (btnRa) {
    const ocultarTabRa = esEstudiante();
    btnRa.classList.toggle("is-hidden", ocultarTabRa);
    btnRa.hidden = ocultarTabRa;
    if (!ocultarTabRa) {
      btnRa.classList.toggle("is-active", showRa);
      btnRa.setAttribute("aria-selected", showRa ? "true" : "false");
    }
  }
  if (btnMg) {
    const ocultarTabMg = esEstudiante();
    btnMg.classList.toggle("is-hidden", ocultarTabMg);
    btnMg.hidden = ocultarTabMg;
    if (!ocultarTabMg) {
      btnMg.classList.toggle("is-active", showMg);
      btnMg.setAttribute("aria-selected", showMg ? "true" : "false");
    }
  }

  try {
    localStorage.setItem(STORAGE.modo, state.appModo);
  } catch (_) {
    /* ignore */
  }

  actualizarVisibilidadCargaRapida();
  aplicarModoSoloLecturaCoordinacionUI();

  if (showCal && esEstudiante() && state.cacheEstudiantes.length) {
    void cargarBoletinEstudiante();
  }
}

function renderInasistenciaSegHtml(falta) {
  const ok = Number(falta) === 0;
  return `<div class="ina-seg" role="group" aria-label="Asistencia ese día">
    <button type="button" class="ina-seg__btn ina-seg__btn--ok${ok ? " is-pressed" : ""}" data-ina-set="0">Asistió</button>
    <button type="button" class="ina-seg__btn ina-seg__btn--no${!ok ? " is-pressed" : ""}" data-ina-set="1">Faltó</button>
  </div>`;
}

function sincronizarSelectFechasConListaDesdeInput() {
  const sel = document.getElementById("ina-fechas-con-lista");
  const inp = document.getElementById("ina-fecha");
  if (!sel || !inp) return;
  const v = inp.value;
  if (!v) {
    sel.value = "";
    return;
  }
  const has = Array.from(sel.options).some((o) => o.value === v);
  sel.value = has ? v : "";
}

/**
 * Carga del servidor las fechas que ya tienen registro (lista) para grupo + materia.
 */
async function rellenarSelectFechasConListaInasistencias() {
  const sel = document.getElementById("ina-fechas-con-lista");
  if (!sel) return;
  if (!state.cacheEstudiantes.length) {
    sel.innerHTML = '<option value="">Días con lista — cargue el grupo primero</option>';
    return;
  }
  sel.innerHTML = '<option value="">… cargando fechas …</option>';
  try {
    const data = await apiGet({
      accion: "inasistenciasFechasGrupo",
      grupo: getGrupo(),
      materia: getMateria()
    });
    const fechas = Array.isArray(data) ? data : [];
    sel.innerHTML = '<option value="">— Ir a un día ya guardado —</option>';
    fechas.forEach((iso) => {
      const o = document.createElement("option");
      o.value = iso;
      const larga = formatoFechaLargaEs(iso);
      o.textContent = larga ? `${iso} · ${larga}` : iso;
      sel.appendChild(o);
    });
    sincronizarSelectFechasConListaDesdeInput();
  } catch (_) {
    sel.innerHTML = '<option value="">No se pudo cargar el listado de fechas</option>';
  }
}

function renderInasistenciasBodyFromCache() {
  const tbody = document.getElementById("ina-body");
  if (!tbody || !state.cacheInasistenciasDia) return;

  const soloActivos = document.getElementById("ina-solo-activos")?.checked !== false;
  const q = normalizarTextoSimple(state.filtroBusqueda);

  const list = state.cacheInasistenciasDia.filter((row) => {
    if (soloActivos && !estadoEsActivo(row.estadoMatricula)) return false;
    if (q && !normalizarTextoSimple(row.nombre).includes(q)) return false;
    return true;
  });

  if (!list.length) {
    tbody.innerHTML = `<tr><td colspan="4"><div class="state-msg">Ningún estudiante coincide con el filtro. Ajuste la búsqueda o desactive «Solo activos».</div></td></tr>`;
    return;
  }

  let idx = 0;
  tbody.innerHTML = list
    .map((row) => {
      idx++;
      const falta = Number(row.falta) === 1 ? 1 : 0;
      const est = escapeHtml(row.estadoMatricula || "—");
      const cl = clasesEstadoMatriculaChip(row.estadoMatricula);
      return `<tr data-estudiante-id="${escapeHtml(String(row.id))}" data-falta="${falta}">
        <td>${idx}</td>
        <td>
          <strong>${escapeHtml(row.nombre || "")}</strong>
          <button type="button" class="btn btn--ghost btn--sm ina-ficha-btn" data-estudiante-id="${escapeHtml(String(row.id))}">Ficha</button>
        </td>
        <td class="col-estado"><span class="${cl}">${est}</span></td>
        <td class="ina-cell-toggle">${renderInasistenciaSegHtml(falta)}</td>
      </tr>`;
    })
    .join("");
}

async function refrescarTablaInasistencias() {
  const tbody = document.getElementById("ina-body");
  if (!tbody) return;
  const fechaEl = document.getElementById("ina-fecha");
  const fecha = fechaEl ? fechaEl.value : "";
  if (!fecha) return;

  if (!state.cacheEstudiantes.length) {
    tbody.innerHTML = `<tr><td colspan="4"><div class="state-msg">Cargue el grupo con el botón de arriba.</div></td></tr>`;
    state.cacheInasistenciasDia = null;
    return;
  }

  const seq = ++state.inasistenciasDiaFetchSeq;
  tbody.innerHTML = `<tr><td colspan="4"><div class="state-msg state-msg--loading"><span class="spinner"></span> Cargando día…</div></td></tr>`;

  try {
    const data = await apiGet({
      accion: "inasistenciasDia",
      grupo: getGrupo(),
      materia: getMateria(),
      fecha
    });
    if (seq !== state.inasistenciasDiaFetchSeq) return;
    state.cacheInasistenciasDia = Array.isArray(data) ? data : [];
    renderInasistenciasBodyFromCache();
  } catch (err) {
    if (seq !== state.inasistenciasDiaFetchSeq) return;
    state.cacheInasistenciasDia = null;
    tbody.innerHTML = `<tr><td colspan="4"><div class="state-msg state-msg--warn">${escapeHtml(err.message)}</div></td></tr>`;
  }
}

function actualizarEtiquetaFechaInasistencias() {
  const label = document.getElementById("ina-fecha-larga");
  const input = document.getElementById("ina-fecha");
  if (!label || !input) return;
  label.textContent = formatoFechaLargaEs(input.value);
}

/** Recuerda en pantalla curso + materia de la barra (donde se guarda la lista). */
function actualizarBannerContextoInasistencias_() {
  const el = document.getElementById("ina-contexto-grupo-materia");
  if (!el) return;
  const g = String(getGrupo() || "").trim();
  const m = String(getMateria() || "").trim();
  if (!g && !m) {
    el.textContent = "Seleccione curso y materia antes de marcar asistencia o faltas.";
  } else if (!m) {
    el.textContent = `Curso «${g}». Seleccione también la materia.`;
  } else if (!g) {
    el.textContent = `Materia «${m}». Seleccione también el curso.`;
  } else {
    el.textContent = `Lista para la materia «${m}» · curso «${g}».`;
  }
}

function usarFechaHoyInasistencias() {
  const input = document.getElementById("ina-fecha");
  if (!input) return;
  input.value = isoFechaLocalHoy();
  actualizarEtiquetaFechaInasistencias();
  sincronizarSelectFechasConListaDesdeInput();
  if (state.appModo === "inasistencias" && state.cacheEstudiantes.length) {
    void refrescarTablaInasistencias();
  }
}

/** Marca asistencia/falta para todos los estudiantes del día cargado (misma fecha en servidor al guardar). */
function aplicarInasistenciaMasivaGrupo(faltaVal) {
  if (esModoSoloConsulta()) {
    toast(
      esEstudiante()
        ? "El acceso de estudiante es solo consulta."
        : "El acceso de coordinación es solo consulta.",
      "err"
    );
    return;
  }
  if (!state.cacheInasistenciasDia || !state.cacheInasistenciasDia.length) {
    toast("Primero cargue el grupo y espere a que cargue la tabla del día.", "err");
    return;
  }
  const v = faltaVal === 1 ? 1 : 0;
  state.cacheInasistenciasDia.forEach((row) => {
    row.falta = v;
  });
  renderInasistenciasBodyFromCache();
}

async function guardarInasistenciasMasivoFront() {
  if (esModoSoloConsulta()) {
    toast(
      esEstudiante()
        ? "El acceso de estudiante es solo consulta."
        : "El acceso de coordinación es solo consulta.",
      "err"
    );
    return;
  }
  if (!state.cacheInasistenciasDia || !state.cacheInasistenciasDia.length) {
    toast("Primero cargue el día (espere la tabla o cambie la fecha).", "err");
    return;
  }
  const fecha = document.getElementById("ina-fecha")?.value;
  if (!fecha) {
    toast("Elija la fecha de la clase.", "err");
    return;
  }

  const mapDom = {};
  document.querySelectorAll("#ina-body tr[data-estudiante-id]").forEach((tr) => {
    mapDom[tr.getAttribute("data-estudiante-id")] = tr.getAttribute("data-falta") === "1" ? 1 : 0;
  });

  const items = state.cacheInasistenciasDia.map((row) => ({
    id: row.id,
    falta:
      mapDom[String(row.id)] !== undefined ? mapDom[String(row.id)] : Number(row.falta) === 1
        ? 1
        : 0
  }));

  const btn = document.getElementById("ina-guardar");
  if (btn) btn.disabled = true;
  mostrarOverlayEsperaApp_(
    "Guardando asistencia del día… Espere; no cierre la ventana. Evite pulsar otras acciones del sistema hasta que termine."
  );
  try {
    await apiPost({
      accion: "inasistenciasGuardar",
      grupo: getGrupo(),
      materia: getMateria(),
      fecha,
      items
    });
    toast("Informacion guardada: asistencia del dia registrada correctamente.", "ok");
    await refrescarTablaInasistencias();
    void rellenarSelectFechasConListaInasistencias();
    if (state.appModo === "resumen-faltas") {
      void cargarResumenFaltasPeriodo();
    }
    if (state.estudianteActual) await abrirEstudiante(state.estudianteActual);
  } catch (e) {
    toast(e.message || "No se pudo guardar.", "err");
  } finally {
    ocultarOverlayEsperaApp_();
    if (btn) btn.disabled = false;
  }
}

function renderInasistenciasEstudiantePanel(items) {
  const el = document.getElementById("ina-estudiante-lista");
  if (!el) return;
  if (!items || !items.length) {
    el.innerHTML = `<p class="empty-hint">No hay días registrados aún para esta materia.</p>`;
    return;
  }
  const soloLect = esModoSoloConsulta();
  el.innerHTML = items
    .map((it) => {
      const larga = formatoFechaLargaEs(it.fecha);
      const badge =
        Number(it.falta) === 1
          ? `<span class="ina-badge ina-badge--falta">Faltó</span>`
          : `<span class="ina-badge ina-badge--ok">Asistió</span>`;
      const btnCorr = soloLect
        ? ""
        : `<button type="button" class="btn btn--secondary btn--sm ina-hist-card__btn"
          data-ina-fecha="${escapeHtml(it.fecha)}" data-ina-falta="${Number(it.falta) === 1 ? 1 : 0}">Corregir</button>`;
      return `<article class="ina-hist-card">
        <div class="ina-hist-card__main">
          <time class="ina-hist-card__fecha" datetime="${escapeHtml(it.fecha)}">${escapeHtml(larga)}</time>
          ${badge}
        </div>
        ${btnCorr}
      </article>`;
    })
    .join("");

  el.querySelectorAll(".ina-hist-card__btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      if (!state.estudianteActual) return;
      const f = btn.getAttribute("data-ina-fecha");
      const fal = btn.getAttribute("data-ina-falta") === "1" ? 1 : 0;
      abrirCorregirInasistenciaModal(state.estudianteActual, f, fal);
    });
  });
}

function abrirCorregirInasistenciaModal(est, fechaIso, faltaActual) {
  if (esModoSoloConsulta()) {
    toast(
      esEstudiante()
        ? "El acceso de estudiante es solo consulta."
        : "El acceso de coordinación es solo consulta.",
      "err"
    );
    return;
  }
  const larga = formatoFechaLargaEs(fechaIso);
  let sel = faltaActual === 1 ? 1 : 0;

  const htmlOk = sel === 0 ? " is-pressed" : "";
  const htmlNo = sel === 1 ? " is-pressed" : "";

  abrirModal("Corregir asistencia", `
    <p class="ina-modal-fecha">${escapeHtml(larga)}</p>
    <p class="card__hint card__hint--compact">Materia: ${escapeHtml(getMateria())}</p>
    <div class="ina-seg ina-seg--modal" role="group" aria-label="Corregir registro">
      <button type="button" class="ina-seg__btn ina-seg__btn--ok${htmlOk}" id="ina-m-ok" data-v="0">Asistió</button>
      <button type="button" class="ina-seg__btn ina-seg__btn--no${htmlNo}" id="ina-m-no" data-v="1">Faltó</button>
    </div>
    <button type="button" class="btn btn--primary btn--block" id="ina-m-guardar">Guardar cambio</button>
  `);

  const sync = (v) => {
    sel = v;
    const b0 = document.getElementById("ina-m-ok");
    const b1 = document.getElementById("ina-m-no");
    if (b0) b0.classList.toggle("is-pressed", v === 0);
    if (b1) b1.classList.toggle("is-pressed", v === 1);
  };

  document.getElementById("ina-m-ok")?.addEventListener("click", () => sync(0));
  document.getElementById("ina-m-no")?.addEventListener("click", () => sync(1));

  document.getElementById("ina-m-guardar")?.addEventListener("click", async () => {
    const guardarBtn = document.getElementById("ina-m-guardar");
    if (guardarBtn) guardarBtn.disabled = true;
    mostrarOverlayEsperaApp_(
      "Guardando asistencia… Espere un momento; no cierre la ventana."
    );
    try {
      await apiPost({
        accion: "inasistenciasGuardar",
        grupo: est.grupo || getGrupo(),
        materia: getMateria(),
        fecha: fechaIso,
        items: [{ id: est.id, falta: sel }]
      });
      cerrarModal();
      toast("Informacion guardada: registro de asistencia actualizado.", "ok");
      await abrirEstudiante(est);
    } catch (e) {
      toast(e.message || "No se pudo guardar.", "err");
    } finally {
      ocultarOverlayEsperaApp_();
      if (guardarBtn) guardarBtn.disabled = false;
    }
  });
}

/**
 * Periodo por defecto: el que en PeriodosAcademicos contiene la fecha de hoy (desde/hasta en ISO).
 * Si hoy no cae en ninguno, se usa el primero de la lista.
 */
function periodoPredeterminadoDesdeLista(periodos) {
  if (!Array.isArray(periodos) || !periodos.length) return "1";
  const hoy = isoFechaLocalHoy();
  for (const p of periodos) {
    if (p.desde && p.hasta && hoy >= p.desde && hoy <= p.hasta) {
      return String(p.id);
    }
  }
  return String(periodos[0].id);
}

/** Texto corto del periodo académico elegido (portal estudiante). */
function etiquetaPeriodoActivoEstudiante_() {
  const sel = document.getElementById("est-periodo-cal");
  let id = sel && sel.value ? String(sel.value).trim() : "";
  if (!id) {
    try {
      id = String(sessionStorage.getItem(STORAGE.estudiantePeriodoCal) || "").trim();
    } catch (_) {
      id = "";
    }
  }
  if (!id) return "";
  const arr = state.cachePeriodosAcademicos;
  if (Array.isArray(arr) && arr.length) {
    const p = arr.find((x) => String(x.id) === String(id));
    if (p) {
      const et = String(p.etiqueta || "").trim();
      return et || `Periodo ${id}`;
    }
  }
  return `Periodo ${id}`;
}

function actualizarTituloNotasGeneralesEstudiante_() {
  const chip = document.getElementById("est-boletin-periodo-chip");
  if (!chip || !esEstudiante()) return;
  const lab = etiquetaPeriodoActivoEstudiante_();
  chip.textContent = lab ? ` · ${lab}` : "";
}

/** Evita que #est-periodo-cal se vea vacío mientras llega periodosAcademicos (la API es async). */
function prepararSelectPeriodoEstudianteCargando() {
  const sel = document.getElementById("est-periodo-cal");
  if (!sel || sel.options.length > 0) return;
  const o = document.createElement("option");
  o.value = "";
  o.disabled = true;
  o.selected = true;
  o.textContent = "Cargando periodos…";
  sel.appendChild(o);
}

function poblarOpcionesPeriodoEnSelect_(sel) {
  if (!sel) return;
  sel.innerHTML = "";
  state.cachePeriodosAcademicos.forEach((p) => {
    const o = document.createElement("option");
    o.value = String(p.id);
    const rango = p.desde && p.hasta ? `${p.desde} → ${p.hasta}` : "";
    o.textContent = `${p.etiqueta || `Periodo ${p.id}`}${rango ? ` · ${rango}` : ""}`;
    sel.appendChild(o);
  });
  if (!state.cachePeriodosAcademicos.length) {
    sel.innerHTML = '<option value="1">Periodo 1</option>';
  }
}

async function refrescarPeriodosAcademicosSelect() {
  const selRf = document.getElementById("rf-periodo");
  const selEst = document.getElementById("est-periodo-cal");
  const selDoc = document.getElementById("doc-periodo-cal");
  if (!selRf && !selEst && !selDoc) return;
  try {
    const data = await apiGet({ accion: "periodosAcademicos" });
    state.cachePeriodosAcademicos = Array.isArray(data) ? data : [];
    poblarOpcionesPeriodoEnSelect_(selRf);
    poblarOpcionesPeriodoEnSelect_(selEst);
    poblarOpcionesPeriodoEnSelect_(selDoc);
    const def = periodoPredeterminadoDesdeLista(state.cachePeriodosAcademicos);
    if (selRf) {
      selRf.value = def;
      actualizarBannerPeriodoRf();
    }
    if (selEst) {
      let guardado = "";
      try {
        guardado = String(sessionStorage.getItem(STORAGE.estudiantePeriodoCal) || "").trim();
      } catch (_) {
        /* ignore */
      }
      if (guardado && [...selEst.options].some((o) => o.value === guardado)) {
        selEst.value = guardado;
      } else {
        selEst.value = def;
      }
      try {
        sessionStorage.setItem(STORAGE.estudiantePeriodoCal, selEst.value);
      } catch (_) {
        /* ignore */
      }
    }
    if (selDoc) {
      let guardadoDoc = "";
      try {
        guardadoDoc = String(localStorage.getItem(STORAGE.docentePeriodoCal) || "").trim();
      } catch (_) {
        /* ignore */
      }
      if (guardadoDoc && [...selDoc.options].some((o) => o.value === guardadoDoc)) {
        selDoc.value = guardadoDoc;
      } else {
        selDoc.value = def;
      }
      try {
        localStorage.setItem(STORAGE.docentePeriodoCal, selDoc.value);
      } catch (_) {
        /* ignore */
      }
    }
    actualizarHintRangoFechaFichaDocente_();
    if (esEstudiante()) actualizarTituloNotasGeneralesEstudiante_();
  } catch (_) {
    if (selRf) selRf.innerHTML = '<option value="1">Periodo 1</option>';
    if (selEst) selEst.innerHTML = '<option value="1">Periodo 1</option>';
    if (selDoc) selDoc.innerHTML = '<option value="1">Periodo 1</option>';
    state.cachePeriodosAcademicos = [];
    actualizarHintRangoFechaFichaDocente_();
    if (esEstudiante()) actualizarTituloNotasGeneralesEstudiante_();
  }
}

function actualizarBannerPeriodoRf() {
  const sel = document.getElementById("rf-periodo");
  const banner = document.getElementById("rf-periodo-rango");
  if (!banner) return;
  const id = sel ? sel.value : "1";
  const p = state.cachePeriodosAcademicos.find((x) => String(x.id) === String(id));
  if (!p || !p.desde || !p.hasta) {
    banner.textContent = "";
    return;
  }
  const di = formatoFechaLargaEs(p.desde);
  const df = formatoFechaLargaEs(p.hasta);
  banner.textContent = `Periodo del ${di} al ${df}`;
}

/** Rango de fechas del periodo docente (barra): evita guardar con fecha fuera de P1/P2/P3. */
function actualizarHintRangoFechaFichaDocente_() {
  const el = document.getElementById("ficha-nueva-fecha-rango-periodo");
  if (!el) return;
  const sel = document.getElementById("doc-periodo-cal");
  const id = sel ? String(sel.value || "1").trim() : "1";
  const p = state.cachePeriodosAcademicos.find((x) => String(x.id) === String(id));
  if (!p || !p.desde || !p.hasta) {
    el.textContent = "";
    el.hidden = true;
    return;
  }
  const di = formatoFechaLargaEs(p.desde);
  const df = formatoFechaLargaEs(p.hasta);
  el.textContent = `Según Periodos académicos, el periodo ${id} va del ${di} al ${df}. La fecha de la nota debe estar en ese intervalo (mismo periodo que la barra superior).`;
  el.hidden = false;
}

function renderTablaResumenFaltasFiltrada() {
  const thead = document.getElementById("rf-head");
  const tbody = document.getElementById("rf-body");
  const meta = document.getElementById("rf-meta-lista");
  const data = state.cacheResumenFaltas;
  if (!tbody) return;

  if (!data || !Array.isArray(data.estudiantes)) {
    if (thead) thead.innerHTML = "";
    tbody.innerHTML = `<tr><td colspan="${rfColspanTabla()}"><div class="state-msg">Sin datos.</div></td></tr>`;
    return;
  }

  const fechas = Array.isArray(data.fechasColumnas) ? data.fechasColumnas : [];
  const colspan = 4 + fechas.length;

  const q = normalizarTextoSimple(state.filtroBusqueda);
  let list = data.estudiantes.filter((row) => {
    if (!q) return true;
    return normalizarTextoSimple(row.nombre).includes(q);
  });

  list = [...list].sort((a, b) =>
    String(a.nombre || "").localeCompare(String(b.nombre || ""), "es", { sensitivity: "base" })
  );

  if (!list.length) {
    if (thead) thead.innerHTML = "";
    tbody.innerHTML = `<tr><td colspan="${colspan}"><div class="state-msg">Nadie coincide con la búsqueda.</div></td></tr>`;
    return;
  }

  const diasGrupo = data.diasListaGrupoEnPeriodo;
  if (meta && diasGrupo != null) {
    const cero =
      diasGrupo === 0
        ? " Si es 0, no hay listas guardadas en las fechas de este periodo para esta materia (revise el periodo, la materia o que existan registros en Inasistencias). "
        : " ";
    meta.textContent = `Días con lista en este periodo (grupo y materia): ${diasGrupo}.${cero}Deslice horizontalmente para revisar cada día con registro y las marcas de falta (F).`;
  }

  if (thead) {
    thead.innerHTML = `<tr>
      <th class="rf__num rf__sticky" scope="col">#</th>
      <th class="rf__student rf__sticky" scope="col">Estudiante</th>
      <th class="rf__estado rf__sticky" scope="col">Estado</th>
      ${fechas
        .map(
          (f) =>
            `<th class="rf__fecha" scope="col" title="${escapeHtml(formatoFechaLargaEs(f))}">${escapeHtml(formatoFechaCortaRf(f))}</th>`
        )
        .join("")}
      <th class="rf__total rf__sticky-end" scope="col">Total faltas</th>
    </tr>`;
  }

  let idx = 0;
  tbody.innerHTML = list
    .map((row) => {
      idx++;
      const est = escapeHtml(row.estadoMatricula || "—");
      const cl = clasesEstadoMatriculaChip(row.estadoMatricula);
      const sinAc = row.sinAcumuloFaltas === true;
      const falTxt = sinAc || row.totalFaltas == null ? "—" : String(row.totalFaltas);
      const title = sinAc
        ? "Sin acumulación institucional de faltas (cancelado o retirado)"
        : "";
      const faltaDia = sinAc ? null : row.faltaDia && typeof row.faltaDia === "object" ? row.faltaDia : {};
      const celdasFecha = fechas
        .map((f) => {
          if (sinAc) {
            return `<td class="rf__fecha rf__fecha--muted" title="${escapeHtml(title)}">—</td>`;
          }
          const falta = faltaDia && faltaDia[f] === true;
          const tit = falta ? `Falta · ${formatoFechaLargaEs(f)}` : `Sin falta · ${formatoFechaLargaEs(f)}`;
          const inner = falta
            ? `<span class="ina-badge ina-badge--falta" title="${escapeHtml(tit)}">F</span>`
            : `<span class="rf__celda-vacia" title="${escapeHtml(tit)}">·</span>`;
          return `<td class="rf__fecha">${inner}</td>`;
        })
        .join("");
      const btnFicha = `<button type="button" class="btn btn--ghost btn--sm ina-ficha-btn" data-estudiante-id="${escapeHtml(String(row.id))}">Ficha</button>`;
      return `<tr${sinAc ? ` class="rf-row--excl"` : ""}>
        <td class="rf__num rf__sticky">${idx}</td>
        <td class="rf__student rf__sticky">
          <strong>${escapeHtml(row.nombre || "")}</strong>
          ${btnFicha}
        </td>
        <td class="col-estado rf__estado rf__sticky"><span class="${cl}">${est}</span></td>
        ${celdasFecha}
        <td class="rf__total rf__sticky-end" title="${escapeHtml(title)}"><strong>${escapeHtml(falTxt)}</strong></td>
      </tr>`;
    })
    .join("");
}

async function cargarResumenFaltasPeriodo() {
  const tbody = document.getElementById("rf-body");
  const meta = document.getElementById("rf-meta-lista");
  if (!tbody) return;

  const periodo = document.getElementById("rf-periodo")?.value || "1";

  if (!state.cacheEstudiantes.length) {
    const rfHead = document.getElementById("rf-head");
    if (rfHead) rfHead.innerHTML = "";
    tbody.innerHTML = `<tr><td colspan="5"><div class="state-msg">Cargue el grupo con el botón de arriba.</div></td></tr>`;
    if (meta) meta.textContent = "";
    state.cacheResumenFaltas = null;
    return;
  }

  tbody.innerHTML = `<tr><td colspan="${rfColspanTabla()}"><div class="state-msg state-msg--loading"><span class="spinner"></span> Cargando resumen…</div></td></tr>`;

  try {
    const data = await apiGet({
      accion: "inasistenciasResumenPeriodo",
      grupo: getGrupo(),
      materia: getMateria(),
      periodo
    });
    state.cacheResumenFaltas = data;
    renderTablaResumenFaltasFiltrada();
  } catch (err) {
    state.cacheResumenFaltas = null;
    if (meta) meta.textContent = "";
    const rfHeadErr = document.getElementById("rf-head");
    if (rfHeadErr) rfHeadErr.innerHTML = "";
    tbody.innerHTML = `<tr><td colspan="5"><div class="state-msg state-msg--warn">${escapeHtml(err.message)}</div></td></tr>`;
  }
}

function renderTablaResumenAreaFiltrada() {
  const thead = document.getElementById("ra-head");
  const tbody = document.getElementById("ra-body");
  const data = state.cacheResumenArea;
  if (!tbody) return;

  const nMat =
    data && Array.isArray(data.titulosColumnasMateria) ? data.titulosColumnasMateria.length : 3;
  const sinTotal = Boolean(data && data.sinColumnaTotalArea);
  const colspan = raColspanTablaResumenArea(nMat, sinTotal);

  if (!data || !Array.isArray(data.filas)) {
    if (thead) thead.innerHTML = "";
    tbody.innerHTML = `<tr><td colspan="${colspan}"><div class="state-msg">Sin datos.</div></td></tr>`;
    return;
  }

  const q = normalizarTextoSimple(state.filtroBusqueda);
  let list = data.filas.filter((row) => {
    if (!q) return true;
    return normalizarTextoSimple(row.nombre).includes(q);
  });
  list = [...list].sort((a, b) =>
    String(a.nombre || "").localeCompare(String(b.nombre || ""), "es", { sensitivity: "base" })
  );

  const titulos = data.titulosColumnasMateria || [];
  const nomArea = data.areaNombre ? escapeHtml(String(data.areaNombre)) : "Área";
  if (thead) {
    const thMats = titulos
      .map((t) => `<th class="ra__mat" scope="col">${escapeHtml(String(t))}</th>`)
      .join("");
    const thTotal = sinTotal
      ? ""
      : `<th class="ra__area-col" scope="col">${nomArea} (orientativa)</th>`;
    thead.innerHTML = `<tr>
      <th class="ra__num" scope="col">#</th>
      <th scope="col">Estudiante</th>
      <th class="col-estado" scope="col">Estado</th>
      ${thMats}
      ${thTotal}
      <th scope="col">Observaciones</th>
    </tr>`;
  }

  if (!list.length) {
    tbody.innerHTML = `<tr><td colspan="${raColspanTablaResumenArea(titulos.length, sinTotal)}"><div class="state-msg">Nadie coincide con la búsqueda.</div></td></tr>`;
    return;
  }

  let idx = 0;
  tbody.innerHTML = list
    .map((row) => {
      idx++;
      const est = escapeHtml(row.estadoMatricula || "—");
      const cl = clasesEstadoMatriculaChip(row.estadoMatricula);
      const defTxt = formatearDefinitivaDisplayBoletinEstudiante_(row.definitivaArea);
      const obs = escapeHtml(row.obs || "");
      const warn = row.incompleto ? " boletin-row--warn" : "";
      const btnFicha = `<button type="button" class="btn btn--ghost btn--sm ina-ficha-btn" data-estudiante-id="${escapeHtml(String(row.id))}">Ficha</button>`;
      const comps = Array.isArray(row.componentes) ? row.componentes : [];
      const celdasMat = titulos
        .map((_, i) => {
          const c = comps[i];
          const txt = c
            ? formatearDefinitivaDisplayBoletinEstudiante_(c.definitiva)
            : "—";
          const inc = c && c.incompleto;
          return `<td class="ra__mat${inc ? " ra__mat--incompleto" : ""}"><strong>${escapeHtml(txt)}</strong></td>`;
        })
        .join("");
      const celdaTotal = sinTotal
        ? ""
        : `<td class="ra__area-col"><strong>${escapeHtml(defTxt)}</strong></td>`;
      return `<tr class="${warn.trim()}">
        <td class="ra__num">${idx}</td>
        <td><strong>${escapeHtml(row.nombre || "")}</strong> ${btnFicha}</td>
        <td class="col-estado"><span class="${cl}">${est}</span></td>
        ${celdasMat}
        ${celdaTotal}
        <td class="ra__obs">${obs}</td>
      </tr>`;
    })
    .join("");
}

async function cargarResumenAreaGrupoDocente_() {
  const thead = document.getElementById("ra-head");
  const tbody = document.getElementById("ra-body");
  const formula = document.getElementById("ra-formula");
  const tituloArea = document.getElementById("ra-titulo-area");
  if (!tbody) return;
  if (esEstudiante()) return;

  const colspanVacío = raColspanTabla(3);

  if (!state.cacheEstudiantes.length) {
    state.cacheResumenArea = null;
    if (tituloArea) tituloArea.textContent = "—";
    if (formula) formula.textContent = "";
    if (thead) thead.innerHTML = "";
    tbody.innerHTML = `<tr><td colspan="${colspanVacío}"><div class="state-msg">Cargue el grupo con el botón de arriba.</div></td></tr>`;
    return;
  }

  const area = areaResumenDocenteParaMateriaActual_();
  if (!area) {
    state.cacheResumenArea = null;
    if (tituloArea) tituloArea.textContent = "—";
    if (formula) {
      formula.textContent = `Esta materia no entra en MEG ni Humanidades. Elija Matemáticas, Geometría, Estadística, Lengua castellana o Competencia lectora.`;
    }
    if (thead) thead.innerHTML = "";
    tbody.innerHTML = `<tr><td colspan="${colspanVacío}"><div class="state-msg state-msg--warn">No hay resumen de área para esta asignatura.</div></td></tr>`;
    return;
  }

  const seq = ++resumenAreaFetchSeq;
  const materiasFetch = materiasFetchResumenAreaDocente_(area);
  const nMateriasArea = materiasApiParaArea_(area).length;
  /** Docente con solo parte del área (p. ej. Geometría y Estadística): sin columna MEG. Con las tres de MEG o las dos de Humanidades: igual que coordinación. */
  const sinColumnaTotalArea =
    getAuthTipo() === "docente" && materiasFetch.length < nMateriasArea;
  const titulosColumnasMateria = materiasFetch.map(tituloColumnaResumenArea_);
  if (tituloArea) {
    tituloArea.textContent = sinColumnaTotalArea
      ? titulosColumnasMateria.join(" · ") || area.nombre
      : area.nombre;
  }
  if (formula) {
    if (sinColumnaTotalArea) {
      formula.textContent =
        "Definitiva orientativa por asignatura (misma regla que el boletín: 70 % seguimiento, 10 % actitudinal, 20 % prueba). " +
        "No se muestra total del área porque en este curso solo figuran parte de las asignaturas del bloque en su asignación. " +
        "Periodo: barra superior.";
    } else {
      formula.textContent = `${area.formulaCorta} · Definitivas por asignatura como en el boletín. Si faltan materias del área en los datos del estudiante, el total es parcial (pesos oficiales reescalados entre las notas disponibles). Curso y periodo: barra superior.`;
    }
  }

  if (thead) thead.innerHTML = "";
  tbody.innerHTML = `<tr><td colspan="${raColspanTablaResumenArea(materiasFetch.length, sinColumnaTotalArea)}"><div class="state-msg state-msg--loading"><span class="spinner"></span> Cargando notas del área…</div></td></tr>`;

  try {
    const periodoPayload = periodoParamParaApiDocente();
    const grupo = getGrupo();
    const results = await Promise.all(
      materiasFetch.map((mat) =>
        apiGet({
          accion: "notasGrupo",
          grupo,
          materia: mat,
          ...periodoPayload
        })
      )
    );

    if (seq !== resumenAreaFetchSeq) return;

    const porMateria = materiasFetch.map((mat, i) => ({
      materia: mat,
      flat: Array.isArray(results[i]) ? results[i] : []
    }));

    const filas = state.cacheEstudiantes.map((est) => {
      const items = porMateria.map(({ materia, flat }) => {
        const metaM = calcularMetaBoletinDesdeNotasArr(notasArrayParaBoletinDesdeFilasGrupo_(flat, est.id));
        return {
          materia,
          definitiva: metaM.definitiva,
          incompleto: metaM.incompleto,
          avisosMateria: metaM.avisos
        };
      });
      if (sinColumnaTotalArea) {
        const obsParts = [];
        items.forEach((it) => {
          if (it.incompleto && it.avisosMateria && it.avisosMateria.length) {
            obsParts.push(`${it.materia}: ${it.avisosMateria.join(", ")}`);
          }
        });
        return {
          id: est.id,
          nombre: est.nombre,
          estadoMatricula: est.estadoMatricula,
          definitivaArea: "—",
          incompleto: items.some((it) => it.incompleto),
          componentes: items.map((it) => ({
            definitiva: it.definitiva,
            incompleto: it.incompleto
          })),
          obs: obsParts.join(" · ")
        };
      }
      const areaRes = calcularDefinitivaAreaUnicaDocente_(area, items);
      const obsParts = [...areaRes.avisos];
      items.forEach((it) => {
        if (it.incompleto && it.avisosMateria && it.avisosMateria.length) {
          obsParts.push(`${it.materia}: ${it.avisosMateria.join(", ")}`);
        }
      });
      return {
        id: est.id,
        nombre: est.nombre,
        estadoMatricula: est.estadoMatricula,
        definitivaArea: areaRes.definitiva,
        incompleto: areaRes.incompleto,
        componentes: items.map((it) => ({
          definitiva: it.definitiva,
          incompleto: it.incompleto
        })),
        obs: obsParts.join(" · ")
      };
    });

    state.cacheResumenArea = {
      sinColumnaTotalArea,
      areaNombre: area.nombre,
      formulaCorta: area.formulaCorta,
      titulosColumnasMateria,
      filas
    };
    renderTablaResumenAreaFiltrada();
  } catch (err) {
    if (seq !== resumenAreaFetchSeq) return;
    state.cacheResumenArea = null;
    if (formula) formula.textContent = "";
    if (thead) thead.innerHTML = "";
    tbody.innerHTML = `<tr><td colspan="${raColspanTablaResumenArea(materiasFetch.length, sinColumnaTotalArea)}"><div class="state-msg state-msg--warn">${escapeHtml(err.message)}</div></td></tr>`;
  }
}

let mosaicoGrupoFetchSeq = 0;

/** Pestaña Mosaico: imagen única por grupo desde hoja MosaicosGrupo. */
async function cargarMosaicoGrupoDocente_() {
  const host = document.getElementById("mg-host");
  const tituloEl = document.getElementById("mg-titulo-grupo");
  const metaEl = document.getElementById("mg-meta");
  if (!host) return;
  if (esEstudiante()) {
    host.innerHTML = `<div class="state-msg state-msg--warn">El mosaico es solo para coordinación y docentes.</div>`;
    return;
  }

  const grupo = String(getGrupo() || "").trim();
  if (tituloEl) tituloEl.textContent = grupo || "—";
  if (metaEl) metaEl.textContent = "";

  if (!grupo) {
    host.innerHTML = `<div class="state-msg">Elija un grupo en la barra de arriba.</div>`;
    return;
  }

  const seq = ++mosaicoGrupoFetchSeq;
  host.innerHTML = `<div class="state-msg state-msg--loading"><span class="spinner" aria-hidden="true"></span> Cargando mosaico…</div>`;

  try {
    const data = await apiGet({ accion: "mosaicoGrupo", grupo });
    if (seq !== mosaicoGrupoFetchSeq) return;

    const imagen = data && typeof data === "object" ? String(data.imagen || "").trim() : "";
    const actualizado = data && typeof data === "object" ? String(data.actualizado || "").trim() : "";
    const mensaje = data && typeof data === "object" ? String(data.mensaje || "").trim() : "";

    state.cacheMosaicoGrupo = { grupo, imagen, actualizado };

    if (!imagen) {
      const aviso = mensaje
        ? mensaje
        : `No hay mosaico configurado para «${grupo}». Agregue una fila en la hoja «MosaicosGrupo» con A=${grupo}, B=URL pública (o data URI base64).`;
      host.innerHTML = `<div class="state-msg state-msg--warn">${escapeHtml(aviso)}</div>`;
      if (metaEl) metaEl.textContent = "";
      return;
    }

    host.innerHTML = `<img src="${escapeHtml(imagen)}" alt="Mosaico del grupo ${escapeHtml(grupo)}" loading="lazy" decoding="async">`;
    if (metaEl) {
      metaEl.textContent = actualizado ? `Actualizado: ${actualizado}` : "";
    }
  } catch (err) {
    if (seq !== mosaicoGrupoFetchSeq) return;
    state.cacheMosaicoGrupo = null;
    host.innerHTML = `<div class="state-msg state-msg--warn">${escapeHtml(`No se pudo cargar el mosaico: ${err.message}`)}</div>`;
    if (metaEl) metaEl.textContent = "";
  }
}

/** Imprime solo el bloque del mosaico del grupo (similar a impresión de resumen de área). */
function imprimirMosaicoGrupoVista() {
  if (state.appModo !== "mosaico-grupo") {
    toast("Abra la pestaña «Mosaico» antes de imprimir.", "err");
    return;
  }
  const cache = state.cacheMosaicoGrupo;
  if (!cache || !cache.imagen) {
    toast("Aún no hay mosaico cargado para este grupo.", "err");
    return;
  }
  const line = document.getElementById("print-line-mosaico-grupo");
  const fe = document.getElementById("print-fecha-mosaico-grupo");
  if (line) line.textContent = String(cache.grupo || getGrupo() || "");
  if (fe) {
    fe.textContent = new Date().toLocaleString("es-CO", { dateStyle: "long", timeStyle: "short" });
  }
  document.documentElement.classList.add("print-mode-mosaico-grupo");
  window.addEventListener(
    "afterprint",
    () => {
      document.documentElement.classList.remove("print-mode-mosaico-grupo");
    },
    { once: true }
  );
  window.print();
}

function guardarContexto() {
  try {
    localStorage.setItem(STORAGE.grupo, getGrupo());
    localStorage.setItem(STORAGE.materia, getMateria());
  } catch (_) {
    /* ignore */
  }
}

function leerContexto() {
  try {
    return {
      grupo: localStorage.getItem(STORAGE.grupo),
      materia: localStorage.getItem(STORAGE.materia)
    };
  } catch (_) {
    return { grupo: null, materia: null };
  }
}

function getStoredClave() {
  try {
    return sessionStorage.getItem(STORAGE.authSession) || localStorage.getItem(STORAGE.authPersist) || "";
  } catch (_) {
    return "";
  }
}

/** "coordinador" | "docente" | "estudiante" */
function getAuthTipo() {
  try {
    return (
      sessionStorage.getItem(STORAGE.authTipo) ||
      localStorage.getItem(STORAGE.authTipo) ||
      "coordinador"
    );
  } catch (_) {
    return "coordinador";
  }
}

function setAuthTipo(tipo, persistirEnEquipo) {
  let t = "coordinador";
  if (tipo === "docente") t = "docente";
  else if (tipo === "estudiante") t = "estudiante";
  else if (tipo === "coordinador") t = "coordinador";
  try {
    sessionStorage.setItem(STORAGE.authTipo, t);
    if (persistirEnEquipo) {
      localStorage.setItem(STORAGE.authTipo, t);
    } else {
      localStorage.removeItem(STORAGE.authTipo);
    }
  } catch (_) {
    /* ignore */
  }
}

function esEstudiante() {
  return getAuthTipo() === "estudiante";
}

/** Coordinación o estudiante: no editar calificaciones ni listas desde la UI. */
function esModoSoloConsulta() {
  return esSoloLecturaCoordinacion() || esEstudiante();
}

function getStoredDocenteId() {
  try {
    return (
      sessionStorage.getItem(STORAGE.docenteId) ||
      localStorage.getItem(STORAGE.docentePersist) ||
      ""
    );
  } catch (_) {
    return "";
  }
}

function setStoredDocenteId(codigo, persistirEnEquipo) {
  const c = String(codigo || "").trim();
  try {
    sessionStorage.setItem(STORAGE.docenteId, c);
    if (persistirEnEquipo) {
      localStorage.setItem(STORAGE.docentePersist, c);
    } else {
      localStorage.removeItem(STORAGE.docentePersist);
    }
  } catch (_) {
    /* ignore */
  }
}

function setStoredDocenteNombre(nombre, persistirEnEquipo) {
  const n = String(nombre || "").trim();
  try {
    sessionStorage.setItem(STORAGE.docenteNombre, n);
    if (persistirEnEquipo) {
      localStorage.setItem(STORAGE.docenteNombrePersist, n);
    } else {
      localStorage.removeItem(STORAGE.docenteNombrePersist);
    }
  } catch (_) {
    /* ignore */
  }
}

function getStoredDocenteNombre() {
  try {
    return (
      sessionStorage.getItem(STORAGE.docenteNombre) ||
      localStorage.getItem(STORAGE.docenteNombrePersist) ||
      ""
    );
  } catch (_) {
    return "";
  }
}

function setStoredClave(clave, persistirEnEquipo) {
  const c = String(clave || "").trim();
  try {
    sessionStorage.setItem(STORAGE.authSession, c);
    if (persistirEnEquipo) {
      localStorage.setItem(STORAGE.authPersist, c);
    } else {
      localStorage.removeItem(STORAGE.authPersist);
    }
  } catch (_) {
    /* ignore */
  }
}

function getStoredEstudianteDocumento() {
  try {
    return (
      sessionStorage.getItem(STORAGE.estudianteDoc) ||
      localStorage.getItem(STORAGE.estudianteDocPersist) ||
      ""
    );
  } catch (_) {
    return "";
  }
}

function getStoredEstudiantePin() {
  try {
    return (
      sessionStorage.getItem(STORAGE.estudiantePin) ||
      localStorage.getItem(STORAGE.estudiantePinPersist) ||
      ""
    );
  } catch (_) {
    return "";
  }
}

function setStoredEstudianteDocumento(doc, persistirEnEquipo) {
  const d = String(doc || "").trim();
  try {
    sessionStorage.setItem(STORAGE.estudianteDoc, d);
    if (persistirEnEquipo) {
      localStorage.setItem(STORAGE.estudianteDocPersist, d);
    } else {
      localStorage.removeItem(STORAGE.estudianteDocPersist);
    }
  } catch (_) {
    /* ignore */
  }
}

function setStoredEstudiantePin(pin, persistirEnEquipo) {
  const p = String(pin || "").trim();
  try {
    sessionStorage.setItem(STORAGE.estudiantePin, p);
    if (persistirEnEquipo) {
      localStorage.setItem(STORAGE.estudiantePinPersist, p);
    } else {
      localStorage.removeItem(STORAGE.estudiantePinPersist);
    }
  } catch (_) {
    /* ignore */
  }
}

function clearCredencialesStaff() {
  try {
    sessionStorage.removeItem(STORAGE.authSession);
    localStorage.removeItem(STORAGE.authPersist);
    sessionStorage.removeItem(STORAGE.docenteId);
    localStorage.removeItem(STORAGE.docentePersist);
    sessionStorage.removeItem(STORAGE.docenteNombre);
    localStorage.removeItem(STORAGE.docenteNombrePersist);
  } catch (_) {
    /* ignore */
  }
}

function clearCredencialesEstudiante() {
  try {
    sessionStorage.removeItem(STORAGE.estudianteDoc);
    localStorage.removeItem(STORAGE.estudianteDocPersist);
    sessionStorage.removeItem(STORAGE.estudiantePin);
    localStorage.removeItem(STORAGE.estudiantePinPersist);
    sessionStorage.removeItem(STORAGE.estudianteMateriaExp);
    sessionStorage.removeItem(STORAGE.estudiantePeriodoCal);
  } catch (_) {
    /* ignore */
  }
}

function clearStoredClave() {
  try {
    clearCredencialesStaff();
    clearCredencialesEstudiante();
    sessionStorage.removeItem(STORAGE.authTipo);
    localStorage.removeItem(STORAGE.authTipo);
  } catch (_) {
    /* ignore */
  }
}

/** Parámetro `periodo` para API de calificaciones en sesión estudiante. */
function periodoParamParaApiEstudiante() {
  if (!esEstudiante()) return {};
  const sel = document.getElementById("est-periodo-cal");
  let v = sel && sel.value ? String(sel.value).trim() : "";
  if (!v) {
    try {
      v = String(sessionStorage.getItem(STORAGE.estudiantePeriodoCal) || "").trim();
    } catch (_) {
      v = "";
    }
  }
  if (!v) return {};
  return { periodo: v };
}

/** Periodo para tabla del grupo y expediente en sesión docente/coordinación. */
function periodoParamParaApiDocente() {
  if (esEstudiante()) return {};
  const sel = document.getElementById("doc-periodo-cal");
  let v = sel && sel.value ? String(sel.value).trim() : "";
  if (!v) {
    try {
      v = String(localStorage.getItem(STORAGE.docentePeriodoCal) || "").trim();
    } catch (_) {
      v = "";
    }
  }
  if (!v) return { periodo: "1" };
  return { periodo: v };
}

/** Periodo para notas y resumen según rol (estudiante o barra docente). */
function periodoParamParaApiCalificaciones() {
  return esEstudiante() ? periodoParamParaApiEstudiante() : periodoParamParaApiDocente();
}

function mergeAuthParams(params = {}) {
  const p = params && typeof params === "object" && !Array.isArray(params) ? params : {};
  const docParam =
    p.documento !== undefined && p.documento !== null && String(p.documento).trim() !== ""
      ? String(p.documento).trim()
      : "";
  const pinParam =
    p.pin !== undefined && p.pin !== null && String(p.pin).trim() !== ""
      ? String(p.pin).trim()
      : "";
  const docStore = String(getStoredEstudianteDocumento() || "").trim();
  const pinStore = String(getStoredEstudiantePin() || "").trim();
  const doc = docParam || docStore;
  const pin = pinParam || pinStore;

  const accionTrim = String(p.accion || "").trim();
  /** Ping de login estudiante antes de fijar `authTipo` en almacenamiento. */
  const pingExplicitoEstudiante = accionTrim === "ping" && Boolean(docParam) && Boolean(pinParam);

  /**
   * Solo autenticar como estudiante si la sesión lo es o es el ping explícito de entrada.
   * Antes: cualquier documento/PIN residual en localStorage (p. ej. tras probar el portal estudiante)
   * forzaba `loginTipo: estudiante` en **todas** las peticiones del docente → el servidor no recibía
   * `clave` en la rama de personal y respondía «Falta clave de acceso» al cargar grupo / guardar.
   */
  if (esEstudiante() || pingExplicitoEstudiante) {
    return {
      ...p,
      loginTipo: "estudiante",
      documento: doc,
      pin
    };
  }

  const merged = {
    ...p,
    clave: String(getStoredClave() || "").trim()
  };
  let did = String(getStoredDocenteId() || "").trim();
  if (!did && getAuthTipo() === "coordinador") {
    did = String(CODIGO_COORDINADOR || "").trim();
  }
  if (did) {
    merged.docenteId = did;
  }
  return merged;
}

/** Evita UI «logueada como docente» con credenciales inconsistentes (sin clave persistida). */
function sanearSesionStaffSiFaltaClave_() {
  let tipoGuardado = "";
  try {
    tipoGuardado =
      sessionStorage.getItem(STORAGE.authTipo) || localStorage.getItem(STORAGE.authTipo) || "";
  } catch (_) {
    tipoGuardado = "";
  }
  if (tipoGuardado !== "docente" && tipoGuardado !== "coordinador") return;
  if (String(getStoredClave() || "").trim()) return;
  clearCredencialesStaff();
  try {
    sessionStorage.removeItem(STORAGE.authTipo);
    localStorage.removeItem(STORAGE.authTipo);
  } catch (_) {
    /* ignore */
  }
  state.cacheDocenteAsignaciones = null;
}

/** Alinea el modo de sesión con la respuesta del servidor (y compatibilidad si el script aún no envía `tipo`). */
function tipoAuthDesdePing(ping, codigoAlmacenado) {
  if (ping && ping.tipo === "docente") return "docente";
  if (ping && ping.tipo === "coordinador") return "coordinador";
  if (ping && ping.tipo === "estudiante") return "estudiante";
  const cod = String(codigoAlmacenado || "").trim();
  if (cod && normalizarTextoSimple(cod) === normalizarTextoSimple(CODIGO_COORDINADOR)) {
    return "coordinador";
  }
  return cod ? "docente" : "coordinador";
}

/**
 * Inicia sesión desde la URL (?clave=… y ?docente=… con código docente o código de coordinación).
 * Quita los parámetros sensibles de la barra de dirección al leerlos.
 */
function aplicarSesionDesdeQueryString() {
  try {
    const sp = new URLSearchParams(window.location.search);
    const claveTrim = sp.get("clave") ? String(sp.get("clave")).trim() : "";
    const docTrim = sp.get("docente") ? String(sp.get("docente")).trim() : "";
    if (!claveTrim || !docTrim) return;
    setStoredClave(claveTrim, false);
    setStoredDocenteId(docTrim, false);
    if (normalizarTextoSimple(docTrim) === normalizarTextoSimple(CODIGO_COORDINADOR)) {
      setAuthTipo("coordinador", false);
      state.cacheDocenteAsignaciones = null;
    } else {
      setAuthTipo("docente", false);
    }
    const path = window.location.pathname || "/index.html";
    history.replaceState({}, "", path + (window.location.hash || ""));
  } catch (_) {
    /* ignore */
  }
}

/** Sesión de coordinación (código de coordinación + clave institucional): solo consulta en servidor y en pantalla. */
function esSoloLecturaCoordinacion() {
  return Boolean(getStoredClave()) && getAuthTipo() === "coordinador";
}

const TITULO_VISTA_CAL_DOCENTE = "Vista general del grupo";
const TITULO_VISTA_CAL_ESTUDIANTE = "Tu informe académico";
const TEXTO_IMPRIMIR_CAL_DOCENTE = "Imprimir calificaciones del grupo";
const TEXTO_IMPRIMIR_CAL_ESTUDIANTE = "Imprimir mi informe académico";

/** Títulos del bloque calificaciones: docente (grupo) vs estudiante (personal). */
function actualizarTextosVistaCalificaciones() {
  const titulo = document.getElementById("titulo-vista-cal");
  const btnTxt = document.getElementById("btn-text-imprimir-cal");
  if (!titulo || !btnTxt) return;
  if (esEstudiante()) {
    titulo.textContent = TITULO_VISTA_CAL_ESTUDIANTE;
    btnTxt.textContent = TEXTO_IMPRIMIR_CAL_ESTUDIANTE;
  } else {
    titulo.textContent = TITULO_VISTA_CAL_DOCENTE;
    btnTxt.textContent = TEXTO_IMPRIMIR_CAL_DOCENTE;
  }
}

const LABEL_MODO_CAL_DOCENTE = "Calificaciones";
const LABEL_MODO_CAL_ESTUDIANTE = "Boletín general de materias";

function actualizarLabelModoCalificaciones() {
  const span = document.getElementById("label-modo-calificaciones");
  if (!span) return;
  span.textContent = esEstudiante() ? LABEL_MODO_CAL_ESTUDIANTE : LABEL_MODO_CAL_DOCENTE;
  const btn = document.getElementById("modo-calificaciones");
  if (btn) {
    btn.setAttribute(
      "title",
      esEstudiante()
        ? "Resumen de todas tus materias con la definitiva orientativa"
        : "Vista de calificaciones del grupo"
    );
  }
}

let boletinFetchSeq = 0;
let resumenAreaFetchSeq = 0;

const AVISO_BOLETIN_CLIENTE =
  "La «definitiva orientativa» usa la fórmula del sistema (70% seguimiento, 10% actitudinal, 20% prueba). Si faltan registros, no es la calificación final oficial. Para el detalle de una materia, elíjala arriba y use «Abrir expediente» al lado del selector.";

/** Si solo corre el legado (fallan boletín y notasEstudianteAgregado en el servidor desplegado). */
const AVISO_BOLETIN_LEGACY =
  "Respaldo del navegador: publique la versión actual de Code.gs en Apps Script para tomar las materias del curso desde la hoja DocenteAsignaciones. ";

function parseDefinitivaBoletinNum(row) {
  const raw = String(row?.definitiva ?? "").trim();
  if (raw === "—" || raw === "") return NaN;
  const n = parseFloat(raw.replace(",", "."));
  return Number.isFinite(n) ? n : NaN;
}

/** Una decimal en resumen estudiante (redondeo estándar, igual que en campos de nota). */
function formatearDefinitivaDisplayBoletinEstudiante_(val) {
  if (val == null) return "—";
  const s = String(val).trim();
  if (s === "" || s === "—") return "—";
  const n = parseFloat(s.replace(",", "."));
  if (!Number.isFinite(n)) return s;
  return redondearNotaPasoDecima(n).toFixed(1);
}

function resolverFilaComponenteArea(mapPorClave, etiquetas) {
  const list = Array.isArray(etiquetas) ? etiquetas : [];
  for (const et of list) {
    const k = normalizarMateriaClaveIEMFS(et);
    const r = mapPorClave.get(k);
    if (!r) continue;
    const n = parseDefinitivaBoletinNum(r);
    if (!Number.isNaN(n)) return { valor: n, etiqueta: r.materia };
  }
  return null;
}

/** Añade filas de sección y áreas MEG / Humanidades tras las asignaturas. */
function compilarFilasBoletinEstudianteCompleto(itemsMaterias) {
  const materiasRows = Array.isArray(itemsMaterias)
    ? itemsMaterias.map((r) => ({ filaTipo: "asignatura", ...r }))
    : [];
  const mapPorClave = new Map();
  for (const row of materiasRows) {
    const lab = String(row.materia || "").trim();
    if (!lab) continue;
    mapPorClave.set(normalizarMateriaClaveIEMFS(lab), row);
  }
  const areaFilas = [];
  for (const area of AREAS_BOLETIN_ESTUDIANTE) {
    const rArea = calcularDefinitivaAreaDesdeMapPorClave_(area, mapPorClave);
    areaFilas.push({
      filaTipo: "area",
      materia: area.nombre,
      formulaCorta: area.formulaCorta,
      definitiva: rArea.definitiva,
      incompleto: rArea.incompleto,
      avisos: rArea.avisos
    });
  }
  return [
    ...materiasRows,
    { filaTipo: "seccion", titulo: "Áreas curriculares (ponderación orientativa)" },
    ...areaFilas
  ];
}

/** MEG o Humanidades según la materia elegida en la barra (docente / coordinación). */
function areaResumenDocenteParaMateriaActual_() {
  const m = String(getMateria() || "").trim();
  if (!m) return null;
  const k = normalizarMateriaClaveIEMFS(m);
  for (const area of AREAS_BOLETIN_ESTUDIANTE) {
    for (const comp of area.componentes) {
      for (const et of comp.etiquetas) {
        if (normalizarMateriaClaveIEMFS(et) === k) return area;
      }
    }
  }
  return null;
}

function materiasApiParaArea_(area) {
  return area.componentes
    .map((c) => String(c.etiquetas && c.etiquetas[0] ? c.etiquetas[0] : "").trim())
    .filter(Boolean);
}

/**
 * Docente: columnas del resumen de área solo para materias que imparte en el curso actual
 * (p. ej. Geometría y Estadística sin Matemáticas). Coordinación u otros: todas las del área.
 */
function materiasFetchResumenAreaDocente_(area) {
  const ordenCompleto = materiasApiParaArea_(area);
  if (getAuthTipo() !== "docente" || !Array.isArray(state.cacheDocenteAsignaciones)) {
    return ordenCompleto;
  }
  const gNorm = normalizarGrupoCurso(getGrupo());
  const matsDocente = new Set();
  state.cacheDocenteAsignaciones.forEach((row) => {
    if (normalizarGrupoCurso(row.grupo || "") !== gNorm) return;
    const mk = normalizarMateriaClaveIEMFS(String(row.materia || "").trim());
    if (mk) matsDocente.add(mk);
  });
  if (!matsDocente.size) return ordenCompleto;
  const filtradas = ordenCompleto.filter((m) => matsDocente.has(normalizarMateriaClaveIEMFS(m)));
  return filtradas.length ? filtradas : ordenCompleto;
}

/** Encabezado visible en la tabla (API puede decir «Lengua Castellana» → «Castellano»). */
function tituloColumnaResumenArea_(etiquetaApi) {
  const s = String(etiquetaApi || "").trim();
  if (!s) return "—";
  const k = normalizarMateriaClaveIEMFS(s);
  if (k === normalizarMateriaClaveIEMFS("Lengua Castellana")) return "Castellano";
  if (k === normalizarMateriaClaveIEMFS("Castellano")) return "Castellano";
  if (k === normalizarMateriaClaveIEMFS("Competencia lectora")) return "Competencia lectora";
  return s;
}

/** # + Estudiante + Estado + n materias + área + Obs */
function raColspanTabla(nMaterias) {
  const n = Math.max(0, Number(nMaterias) || 0);
  return 5 + n;
}

/** Resumen área: docente sin columna de total del área → # + Estudiante + Estado + materias + Obs. */
function raColspanTablaResumenArea(nMaterias, sinColumnaTotalArea) {
  const n = Math.max(0, Number(nMaterias) || 0);
  return sinColumnaTotalArea ? 4 + n : 5 + n;
}

function notasArrayParaBoletinDesdeFilasGrupo_(flat, estudianteId) {
  const id = String(estudianteId);
  return (flat || [])
    .filter((n) => String(n.estudiante) === id)
    .map((n) => ({ tipo: n.tipo, nota: n.nota }));
}

/**
 * Ponderación del área: si hay definitiva en todas las asignaturas del bloque, usa los pesos oficiales.
 * Si faltan notas en alguna, recalcula solo con las disponibles (misma proporción relativa entre sí; total parcial).
 */
function calcularDefinitivaAreaDesdeMapPorClave_(area, mapPorClave) {
  const hits = [];
  const avisosFalt = [];
  for (const comp of area.componentes) {
    const hit = resolverFilaComponenteArea(mapPorClave, comp.etiquetas);
    if (!hit) {
      avisosFalt.push(`Falta ${comp.etiquetas[0]}`);
    } else {
      hits.push({ valor: hit.valor, peso: comp.peso });
    }
  }
  if (!hits.length) {
    return { definitiva: "—", incompleto: true, avisos: avisosFalt };
  }
  const completa = hits.length === area.componentes.length;
  let sumP = 0;
  for (let i = 0; i < hits.length; i++) sumP += hits[i].peso;
  let sum = 0;
  for (let i = 0; i < hits.length; i++) sum += hits[i].valor * (hits[i].peso / sumP);
  return {
    definitiva: redondearNotaPasoDecima(sum).toFixed(1),
    incompleto: !completa,
    avisos: avisosFalt
  };
}

/** Ponderación de un solo bloque para la tabla docente (misma lógica que boletín estudiante). */
function calcularDefinitivaAreaUnicaDocente_(area, itemsMateriasRows) {
  const mapPorClave = new Map();
  for (const row of itemsMateriasRows) {
    const lab = String(row.materia || "").trim();
    if (!lab) continue;
    mapPorClave.set(normalizarMateriaClaveIEMFS(lab), { materia: lab, definitiva: row.definitiva });
  }
  return calcularDefinitivaAreaDesdeMapPorClave_(area, mapPorClave);
}

function renderBoletinTablaHtml(items) {
  const rows = Array.isArray(items) ? items : [];
  return `
      <table class="boletin-table">
        <thead>
          <tr>
            <th scope="col">Materia / área</th>
            <th scope="col">Definitiva orientativa</th>
            <th scope="col">Observaciones</th>
          </tr>
        </thead>
        <tbody>
          ${rows
            .map((row) => {
              if (row.filaTipo === "seccion") {
                return `<tr class="boletin-row--section"><th colspan="3" scope="colgroup">${escapeHtml(
                  row.titulo || ""
                )}</th></tr>`;
              }
              if (row.filaTipo === "area") {
                const obs = row.incompleto
                  ? Array.isArray(row.avisos) && row.avisos.length
                    ? row.avisos.join(" · ")
                    : "Faltan notas de alguna asignatura del área"
                  : String(row.formulaCorta || "");
                const cls = `boletin-row--area${row.incompleto ? " boletin-row--warn" : ""}`;
                return `<tr class="${cls}">
                <td>${escapeHtml(row.materia)}</td>
                <td><strong>${escapeHtml(formatearDefinitivaDisplayBoletinEstudiante_(row.definitiva))}</strong></td>
                <td>${escapeHtml(obs)}</td>
              </tr>`;
              }
              const obs = row.incompleto
                ? Array.isArray(row.avisos) && row.avisos.length
                  ? row.avisos.join(" · ")
                  : "Evaluación incompleta u orientativa"
                : "Al día (detalle en expediente)";
              return `<tr class="${row.incompleto ? "boletin-row--warn" : ""}">
                <td>${escapeHtml(row.materia)}</td>
                <td><strong>${escapeHtml(formatearDefinitivaDisplayBoletinEstudiante_(row.definitiva))}</strong></td>
                <td>${escapeHtml(obs)}</td>
              </tr>`;
            })
            .join("")}
        </tbody>
      </table>`;
}

/** Misma lógica que Code.gs calcularResumenBoletinMetaDesdeFilas_ usando filas de `notas`. */
function calcularMetaBoletinDesdeNotasArr(notas) {
  const arr = Array.isArray(notas) ? notas : [];
  const seguimiento = [];
  let actitudinal = 0;
  let prueba = 0;
  let tieneActitudinal = false;
  let tienePrueba = false;

  for (const n of arr) {
    const tipo = String(n.tipo || "").trim();
    const val = parseFloat(n.nota);
    if (tipo === "Seguimiento" && !isNaN(val)) seguimiento.push(val);
    if (tipo === "Actitudinal") {
      tieneActitudinal = true;
      actitudinal = isNaN(val) ? 0 : val;
    }
    if (tipo === "Prueba") {
      tienePrueba = true;
      prueba = isNaN(val) ? 0 : val;
    }
  }

  const promSeg =
    seguimiento.length > 0
      ? seguimiento.reduce((a, b) => a + b, 0) / seguimiento.length
      : 0;
  const final = promSeg * 0.7 + actitudinal * 0.1 + prueba * 0.2;
  const incompleto = seguimiento.length === 0 || !tieneActitudinal || !tienePrueba;
  const avisos = [];
  if (seguimiento.length === 0) avisos.push("Falta seguimiento");
  if (!tieneActitudinal) avisos.push("Falta Actitudinal");
  if (!tienePrueba) avisos.push("Falta Prueba");

  return {
    definitiva: redondearNotaPasoDecima(final).toFixed(1),
    incompleto,
    avisos
  };
}

/** Si falla `boletinResumenEstudiante`: una petición `notasEstudianteAgregado` (una lectura en servidor). */
async function cargarBoletinEstudianteFallback(estudianteId) {
  try {
    const agg = await apiGet({ accion: "notasEstudianteAgregado", ...periodoParamParaApiEstudiante() });
    const materias = Array.isArray(agg.materias) ? agg.materias : [];
    const items = materias.map(({ materia, notas }) => {
      const meta = calcularMetaBoletinDesdeNotasArr(notas);
      return {
        materia,
        definitiva: meta.definitiva,
        incompleto: meta.incompleto,
        avisos: meta.avisos
      };
    });
    return { items, aviso: AVISO_BOLETIN_CLIENTE };
  } catch {
    return cargarBoletinEstudianteFallbackLegacy(estudianteId);
  }
}

/** Despliegues muy antiguos sin `notasEstudianteAgregado`: una petición `notas` por materia. */
async function cargarBoletinEstudianteFallbackLegacy(estudianteId) {
  const id = String(estudianteId || "").trim();
  const mats = materiasListaEstudianteActual();
  const items = await Promise.all(
    mats.map(async (mat) => {
      try {
        const notas = await apiGet({ accion: "notas", id, materia: mat, ...periodoParamParaApiEstudiante() });
        const meta = calcularMetaBoletinDesdeNotasArr(notas);
        return {
          materia: mat,
          definitiva: meta.definitiva,
          incompleto: meta.incompleto,
          avisos: meta.avisos
        };
      } catch {
        return {
          materia: mat,
          definitiva: "—",
          incompleto: true,
          avisos: ["Sin datos"]
        };
      }
    })
  );
  return { items, aviso: AVISO_BOLETIN_LEGACY + AVISO_BOLETIN_CLIENTE };
}

/** Tabla resumen en la tarjeta del estudiante (host #estudiante-boletin-inline-host). */
async function cargarBoletinEstudiante() {
  if (!esEstudiante()) return;
  if (state.appModo !== "calificaciones") return;

  const host = document.getElementById("estudiante-boletin-inline-host");
  const avisoEl = document.getElementById("estudiante-boletin-aviso");
  if (!host) return;

  actualizarTituloNotasGeneralesEstudiante_();

  const seq = ++boletinFetchSeq;
  host.innerHTML =
    '<div class="state-msg state-msg--loading"><span class="spinner" aria-hidden="true"></span> Cargando notas generales…</div>';
  if (avisoEl) avisoEl.textContent = "";

  try {
    let data;
    try {
      data = await apiGet({ accion: "boletinResumenEstudiante", ...periodoParamParaApiEstudiante() });
    } catch (primaryErr) {
      const id = String(state.cacheEstudiantes[0]?.id || "").trim();
      if (!id) throw primaryErr;
      data = await cargarBoletinEstudianteFallback(id);
    }
    if (seq !== boletinFetchSeq) return;
    if (avisoEl && data.aviso) avisoEl.textContent = data.aviso;
    const filas = compilarFilasBoletinEstudianteCompleto(data.items);
    host.innerHTML = renderBoletinTablaHtml(filas);
  } catch (err) {
    if (seq !== boletinFetchSeq) return;
    host.innerHTML = `<p class="state-msg state-msg--warn">${escapeHtml(err.message || "No se pudo cargar el resumen.")}</p>`;
    if (avisoEl) avisoEl.textContent = "";
  }
}

function tituloSugeridoFichaNuevaNota(tipo) {
  const t = String(tipo || "Seguimiento").trim();
  const m = String(getMateria() || "").trim();
  return m ? `${t} ${m}` : t;
}

/** Separador interno valor option (no suele aparecer en títulos). */
const FICHA_EVAL_SEP = "\x1e";

function encodeEvalSelectVal(tipo, titulo) {
  return `${String(tipo || "").trim()}${FICHA_EVAL_SEP}${String(titulo || "").trim()}`;
}

function decodeEvalSelectVal(val) {
  const s = String(val || "");
  const i = s.indexOf(FICHA_EVAL_SEP);
  if (i < 0) return { tipo: "", titulo: "" };
  return { tipo: s.slice(0, i).trim(), titulo: s.slice(i + FICHA_EVAL_SEP.length).trim() };
}

/**
 * Pares tipo+título que existen en el grupo (notasGrupo) o, si no hubo carga, en el expediente actual;
 * más sugerencias «Seguimiento Materia» para la materia activa.
 */
function evaluacionesDisponiblesParaFicha_(listaNotas) {
  const seen = new Set();
  const out = [];
  const add = (tipo, titulo) => {
    const t = String(tipo || "").trim();
    const ti = String(titulo || "").trim();
    if (!t || !ti) return;
    const k = `${t}\n${ti}`;
    if (seen.has(k)) return;
    seen.add(k);
    out.push({ tipo: t, titulo: ti });
  };
  (state.cacheNotasGrupo || []).forEach((n) => add(n.tipo, n.titulo));
  if (!out.length && Array.isArray(listaNotas)) {
    listaNotas.forEach((n) => add(n.tipo, n.titulo));
  }
  const m = String(getMateria() || "").trim();
  if (m) {
    add("Seguimiento", `Seguimiento ${m}`);
    add("Actitudinal", `Actitudinal ${m}`);
    add("Prueba", `Prueba ${m}`);
  }
  out.sort(
    (a, b) =>
      a.tipo.localeCompare(b.tipo, "es") || a.titulo.localeCompare(b.titulo, "es")
  );
  return out;
}

function notaEstudianteMasRecienteParaEval_(listaNotas, tipo, titulo) {
  if (!Array.isArray(listaNotas)) return null;
  const tp = String(tipo || "").trim();
  const tt = String(titulo || "").trim();
  const matches = listaNotas.filter(
    (n) =>
      String(n.tipo || "").trim() === tp && String(n.titulo || "").trim() === tt
  );
  if (!matches.length) return null;
  matches.sort((a, b) => {
    const tb = new Date(b.fecha || 0).getTime();
    const ta = new Date(a.fecha || 0).getTime();
    return tb - ta;
  });
  return matches[0];
}

function actualizarFichaHintListaEvalDesdeGrupo_() {
  const hintGrupo = document.getElementById("ficha-nueva-hint-grupo");
  if (!hintGrupo) return;
  const vacio = !(state.cacheNotasGrupo && state.cacheNotasGrupo.length);
  if (vacio && !esModoSoloConsulta()) {
    hintGrupo.textContent =
      "No hay datos del grupo en memoria: en la vista general pulse «Cargar grupo» con el mismo curso y materia para llenar el selector con los títulos reales de la hoja.";
    hintGrupo.hidden = false;
    hintGrupo.classList.remove("is-hidden");
  } else {
    hintGrupo.textContent = "";
    hintGrupo.hidden = true;
    hintGrupo.classList.add("is-hidden");
  }
}

function poblarSelectorEvaluacionFicha(listaNotas) {
  const sel = document.getElementById("ficha-nueva-eval");
  if (!sel) return;
  sel.innerHTML = `<option value="">— Elija la evaluación del grupo —</option>`;
  evaluacionesDisponiblesParaFicha_(listaNotas).forEach(({ tipo, titulo }) => {
    const opt = document.createElement("option");
    opt.value = encodeEvalSelectVal(tipo, titulo);
    opt.textContent = `${tipo} · ${titulo}`;
    sel.appendChild(opt);
  });
  const oMan = document.createElement("option");
  oMan.value = "__manual__";
  oMan.textContent = "Otro (tipo y título manual…)";
  sel.appendChild(oMan);
  actualizarFichaHintListaEvalDesdeGrupo_();
  sincronizarUiTrasCambioEvalFicha_();
}

function sincronizarUiTrasCambioEvalFicha_() {
  const sel = document.getElementById("ficha-nueva-eval");
  const manualWrap = document.getElementById("ficha-nueva-manual-wrap");
  const hint = document.getElementById("ficha-nueva-accion-hint");
  const btn = document.getElementById("ficha-nueva-guardar");
  const val = sel?.value ?? "";
  if (val === "__manual__") {
    if (manualWrap) {
      manualWrap.hidden = false;
      manualWrap.classList.remove("is-hidden");
    }
    if (hint) {
      hint.textContent =
        "Modo manual: escriba el título exactamente como en el registro masivo para no crear una evaluación duplicada.";
    }
    if (btn) btn.textContent = "Guardar calificación";
    return;
  }
  if (manualWrap) {
    manualWrap.hidden = true;
    manualWrap.classList.add("is-hidden");
  }
  if (!val) {
    const notaEl = document.getElementById("ficha-nueva-nota");
    if (notaEl) notaEl.value = "";
    if (hint) hint.textContent = "";
    if (btn) btn.textContent = "Guardar calificación";
    return;
  }
  const { tipo, titulo } = decodeEvalSelectVal(val);
  const existente = notaEstudianteMasRecienteParaEval_(state.fichaListaNotas, tipo, titulo);
  const notaEl = document.getElementById("ficha-nueva-nota");
  const fechaEl = document.getElementById("ficha-nueva-fecha");
  if (existente && notaEl) {
    const nr = existente.nota;
    const tiene =
      nr !== undefined &&
      nr !== null &&
      String(nr).trim() !== "" &&
      !Number.isNaN(parseFloat(nr));
    notaEl.value = tiene ? String(redondearNotaPasoDecima(parseFloat(nr))) : "";
    if (fechaEl) fechaEl.value = fechaIsoParaInputDate(existente.fecha);
    if (hint) {
      hint.textContent =
        "Ya hay una nota para esta evaluación: al guardar se actualiza la misma fila (corrección o mejora).";
    }
    if (btn) btn.textContent = "Actualizar calificación";
  } else {
    if (notaEl) notaEl.value = "";
    if (fechaEl) {
      const cr = document.getElementById("cr-fecha");
      fechaEl.value = (cr && cr.value) || fechaIsoParaInputDate("");
    }
    if (hint) {
      hint.textContent =
        "No hay nota en este expediente para esta evaluación: se registrará una fila nueva.";
    }
    if (btn) btn.textContent = "Registrar calificación";
  }
}

function resetFormAnadirNotaFicha() {
  const sel = document.getElementById("ficha-nueva-eval");
  if (sel) {
    sel.innerHTML = `<option value="">— Elija la evaluación del grupo —</option>`;
  }
  const tipoEl = document.getElementById("ficha-nueva-tipo");
  const tituloEl = document.getElementById("ficha-nueva-titulo");
  const notaEl = document.getElementById("ficha-nueva-nota");
  const fechaEl = document.getElementById("ficha-nueva-fecha");
  const manualWrap = document.getElementById("ficha-nueva-manual-wrap");
  const hint = document.getElementById("ficha-nueva-accion-hint");
  const hintGrupo = document.getElementById("ficha-nueva-hint-grupo");
  const btn = document.getElementById("ficha-nueva-guardar");
  if (tipoEl) tipoEl.value = "Seguimiento";
  if (tituloEl) tituloEl.value = tituloSugeridoFichaNuevaNota("Seguimiento");
  if (notaEl) notaEl.value = "";
  asegurarFechaCargaRapidaPorDefecto();
  if (fechaEl) {
    const cr = document.getElementById("cr-fecha");
    fechaEl.value = (cr && cr.value) || fechaIsoParaInputDate("");
  }
  if (manualWrap) {
    manualWrap.hidden = true;
    manualWrap.classList.add("is-hidden");
  }
  if (hint) hint.textContent = "";
  if (hintGrupo) {
    hintGrupo.textContent = "";
    hintGrupo.hidden = true;
    hintGrupo.classList.add("is-hidden");
  }
  if (btn) btn.textContent = "Guardar calificación";
}

function aplicarVisibilidadAnadirNotaFicha() {
  const w = document.getElementById("ficha-anadir-nota-wrap");
  if (!w) return;
  const show = !esModoSoloConsulta();
  w.hidden = !show;
  w.classList.toggle("is-hidden", !show);
}

async function guardarNotaDesdeFichaEstudiante() {
  if (esModoSoloConsulta()) {
    toast(
      esEstudiante()
        ? "El acceso de estudiante es solo consulta."
        : "El acceso de coordinación es solo consulta.",
      "err"
    );
    return;
  }
  if (!state.estudianteActual) {
    toast("Abra primero la ficha de un estudiante.", "err");
    return;
  }
  const materia = String(getMateria() || "").trim();
  const grupo = String(state.estudianteActual.grupo || getGrupo() || "").trim();
  if (!materia) {
    toast("Seleccione la materia en la barra superior.", "err");
    return;
  }
  if (!grupo) {
    toast("Falta el curso del estudiante.", "err");
    return;
  }

  const selVal = document.getElementById("ficha-nueva-eval")?.value ?? "";
  let tipo;
  let titulo;
  if (selVal === "__manual__") {
    tipo = String(document.getElementById("ficha-nueva-tipo")?.value || "").trim();
    titulo = String(document.getElementById("ficha-nueva-titulo")?.value || "").trim();
    if (!tipo) {
      toast("Elija el tipo de evaluación.", "err");
      return;
    }
    if (!titulo) {
      toast("Escriba el título de la evaluación.", "err");
      return;
    }
  } else if (selVal) {
    const dec = decodeEvalSelectVal(selVal);
    tipo = dec.tipo;
    titulo = dec.titulo;
    if (!tipo || !titulo) {
      toast("Selección no válida.", "err");
      return;
    }
  } else {
    toast("Elija una evaluación en la lista o «Otro (tipo y título manual)».", "err");
    return;
  }

  const fechaStr = document.getElementById("ficha-nueva-fecha")?.value?.trim() || "";
  const notaRaw = document.getElementById("ficha-nueva-nota")?.value ?? "";
  if (!fechaStr) {
    toast("Indique la fecha de la calificación.", "err");
    return;
  }
  const num = parseNotaDesdeCampo(notaRaw);
  if (num === null) {
    toast("Nota no válida.", "err");
    return;
  }
  if (!notaEnRango(num)) {
    toast("La nota debe estar entre 0 y 5.", "err");
    return;
  }
  const notaEnvio = redondearNotaPasoDecima(num);
  const periodoPayload = periodoParamParaApiDocente();

  const existente = notaEstudianteMasRecienteParaEval_(state.fichaListaNotas, tipo, titulo);
  const puedeActualizar =
    existente && (notaIdDe(existente) || snapshotAntesParaBuscar(existente));

  const btn = document.getElementById("ficha-nueva-guardar");
  if (btn) btn.disabled = true;
  try {
    if (puedeActualizar) {
      const id = notaIdDe(existente);
      const antes = snapshotAntesParaBuscar(existente);
      const payload = {
        accion: "actualizar",
        estudiante: state.estudianteActual.id,
        materia,
        grupo,
        titulo,
        nota: notaEnvio,
        tipo,
        ...periodoPayload,
        fecha: fechaStr,
        fechaCalificacion: fechaStr,
        fechaNota: fechaStr
      };
      if (id) payload.notaId = id;
      else payload.antes = antes;
      await apiPost(payload);
      toast("Informacion guardada: calificacion actualizada.", "ok");
    } else {
      await apiPost({
        accion: "guardar",
        estudiante: state.estudianteActual.id,
        materia,
        grupo,
        tipo: String(tipo).trim(),
        titulo,
        nota: notaEnvio,
        ...periodoPayload,
        fecha: fechaStr,
        fechaCalificacion: fechaStr,
        fechaNota: fechaStr
      });
      toast("Informacion guardada: calificacion registrada.", "ok");
    }
    await abrirEstudiante(state.estudianteActual);
  } catch (error) {
    toast(error.message || "No se pudo guardar.", "err");
  } finally {
    if (btn) btn.disabled = false;
  }
}

function aplicarModoEstudianteUI() {
  const shell = document.getElementById("app-shell");
  const banner = document.getElementById("banner-estudiante");
  const exWrap = document.getElementById("toolbar-expediente-wrap");
  const perWrap = document.getElementById("toolbar-est-periodo-wrap");
  const es = esEstudiante();
  if (shell) {
    shell.classList.toggle("app-shell--estudiante", es);
  }
  if (banner) {
    banner.classList.toggle("is-hidden", !es);
    banner.hidden = !es;
  }
  if (exWrap) {
    exWrap.hidden = !es;
    exWrap.classList.toggle("is-hidden", !es);
  }
  if (perWrap) {
    perWrap.hidden = !es;
    perWrap.classList.toggle("is-hidden", !es);
  }
  const docPerWrap = document.getElementById("toolbar-doc-periodo-wrap");
  if (docPerWrap) {
    docPerWrap.hidden = es;
    docPerWrap.classList.toggle("is-hidden", es);
  }
  if (es) {
    prepararSelectPeriodoEstudianteCargando();
  }
  actualizarTextosVistaCalificaciones();
  actualizarLabelModoCalificaciones();
  actualizarEtiquetaFichaTabEstado();
  if (state.estudianteActual) {
    aplicarVisibilidadEstadoMatriculaPanel();
    actualizarMateriaEnFichaEstudiante();
  }
  aplicarVisibilidadAnadirNotaFicha();
}

function aplicarModoSoloLecturaCoordinacionUI() {
  const shell = document.getElementById("app-shell");
  const banner = document.getElementById("banner-solo-lectura");
  const solo = esSoloLecturaCoordinacion();
  if (shell) {
    shell.classList.toggle("app-shell--solo-lectura", solo);
  }
  if (banner) {
    banner.classList.toggle("is-hidden", !solo);
    banner.hidden = !solo;
  }
  if (state.estudianteActual) {
    aplicarVisibilidadEstadoMatriculaPanel();
    actualizarMateriaEnFichaEstudiante();
  }
  aplicarVisibilidadAnadirNotaFicha();
}

function applyTheme(mode) {
  const root = document.documentElement;
  const meta = document.querySelector("meta[name=\"theme-color\"]");
  const btn = document.getElementById("btn-theme");
  if (mode === "dark") {
    root.setAttribute("data-theme", "dark");
    if (meta) meta.setAttribute("content", "#0c1222");
    if (btn) {
      btn.setAttribute("aria-label", "Activar modo claro");
      btn.setAttribute("title", "Modo claro");
    }
  } else {
    root.removeAttribute("data-theme");
    if (meta) meta.setAttribute("content", "#0c4a6e");
    if (btn) {
      btn.setAttribute("aria-label", "Activar modo oscuro");
      btn.setAttribute("title", "Modo oscuro");
    }
  }
}

function initTheme() {
  let t = "light";
  try {
    t = localStorage.getItem(STORAGE.theme) || "light";
  } catch (_) {
    /* ignore */
  }
  applyTheme(t === "dark" ? "dark" : "light");

  const btnTheme = document.getElementById("btn-theme");
  if (btnTheme) {
    btnTheme.addEventListener("click", () => {
      const isDark = document.documentElement.getAttribute("data-theme") === "dark";
      const next = isDark ? "light" : "dark";
      applyTheme(next);
      try {
        localStorage.setItem(STORAGE.theme, next);
      } catch (_) {
        /* ignore */
      }
    });
  }
}

function actualizarEncabezadoImpresion() {
  const pf = document.getElementById("print-fecha");
  const pg = document.getElementById("print-grupo");
  const pm = document.getElementById("print-materia");
  const pe = document.getElementById("print-estado");
  if (pf) {
    pf.textContent = new Date().toLocaleString("es-CO", { dateStyle: "long", timeStyle: "short" });
  }
  if (pg) pg.textContent = getGrupo();
  if (pm) pm.textContent = getMateria();
  if (pe) {
    pe.textContent = state.estudianteActual
      ? `Matrícula: ${state.estudianteActual.estadoMatricula || "Activo"}`
      : "";
  }
}

function limpiarClasesImpresionVistaListado() {
  document.documentElement.classList.remove(
    "print-mode-grupo-cal",
    "print-mode-estudiante-resumen",
    "print-mode-resumen-faltas",
    "print-mode-resumen-area"
  );
}

/** Imprime la tabla del grupo (docente) o la tarjeta resumen del curso (estudiante). */
function imprimirGrupoCalificacionesVista() {
  const line = document.getElementById("print-line-grupo-cal");
  const fe = document.getElementById("print-fecha-grupo-cal");
  const g = getGrupo();
  const m = getMateria();

  if (esEstudiante()) {
    const lista = document.getElementById("lista");
    const card = lista?.querySelector(".estudiante-solo-card");
    if (!lista || lista.classList.contains("is-hidden") || !card) {
      toast("Espere a que cargue su información del curso.", "err");
      return;
    }
    if (line) {
      const nom = String(state.estudianteNombreSesion || "").trim();
      line.textContent = nom || "Informe académico";
    }
    if (fe) {
      fe.textContent = new Date().toLocaleString("es-CO", { dateStyle: "long", timeStyle: "short" });
    }
    document.documentElement.classList.add("print-mode-grupo-cal", "print-mode-estudiante-resumen");
    window.addEventListener("afterprint", limpiarClasesImpresionVistaListado, { once: true });
    window.print();
    return;
  }

  const wrap = document.getElementById("grupo-board-wrap");
  if (!wrap || wrap.classList.contains("is-hidden")) {
    toast("Cargue el grupo y espere a que aparezcan las calificaciones.", "err");
    return;
  }
  if (line) {
    line.textContent = g && m ? `${g} · ${m}` : g || m || "—";
  }
  if (fe) {
    fe.textContent = new Date().toLocaleString("es-CO", { dateStyle: "long", timeStyle: "short" });
  }
  document.documentElement.classList.add("print-mode-grupo-cal");
  window.addEventListener("afterprint", limpiarClasesImpresionVistaListado, { once: true });
  window.print();
}

/** Imprime el resumen de faltas por periodo (tabla del grupo ya cargada). */
function imprimirResumenFaltasVista() {
  if (state.appModo !== "resumen-faltas" || !state.cacheEstudiantes.length) {
    toast("Cargue el grupo, abra «Resumen faltas» y actualice la tabla.", "err");
    return;
  }
  const data = state.cacheResumenFaltas;
  if (!data || !Array.isArray(data.estudiantes) || !data.estudiantes.length) {
    toast("Pulse «Actualizar tabla» cuando el resumen haya cargado.", "err");
    return;
  }
  const line = document.getElementById("print-line-resumen-faltas");
  const fe = document.getElementById("print-fecha-resumen-faltas");
  const rango = document.getElementById("rf-periodo-rango");
  if (line) {
    const per = rango && rango.textContent.trim() ? rango.textContent.trim() : "";
    line.textContent = [getGrupo(), getMateria(), per].filter(Boolean).join(" · ");
  }
  if (fe) {
    fe.textContent = new Date().toLocaleString("es-CO", { dateStyle: "long", timeStyle: "short" });
  }
  document.documentElement.classList.add("print-mode-resumen-faltas");
  window.addEventListener("afterprint", limpiarClasesImpresionVistaListado, { once: true });
  window.print();
}

/** Imprime la tabla de resumen de área (MEG / Humanidades) del grupo cargado. */
function imprimirResumenAreaVista() {
  if (state.appModo !== "resumen-area" || !state.cacheEstudiantes.length) {
    toast("Cargue el grupo, abra «Resumen de área» y espere la tabla.", "err");
    return;
  }
  const data = state.cacheResumenArea;
  if (!data || !Array.isArray(data.filas) || !data.filas.length) {
    toast("Pulse «Actualizar tabla» cuando el resumen haya cargado.", "err");
    return;
  }
  const line = document.getElementById("print-line-resumen-area");
  const fe = document.getElementById("print-fecha-resumen-area");
  const sel = document.getElementById("doc-periodo-cal");
  const per = sel && sel.selectedOptions && sel.selectedOptions[0] ? sel.selectedOptions[0].textContent.trim() : "";
  if (line) {
    if (data.sinColumnaTotalArea) {
      const mats = (data.titulosColumnasMateria || []).filter(Boolean).join(", ");
      line.textContent = [getGrupo(), mats, per].filter(Boolean).join(" · ");
    } else {
      const piezas = [getGrupo(), data.areaNombre || "", per].filter(Boolean);
      line.textContent = piezas.join(" · ");
    }
  }
  if (fe) {
    fe.textContent = new Date().toLocaleString("es-CO", { dateStyle: "long", timeStyle: "short" });
  }
  document.documentElement.classList.add("print-mode-resumen-area");
  window.addEventListener("afterprint", limpiarClasesImpresionVistaListado, { once: true });
  window.print();
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Reintenta una vez ante cortes típicos en móvil / Wi‑Fi inestable (no ante errores de negocio JSON). */
function esFalloRedRecuperable(err) {
  if (!err) return false;
  if (err instanceof TypeError) return true;
  const s = String(err.message || err);
  return /failed to fetch|load failed|networkerror|network error|aborted|timeout|timed out/i.test(s);
}

async function conUnReintentoRed(fn) {
  try {
    return await fn();
  } catch (e) {
    if (esFalloRedRecuperable(e)) {
      await sleep(450);
      return await fn();
    }
    throw e;
  }
}

/**
 * POST a Apps Script: el navegador no puede manejar bien el 302 (cuerpo perdido u opaqueredirect → HTTP 0).
 * En localhost `API` es /notas-gas-proxy (servidor_local.py + curl -L). Fuera: API_RELAY_URL o API_DIRECT en config.
 */
async function fetchAppsScriptPost(bodyObj, timeoutMs) {
  const bodyStr = JSON.stringify(bodyObj);
  const hdr = { "Content-Type": "text/plain;charset=utf-8" };
  const ctrl = new AbortController();
  const ms = timeoutMs != null ? timeoutMs : 35000;
  const timeoutId = setTimeout(() => ctrl.abort(), ms);
  try {
    return await fetch(urlApiNotas(), {
      method: "POST",
      headers: hdr,
      body: bodyStr,
      cache: "no-store",
      signal: ctrl.signal
    });
  } catch (error) {
    if (error && error.name === "AbortError") {
      throw new Error(
        "La solicitud tardo demasiado y fue cancelada. Revise la conexion o vuelva a intentarlo."
      );
    }
    throw error;
  } finally {
    clearTimeout(timeoutId);
  }
}

async function apiGet(params = {}) {
  return conUnReintentoRed(async () => {
    const merged = mergeAuthParams(params);
    asegurarCredencialesParaApiSalvoPing_(merged);
    /**
     * No usar GET con query: se pierde tras el 302. POST JSON + __readViaPost → doGet en servidor.
     * El envío real va por fetchAppsScriptPost para no perder el body en el redirect.
     */
    const body = { ...merged, __readViaPost: true };
    const res = await fetchAppsScriptPost(body);

    if (!res.ok) {
      const st = res.status;
      if (st === 0) {
        throw new Error(
          "Sin respuesta del servidor (HTTP 0). Use iniciar-servidor.bat (incluye proxy con Python + curl) " +
            "o defina API_RELAY_URL en src/config.js si publica la web fuera de localhost."
        );
      }
      throw new Error(`Error HTTP ${st}`);
    }

    const text = await res.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      throw new Error("La respuesta del servidor no es JSON válido.");
    }

    if (data && typeof data === "object" && !Array.isArray(data) && data.ok === false) {
      throw new Error(data.error || "Error del servidor");
    }

    return data;
  });
}

async function apiPost(payload, opts) {
  return conUnReintentoRed(async () => {
    const body = mergeAuthParams(payload);
    asegurarCredencialesParaApiSalvoPing_(body);
    const timeoutMs = opts && opts.timeoutMs != null ? opts.timeoutMs : undefined;
    // text/plain evita el preflight CORS de application/json; útil al abrir index.html como file://
    const res = await fetchAppsScriptPost(body, timeoutMs);

    if (!res.ok) {
      const st = res.status;
      if (st === 0) {
        throw new Error(
          "Sin respuesta del servidor (HTTP 0). Use iniciar-servidor.bat (proxy local) o API_RELAY_URL en config."
        );
      }
      throw new Error(`Error HTTP ${st}`);
    }

    const text = await res.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      throw new Error("La respuesta del servidor no es JSON válido.");
    }

    if (data && data.ok === false) {
      throw new Error(data.error || "Error desconocido");
    }

    return data;
  });
}

// =====================
// LISTA ESTUDIANTES
// =====================
function filtrarEstudiantes(lista, q) {
  const needle = normalizarTextoSimple(q);
  if (!needle) return lista;
  return lista.filter((e) => normalizarTextoSimple(e.nombre || "").includes(needle));
}

/** Orden alfabético por nombre (no depende del orden de filas en la hoja Estudiantes). */
function ordenarEstudiantesPorNombreAsc(lista) {
  if (!Array.isArray(lista) || lista.length < 2) return lista ? [...lista] : [];
  return [...lista].sort((a, b) =>
    String(a.nombre || "").localeCompare(String(b.nombre || ""), "es", { sensitivity: "base" })
  );
}

function construirTitulosGrupo() {
  const unicos = new Set();
  (state.cacheNotasGrupo || []).forEach((n) => {
    const titulo = String(n.titulo || "").trim();
    if (titulo) unicos.add(titulo);
  });
  state.cacheTitulosGrupo = Array.from(unicos);
}

function construirMapaNotasPorEstudiante() {
  const mapa = new Map();
  (state.cacheNotasGrupo || []).forEach((n) => {
    const id = String(n.estudiante || "");
    if (!id) return;
    if (!mapa.has(id)) mapa.set(id, new Map());
    const byTitulo = mapa.get(id);
    const titulo = String(n.titulo || "").trim();
    if (!titulo) return;
    const prev = byTitulo.get(titulo);
    if (!prev) {
      byTitulo.set(titulo, n);
      return;
    }
    const ta = new Date(prev.fecha || 0).getTime();
    const tb = new Date(n.fecha || 0).getTime();
    if (tb >= ta) byTitulo.set(titulo, n);
  });
  return mapa;
}

function renderListaEstudiantes(data) {
  const lista = document.getElementById("lista");
  const wrap = document.getElementById("grupo-board-wrap");
  const head = document.getElementById("grupo-board-head");
  const body = document.getElementById("grupo-board-body");
  if (!lista || !wrap || !head || !body) return;

  const filtrados = filtrarEstudiantes(data, state.filtroBusqueda);

  if (esEstudiante() && Array.isArray(data) && data.length > 0) {
    lista.classList.remove("is-hidden");
    wrap.classList.add("is-hidden");
    lista.innerHTML = `<div class="estudiante-solo-card card">
      <div id="est-panel-boletin" class="estudiante-cal-panel">
        <div class="boletin-inline__head no-print">
          <h3 class="boletin-inline__title">Notas generales<span id="est-boletin-periodo-chip" class="boletin-inline__periodo"></span></h3>
          <button type="button" id="btn-actualizar-boletin-est" class="btn btn--secondary btn--sm" title="Vuelve a pedir el resumen al servidor (útil si acaban de actualizar una nota)">
            Actualizar resumen
          </button>
        </div>
        <p id="estudiante-boletin-aviso" class="boletin-aviso"></p>
        <div id="estudiante-boletin-inline-host" class="boletin-tabla-host">
          <div class="state-msg state-msg--loading"><span class="spinner" aria-hidden="true"></span> Cargando notas generales…</div>
        </div>
      </div>
    </div>`;
    document.getElementById("btn-actualizar-boletin-est")?.addEventListener("click", () => {
      void cargarBoletinEstudiante();
    });
    actualizarTituloNotasGeneralesEstudiante_();
    void cargarBoletinEstudiante();
    return;
  }

  if (!Array.isArray(data) || data.length === 0) {
    wrap.classList.add("is-hidden");
    lista.classList.remove("is-hidden");
    lista.innerHTML = `No hay estudiantes para mostrar con este criterio.`;
    return;
  }

  if (filtrados.length === 0) {
    wrap.classList.add("is-hidden");
    lista.classList.remove("is-hidden");
    lista.innerHTML = `Ningún nombre coincide con la búsqueda.`;
    return;
  }

  const cols = state.cacheTitulosGrupo || [];
  const mapa = construirMapaNotasPorEstudiante();

  head.innerHTML = `
    <tr>
      <th class="group-board__student" scope="col">Estudiante</th>
      <th class="group-board__estado" scope="col">Estado</th>
      ${cols.map((t) => `<th class="group-board__nota" scope="col">${escapeHtml(t)}</th>`).join("")}
      <th class="group-board__detalle" scope="col">Detalle</th>
    </tr>
  `;

  body.innerHTML = filtrados.map((e) => {
    const est = e.estadoMatricula || "Activo";
    const notasEst = mapa.get(String(e.id)) || new Map();
    const notasCols = cols.map((t) => {
      const n = notasEst.get(t);
      const val = n && n.nota !== undefined && n.nota !== null && String(n.nota).trim() !== ""
        ? String(n.nota)
        : "";
      return `<td class="note-cell group-board__nota ${val ? "" : "is-empty"}">${escapeHtml(val || "—")}</td>`;
    }).join("");
    const celdaDetalle = `<td class="group-board__detalle"><button type="button" class="btn btn--secondary btn--sm btn-ver-detalle" data-estudiante-id="${escapeHtml(e.id)}">Ver detalle</button></td>`;
    return `
      <tr>
        <td class="group-board__student"><strong>${escapeHtml(e.nombre || "")}</strong></td>
        <td class="group-board__estado"><span class="${clasesEstadoMatriculaChip(est)}">${escapeHtml(est)}</span></td>
        ${notasCols}
        ${celdaDetalle}
      </tr>
    `;
  }).join("");

  lista.classList.add("is-hidden");
  wrap.classList.remove("is-hidden");
}

function aplicarFiltroBusqueda() {
  const buscar = document.getElementById("buscar");
  state.filtroBusqueda = buscar ? buscar.value : "";
  renderListaEstudiantes(state.cacheEstudiantes);
  if (state.appModo === "inasistencias" && state.cacheInasistenciasDia) {
    renderInasistenciasBodyFromCache();
  }
  if (state.appModo === "resumen-faltas" && state.cacheResumenFaltas) {
    renderTablaResumenFaltasFiltrada();
  }
  if (state.appModo === "resumen-area" && state.cacheResumenArea) {
    renderTablaResumenAreaFiltrada();
  }
}

async function cargarEstudiantes() {
  const lista = document.getElementById("lista");
  const wrap = document.getElementById("grupo-board-wrap");
  const btn = document.getElementById("btn-cargar");
  document.getElementById("bulk-area")?.classList.add("is-hidden");
  if (lista) {
    lista.classList.remove("is-hidden");
    const msgCarga = esEstudiante()
      ? "Cargando notas generales…"
      : "Cargando grupo…";
    lista.innerHTML = `<span class="state-msg state-msg--loading"><span class="spinner" aria-hidden="true"></span> ${msgCarga}</span>`;
  }
  if (wrap) wrap.classList.add("is-hidden");
  if (btn) {
    btn.disabled = true;
    btn.setAttribute("aria-busy", "true");
  }

  try {
    if (esEstudiante()) {
      const [data] = await Promise.all([
        apiGet({
          accion: "estudiantes",
          grupo: getGrupo()
        }),
        refrescarMateriasEstudianteCache()
      ]);
      const g = getGrupo();
      state.cacheEstudiantes = ordenarEstudiantesPorNombreAsc(
        (Array.isArray(data) ? data : []).map((s) => {
          const row = { ...s };
          if (!row.grupo) row.grupo = g;
          return row;
        })
      );
      state.cacheNotasGrupo = [];
      state.cacheTitulosGrupo = [];
      construirTitulosGrupo();
      llenarSelectMateriaEstudiante(leerMateriaExpedienteEstudiante());
      renderListaEstudiantes(state.cacheEstudiantes);
      renderCargaRapida();
      await refrescarGruposSelect();
      actualizarVisibilidadCargaRapida();
    } else {
      const [data] = await Promise.all([
        apiGet({
          accion: "estudiantes",
          grupo: getGrupo()
        }),
        cargarNotasGrupoMateria()
      ]);

      const g = getGrupo();
      state.cacheEstudiantes = ordenarEstudiantesPorNombreAsc(
        (Array.isArray(data) ? data : []).map((s) => {
          const row = { ...s };
          if (!row.grupo) row.grupo = g;
          return row;
        })
      );
      construirTitulosGrupo();
      renderListaEstudiantes(state.cacheEstudiantes);
      renderCargaRapida();
      await refrescarGruposSelect();
      actualizarVisibilidadCargaRapida();
    }
  } catch (error) {
    state.cacheEstudiantes = [];
    if (lista) {
      lista.classList.remove("is-hidden");
      lista.innerHTML = `<span class="state-msg state-msg--warn">${escapeHtml(`No se pudo cargar el grupo: ${error.message}`)}</span>`;
    }
    if (wrap) wrap.classList.add("is-hidden");
    renderCargaRapida();
    actualizarVisibilidadCargaRapida();
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.removeAttribute("aria-busy");
    }
  }
}

// =====================
// VISTAS
// =====================
function mostrarLista() {
  const vLista = document.getElementById("vista-lista");
  const vPanel = document.getElementById("vista-panel");
  if (vLista) vLista.classList.remove("is-hidden");
  if (vPanel) {
    vPanel.classList.add("is-hidden");
    vPanel.hidden = true;
  }
  state.estudianteActual = null;
  limpiarFotoFichaEstudiante_();
  document.getElementById("buscar")?.focus();
}

function mostrarPanel() {
  const vLista = document.getElementById("vista-lista");
  const vPanel = document.getElementById("vista-panel");
  if (vLista) vLista.classList.add("is-hidden");
  if (vPanel) {
    vPanel.classList.remove("is-hidden");
    vPanel.hidden = false;
  }
}

/** Muestra el nombre de la materia vigente en las pestañas Calificaciones y Faltas de la ficha. */
function actualizarMateriaEnFichaEstudiante() {
  const m = String(getMateria() || "").trim() || "—";
  const cal = document.getElementById("ficha-cal-materia-nombre");
  const fal = document.getElementById("ficha-faltas-materia-nombre");
  if (cal) cal.textContent = m;
  if (fal) fal.textContent = m;
  const trailer = document.getElementById("ficha-faltas-hint-trailer");
  if (trailer) {
    if (esModoSoloConsulta()) {
      trailer.innerHTML = " · Historial de días registrados.";
    } else {
      trailer.innerHTML =
        " · Historial de días registrados. Use <strong>Corregir</strong> si hubo un error en una fecha.";
    }
  }
}

/** Pestañas del panel: estado (y traslado solo staff) · calificaciones · faltas (esta materia). */
function aplicarFichaTab(which) {
  const map = [
    { key: "estado", btn: "ficha-tab-estado", panel: "ficha-panel-estado" },
    { key: "cal", btn: "ficha-tab-cal", panel: "ficha-panel-cal" },
    { key: "faltas", btn: "ficha-tab-faltas", panel: "ficha-panel-faltas" }
  ];
  const active = map.some((x) => x.key === which) ? which : "estado";
  map.forEach(({ key, btn, panel }) => {
    const b = document.getElementById(btn);
    const p = document.getElementById(panel);
    const on = key === active;
    if (b) {
      b.classList.toggle("is-active", on);
      b.setAttribute("aria-selected", on ? "true" : "false");
    }
    if (p) {
      if (on) p.removeAttribute("hidden");
      else p.setAttribute("hidden", "");
    }
  });
}

// =====================
// FICHA: FOTO DEL ESTUDIANTE (columna G / fotoUrl)
// =====================
/** En false: no se muestra la foto ni mensajes de diagnóstico; ocultar también `#boletin .student-photo-frame` en index.html si hace falta. */
const FICHA_ESTUDIANTE_FOTO_ACTIVA = false;

function inicialesDesdeNombreFicha_(nombre) {
  const parts = String(nombre || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  if (parts.length >= 2) {
    const a = parts[0][0] || "";
    const b = parts[parts.length - 1][0] || "";
    return (a + b).toUpperCase();
  }
  if (parts.length === 1 && parts[0].length >= 2) return parts[0].slice(0, 2).toUpperCase();
  if (parts.length === 1 && parts[0].length === 1) return parts[0].toUpperCase();
  return "?";
}

/**
 * IDs de archivo en Drive → varias URLs (Drive a menudo devuelve HTML con uc?export=view; la miniatura suele ir mejor en <img>).
 * El archivo debe estar como «Cualquiera con el enlace puede ver» (lector).
 */
function urlsDriveImgAlternativas_(fileId) {
  const id = String(fileId || "").trim();
  if (!id) return [];
  const enc = encodeURIComponent(id);
  return [
    `https://drive.google.com/thumbnail?id=${enc}&sz=w1000`,
    `https://drive.google.com/thumbnail?id=${enc}&sz=w400`,
    `https://drive.google.com/uc?export=view&id=${enc}`,
    `https://drive.google.com/uc?id=${enc}`,
    `https://drive.usercontent.google.com/uc?id=${enc}&export=view`
  ];
}

/** Valor de celda (hoja o JSON): quita comillas envolventes y unifica separadores para rutas locales. */
function normalizarFotoUrlCelda_(raw) {
  let s = String(raw ?? "").trim();
  if (!s) return "";
  if (
    (s.startsWith('"') && s.endsWith('"') && s.length >= 2) ||
    (s.startsWith("'") && s.endsWith("'") && s.length >= 2)
  ) {
    s = s.slice(1, -1).trim();
  }
  s = s.replace(/\\/g, "/");
  if (s.startsWith("./")) s = s.slice(2).trim();
  return s;
}

function fotoUrlCrudaDesdeEst_(row) {
  if (!row || typeof row !== "object") return "";
  const v = row.fotoUrl ?? row.foto_url;
  return v == null ? "" : v;
}

/** fotoUrl del objeto o, si falta, la misma fila en cacheEstudiantes (p. ej. expediente con objeto parcial). */
function fotoUrlResueltaParaFicha_(est) {
  const norm = (v) => normalizarFotoUrlCelda_(v);
  if (!est) return "";
  let u = norm(fotoUrlCrudaDesdeEst_(est));
  if (u) return u;
  const id = est.id != null ? String(est.id) : "";
  if (!id || !Array.isArray(state.cacheEstudiantes)) return "";
  const row = state.cacheEstudiantes.find((x) => String(x.id) === id);
  return row ? norm(fotoUrlCrudaDesdeEst_(row)) : "";
}

/** true si el JSON ya trae la propiedad (aunque esté vacía): distingue API vieja de celda vacía. */
function estudianteApiIncluyeCampoFoto_(est) {
  if (!est || typeof est !== "object") return false;
  if ("fotoUrl" in est || "foto_url" in est) return true;
  const id = est.id != null ? String(est.id) : "";
  if (!id || !Array.isArray(state.cacheEstudiantes)) return false;
  const row = state.cacheEstudiantes.find((x) => String(x.id) === id);
  return !!(row && (("fotoUrl" in row) || ("foto_url" in row)));
}

/** Rutas tipo images/... → URL absoluta según la página actual (evita fallos si la base del documento no es la esperada). */
function absolutizarRutaFotoSiRelativa_(s) {
  if (!s || /^https?:\/\//i.test(s) || /^data:/i.test(s) || /^\/\//.test(s)) return s;
  try {
    return new URL(s, window.location.href).href;
  } catch {
    return s;
  }
}

/** Lista de URLs a probar en orden (una sola si no es enlace de archivo Drive). */
function listaUrlsFotoEstudianteParaImg_(raw) {
  const s = normalizarFotoUrlCelda_(raw);
  if (!s) return [];
  if (/^https?:\/\//i.test(s)) {
    if (/drive\.google\.com\/thumbnail\?/i.test(s) || /drive\.google\.com\/uc\?/i.test(s)) {
      return [s];
    }
    try {
      const u = new URL(s);
      const host = u.hostname.replace(/^www\./, "").toLowerCase();
      if (host === "drive.google.com") {
        const mFile = u.pathname.match(/\/file\/d\/([^/]+)/);
        let id = mFile ? mFile[1] : "";
        if (!id) {
          const qp = u.searchParams.get("id");
          if (qp) id = String(qp).trim();
        }
        if (id) return urlsDriveImgAlternativas_(id);
      }
    } catch {
      /* no es URL absoluta válida */
    }
    return [s];
  }
  // Ruta local (ej. images/8-9B/2.webp): absoluta respecto a la página y, si hace falta, versión encodeURI.
  const abs = absolutizarRutaFotoSiRelativa_(s);
  const encoded = encodeURI(abs);
  return encoded !== abs ? [abs, encoded] : [abs];
}

function setFichaFotoHint_(texto) {
  const el = document.getElementById("ficha-foto-hint");
  if (!el) return;
  if (!texto) {
    el.textContent = "";
    el.classList.add("is-hidden");
    el.hidden = true;
    return;
  }
  el.textContent = texto;
  el.classList.remove("is-hidden");
  el.hidden = false;
}

function limpiarFotoFichaEstudiante_() {
  if (!FICHA_ESTUDIANTE_FOTO_ACTIVA) return;
  setFichaFotoHint_("");
  const img = document.getElementById("ficha-estudiante-foto");
  const fb = document.getElementById("ficha-estudiante-foto-fallback");
  if (img) {
    img.onload = null;
    img.onerror = null;
    img.removeAttribute("src");
    img.removeAttribute("referrerpolicy");
    img.classList.add("is-hidden");
    img.setAttribute("hidden", "");
    img.alt = "";
  }
  if (fb) {
    fb.textContent = "";
    fb.removeAttribute("hidden");
    fb.classList.remove("is-hidden");
  }
}

const MSG_FICHA_FOTO_API_SIN_CAMPO =
  "Foto: el despliegue de Apps Script no devuelve «fotoUrl». Abra el proyecto de la **misma** hoja NotasIEMFS → pegue el Code.gs actual → Implementar (web). En src/config.js la URL /exec debe ser **ese** proyecto. Luego «Cargar grupo».";
const MSG_FICHA_FOTO_CELDA_VACIA =
  "Foto: la API ya envía fotoUrl pero está vacío para este alumno. En Estudiantes, columna fotoUrl, **la fila de su id** debe tener la ruta (ej. images/8-9B/2.webp). Guarde la hoja y pulse «Cargar grupo».";
const MSG_FICHA_FOTO_RUTA_INVALIDA =
  "Foto: la ruta recibida no es usable. Revise texto en la celda fotoUrl (sin caracteres raros).";
const MSG_FICHA_FOTO_NO_CARGA =
  "Foto: no se pudo cargar el archivo. Use iniciar-servidor.bat (http://127.0.0.1:8080), compruebe que el .webp exista en la carpeta del proyecto y que la fila sea la del estudiante correcto. Si usa Drive: enlace con permiso de lector.";

function aplicarFotoFichaEstudiante_(est) {
  if (!FICHA_ESTUDIANTE_FOTO_ACTIVA) return;
  const img = document.getElementById("ficha-estudiante-foto");
  const fb = document.getElementById("ficha-estudiante-foto-fallback");
  if (!img || !fb) return;

  setFichaFotoHint_("");

  const mostrarFallback = (mensajeHint) => {
    img.removeAttribute("src");
    img.classList.add("is-hidden");
    img.setAttribute("hidden", "");
    img.alt = "";
    fb.removeAttribute("hidden");
    fb.classList.remove("is-hidden");
    if (mensajeHint) setFichaFotoHint_(mensajeHint);
  };

  const mostrarImagen = () => {
    setFichaFotoHint_("");
    img.classList.remove("is-hidden");
    img.removeAttribute("hidden");
    fb.setAttribute("hidden", "");
    fb.classList.add("is-hidden");
  };

  fb.textContent = inicialesDesdeNombreFicha_(est.nombre);
  const url = fotoUrlResueltaParaFicha_(est);
  if (!url) {
    mostrarFallback(
      estudianteApiIncluyeCampoFoto_(est) ? MSG_FICHA_FOTO_CELDA_VACIA : MSG_FICHA_FOTO_API_SIN_CAMPO
    );
    return;
  }

  const candidatas = listaUrlsFotoEstudianteParaImg_(url);
  if (!candidatas.length) {
    mostrarFallback(MSG_FICHA_FOTO_RUTA_INVALIDA);
    return;
  }

  img.alt = `Fotografía de ${String(est.nombre || "estudiante").trim()}`;
  img.removeAttribute("referrerpolicy");

  let idx = 0;
  let ultimaUrlProbada = "";
  const tryNext = () => {
    img.onload = null;
    img.onerror = null;
    if (idx >= candidatas.length) {
      const detalle = ultimaUrlProbada ? " Detalle: " + ultimaUrlProbada : "";
      mostrarFallback(MSG_FICHA_FOTO_NO_CARGA + detalle);
      return;
    }
    const u = candidatas[idx];
    idx += 1;
    ultimaUrlProbada = u;
    img.onload = () => {
      mostrarImagen();
    };
    img.onerror = () => {
      tryNext();
    };
    img.removeAttribute("src");
    img.src = u;
    if (img.complete && img.naturalWidth > 0) {
      mostrarImagen();
    }
  };

  tryNext();
}

// =====================
// ABRIR ESTUDIANTE
// =====================
async function abrirEstudiante(e) {
  if (esEstudiante()) {
    const mat = String(getMateria() || "").trim();
    if (!mat) {
      toast("Seleccione una materia en la barra superior antes de abrir el expediente.", "err");
      return;
    }
  }
  if (!e.estadoMatricula) e.estadoMatricula = "Activo";
  state.estudianteActual = e;
  if (!state.estudianteActual.grupo) state.estudianteActual.grupo = getGrupo();
  mostrarPanel();
  actualizarEncabezadoImpresion();

  document.getElementById("nombre").textContent = e.nombre;
  aplicarFotoFichaEstudiante_(e);
  actualizarChipEstadoEnPanel(e.estadoMatricula);
  const selEst = document.getElementById("estado-matricula");
  if (selEst) {
    selEst.value = e.estadoMatricula;
    if (selEst.value !== e.estadoMatricula) selEst.value = "Activo";
  }
  const motEst = document.getElementById("estado-motivo");
  if (motEst) motEst.value = "";
  const claveEst = document.getElementById("estado-clave-admin");
  if (claveEst) claveEst.value = "";
  aplicarVisibilidadEstadoMatriculaPanel();

  const cg = document.getElementById("ctx-grupo");
  const cm = document.getElementById("ctx-materia");
  if (cg) cg.textContent = state.estudianteActual.grupo || getGrupo();
  if (cm) cm.textContent = getMateria();
  actualizarMateriaEnFichaEstudiante();
  const trasladoActual = document.getElementById("traslado-grupo-actual-label");
  if (trasladoActual) trasladoActual.textContent = state.estudianteActual.grupo || "—";
  poblarSelectTrasladoDestino();
  aplicarVisibilidadTrasladoCurso();
  aplicarFichaTab("estado");
  resetFormAnadirNotaFicha();

  const resumenEl = document.getElementById("resumen");
  const notasEl = document.getElementById("notas");
  const inaAlumnoEl = document.getElementById("ina-estudiante-lista");
  if (resumenEl) {
    resumenEl.innerHTML = `<div class="state-msg state-msg--loading"><span class="spinner"></span> Cargando resumen…</div>`;
  }
  if (notasEl) {
    notasEl.innerHTML = `<div class="state-msg state-msg--loading"><span class="spinner"></span> Cargando calificaciones…</div>`;
  }
  if (inaAlumnoEl) {
    inaAlumnoEl.innerHTML = `<div class="state-msg state-msg--loading"><span class="spinner"></span> Cargando inasistencias…</div>`;
  }

  const materia = getMateria();

  try {
    const perCal = periodoParamParaApiCalificaciones();
    const [r, listaNotas, inaHist] = await Promise.all([
      apiGet({ accion: "resumen", id: e.id, materia, ...perCal }),
      apiGet({ accion: "notas", id: e.id, materia, ...perCal }),
      apiGet({
        accion: "inasistenciasEstudiante",
        estudiante: e.id,
        id: e.id,
        materia,
        limite: 48,
        ...perCal
      })
    ]);

    if (resumenEl) {
      resumenEl.innerHTML = `
        <div class="resumen-grid">
          <div class="resumen-item">
            <div class="resumen-item__label">Seguimiento</div>
            <div class="resumen-item__value">${escapeHtml(r.seguimiento ?? "—")}</div>
          </div>
          <div class="resumen-item">
            <div class="resumen-item__label">Actitudinal</div>
            <div class="resumen-item__value">${escapeHtml(r.actitudinal ?? "—")}</div>
          </div>
          <div class="resumen-item">
            <div class="resumen-item__label">Prueba</div>
            <div class="resumen-item__value">${escapeHtml(r.prueba ?? "—")}</div>
          </div>
          <div class="resumen-item resumen-item--final">
            <div class="resumen-item__label">Definitiva</div>
            <div class="resumen-item__value">${escapeHtml(r.final ?? "—")}</div>
          </div>
        </div>
      `;
    }

    state.fichaListaNotas = Array.isArray(listaNotas) ? listaNotas : [];
    if (!esModoSoloConsulta()) {
      await cargarNotasGrupoMateria();
      poblarSelectorEvaluacionFicha(listaNotas);
      actualizarHintRangoFechaFichaDocente_();
    }

    renderNotas(listaNotas);
    renderInasistenciasEstudiantePanel(Array.isArray(inaHist) ? inaHist : []);
  } catch (error) {
    state.fichaListaNotas = [];
    const msg = escapeHtml(error.message);
    if (resumenEl) {
      resumenEl.innerHTML = `<div class="state-msg state-msg--warn">No se pudo cargar el resumen: ${msg}</div>`;
    }
    if (notasEl) {
      notasEl.innerHTML = `<div class="state-msg state-msg--warn">No se pudieron cargar las calificaciones: ${msg}</div>`;
    }
    if (inaAlumnoEl) {
      inaAlumnoEl.innerHTML = `<div class="state-msg state-msg--warn">No se pudieron cargar las inasistencias: ${msg}</div>`;
    }
  }
}

// =====================
// RENDER DE NOTAS
// =====================
function renderNotas(listaNotas) {
  const notasEl = document.getElementById("notas");
  if (!notasEl) return;

  if (!Array.isArray(listaNotas) || listaNotas.length === 0) {
    notasEl.innerHTML = esModoSoloConsulta()
      ? `<div class="card state-msg">Aún no hay calificaciones registradas para esta materia.</div>`
      : `<div class="card state-msg">Aún no hay calificaciones registradas. Puede usar arriba <strong>Registrar o corregir calificación (solo este estudiante)</strong> (elegir la evaluación del grupo), o el <strong>Registro masivo de calificaciones</strong> en la vista del grupo (misma actividad para todos; vacío cuenta como 0 en los cálculos).</div>`;
    return;
  }

  notasEl.innerHTML = "";
  listaNotas.forEach((n) => {
    const id = notaIdDe(n);
    const criterio = snapshotAntesParaBuscar(n);
    const fecha = n.fecha ? fechaTextoMostrarNota_(n.fecha) : "";
    const titulo = escapeHtml(n.titulo || "Sin título");
    const tipo = escapeHtml(n.tipo || "—");
    const notaVal = escapeHtml(n.nota ?? "—");
    const tipoClass = claseTipo(n.tipo);

    const card = document.createElement("article");
    card.className = `note-card ${tipoClass}`;
    card.innerHTML = `
      <h4 class="note-card__title">${titulo}</h4>
      <div class="note-card__meta">
        <span class="badge-tipo">${tipo}</span>
        ${fecha ? `<span> · ${escapeHtml(fecha)}</span>` : ""}
      </div>
      <div class="note-card__score">${notaVal}</div>
      <div class="note-card__actions"></div>
    `;

    const actions = card.querySelector(".note-card__actions");
    const btnHist = document.createElement("button");
    btnHist.type = "button";
    btnHist.className = "btn btn--secondary btn--sm";
    btnHist.textContent = "Historial";
    btnHist.addEventListener("click", () => abrirHistorial(n));

    const btnEdit = document.createElement("button");
    btnEdit.type = "button";
    btnEdit.className = "btn btn--primary btn--sm";
    btnEdit.textContent = "Corregir";
    const puedeCorregir = Boolean(id || criterio);
    btnEdit.disabled = !puedeCorregir;
    if (!puedeCorregir) {
      btnEdit.title = "No disponible para esta nota";
    } else {
      btnEdit.title = "";
    }

    btnEdit.addEventListener("click", () => abrirEdicion(n));

    actions.appendChild(btnHist);
    if (!esModoSoloConsulta()) {
      actions.appendChild(btnEdit);
    }
    notasEl.appendChild(card);
  });
}

// =====================
// MODAL
// =====================
let modalElementoFocoOrigen = null;
let modalKeydownTrampaTab = null;

function listaEnfocablesModal() {
  const bd = document.getElementById("modal-backdrop");
  if (!bd || bd.classList.contains("is-hidden") || bd.hidden) return [];
  const modal = bd.querySelector(".modal");
  if (!modal) return [];
  const sel =
    'button:not([disabled]), a[href], input:not([type="hidden"]):not([disabled]), select:not([disabled]), textarea:not([disabled])';
  return Array.from(modal.querySelectorAll(sel)).filter((el) => {
    if (el.closest("[hidden]")) return false;
    const st = window.getComputedStyle(el);
    return st.visibility !== "hidden" && st.display !== "none";
  });
}

function enfocarContenidoModal() {
  const body = document.getElementById("modal-body");
  const bd = document.getElementById("modal-backdrop");
  if (!bd || bd.hidden) return;
  const enCuerpo = body?.querySelector(
    'input:not([type="hidden"]):not([disabled]), select:not([disabled]), textarea:not([disabled]), button.btn--primary, a[href]'
  );
  const cerrar = document.getElementById("modal-cerrar");
  (enCuerpo || cerrar)?.focus();
}

function cerrarModal() {
  if (modalKeydownTrampaTab) {
    document.removeEventListener("keydown", modalKeydownTrampaTab, true);
    modalKeydownTrampaTab = null;
  }
  const bd = document.getElementById("modal-backdrop");
  if (bd) {
    bd.classList.add("is-hidden");
    bd.hidden = true;
  }
  document.getElementById("modal-body").innerHTML = "";
  const prev = modalElementoFocoOrigen;
  modalElementoFocoOrigen = null;
  if (prev && typeof prev.focus === "function") {
    try {
      prev.focus();
    } catch (_) {
      /* ignore */
    }
  }
}

function abrirModal(titulo, innerHtml) {
  if (modalKeydownTrampaTab) {
    document.removeEventListener("keydown", modalKeydownTrampaTab, true);
    modalKeydownTrampaTab = null;
  }
  const t = document.getElementById("modal-title");
  const body = document.getElementById("modal-body");
  const bd = document.getElementById("modal-backdrop");
  const backdropYaAbierto = bd && !bd.hidden && !bd.classList.contains("is-hidden");
  if (!backdropYaAbierto) {
    modalElementoFocoOrigen = document.activeElement;
  }
  if (t) t.textContent = titulo;
  if (body) body.innerHTML = innerHtml;
  if (bd) {
    bd.classList.remove("is-hidden");
    bd.hidden = false;
  }
  modalKeydownTrampaTab = (ev) => {
    if (ev.key !== "Tab") return;
    const lista = listaEnfocablesModal();
    if (lista.length === 0) return;
    const first = lista[0];
    const last = lista[lista.length - 1];
    if (ev.shiftKey) {
      if (document.activeElement === first) {
        ev.preventDefault();
        last.focus();
      }
    } else if (document.activeElement === last) {
      ev.preventDefault();
      first.focus();
    }
  };
  document.addEventListener("keydown", modalKeydownTrampaTab, true);
  requestAnimationFrame(() => enfocarContenidoModal());
}

// =====================
// EDICIÓN (requiere id en cada nota desde Apps Script)
// =====================
function abrirEdicion(n) {
  if (esModoSoloConsulta()) {
    toast(
      esEstudiante()
        ? "El acceso de estudiante es solo consulta."
        : "El acceso de coordinación es solo consulta.",
      "err"
    );
    return;
  }
  const id = notaIdDe(n);
  const antes = snapshotAntesParaBuscar(n);
  if (!id && !antes) {
    toast("No se puede corregir esta nota desde aquí.", "err");
    return;
  }

  const titulo = escapeHtml(n.titulo || "");
  const nota = escapeHtml(n.nota ?? "");
  const tipo = escapeHtml(n.tipo || "Seguimiento");
  const fechaInputVal = escapeHtml(fechaIsoParaInputDate(n.fecha));

  abrirModal("Corregir calificación", `
    <div class="form-grid">
      <div class="field-group field-group--full">
        <label for="edit-titulo">Título</label>
        <input type="text" id="edit-titulo" class="input" value="${titulo}">
      </div>
      <div class="field-group">
        <label for="edit-nota">Nota</label>
        <input type="number" id="edit-nota" class="input" step="0.1" min="0" max="5" value="${nota}">
      </div>
      <div class="field-group">
        <label for="edit-tipo">Tipo</label>
        <select id="edit-tipo" class="input">
          <option value="Seguimiento" ${tipo === "Seguimiento" ? "selected" : ""}>Seguimiento</option>
          <option value="Actitudinal" ${tipo === "Actitudinal" ? "selected" : ""}>Actitudinal</option>
          <option value="Prueba" ${tipo === "Prueba" ? "selected" : ""}>Prueba</option>
        </select>
      </div>
      <div class="field-group field-group--full">
        <label for="edit-fecha">Fecha de la calificación</label>
        <input type="date" id="edit-fecha" class="input" value="${fechaInputVal}" required>
        <p class="empty-hint" style="margin-top:0.35rem">En modo por periodos, la fecha debe caer dentro del rango de esa hoja (P1, P2 o P3).</p>
      </div>
      <div class="field-group field-group--full">
        <label for="edit-motivo">Motivo del cambio (opcional)</label>
        <input type="text" id="edit-motivo" class="input" placeholder="Ej. Revisión de taller">
      </div>
    </div>
    <button type="button" class="btn btn--primary btn--block" id="edit-guardar">Guardar cambios</button>
  `);

  vincularValidacionNotaInput(document.getElementById("edit-nota"));

  document.getElementById("edit-guardar").addEventListener("click", async () => {
    const nt = document.getElementById("edit-titulo").value.trim();
    const nn = document.getElementById("edit-nota").value;
    const tt = document.getElementById("edit-tipo").value;
    const fechaStr = document.getElementById("edit-fecha")?.value?.trim() || "";
    const motivo = document.getElementById("edit-motivo").value.trim();

    if (!nt) {
      toast("El título no puede quedar vacío.", "err");
      return;
    }
    const num = parseNotaDesdeCampo(nn);
    if (num === null) {
      toast("Nota no válida.", "err");
      return;
    }
    if (!notaEnRango(num)) {
      toast("La nota debe estar entre 0 y 5.", "err");
      return;
    }
    if (!fechaStr) {
      toast("Indique la fecha de la calificación.", "err");
      return;
    }
    const notaEnvio = redondearNotaPasoDecima(num);

    const btn = document.getElementById("edit-guardar");
    if (btn) btn.disabled = true;
    try {
      const payload = {
        accion: "actualizar",
        estudiante: state.estudianteActual.id,
        materia: getMateria(),
        grupo: state.estudianteActual.grupo || getGrupo(),
        titulo: nt,
        nota: notaEnvio,
        tipo: tt,
        motivo: motivo || undefined,
        ...periodoParamParaApiDocente(),
        fecha: fechaStr,
        fechaCalificacion: fechaStr,
        fechaNota: fechaStr
      };
      if (id) {
        payload.notaId = id;
      } else {
        payload.antes = antes;
      }

      await apiPost(payload);
      cerrarModal();
      toast("Informacion guardada: cambio aplicado.", "ok");
      await abrirEstudiante(state.estudianteActual);
    } catch (error) {
      toast(error.message, "err");
    } finally {
      if (btn) btn.disabled = false;
    }
  });
}

// =====================
// HISTORIAL (acción "historial" en Apps Script)
// =====================
async function abrirHistorial(n) {
  const id = notaIdDe(n);
  abrirModal("Historial de cambios", `<div class="state-msg state-msg--loading"><span class="spinner"></span> Cargando…</div>`);

  if (!id) {
    document.getElementById("modal-body").innerHTML = `
      <p class="empty-hint">El historial no está disponible para esta nota.</p>
    `;
    return;
  }

  try {
    const data = await apiGet({ accion: "historial", notaId: id });
    const items = Array.isArray(data) ? data : data?.items ?? data?.registros ?? [];

    if (!items.length) {
      document.getElementById("modal-body").innerHTML = `
        <p class="empty-hint">No hay cambios registrados aún.</p>
      `;
      return;
    }

    const lis = items
      .map((row) => {
        const fecha = row.fecha || row.cuando || row.timestamp || "";
        const txt = row.detalle || row.mensaje || row.texto || row.descripcion || JSON.stringify(row);
        const fLabel = fecha ? new Date(fecha).toLocaleString("es-CO") : "—";
        return `<li><time>${escapeHtml(fLabel)}</time><p>${escapeHtml(String(txt))}</p></li>`;
      })
      .join("");

    document.getElementById("modal-body").innerHTML = `<ul class="timeline">${lis}</ul>`;
  } catch (error) {
    document.getElementById("modal-body").innerHTML = `
      <p class="state-msg state-msg--warn">${escapeHtml(error.message)}</p>
    `;
  }
}

async function confirmarTrasladoEstudiante() {
  if (getAuthTipo() === "docente") {
    toast("Solo coordinación puede realizar traslados de curso.", "err");
    return;
  }
  if (!state.estudianteActual) return;
  const sel = document.getElementById("traslado-grupo-destino");
  const btn = document.getElementById("btn-traslado-curso");
  const mot = document.getElementById("reubicacion-motivo");
  const grupoDestino = sel && sel.value ? String(sel.value).trim() : "";
  const motivo = mot && mot.value.trim() ? mot.value.trim() : undefined;
  const grupoOrigen = state.estudianteActual.grupo || getGrupo();

  if (!grupoDestino) {
    toast("Seleccione el curso de destino.", "err");
    return;
  }
  if (normalizarTextoSimple(grupoDestino) === normalizarTextoSimple(grupoOrigen)) {
    toast("El curso de destino debe ser distinto al actual.", "err");
    return;
  }

  const claveAdminEl = document.getElementById("traslado-clave-admin");
  const claveAdmin = claveAdminEl && claveAdminEl.value ? String(claveAdminEl.value).trim() : "";
  if (!claveAdmin) {
    toast("Ingrese la clave de administración para confirmar el traslado.", "err");
    if (claveAdminEl) claveAdminEl.focus();
    return;
  }

  if (btn) btn.disabled = true;
  try {
    const data = await apiPost({
      accion: "migrarGrupoNotas",
      estudiante: state.estudianteActual.id,
      grupoNuevo: grupoDestino,
      grupoAnterior: grupoOrigen,
      motivo,
      claveAdmin
    });
    const n = data && data.actualizadas != null ? Number(data.actualizadas) : 0;
    state.estudianteActual.grupo = grupoDestino;

    const ix = state.cacheEstudiantes.findIndex((s) => String(s.id) === String(state.estudianteActual.id));
    if (ix >= 0) {
      if (normalizarTextoSimple(grupoDestino) !== normalizarTextoSimple(getGrupo())) {
        state.cacheEstudiantes.splice(ix, 1);
      } else {
        state.cacheEstudiantes[ix].grupo = grupoDestino;
      }
    }

    let msg =
      n > 0
        ? `Traslado completado: ${n} registro(s) de calificación actualizados. Curso del estudiante: «${grupoDestino}».`
        : `Traslado completado. Curso del estudiante: «${grupoDestino}» (las calificaciones ya estaban en ese curso).`;
    if (normalizarTextoSimple(grupoDestino) !== normalizarTextoSimple(getGrupo())) {
      msg += ` Para verlo en la tabla del grupo, arriba elija «${grupoDestino}» y pulse «Cargar grupo».`;
    }
    toast(`Informacion guardada. ${msg}`, "ok");

    if (mot) mot.value = "";
    if (claveAdminEl) claveAdminEl.value = "";
    await cargarNotasGrupoMateria();
    construirTitulosGrupo();
    state.cacheEstudiantes = ordenarEstudiantesPorNombreAsc(state.cacheEstudiantes);
    renderListaEstudiantes(state.cacheEstudiantes);
    renderCargaRapida();
    await abrirEstudiante(state.estudianteActual);
  } catch (error) {
    toast(error.message || "No se pudo completar el traslado.", "err");
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function guardarEstadoMatricula() {
  if (getAuthTipo() === "docente") {
    toast("Solo administración puede cambiar el estado de matrícula (valor único en la institución).", "err");
    return;
  }
  if (!state.estudianteActual) return;
  const claveAdminEl = document.getElementById("estado-clave-admin");
  const claveAdmin = claveAdminEl && claveAdminEl.value ? String(claveAdminEl.value).trim() : "";
  if (!claveAdmin) {
    toast("Ingrese la clave de administración para registrar el cambio de estado.", "err");
    if (claveAdminEl) claveAdminEl.focus();
    return;
  }
  const sel = document.getElementById("estado-matricula");
  const motivoEl = document.getElementById("estado-motivo");
  const estadoMatricula = sel ? sel.value : "Activo";
  const btn = document.getElementById("btn-guardar-estado");
  if (btn) btn.disabled = true;
  try {
    await apiPost({
      accion: "estadoEstudiante",
      estudiante: state.estudianteActual.id,
      grupo: state.estudianteActual.grupo || getGrupo(),
      estadoMatricula,
      motivo: motivoEl && motivoEl.value.trim() ? motivoEl.value.trim() : undefined,
      claveAdmin
    });
    state.estudianteActual.estadoMatricula = estadoMatricula;
    const ix = state.cacheEstudiantes.findIndex((s) => String(s.id) === String(state.estudianteActual.id));
    if (ix >= 0) state.cacheEstudiantes[ix].estadoMatricula = estadoMatricula;
    if (motivoEl) motivoEl.value = "";
    if (claveAdminEl) claveAdminEl.value = "";
    const txEst = document.getElementById("estado-valor-texto");
    if (txEst) txEst.textContent = estadoMatricula;
    actualizarChipEstadoEnPanel(estadoMatricula);
    actualizarEncabezadoImpresion();
    renderListaEstudiantes(state.cacheEstudiantes);
    toast("Informacion guardada: situacion de matricula actualizada.", "ok");
  } catch (error) {
    toast(error.message, "err");
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function abrirHistorialEstadoMatricula() {
  if (!state.estudianteActual) return;
  abrirModal("Historial de matrícula", `<div class="state-msg state-msg--loading"><span class="spinner"></span> Cargando…</div>`);
  try {
    const data = await apiGet({ accion: "historialEstudiante", estudiante: state.estudianteActual.id });
    const items = Array.isArray(data) ? data : data?.items ?? data?.registros ?? [];

    if (!items.length) {
      document.getElementById("modal-body").innerHTML = `
        <p class="empty-hint">No hay cambios de matrícula registrados.</p>
      `;
      return;
    }

    const lis = items
      .map((row) => {
        const fecha = row.fecha || row.cuando || row.timestamp || "";
        const txt = row.detalle || row.mensaje || row.texto || row.descripcion || JSON.stringify(row);
        const fLabel = fecha ? new Date(fecha).toLocaleString("es-CO") : "—";
        return `<li><time>${escapeHtml(fLabel)}</time><p>${escapeHtml(String(txt))}</p></li>`;
      })
      .join("");

    document.getElementById("modal-body").innerHTML = `<ul class="timeline">${lis}</ul>`;
  } catch (error) {
    document.getElementById("modal-body").innerHTML = `
      <p class="state-msg state-msg--warn">${escapeHtml(error.message)}</p>
    `;
  }
}

// =====================
// ACCESO
// =====================
function mensajeListaInicialTrasEntrar_() {
  if (esEstudiante()) {
    return `<span class="state-msg state-msg--loading"><span class="spinner" aria-hidden="true"></span> Cargando tus notas generales…</span>`;
  }
  return `Selecciona grupo y pulsa “Cargar grupo”.`;
}

function entrarAlSistema() {
  const gate = document.getElementById("login-gate");
  const shell = document.getElementById("app-shell");
  if (gate) {
    gate.classList.add("is-hidden");
    gate.hidden = true;
  }
  if (shell) shell.hidden = false;
  const lista = document.getElementById("lista");
  const wrap = document.getElementById("grupo-board-wrap");
  if (lista) {
    lista.classList.remove("is-hidden");
    lista.innerHTML = mensajeListaInicialTrasEntrar_();
  }
  if (wrap) wrap.classList.add("is-hidden");
}

async function aplicarEntradaSesionTrasPing_(tipo, ping) {
  let persistTipo = false;
  try {
    persistTipo =
      localStorage.getItem(STORAGE.authTipo) !== null ||
      localStorage.getItem(STORAGE.authPersist) !== null;
  } catch (_) {
    /* ignore */
  }
  setAuthTipo(tipo, persistTipo);
  if (tipo === "estudiante") {
    state.estudianteNombreSesion = String(ping.nombre || "").trim();
    state.cacheDocenteAsignaciones = null;
  } else if (tipo === "docente" && getStoredDocenteId()) {
    await cargarAsignacionesDocenteCache();
    state.estudianteNombreSesion = "";
  } else {
    state.cacheDocenteAsignaciones = null;
    state.estudianteNombreSesion = "";
  }
  entrarAlSistema();
  actualizarBadgeSesion();
  aplicarModoSoloLecturaCoordinacionUI();
  aplicarModoEstudianteUI();
  aplicarVisibilidadTrasladoCurso();
  await refrescarGruposSelect();
  aplicarMateriasSegunGrupoDesdeCache();
  await refrescarPeriodosAcademicosSelect();
  aplicarAppModoUI();
  if (tipo === "estudiante") {
    await cargarEstudiantes();
  }
}

/** Docente/coordinación: solo si hay clave de staff guardada. */
async function verificarSesionGuardada() {
  if (!getStoredClave()) return;
  try {
    const ping = await apiGet({ accion: "ping" });
    const tipo = tipoAuthDesdePing(ping, getStoredDocenteId());
    await aplicarEntradaSesionTrasPing_(tipo, ping);
    const buildSrv = ping && typeof ping === "object" ? String(ping.notasIemfsBuild || "").trim() : "";
    if (buildSrv !== NOTAS_IEMFS_BUILD_ESPERADO) {
      let yaAviso;
      try {
        yaAviso = sessionStorage.getItem("notas_iemfs_warn_gas_build");
      } catch {
        yaAviso = null;
      }
      if (!yaAviso) {
        try {
          sessionStorage.setItem("notas_iemfs_warn_gas_build", "1");
        } catch {
          /* ignore */
        }
        toast(
          "El despliegue de Apps Script no coincide con esta app (falta guardado en bloque corregido). Pegue el Code.gs actual en el editor, despliegue una nueva versión y, si cambia la URL, actualice API en config.js.",
          "err"
        );
      }
    }
  } catch (_) {
    clearCredencialesStaff();
    state.cacheDocenteAsignaciones = null;
    state.estudianteNombreSesion = "";
    const puedeEst =
      String(getStoredEstudianteDocumento() || "").trim() &&
      String(getStoredEstudiantePin() || "").trim();
    if (!puedeEst) {
      try {
        sessionStorage.removeItem(STORAGE.authTipo);
        localStorage.removeItem(STORAGE.authTipo);
      } catch (e2) {
        /* ignore */
      }
    }
  }
}

/**
 * Estudiante con documento/PIN guardados y sin sesión de staff: entrar al recargar (no había clave → antes no corría verificación).
 */
async function verificarSesionEstudianteGuardada() {
  if (getStoredClave()) return;
  const doc = String(getStoredEstudianteDocumento() || "").trim();
  const pin = String(getStoredEstudiantePin() || "").trim();
  if (!doc || !pin) return;

  try {
    const ping = await apiGet({ accion: "ping", documento: doc, pin });
    const tipo = tipoAuthDesdePing(ping, "");
    if (tipo !== "estudiante") return;
    await aplicarEntradaSesionTrasPing_("estudiante", ping);
  } catch (_) {
    /* Fallo de red o credenciales: no borrar almacenamiento persistido aquí. */
  }
}

function loginModoEstudianteActivo() {
  const tab = document.getElementById("login-tab-est");
  return Boolean(tab && tab.classList.contains("is-active"));
}

async function onLoginSubmit(ev) {
  ev.preventDefault();
  const errEl = document.getElementById("login-error");
  const inputClave = document.getElementById("login-clave");
  const inputCodigo = document.getElementById("login-docente-codigo");
  const inputDocEst = document.getElementById("login-est-documento");
  const inputPinEst = document.getElementById("login-est-pin");
  const recordar = document.getElementById("login-recordar");
  const btn = document.getElementById("btn-login");
  if (errEl) {
    errEl.classList.add("is-hidden");
    errEl.textContent = "";
  }

  const persist = !!(recordar && recordar.checked);

  if (loginModoEstudianteActivo()) {
    const docEst = inputDocEst ? inputDocEst.value.trim() : "";
    const pinEst = inputPinEst ? inputPinEst.value.trim() : "";
    if (!docEst || !pinEst) {
      if (errEl) {
        errEl.textContent = "Ingrese documento y PIN.";
        errEl.classList.remove("is-hidden");
      }
      return;
    }

    clearCredencialesStaff();
    setStoredEstudianteDocumento(docEst, persist);
    setStoredEstudiantePin(pinEst, persist);
    setAuthTipo("estudiante", persist);

    if (btn) {
      btn.disabled = true;
      btn.setAttribute("aria-busy", "true");
    }
    try {
      const ping = await apiGet({ accion: "ping", documento: docEst, pin: pinEst });
      const tipo = tipoAuthDesdePing(ping, "");
      setAuthTipo(tipo, persist);
      if (tipo !== "estudiante") {
        throw new Error("No se pudo validar el acceso de estudiante.");
      }
      state.estudianteNombreSesion = String(ping.nombre || "").trim();
      state.cacheDocenteAsignaciones = null;
      entrarAlSistema();
      actualizarBadgeSesion();
      aplicarModoSoloLecturaCoordinacionUI();
      aplicarModoEstudianteUI();
      aplicarVisibilidadTrasladoCurso();
      await refrescarGruposSelect();
      aplicarMateriasSegunGrupoDesdeCache();
      await refrescarPeriodosAcademicosSelect();
      aplicarAppModoUI();
      await cargarEstudiantes();
    } catch (e) {
      clearStoredClave();
      state.cacheDocenteAsignaciones = null;
      state.estudianteNombreSesion = "";
      if (errEl) {
        errEl.textContent = e.message || "No se pudo validar el acceso.";
        errEl.classList.remove("is-hidden");
      }
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.removeAttribute("aria-busy");
      }
    }
    return;
  }

  const clave = inputClave ? inputClave.value.trim() : "";
  const codigoDoc = inputCodigo ? inputCodigo.value.trim() : "";
  if (!clave) return;
  if (!codigoDoc) {
    if (errEl) {
      errEl.textContent =
        "Ingrese su código de usuario: docente (p. ej. D001) o coordinación (p. ej. " + CODIGO_COORDINADOR + ").";
      errEl.classList.remove("is-hidden");
    }
    return;
  }

  clearCredencialesEstudiante();
  setStoredClave(clave, persist);
  setStoredDocenteId(codigoDoc, persist);
  state.estudianteNombreSesion = "";

  if (btn) {
    btn.disabled = true;
    btn.setAttribute("aria-busy", "true");
  }
  try {
    const ping = await apiGet({ accion: "ping" });
    const tipo = tipoAuthDesdePing(ping, codigoDoc);
    setAuthTipo(tipo, persist);
    if (tipo === "coordinador") {
      state.cacheDocenteAsignaciones = null;
      setStoredDocenteNombre("", persist);
    } else {
      await cargarAsignacionesDocenteCache();
    }
    entrarAlSistema();
    actualizarBadgeSesion();
    aplicarModoSoloLecturaCoordinacionUI();
    aplicarModoEstudianteUI();
    aplicarVisibilidadTrasladoCurso();
    await refrescarGruposSelect();
    await refrescarPeriodosAcademicosSelect();
    aplicarAppModoUI();
  } catch (e) {
    clearStoredClave();
    state.cacheDocenteAsignaciones = null;
    state.estudianteNombreSesion = "";
    if (errEl) {
      errEl.textContent = e.message || "No se pudo validar el acceso.";
      errEl.classList.remove("is-hidden");
    }
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.removeAttribute("aria-busy");
    }
  }
}

function configurarPestanasLogin() {
  const tabStaff = document.getElementById("login-tab-staff");
  const tabEst = document.getElementById("login-tab-est");
  const panelStaff = document.getElementById("login-fields-staff");
  const panelEst = document.getElementById("login-fields-est");

  function activarStaff() {
    tabStaff?.classList.add("is-active");
    tabEst?.classList.remove("is-active");
    tabStaff?.setAttribute("aria-selected", "true");
    tabEst?.setAttribute("aria-selected", "false");
    panelStaff?.removeAttribute("hidden");
    panelEst?.setAttribute("hidden", "");
    panelEst?.classList.add("is-hidden");
    panelStaff?.classList.remove("is-hidden");
    document.getElementById("login-docente-codigo")?.setAttribute("required", "");
    document.getElementById("login-clave")?.setAttribute("required", "");
    document.getElementById("login-est-documento")?.removeAttribute("required");
    document.getElementById("login-est-pin")?.removeAttribute("required");
    const chkPin = document.getElementById("login-mostrar-pin");
    if (chkPin) chkPin.checked = false;
    const pinEl = document.getElementById("login-est-pin");
    if (pinEl instanceof HTMLInputElement) pinEl.type = "password";
  }

  function activarEst() {
    tabEst?.classList.add("is-active");
    tabStaff?.classList.remove("is-active");
    tabEst?.setAttribute("aria-selected", "true");
    tabStaff?.setAttribute("aria-selected", "false");
    panelEst?.removeAttribute("hidden");
    panelStaff?.setAttribute("hidden", "");
    panelStaff?.classList.add("is-hidden");
    panelEst?.classList.remove("is-hidden");
    document.getElementById("login-est-documento")?.setAttribute("required", "");
    document.getElementById("login-est-pin")?.setAttribute("required", "");
    document.getElementById("login-docente-codigo")?.removeAttribute("required");
    document.getElementById("login-clave")?.removeAttribute("required");
    const chkCl = document.getElementById("login-mostrar-clave");
    if (chkCl) chkCl.checked = false;
    const clv = document.getElementById("login-clave");
    if (clv instanceof HTMLInputElement) clv.type = "password";
  }

  tabStaff?.addEventListener("click", activarStaff);
  tabEst?.addEventListener("click", activarEst);
}

// =====================
// INICIO
// =====================
window.addEventListener("DOMContentLoaded", () => {
  aplicarSesionDesdeQueryString();
  sanearSesionStaffSiFaltaClave_();

  initTheme();

  configurarPestanasLogin();

  const formLogin = document.getElementById("login-form");
  if (formLogin) {
    formLogin.addEventListener("submit", onLoginSubmit);
  }

  document.getElementById("login-mostrar-clave")?.addEventListener("change", (ev) => {
    const on = ev.target && ev.target.checked === true;
    const inp = document.getElementById("login-clave");
    if (inp instanceof HTMLInputElement) inp.type = on ? "text" : "password";
  });

  document.getElementById("login-mostrar-pin")?.addEventListener("change", (ev) => {
    const on = ev.target && ev.target.checked === true;
    const pin = document.getElementById("login-est-pin");
    if (pin instanceof HTMLInputElement) pin.type = on ? "text" : "password";
  });

  const ctx = leerContexto();
  llenarSelectGrupoPlaceholder("Cargando cursos…");
  if (getAuthTipo() === "estudiante") {
    llenarSelectMateriaEstudiante(leerMateriaExpedienteEstudiante());
  } else {
    llenarSelect("materia", MATERIAS, ctx.materia || "Matemáticas");
  }
  llenarSelectEstadoMatricula();

  try {
    const guardado = localStorage.getItem(STORAGE.modo);
    if (guardado === "inasistencias") state.appModo = "inasistencias";
    if (guardado === "resumen-faltas") state.appModo = "resumen-faltas";
    if (guardado === "resumen-area") state.appModo = "resumen-area";
    if (guardado === "mosaico-grupo") state.appModo = "mosaico-grupo";
  } catch (_) {
    /* ignore */
  }

  const inaFecha = document.getElementById("ina-fecha");
  if (inaFecha && !inaFecha.value) {
    inaFecha.value = isoFechaLocalHoy();
  }
  actualizarEtiquetaFechaInasistencias();

  document.getElementById("btn-expediente-estudiante")?.addEventListener("click", () => {
    if (!esEstudiante()) return;
    const row = state.cacheEstudiantes[0];
    if (!row) {
      toast(
        esEstudiante()
          ? "Espere unos segundos mientras se cargan sus notas generales abajo."
          : "Espere a que cargue su curso o pulse «Cargar grupo».",
        "err"
      );
      return;
    }
    void abrirEstudiante(row);
  });

  document.getElementById("modo-calificaciones")?.addEventListener("click", () => {
    state.appModo = "calificaciones";
    aplicarAppModoUI();
  });
  document.getElementById("modo-inasistencias")?.addEventListener("click", () => {
    if (esEstudiante()) {
      toast("Como estudiante, use su expediente (materia arriba) para ver inasistencias de cada materia.", "err");
      return;
    }
    if (esSoloLecturaCoordinacion()) {
      toast(
        "El registro diario de asistencia es solo para docentes. Use «Resumen faltas» para consultar inasistencias por periodo.",
        "err"
      );
      return;
    }
    state.appModo = "inasistencias";
    aplicarAppModoUI();
  });
  document.getElementById("modo-resumen-faltas")?.addEventListener("click", () => {
    state.appModo = "resumen-faltas";
    aplicarAppModoUI();
  });
  document.getElementById("modo-resumen-area")?.addEventListener("click", () => {
    if (esEstudiante()) {
      toast("Como estudiante, use el boletín general en «Calificaciones».", "err");
      return;
    }
    state.appModo = "resumen-area";
    aplicarAppModoUI();
  });
  document.getElementById("modo-mosaico-grupo")?.addEventListener("click", () => {
    if (esEstudiante()) {
      toast("La pestaña «Mosaico» es solo para coordinación y docentes.", "err");
      return;
    }
    state.appModo = "mosaico-grupo";
    aplicarAppModoUI();
  });

  document.getElementById("rf-periodo")?.addEventListener("change", () => {
    actualizarBannerPeriodoRf();
    if (state.appModo === "resumen-faltas" && state.cacheEstudiantes.length) {
      void cargarResumenFaltasPeriodo();
    }
  });

  document.getElementById("est-periodo-cal")?.addEventListener("change", () => {
    try {
      const sel = document.getElementById("est-periodo-cal");
      if (sel && sel.value) sessionStorage.setItem(STORAGE.estudiantePeriodoCal, sel.value);
    } catch (_) {
      /* ignore */
    }
    if (!esEstudiante()) return;
    actualizarTituloNotasGeneralesEstudiante_();
    void cargarBoletinEstudiante();
    if (state.estudianteActual) {
      void abrirEstudiante(state.estudianteActual);
    }
  });

  document.getElementById("rf-actualizar")?.addEventListener("click", () => {
    void cargarResumenFaltasPeriodo();
  });
  document.getElementById("ra-actualizar")?.addEventListener("click", () => {
    void cargarResumenAreaGrupoDocente_();
  });
  document.getElementById("mg-actualizar")?.addEventListener("click", () => {
    state.cacheMosaicoGrupo = null;
    void cargarMosaicoGrupoDocente_();
  });

  document.getElementById("rf-body")?.addEventListener("click", (ev) => {
    const ficha = ev.target.closest(".ina-ficha-btn");
    if (!ficha) return;
    const sid = ficha.getAttribute("data-estudiante-id");
    const est = state.cacheEstudiantes.find((x) => String(x.id) === String(sid));
    if (est) abrirEstudiante(est);
  });

  document.getElementById("ra-body")?.addEventListener("click", (ev) => {
    const ficha = ev.target.closest(".ina-ficha-btn");
    if (!ficha) return;
    const sid = ficha.getAttribute("data-estudiante-id");
    const est = state.cacheEstudiantes.find((x) => String(x.id) === String(sid));
    if (est) abrirEstudiante(est);
  });

  inaFecha?.addEventListener("change", () => {
    actualizarEtiquetaFechaInasistencias();
    sincronizarSelectFechasConListaDesdeInput();
    if (state.appModo === "inasistencias" && state.cacheEstudiantes.length) {
      void refrescarTablaInasistencias();
    }
  });

  document.getElementById("ina-fechas-con-lista")?.addEventListener("change", (ev) => {
    const iso = ev.target.value;
    if (!iso) return;
    if (inaFecha) inaFecha.value = iso;
    actualizarEtiquetaFechaInasistencias();
    if (state.appModo === "inasistencias" && state.cacheEstudiantes.length) {
      void refrescarTablaInasistencias();
    }
  });

  document.getElementById("ina-btn-hoy")?.addEventListener("click", () => {
    usarFechaHoyInasistencias();
  });

  document.getElementById("ina-todos-asistio")?.addEventListener("click", () => {
    aplicarInasistenciaMasivaGrupo(0);
  });

  document.getElementById("ina-todos-falto")?.addEventListener("click", () => {
    aplicarInasistenciaMasivaGrupo(1);
  });

  document.getElementById("ina-solo-activos")?.addEventListener("change", () => {
    if (state.cacheInasistenciasDia) renderInasistenciasBodyFromCache();
  });

  document.getElementById("ina-guardar")?.addEventListener("click", () => {
    void guardarInasistenciasMasivoFront();
  });

  document.getElementById("ina-body")?.addEventListener("click", (ev) => {
    const ficha = ev.target.closest(".ina-ficha-btn");
    if (ficha) {
      const sid = ficha.getAttribute("data-estudiante-id");
      const est = state.cacheEstudiantes.find((x) => String(x.id) === String(sid));
      if (est) abrirEstudiante(est);
      return;
    }
    const btn = ev.target.closest("[data-ina-set]");
    if (!btn) return;
    if (esModoSoloConsulta()) {
      toast(
        esEstudiante()
          ? "El acceso de estudiante es solo consulta."
          : "El acceso de coordinación es solo consulta.",
        "err"
      );
      return;
    }
    const tr = btn.closest("tr[data-estudiante-id]");
    if (!tr) return;
    const v = btn.getAttribute("data-ina-set") === "1" ? 1 : 0;
    tr.setAttribute("data-falta", String(v));
    const cell = tr.querySelector(".ina-cell-toggle");
    if (cell) cell.innerHTML = renderInasistenciaSegHtml(v);
    const sid = tr.getAttribute("data-estudiante-id");
    if (sid && state.cacheInasistenciasDia) {
      const row = state.cacheInasistenciasDia.find((r) => String(r.id) === String(sid));
      if (row) row.falta = v;
    }
  });

  document.getElementById("grupo").addEventListener("change", () => {
    aplicarMateriasSegunGrupoDesdeCache();
    guardarContexto();
    if (state.estudianteActual) mostrarLista();
    resetVistaGrupoSinCargar();
    if (esEstudiante()) {
      void cargarEstudiantes();
    }
  });
  document.getElementById("materia").addEventListener("change", () => {
    if (esEstudiante()) {
      guardarMateriaExpedienteEstudiante(getMateria());
    }
    guardarContexto();
    actualizarTituloCargaRapida();
    if (state.estudianteActual) abrirEstudiante(state.estudianteActual);
    if (state.appModo === "inasistencias" && state.cacheEstudiantes.length) {
      void rellenarSelectFechasConListaInasistencias();
      void refrescarTablaInasistencias();
    }
    if (state.appModo === "resumen-faltas" && state.cacheEstudiantes.length) {
      void cargarResumenFaltasPeriodo();
    }
    if (state.appModo === "resumen-area" && state.cacheEstudiantes.length) {
      void cargarResumenAreaGrupoDocente_();
    }
  });

  document.getElementById("doc-periodo-cal")?.addEventListener("change", () => {
    const sel = document.getElementById("doc-periodo-cal");
    try {
      if (sel && sel.value) localStorage.setItem(STORAGE.docentePeriodoCal, sel.value);
    } catch (_) {
      /* ignore */
    }
    actualizarHintRangoFechaFichaDocente_();
    if (esEstudiante()) return;
    if (state.appModo === "calificaciones" && state.cacheEstudiantes.length) {
      void (async () => {
        await cargarNotasGrupoMateria();
        construirTitulosGrupo();
        renderListaEstudiantes(state.cacheEstudiantes);
      })();
    }
    if (state.estudianteActual) void abrirEstudiante(state.estudianteActual);
    if (state.appModo === "resumen-area" && state.cacheEstudiantes.length) {
      void cargarResumenAreaGrupoDocente_();
    }
  });

  document.getElementById("btn-cargar").addEventListener("click", () => cargarEstudiantes({ silent: false }));
  document.getElementById("btn-volver").addEventListener("click", mostrarLista);
  document.getElementById("btn-imprimir").addEventListener("click", () => {
    actualizarEncabezadoImpresion();
    window.print();
  });
  document.getElementById("btn-imprimir-grupo-cal")?.addEventListener("click", imprimirGrupoCalificacionesVista);
  document.getElementById("btn-imprimir-resumen-faltas")?.addEventListener("click", imprimirResumenFaltasVista);
  document.getElementById("btn-imprimir-resumen-area")?.addEventListener("click", imprimirResumenAreaVista);
  document.getElementById("btn-imprimir-mosaico-grupo")?.addEventListener("click", imprimirMosaicoGrupoVista);
  document.getElementById("btn-guardar-estado").addEventListener("click", guardarEstadoMatricula);
  document.getElementById("btn-historial-estado").addEventListener("click", abrirHistorialEstadoMatricula);
  document.getElementById("btn-traslado-curso")?.addEventListener("click", confirmarTrasladoEstudiante);
  document.getElementById("ficha-tab-estado")?.addEventListener("click", () => aplicarFichaTab("estado"));
  document.getElementById("ficha-tab-cal")?.addEventListener("click", () => aplicarFichaTab("cal"));
  document.getElementById("ficha-tab-faltas")?.addEventListener("click", () => aplicarFichaTab("faltas"));
  document.getElementById("cr-guardar").addEventListener("click", guardarCargaRapida);
  document.getElementById("cr-limpiar").addEventListener("click", limpiarCargaRapida);
  document.getElementById("cr-definitiva-triple")?.addEventListener("change", aplicarEstadoUiDefinitivaTriple);

  document.getElementById("ficha-nueva-guardar")?.addEventListener("click", () => {
    void guardarNotaDesdeFichaEstudiante();
  });
  document.getElementById("ficha-nueva-eval")?.addEventListener("change", () => {
    sincronizarUiTrasCambioEvalFicha_();
  });
  document.getElementById("ficha-nueva-tipo")?.addEventListener("change", (ev) => {
    const el = document.getElementById("ficha-nueva-titulo");
    const t = ev.target && ev.target.value ? String(ev.target.value) : "Seguimiento";
    if (el) el.value = tituloSugeridoFichaNuevaNota(t);
  });
  const fichaNotaInp = document.getElementById("ficha-nueva-nota");
  if (fichaNotaInp) vincularValidacionNotaInput(fichaNotaInp);

  asegurarFechaCargaRapidaPorDefecto();

  document.getElementById("cr-titulo").addEventListener("input", actualizarTituloCargaRapida);
  document.getElementById("cr-tipo").addEventListener("change", renderCargaRapida);
  document.getElementById("cr-solo-activos").addEventListener("change", renderCargaRapida);
  document.getElementById("cr-col-estado").addEventListener("change", aplicarVisibilidadColumnasCargaRapida);
  document.getElementById("cr-col-tipo").addEventListener("change", aplicarVisibilidadColumnasCargaRapida);
  document.getElementById("cr-col-nota").addEventListener("change", aplicarVisibilidadColumnasCargaRapida);

  document.getElementById("cr-body").addEventListener("keydown", (ev) => {
    if (ev.key !== "Enter") return;
    const target = ev.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (!target.classList.contains("cr-nota-input")) return;
    ev.preventDefault();
    target.blur();
    moverFocoNotaSiguiente(target);
  });

  const btnLogout = document.getElementById("btn-logout");
  if (btnLogout) {
    btnLogout.addEventListener("click", () => {
      clearStoredClave();
      location.reload();
    });
  }

  const buscar = document.getElementById("buscar");
  const aplicarFiltroBusquedaDebounced = debounce(aplicarFiltroBusqueda, 160);
  state.cancelDebouncedFiltroBusqueda = () => aplicarFiltroBusquedaDebounced.cancel();
  buscar.addEventListener("input", () => aplicarFiltroBusquedaDebounced());
  buscar.addEventListener("blur", () => {
    aplicarFiltroBusquedaDebounced.cancel();
    aplicarFiltroBusqueda();
  });

  const boardBody = document.getElementById("grupo-board-body");
  if (boardBody) {
    boardBody.addEventListener("click", (ev) => {
      const btn = ev.target.closest(".btn-ver-detalle");
      if (!btn) return;
      const id = btn.getAttribute("data-estudiante-id");
      if (!id) return;
      const est = state.cacheEstudiantes.find((x) => String(x.id) === String(id));
      if (est) abrirEstudiante(est);
    });
  }

  document.getElementById("modal-cerrar").addEventListener("click", cerrarModal);
  document.getElementById("modal-backdrop").addEventListener("click", (ev) => {
    if (ev.target.id === "modal-backdrop") cerrarModal();
  });

  document.addEventListener("keydown", (ev) => {
    if (ev.key !== "Escape") return;
    const bd = document.getElementById("modal-backdrop");
    if (bd && !bd.classList.contains("is-hidden") && !bd.hidden) {
      cerrarModal();
    }
  });

  actualizarTituloCargaRapida();
  aplicarVisibilidadColumnasCargaRapida();
  renderCargaRapida();
  aplicarAppModoUI();
  aplicarVisibilidadAnadirNotaFicha();
  actualizarLabelModoCalificaciones();
  actualizarEtiquetaFichaTabEstado();
  void (async () => {
    await verificarSesionGuardada();
    await verificarSesionEstudianteGuardada();
  })();
});
