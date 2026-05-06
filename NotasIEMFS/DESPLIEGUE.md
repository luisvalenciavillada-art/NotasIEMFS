# Publicar el Sistema de Notas en la nube (web + clave)

La **interfaz** (`index.html`, `styles.css`, `app.js`, `icon.svg`, `manifest.webmanifest`) puede vivir en cualquier hosting estático. Los **datos y la clave real** siguen en **Google Apps Script** y en tu **hoja de cálculo**.

## Hoja Grupos (desplegable «Grupo» en la web)

Si existe una pestaña llamada **`Grupos`**, el script usa la **columna B** (`Nombre_Grupo`) como lista de cursos en la app (fila 1 = encabezados `ID_Grupo` / `Nombre_Grupo`). Si no hay esa hoja o está vacía, se usan los valores únicos de la columna **C** de **Estudiantes**. Los textos en **Estudiantes** columna **C** deben coincidir con **`Nombre_Grupo`** (misma escritura, salvo mayúsculas/acentos al comparar).

## Hoja Estudiantes (columnas)

| A | B | C | D |
|---|---|---|---|
| Id estudiante | Nombre | Grupo | **Estado de matrícula** (`Activo`, `Inactivo`, `Desertor`, `Matrícula cancelada`) |

Los valores antiguos `activo` en la columna D se muestran como **Activo**. La lista incluye **todos** los estudiantes del grupo (también desertores o cancelados), con etiqueta de color.

### Traslado de curso (automático desde la web)

En el **detalle del estudiante**, la tarjeta **Cambio de curso (traslado)** permite elegir el **curso de destino** (misma lista que el desplegable superior), un **motivo** opcional y **Confirmar traslado**. El script (`POST` `migrarGrupoNotas`):

- Actualiza la **columna C** (grupo) en **Estudiantes** para ese id.
- Actualiza la **columna D** (grupo) en **Notas** en **todas** las filas de ese estudiante (todas las materias).
- Registra el hecho en **Auditoría** (detalle del traslado + motivo).

El curso de destino debe existir en la hoja **Grupos** (columna B) o, si no hay hoja Grupos, en el listado de grupos que arma el script desde **Estudiantes**. Cursos nuevos: créelos primero en **Grupos**.

Tras cambiar `Code.gs`, vuelve a **implementar** la aplicación web.

### Inasistencias (hoja nueva)

La primera vez que se guarda asistencia, el script crea la hoja **`Inasistencias`** con columnas: `timestamp`, `estudiante`, `grupo`, `materia`, `fecha` (día), `falta` (1 = faltó, 0 = asistió). El registro masivo en la web usa la **materia** y el **grupo** del panel superior; la **fecha** se elige en el módulo Inasistencias (por defecto el día actual). Endpoints: `GET inasistenciasDia`, `GET inasistenciasEstudiante`, `POST inasistenciasGuardar`.

### Periodos académicos (tres periodos)

El script crea la hoja **`PeriodosAcademicos`** con columnas `periodo` (1, 2 o 3), `fecha_inicio`, `fecha_fin`, `etiqueta`. **Edite las fechas** para que coincidan con el calendario de la institución (los valores iniciales son solo ejemplo). La vista **Resumen faltas** usa estos rangos para totalizar por periodo. Endpoints: `GET periodosAcademicos`, `GET inasistenciasResumenPeriodo`.

---

## 1. Configurar la clave en Apps Script (obligatorio)

1. Abre el proyecto en [script.google.com](https://script.google.com) vinculado a tu hoja.
2. Pega o actualiza el archivo `Code.gs` (con validación de clave).
3. En **`Código.gs`**, abre la función **`EJECUTAR_PARA_GUARDAR_MI_CLAVE`**, escribe tu clave en la línea `var miClave = "..."` (mínimo 6 caracteres), guarda, y en el menú de funciones elige **esa** función (no `establecerClaveAcceso`) y pulsa **Ejecutar** ▶.  
   **Importante:** Si ejecutas solo `establecerClaveAcceso`, Google no puede pasar la clave y dará error.
4. La clave queda guardada en **Propiedades del proyecto** (`NOTAS_IEMFS_CLAVE`). No la pongas en el código de la web.

5. **Despliega** la aplicación web: **Implementar** → **Nueva implementación** → tipo **Aplicación web**.  
   - Ejecutar como: **Yo**  
   - Quién tiene acceso: **Cualquier usuario** (así el front en HTTPS puede llamar a la API; la protección es la **clave**, no el anonimato de Google).

6. Copia la **URL del script** (`.../exec`) y pégala en `app.js` en la constante **`API`**.

## 2. Subir la web a la nube (elige una opción)

### Opción A — GitHub Pages (gratis)

1. Crea un repositorio en GitHub y sube la carpeta del proyecto (los archivos estáticos).
2. En el repo: **Settings** → **Pages** → **Source**: rama `main`, carpeta `/ (root)`.
3. Tras unos minutos tendrás una URL del tipo `https://TU_USUARIO.github.io/TU_REPO/`.
4. Abre `https://.../index.html` e inicia sesión con la clave.

### Opción B — Netlify / Cloudflare Pages (gratis)

1. Arrastra la carpeta del sitio en [Netlify Drop](https://app.netlify.com/drop) o conecta el repositorio.
2. Usa la URL que te asignen (`https://algo.netlify.app`).
3. Asegúrate de que `index.html` esté en la raíz del sitio publicado.

### Opción C — Servidor del colegio

Sube por FTP o panel los mismos archivos estáticos a un virtual host con **HTTPS** (recomendado).

**Importante:** La URL de la API en `app.js` debe seguir siendo la de tu Apps Script. La web solo es “cáscara”; el guardado de notas sigue yendo a Google.

## 3. Seguridad (lectura honesta)

- La clave viaja en **HTTPS**; no la incrustes en repositorios públicos como texto plano en el front (no va en `app.js`; solo la URL del script).
- Cualquiera con la **clave** puede usar la API como docente. Para más seguridad a futuro: dominio restringido, OAuth institucional o listas blancas en el script.
- “Recordar en este equipo” guarda la clave en el navegador (`localStorage` / `sessionStorage`). En equipos compartidos conviene **no** marcarlo.

## 4. Comprobar que todo funciona

1. Abre la web publicada.
2. Debe aparecer la pantalla de **clave**.
3. Tras entrar, **Actualizar lista** debe cargar estudiantes y poder **guardar** una nota.

Si ves “Falta configurar la clave en Apps Script”, ejecuta de nuevo **`EJECUTAR_PARA_GUARDAR_MI_CLAVE`** en el editor.
