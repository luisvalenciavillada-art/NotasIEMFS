# Sistema de Notas · IEMFS

Aplicación web para gestión de calificaciones, inasistencias, resumen de faltas, resumen de área y mosaico por grupo.

## URL pública

- App: `https://luisvalenciavillada-art.github.io/NotasIEMFS/`

## Mosaico por grupo (hoja `MosaicosGrupo`)

En Google Sheets, use la hoja `MosaicosGrupo` con estas columnas:

- **A:** `grupo`
- **B:** `imagen` (URL pública directa)
- **C:** `actualizado_en` (opcional, formato fecha)

### Ejemplo

- `603` | `https://cdn.jsdelivr.net/gh/luisvalenciavillada-art/NotasIEMFS@main/images/mosaico-603.webp` | `2026-05-06`
- `6-7 C` | `https://cdn.jsdelivr.net/gh/luisvalenciavillada-art/NotasIEMFS@main/images/mosaico-6-7c.webp` | `2026-05-06`

## Carpeta de imágenes

Guardar los mosaicos en:

- `images/mosaico-603.webp`
- `images/mosaico-6-7c.webp`
- `images/mosaico-6-7d.webp`
- `images/mosaico-6-7e.webp`
- etc.

## Cómo actualizar un mosaico

1. Subir/actualizar archivo en `images/` (mismo nombre recomendado).
2. Verificar o actualizar URL en `MosaicosGrupo` columna B.
3. Recargar app (`Ctrl+F5`) y usar botón **Actualizar mosaico**.

## Desarrollo local

```bash
npm run build
