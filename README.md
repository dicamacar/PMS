# PMS

## Despliegue estático en Render

El repositorio incluye `render.yaml` para crear el tablero como un sitio estático en Render.

1. En Render, selecciona **New + → Blueprint**.
2. Conecta la cuenta `dicamacar` y el repositorio `dicamacar/PMS`.
3. Selecciona la rama `main` y confirma el Blueprint `render.yaml`.
4. Revisa los detalles del servicio `pms-tablero-agendamiento` y selecciona **Apply**.

La carpeta publicada es `outputs/01a16a48-5490-7a73-bc55-755f757c3f53/interface/dist`. No requiere instalación ni compilación; el build command es `true`. Render asignará una URL `onrender.com` y podrá volver a desplegar cada vez que se actualice `main`.

**Acceso a los datos:** el tablero carga el archivo `data.js` con registros de citas, usuarios que las registraron y fechas. Un sitio estático de Render queda disponible por URL pública; las cabeceras de caché incluidas no restringen el acceso. Antes de aplicar el Blueprint, confirma que esa audiencia está autorizada o prepara una versión anonimizada del tablero.
