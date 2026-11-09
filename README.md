# PMS — tablero de agendamiento

El sitio estático está en `outputs/dist/`:

- `outputs/dist/index.html`: interfaz del tablero.
- `outputs/dist/data.js`: datos que carga la interfaz.

El resto de los Excel, scripts y vistas previas permanece en `outputs/`; Render publica únicamente `dist/`.

## Configuración del Static Site en Render

Conecta el repositorio `dicamacar/PMS`, rama `main`, y configura el servicio así:

| Campo | Valor |
| --- | --- |
| Root Directory | `outputs` |
| Build Command | `echo "Static assets already built"` |
| Publish Directory | `dist` |

No ejecutes `npm install` ni `npm run build`: el HTML y JavaScript ya están preparados y no existe un `package.json` en `outputs/`. Si usas `render.yaml` como Blueprint, sus valores corresponden a estos mismos campos.

## Datos y acceso

El archivo `outputs/dist/data.js` contiene registros de citas, usuario que registra y fechas. Un Static Site se sirve mediante una URL pública; `Cache-Control: no-store` reduce el almacenamiento en caché, pero no restringe el acceso. Publica estos datos solo si la audiencia está autorizada.
