# Despliegue cloud MGA en Render

Esta es la arquitectura ideal para que el APK no dependa de la IP de la computadora:

1. El APK guarda capturas y fotos offline en el celular.
2. Cuando el celular tiene internet, sube pendientes a la nube Render.
3. Render guarda las capturas en PostgreSQL para que no dependan de que la PC este encendida.
4. El programa de escritorio sigue siendo la bitacora central.
5. Cuando el programa de escritorio tiene internet, importa desde Render y guarda en la base SQLite local.

## Archivos agregados

- `render_backend/main.py`: API cloud FastAPI.
- `requirements-cloud.txt`: dependencias del servicio cloud.
- `render.yaml`: Blueprint para crear el servicio. `DATABASE_URL` queda como secreto manual para poder reutilizar una base gratuita existente durante pruebas.
- `mobile_app/cloud_config.json`: URL de nube que se empaca dentro del APK.

## Desplegar en Render

Opcion recomendada: usar el `render.yaml` y configurar `DATABASE_URL` manualmente.

1. Sube este proyecto a un repositorio GitHub privado.
2. En Render, crea un Blueprint y selecciona el repositorio.
3. Render detectara `render.yaml` y creara/configurara:
   - Web service `mga-cloud-sync`.
   - Variable `DATABASE_URL` como secreto pendiente de captura manual.
   - Proteccion por `MGA_API_KEY`.
4. Al terminar, abre:
   - `https://TU-SERVICIO.onrender.com/health`
   - Debe responder `service: mga-cloud-sync`.
5. Para revisar capturas alojadas:
   - Abre `MGA Mantenimiento > Nube > Ver Render`.
   - Debe mostrar capturas alojadas, pendientes, importadas y si la base es persistente.
6. Para inventario de filtros del almacen:
   - Abre `https://TU-SERVICIO.onrender.com/almacen-filtros`.
   - Captura la misma API key configurada en Render.
   - El almacenista puede importar Excel, exportar Excel, revisar filtros por equipo y registrar entradas/salidas en el concentrado.

Opcion manual si no usas Blueprint:

- Runtime: Python.
- Build command: `pip install -r requirements-cloud.txt`.
- Start command: `uvicorn render_backend.main:app --host 0.0.0.0 --port $PORT`.
- Variables opcionales:
  - `DATABASE_URL`: conexion a PostgreSQL de Render.
  - `MGA_REQUIRE_API_KEY=true`.
  - `MGA_API_KEY`: debe coincidir con el escritorio y con el APK.

Nota importante: si `DATABASE_URL` no esta configurado, el backend usa `cloud_sync.db` dentro del servicio, que es temporal en Render y puede perderse en reinicios o redeploys. Para uso diario, usa PostgreSQL de Render o un disco persistente.

Para la prueba gratuita actual, `mga-cloud-sync` puede reutilizar una base PostgreSQL gratuita existente. La informacion de MGA queda separada por tablas con prefijo `mga_`: `mga_mobile_capture`, `mga_mobile_photo`, `mga_catalog_snapshot`, `mga_filter_inventory_item` y `mga_filter_inventory_movement`.

Render recomienda para FastAPI usar Uvicorn enlazado a `0.0.0.0` y al puerto `$PORT`. Referencias oficiales:

- [Deploy a FastAPI App - Render Docs](https://render.com/docs/deploy-fastapi)
- [Blueprint YAML Reference - Render Docs](https://render.com/docs/blueprint-spec)
- [Web Services - Render Docs](https://render.com/docs/web-services/)

## Configurar escritorio

1. Abre `MGA Mantenimiento`.
2. Entra a `Red`.
3. Captura:
   - `URL cloud Render`: `https://TU-SERVICIO.onrender.com`
   - `API key cloud`: valor de `MGA_API_KEY` en Render.
4. Guarda.
5. En el boton `Nube`, usa:
   - `Publicar equipos`: sube el catalogo para que el APK tenga equipos actualizados.
   - `Importar capturas`: baja capturas pendientes de la nube a la bitacora local.
   - `Sync inventario`: baja al escritorio el concentrado de filtros actualizado en Render.

## Generar APK con nube

Despues de configurar la URL cloud en el escritorio, genera el APK de nuevo:

```powershell
.\build_release_apk.ps1
```

Ese proceso actualiza `mobile_app/catalog_seed.json` y `mobile_app/cloud_config.json`, sincroniza Capacitor y genera:

```text
installer_output\MGA_Captura_Movil_release.apk
```

## Flujo de trabajo

- En campo sin internet: el supervisor captura normalmente; todo queda pendiente en el APK.
- Con datos moviles o WiFi: el APK sube a Render.
- En oficina: el escritorio importa de Render y guarda en la bitacora, historial y evidencias fotograficas.

La IP local queda como respaldo para redes internas, pero ya no es el camino principal.
