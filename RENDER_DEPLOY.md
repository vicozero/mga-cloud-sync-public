# Despliegue cloud MGA en Render

Esta es la arquitectura ideal para que el APK no dependa de la IP de la computadora:

1. El APK guarda capturas y fotos offline en el celular.
2. Cuando el celular tiene internet, sube pendientes a la nube Render.
3. El programa de escritorio sigue siendo la bitacora central.
4. Cuando el programa de escritorio tiene internet, importa desde Render y guarda en la base SQLite local.

## Archivos agregados

- `render_backend/main.py`: API cloud FastAPI.
- `requirements-cloud.txt`: dependencias del servicio cloud.
- `render.yaml`: Blueprint para crear el servicio y la base PostgreSQL en Render.
- `mobile_app/cloud_config.json`: URL de nube que se empaca dentro del APK.

## Desplegar en Render

Opcion recomendada: usar el `render.yaml`.

1. Sube este proyecto a un repositorio GitHub privado.
2. En Render, crea un Blueprint y selecciona el repositorio.
3. Render detectara `render.yaml` y creara:
   - Web service `mga-cloud-sync`.
   - Base PostgreSQL `mga-cloud-sync-db`.
   - Variable `DATABASE_URL`.
   - Variable `MGA_API_KEY`.
   - En este paquete esta configurado en plan gratis para prueba.
4. Al terminar, abre:
   - `https://TU-SERVICIO.onrender.com/health`
   - Debe responder `service: mga-cloud-sync`.

Opcion manual si no usas Blueprint:

- Runtime: Python.
- Build command: `pip install -r requirements-cloud.txt`.
- Start command: `uvicorn render_backend.main:app --host 0.0.0.0 --port $PORT`.
- Variables:
  - `DATABASE_URL`: URL de PostgreSQL en Render.
  - `MGA_API_KEY`: una clave privada que tambien pondras en el escritorio.

Nota: Render indica que las bases Postgres gratis expiran despues de 30 dias. Para uso diario real, cambia `plan: free` por un plan pagado vigente en `render.yaml`.

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
