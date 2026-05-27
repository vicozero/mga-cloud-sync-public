# Migracion MGA a Cloudflare Workers + D1 + R2

Esta carpeta contiene el backend alterno para reemplazar Render con Cloudflare.
La version actual funciona con Workers + D1. Las evidencias se guardan en D1 para evitar el bloqueo de activacion de R2; cuando R2 quede habilitado se puede volver a agregar el bucket para evidencias grandes.

URL actual desplegada:

```text
https://mga-cloud-sync.mga-vicozero.workers.dev
```

## Que cubre

- Capturas del APK: `POST /api/sync`
- Pendientes para importar a la PC: `GET /api/desktop/pending`
- Confirmacion de importacion PC: `POST /api/desktop/ack`
- Estado de capturas en APK: `POST /api/mobile/status`
- Catalogo publicado desde el programa: `GET/POST /api/catalog`
- Existencias de filtros: `GET/POST /api/filter-inventory/snapshot`
- EPP: `GET/POST /api/epp/snapshot`
- Mangueras: `GET/POST /api/hose-changes`
- Evidencias fotograficas en D1. R2 queda como mejora posterior cuando la cuenta lo permita.

## Despliegue

### Opcion automatica

Cuando `wrangler login` este autorizado, puedes correr:

```powershell
cd "C:\Users\Dell\Documents\New project\cloudflare_worker"
powershell -ExecutionPolicy Bypass -File .\tools\desplegar_cloudflare.ps1
```

El script instala dependencias, crea D1, aplica `schema.sql`, configura `MGA_API_KEY`, despliega y actualiza `mobile_app/cloud_config.json`, `android/app/src/main/assets/public/cloud_config.json` y `mga_config.json`.

### Opcion manual

1. Instalar dependencias:

   ```powershell
   cd "C:\Users\Dell\Documents\New project\cloudflare_worker"
   npm install
   ```

2. Iniciar sesion:

   ```powershell
   npx wrangler login
   npx wrangler whoami
   ```

3. Crear D1:

   ```powershell
   npx wrangler d1 create mga-cloud-db
   ```

4. Copiar el `database_id` que entregue Cloudflare y reemplazarlo en `wrangler.jsonc`.

5. Crear las tablas:

   ```powershell
   npx wrangler d1 execute mga-cloud-db --remote --file=./schema.sql
   ```

6. Configurar la misma clave API que usa actualmente tu APK/programa:

   ```powershell
   npx wrangler secret put MGA_API_KEY
   ```

7. Publicar:

   ```powershell
   npx wrangler deploy
   ```

8. Configurar el programa y el APK con la URL que entregue Cloudflare:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\tools\configurar_cloudflare.ps1 -Url "https://TU-WORKER.workers.dev"
   ```

9. Recompilar APK:

   ```powershell
   cd "C:\Users\Dell\Documents\New project"
   powershell -ExecutionPolicy Bypass -File .\build_release_apk.ps1
   ```

## Prueba recomendada

1. En el programa PC, entrar a `Nube` y publicar catalogo/equipos.
2. En el APK, capturar una prueba con foto y sincronizar.
3. En el programa PC, entrar a `Nube` e importar capturas.
4. Revisar que la captura quede en bitacora y que las evidencias abran.

No borres Render hasta validar esta prueba en campo.
