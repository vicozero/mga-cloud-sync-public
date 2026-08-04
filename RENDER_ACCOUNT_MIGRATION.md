# Migración de MGA Cloud Sync a otra cuenta de Render

Objetivo: mover `https://mga-cloud-sync.onrender.com/almacen-filtros` a una nueva cuenta/workspace de Render conservando código, base de datos y clave de acceso.

## Estado del proyecto

El repositorio ya está preparado para Render mediante `render.yaml`.

- Repositorio GitHub: `https://github.com/vicozero/mga-cloud-sync.git`
- Servicio web definido: `mga-cloud-sync`
- Base PostgreSQL definida: `mga-cloud-sync-db`
- Comando de build: `pip install -r requirements-cloud.txt`
- Comando de inicio: `uvicorn render_backend.main:app --host 0.0.0.0 --port $PORT`
- Health check: `/health`

## Variables necesarias

Render debe tener estas variables en el servicio web:

| Variable | Valor |
| --- | --- |
| `PYTHON_VERSION` | `3.12.8` |
| `DATABASE_URL` | Connection string de la nueva PostgreSQL de Render |
| `MGA_REQUIRE_API_KEY` | `true` |
| `MGA_API_KEY` | La misma clave que usa MGA Mantenimiento / página web / APK |

Importante: si cambia `MGA_API_KEY`, también hay que actualizarla en:

- MGA Mantenimiento > Red > API key cloud
- La página web, campo “Clave para editar”
- APK/configuración móvil si aplica

## Ruta recomendada

### 1. Dar acceso al repositorio

En la cuenta nueva de Render, conectar GitHub y seleccionar:

```text
https://github.com/vicozero/mga-cloud-sync.git
```

Si la cuenta nueva no ve el repositorio, hay que invitar esa cuenta al repositorio de GitHub o transferir/forkear el repositorio.

### 2. Crear Blueprint

En Render:

1. Dashboard > New > Blueprint
2. Seleccionar el repositorio `mga-cloud-sync`
3. Confirmar que Render detecta `render.yaml`
4. Crear los recursos

Esto debe crear:

- Web service `mga-cloud-sync`
- PostgreSQL `mga-cloud-sync-db`
- Variable `DATABASE_URL` conectada automáticamente a la base

### 3. Copiar la clave `MGA_API_KEY`

En la cuenta actual de Render:

1. Abrir el servicio `mga-cloud-sync`
2. Environment
3. Copiar `MGA_API_KEY`

En la cuenta nueva:

1. Abrir el nuevo servicio
2. Environment
3. Pegar el mismo valor en `MGA_API_KEY`

No conviene generar una clave nueva si el escritorio y móviles ya están trabajando con la clave actual.

### 4. Migrar la base de datos

Para conservar capturas, inventario, filtros, preventivos, movimientos y snapshots, hay que copiar PostgreSQL de la cuenta actual a la cuenta nueva.

Flujo seguro:

1. Pausar capturas temporalmente para evitar cambios durante la copia.
2. Obtener `External Database URL` de la base actual.
3. Obtener `External Database URL` de la base nueva.
4. Ejecutar:

```powershell
pg_dump --format=custom --no-owner --no-privileges --file mga_cloud_sync.backup "OLD_DATABASE_URL"
pg_restore --clean --if-exists --no-owner --no-privileges --dbname "NEW_DATABASE_URL" mga_cloud_sync.backup
```

Notas:

- Usar las URLs externas, no las internas, si se ejecuta desde la PC.
- Si `pg_dump` y `pg_restore` no están instalados en la PC, instalarlos con PostgreSQL Tools o hacer el backup desde Render.
- En bases gratuitas de Render, los backups automáticos pueden no estar disponibles; por eso el método con `pg_dump` es el más directo.

### 5. Probar la nueva URL

Abrir:

```text
https://NUEVO-SERVICIO.onrender.com/health
https://NUEVO-SERVICIO.onrender.com/almacen-filtros
```

Verificar:

- `/health` responde correctamente.
- La página carga.
- La clave `MGA_API_KEY` permite editar/capturar.
- Inventario y capturas aparecen con datos.
- Exportar Excel funciona.

### 6. Actualizar MGA Mantenimiento

En el programa:

1. Red
2. Cambiar `URL cloud Render` a:

```text
https://NUEVO-SERVICIO.onrender.com
```

3. Mantener la misma API key si se copió.
4. Guardar.
5. Probar Nube / Sync.

### 7. Corte final

Cuando la nueva cuenta esté validada:

1. Avisar a usuarios que usen la nueva URL.
2. Actualizar accesos directos/favoritos.
3. Mantener el servicio anterior unos días como respaldo.
4. Suspender o eliminar el servicio anterior solo cuando ya no haya capturas pendientes.

## Validaciones técnicas locales

Antes de migrar, validar que el código actual compila:

```powershell
cd "C:\Users\Dell\mga-cloud-sync-fix"
python -m py_compile render_backend\main.py
git status --short
```

El repositorio está listo para Render si `git status --short` no muestra cambios pendientes importantes y `py_compile` no arroja errores.
