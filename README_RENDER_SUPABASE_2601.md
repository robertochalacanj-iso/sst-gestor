# SST Gestor - Patch Render/Supabase 26.0.1

## Qué corrige

- Evita que Render quede en pantalla negra si la base de datos tarda o falla.
- Agrega `/healthz` y `/api/health` para diagnóstico.
- Actualiza Python a 3.11.9.
- Actualiza `gunicorn` con bind explícito a `$PORT`, logs y timeout.
- Deja la app preparada para Supabase/PostgreSQL usando `DATABASE_URL` o variables `PGHOST`, `PGUSER`, `PGPASSWORD`, `PGDATABASE`.
- Mantiene SQLite solo para trabajo local.

## Variables que debes revisar en Render

En Render > sst-gestor > Environment:

```text
ADMIN_USERNAME=admin
ADMIN_PASSWORD=tu_clave_segura
SECRET_KEY=un_valor_largo_seguro
DATABASE_URL=la_cadena_postgresql_de_supabase
```

Si no configuras `DATABASE_URL`, la app puede arrancar con SQLite, pero en Render no es recomendable porque los datos se pueden perder al reiniciar.

## Validación

Después de publicar abre:

```text
https://sst-gestor.onrender.com/healthz
```

Debe responder JSON. Si `database_ok` sale `false`, el problema ya no es Render sino la conexión a Supabase/Postgres.

## Publicación

Reemplaza estos archivos en tu repositorio de `seguridadsst` y sube los cambios a GitHub. Render debe redeployar automáticamente.
