# Conectar la base de datos en la nube (Supabase)

Sin esto, la app corre igual pero guarda la memoria SOLO en tu navegador.
Con esto, todo (tasa, historial, respaldo) se guarda en una base de datos real
en la nube, se sincroniza entre dispositivos y la tasa BCV se auto-actualiza.

## Paso a paso

1. Entra a https://supabase.com y crea una cuenta (gratis).
2. Crea un proyecto nuevo (New project). Anota la contraseña de la base de datos.
3. Cuando el proyecto esté listo, ve a  Project Settings > API  y copia:
   - Project URL          (ej. https://abcd.supabase.co)
   - anon public key       (una clave larga)
4. En este proyecto, crea un archivo llamado  .env  (copia de .env.example) y pega:

       VITE_SUPABASE_URL=https://TU-PROYECTO.supabase.co
       VITE_SUPABASE_ANON_KEY=tu_clave_anon_publica

5. En Supabase, ve a  SQL Editor > New query, pega el contenido de
   supabase/schema.sql  y dale RUN. Repite con  supabase/seed.sql.
6. Instala la nueva dependencia y arranca:

       npm install
       npm run dev

Listo: ahora el sistema guarda todo en la nube. Genera un documento, ábrelo desde
otro dispositivo con la misma app y verás el historial ahí. La tasa BCV se intenta
actualizar sola cada vez que abres la app.

## Qué guarda hoy en la base de datos
- tasa_bcv        (la tasa, auto-actualizada)
- bitacora        (actividad reciente)
- historial       (todos los documentos generados, para respaldo)

## Próximo nivel (cuando quieras)
- Mover la lista de EMPRESAS (hoy en el código) a una tabla editable en Supabase,
  con pantalla para agregar/editar empresas y sus porcentajes.
- Guardar cada balance y desagregación como registros con sus datos, no solo en kv.
- Auto-actualizar la tasa BCV con una función programada (cron) en el servidor.
