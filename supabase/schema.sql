-- Tabla de memoria del sistema (clave-valor). Guarda tasa, bitácora e historial.
create table if not exists kv (
  key text primary key,
  value jsonb,
  updated_at timestamptz default now()
);

-- Acceso abierto (herramienta de una sola oficina). Restringir con auth si se comparte.
alter table kv enable row level security;
drop policy if exists "kv_abierto" on kv;
create policy "kv_abierto" on kv for all using (true) with check (true);
