-- Tasa inicial (luego se auto-actualiza al abrir la app)
insert into kv (key, value) values ('tasa_bcv', '{"valor":623.02}')
on conflict (key) do nothing;
