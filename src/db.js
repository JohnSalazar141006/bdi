import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const haySupabase = !!(url && key);
export const supabase = haySupabase ? createClient(url, key) : null;

// Almacén clave-valor en la nube (tabla kv). Persiste y se sincroniza.
export const dbStorage = {
  get: async (k) => {
    const { data } = await supabase.from("kv").select("value").eq("key", k).maybeSingle();
    return data ? { key: k, value: JSON.stringify(data.value) } : null;
  },
  set: async (k, v) => {
    await supabase.from("kv").upsert({ key: k, value: JSON.parse(v), updated_at: new Date().toISOString() });
    return { key: k, value: v };
  },
  delete: async (k) => { await supabase.from("kv").delete().eq("key", k); return { key: k, deleted: true }; },
  list: async (p = "") => {
    const { data } = await supabase.from("kv").select("key").like("key", p + "%");
    return { keys: (data || []).map((r) => r.key) };
  },
};
