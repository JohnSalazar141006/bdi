import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";
import { dbStorage, haySupabase } from "./db.js";

// Memoria: base de datos en la nube si hay Supabase configurado; si no, el navegador.
const local = {
  get: async (k) => { const v = localStorage.getItem(k); return v == null ? null : { key: k, value: v }; },
  set: async (k, v) => { localStorage.setItem(k, v); return { key: k, value: v }; },
  delete: async (k) => { localStorage.removeItem(k); return { key: k, deleted: true }; },
  list: async (p = "") => ({ keys: Object.keys(localStorage).filter((x) => x.startsWith(p)) }),
};
window.storage = haySupabase ? dbStorage : local;

// Auto-actualizar la tasa BCV al abrir (si la fuente permite el acceso).
(async () => {
  try {
    const r = await fetch("https://ve.dolarapi.com/v1/dolares/oficial");
    const d = await r.json();
    const valor = d.promedio || d.price;
    if (valor) await window.storage.set("tasa_bcv", JSON.stringify({ valor, fecha: new Date().toISOString() }));
  } catch (e) { /* sin conexión: se mantiene la última tasa guardada */ }
})();

createRoot(document.getElementById("root")).render(<App />);
