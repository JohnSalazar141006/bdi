import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    // Proxy opcional para la pestaña "Agente" (chat con IA) en local.
    // 1) crea un archivo .env con:  VITE_ANTHROPIC_KEY=tu_clave
    // 2) en src/App.jsx cambia la URL del fetch de
    //    "https://api.anthropic.com/v1/messages"  ->  "/anthropic/v1/messages"
    proxy: {
      "/anthropic": {
        target: "https://api.anthropic.com",
        changeOrigin: true,
        rewrite: (p) => p.replace(/^\/anthropic/, ""),
        configure: (proxy) => {
          proxy.on("proxyReq", (req) => {
            if (process.env.VITE_ANTHROPIC_KEY) {
              req.setHeader("x-api-key", process.env.VITE_ANTHROPIC_KEY);
              req.setHeader("anthropic-version", "2023-06-01");
            }
          });
        },
      },
    },
  },
});
