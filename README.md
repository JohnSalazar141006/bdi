# Sistema Contable Margarita — cómo correrlo

App web con el agente contable, los balances (comprobación, cierre, persona),
desagregaciones e informes. Todo corre en tu navegador; la memoria (historial,
respaldo, tasa) se guarda en el navegador.

## Requisito único: Node.js
Descarga e instala Node.js LTS desde https://nodejs.org (versión 18 o superior).
Para comprobar que quedó instalado, abre una terminal y escribe:

    node --version

## Paso a paso (3 comandos)
1. Abre una terminal DENTRO de esta carpeta (la que tiene package.json).
2. Instala las dependencias (solo la primera vez):

       npm install

3. Arranca la app:

       npm run dev

4. Abre en el navegador la dirección que aparece (normalmente):

       http://localhost:5173

Listo. Ya puedes usar Panel, Comprobación, Cierre, Persona natural,
Desagregación, Informes e Historial. Cada documento descarga su Excel/PDF.

## Qué funciona sin nada más
- Todos los balances y el análisis inteligente (cuadre, alertas).
- Desagregaciones con proyección.
- Informes (constancias y certificaciones).
- Deshacer/Rehacer (Ctrl+Z / Ctrl+Y).
- Historial + respaldo exportable.

## Activar la pestaña "Agente" (chat con IA) — opcional
El chat necesita tu clave de Anthropic. Dos pasos:
1. Crea un archivo llamado  .env  en esta carpeta con esta línea:

       VITE_ANTHROPIC_KEY=pega_aquí_tu_clave

2. En  src/App.jsx  busca  https://api.anthropic.com/v1/messages
   y cámbialo por  /anthropic/v1/messages
   (el proxy de vite.config.js le pone la clave por detrás).

Reinicia con  npm run dev  y el Agente quedará activo.

## Para publicarlo (opcional)
    npm run build      → genera la carpeta dist/ lista para subir a
                         Vercel, Netlify o cualquier hosting estático.

## Notas
- La tasa BCV viene sembrada; puedes cambiarla desde el chat ("fija la tasa BCV en 630").
- Si algo se borra, entra a Historial → Exportar respaldo para guardar todo en un .json.
