import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base: nombre del repositorio en GitHub Pages.
// Si publicas en https://usuario.github.io/control-docente-integral/
// deja "/control-docente-integral/". Si usas dominio propio, pon "/".
export default defineConfig({
  base: process.env.VITE_BASE || "/control-docente-integral/",
  plugins: [react()],
  build: {
    outDir: "dist",
    chunkSizeWarningLimit: 1500,
    rollupOptions: {
      output: {
        manualChunks: {
          react: ["react", "react-dom"],
          charts: ["recharts"],
          docs: ["jspdf", "jspdf-autotable", "xlsx"],
        },
      },
    },
  },
});
