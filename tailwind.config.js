/* ============================================================
   PALETA DE LA APLICACIÓN
   Un solo lugar para todos los colores. Cambiar un valor aquí
   lo cambia en toda la app, sin tocar las pantallas.

   Los nombres son los de Tailwind, pero los valores son propios:
   más oscuros en los textos y más visibles en las líneas, para que
   todo se lea bien incluso en el patio a pleno sol.
   ============================================================ */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      colors: {
        /* Neutros: papel, líneas y tinta */
        slate: {
          50: "#EDEAE0",   // fondo de la app, tono papel bond
          100: "#E3DFD2",  // rellenos suaves
          200: "#C6C0AF",  // líneas y bordes, ahora sí visibles
          300: "#A8A192",  // bordes marcados
          400: "#5F6862",  // texto terciario (antes casi invisible)
          500: "#474E4A",  // texto secundario
          600: "#363C39",
          700: "#262B29",
          800: "#171C1A",
          900: "#0D1110",  // tinta
        },

        /* Verde: asistencia y acción principal */
        emerald: {
          50: "#E3F0E8",
          100: "#C8E2D2",
          200: "#94C6AB",
          300: "#5FA882",
          400: "#2F8A5D",
          500: "#157347",
          600: "#0F5D39",
          700: "#0B4A2C",
          800: "#083A23",
          900: "#052A19",
        },

        /* Azul: académico, actividades y evaluación */
        sky: {
          50: "#E4EBFA",
          100: "#C9D6F4",
          200: "#9DB3E8",
          300: "#6D8EDB",
          400: "#4067CA",
          500: "#2450BC",
          600: "#1E40AF",
          700: "#1B3A9C",
          800: "#16307F",
          900: "#112563",
        },

        /* Ámbar: ECOEMS y retardos */
        amber: {
          50: "#FBF0DF",
          100: "#F6E1BE",
          200: "#E9C287",
          300: "#D9A24F",
          400: "#C4881F",
          500: "#A87410",
          600: "#92640C",
          700: "#7A530A",
          800: "#5F4108",
          900: "#452F06",
        },

        /* Vino: convivencia y faltas */
        rose: {
          50: "#FBE8EA",
          100: "#F5CFD4",
          200: "#E8A3AC",
          300: "#D6707E",
          400: "#C44257",
          500: "#B0263C",
          600: "#991B31",
          700: "#7F1527",
          800: "#66111F",
          900: "#4D0D18",
        },

        /* Naranja: avisos que piden intervención */
        orange: {
          50: "#FCEDE0",
          100: "#F8D9C0",
          200: "#EFB685",
          300: "#E09248",
          400: "#CB741C",
          500: "#B0600F",
          600: "#96500C",
          700: "#7C420A",
          800: "#5F3308",
          900: "#452506",
        },

        /* Violeta: permisos y actividades escolares */
        violet: {
          50: "#EEE9F7",
          100: "#DCD2EF",
          200: "#BBA8DE",
          300: "#977CC9",
          400: "#7455B0",
          500: "#5C3D97",
          600: "#4C3180",
          700: "#3E2769",
          800: "#2F1D50",
          900: "#221538",
        },
      },
      fontFamily: {
        sans: ["system-ui", "-apple-system", "Segoe UI", "sans-serif"],
      },
    },
  },
  plugins: [],
};
