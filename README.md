# Control Docente Integral

Aplicación para el control docente de grupos de secundaria: asistencia, evaluación,
actividades, permisos, convivencia escolar, bitácora y simulador ECOEMS.

Funciona en computadora, tableta y celular. **Guarda todo en el dispositivo y sirve
sin internet.**

---

## Cómo se guardan los datos

- La información vive en el dispositivo, en una base de datos del navegador (IndexedDB).
- Se guarda sola cada vez que capturas algo. No hay botón de "sincronizar".
- No se pierde al cerrar la app, apagar el teléfono ni quedarte sin señal.
- La app le pide permiso al navegador para que **no borre los datos** aunque el
  dispositivo se quede sin espacio. Puedes revisarlo en `Configuración → Respaldo`.
- **Nada se elimina de golpe.** Todo lo que borras va a la papelera y puede regresar.
  Solo desaparece si lo eliminas de la papelera a propósito.
- Cada día que usas la app se guarda una copia automática. Se conservan las diez
  más recientes y puedes restaurar cualquiera.
- Además puedes descargar un respaldo en archivo para cambiar de dispositivo.

Lo único que sí conviene hacer a mano: **descargar un respaldo al final de cada
trimestre** y antes de cambiar de teléfono.

---

## Reportes en PDF

Todos los reportes se descargan directo como archivo PDF, sin pasar por el cuadro
de impresión. Están en:

| Dónde | Qué descarga |
|---|---|
| Reportes | Asistencia, evaluación, actividades, por materia, permisos, incidencias, convivencia, seguimiento, ECOEMS, comparativo ECOEMS, bitácora e integral del grupo |
| Ficha del alumno | Reporte integral del estudiante |
| Incidencias | Formato oficial de reporte con los espacios de firma |
| Permisos | Listas de autorizados y de quienes no entregaron |
| Evaluación | Concentrado de calificaciones del trimestre |
| ECOEMS | Concentrado del simulador con semáforo |
| Estadísticas | La tabla que estés viendo |

Los PDF llevan el nombre de la escuela, el CCT, el ciclo escolar, la fecha y el
número de página.

---

## Publicarla e instalarla en el celular

La primera carga conviene hacerla desde una computadora, aunque sea prestada:
el navegador del celular no permite subir carpetas completas. Es una sola vez.
Después todo se hace desde el teléfono.

### 1. Subir el proyecto a GitHub

1. Entra a [github.com](https://github.com) y crea un repositorio nuevo.
   Ponle de nombre **control-docente-integral** y déjalo **Public**.

   > GitHub solo publica sitios gratis desde repositorios públicos. Lo que queda
   > visible es el programa, nunca la información de las alumnas y los alumnos:
   > esos datos se guardan solo en tu dispositivo y no se suben a ningún lado.

2. Descomprime este archivo y sube **el contenido** de la carpeta `cdi`, no la
   carpeta misma (botón *Add file → Upload files*). Activa antes la vista de
   archivos ocultos para que se incluya la carpeta `.github`.

### 2. Encender la publicación automática

1. En el repositorio entra a **Settings → Pages**.
2. En *Source* elige **GitHub Actions**.

Listo. Cada vez que subas un cambio, GitHub compila la app solo y la publica.
El proceso tarda unos dos minutos y lo puedes ver en la pestaña **Actions**.

Tu dirección queda así:

```
https://TU-USUARIO.github.io/control-docente-integral/
```

> Si le pones otro nombre al repositorio, no hay que tocar nada: el flujo de
> trabajo toma el nombre automáticamente.

### 3. Instalarla en el teléfono

Abre esa dirección en Chrome y elige **Instalar aplicación** en el menú de los
tres puntos. Queda con su ícono en la pantalla de inicio, a pantalla completa y
funcionando sin internet.

En iPhone: Safari → Compartir → **Agregar a la pantalla de inicio**.

### 4. Generar el APK (opcional)

Si quieres un archivo instalable para repartir entre compañeros:

1. Entra a [pwabuilder.com](https://www.pwabuilder.com).
2. Pega la dirección de tu app.
3. Elige **Android** y descarga el paquete.

Viene un `.apk` para instalar directo (hay que permitir "orígenes desconocidos"
en el teléfono) y un `.aab` por si algún día lo subes a Google Play.

---

## Trabajar en la app desde la computadora

Solo si quieres modificar el código:

```bash
npm install
npm run dev      # abre en http://localhost:5173
npm run build    # genera la carpeta dist/
```

---

## Estructura

```
src/
  App.jsx       Toda la aplicación y sus módulos
  almacen.js    Guardado permanente, papelera y copias automáticas
  pdf.js        Generación de los reportes en PDF
  main.jsx      Arranque y registro del service worker
public/
  manifest.webmanifest   Datos de la app instalable
  sw.js                  Permite abrir sin internet
  icono-*.png            Íconos
```

---

## Datos de prueba

La primera vez, la app se abre con un grupo ficticio **3° D** de diez estudiantes
inventados, con asistencia, actividades, calificaciones, un permiso, una
incidencia, una bitácora y dos aplicaciones del simulador ECOEMS.

Sirve para que explores todo antes de capturar información real. Cuando vayas a
empezar en serio, crea tu propio grupo y elimina el de prueba.

---

## Aviso

El sistema es una herramienta de apoyo y seguimiento docente. No sustituye la
normativa, los protocolos ni las decisiones de la autoridad escolar. La
información de las alumnas y los alumnos es confidencial y su uso se limita a
fines educativos.
