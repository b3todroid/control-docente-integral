/* ============================================================
   GENERACIÓN DE PDF
   Descarga directa, sin pasar por el cuadro de impresión.
   Funciona sin internet: la librería viaja dentro de la app.
   ============================================================ */

import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";

const VERDE = [4, 120, 87];
const GRIS = [100, 116, 139];
const GRIS_CLARO = [241, 245, 249];
const TINTA = [15, 23, 42];

const fFecha = (s) => {
  if (!s) return "—";
  const [y, m, d] = String(s).split("-");
  return `${d}/${m}/${y}`;
};

const hoyLargo = () =>
  new Date().toLocaleDateString("es-MX", { day: "numeric", month: "long", year: "numeric" });

function nombreArchivo(titulo, extra = "") {
  const limpio = titulo
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "_")
    .replace(/^_|_$/g, "")
    .toLowerCase();
  const fecha = new Date().toISOString().slice(0, 10);
  return `${limpio}${extra ? "_" + extra : ""}_${fecha}.pdf`;
}

/* ---------- encabezado y pie institucionales ---------- */

function encabezado(doc, { titulo, subtitulo, config, ciclo }) {
  const ancho = doc.internal.pageSize.getWidth();
  doc.setFillColor(...VERDE);
  doc.rect(0, 0, ancho, 3, "F");

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(...TINTA);
  doc.text((config?.escuela || "Escuela").toUpperCase(), ancho / 2, 14, { align: "center", maxWidth: ancho - 28 });

  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(...GRIS);
  const datos = [config?.cct ? `CCT ${config.cct}` : "", ciclo ? `Ciclo escolar ${ciclo}` : "", config?.turno ? `Turno ${config.turno}` : ""]
    .filter(Boolean)
    .join("   ·   ");
  if (datos) doc.text(datos, ancho / 2, 19.5, { align: "center" });

  doc.setDrawColor(...VERDE);
  doc.setLineWidth(0.4);
  doc.line(14, 23, ancho - 14, 23);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  doc.setTextColor(...TINTA);
  doc.text(titulo, 14, 31);

  let y = 31;
  if (subtitulo) {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(...GRIS);
    doc.text(subtitulo, 14, 36.5, { maxWidth: ancho - 28 });
    y = 36.5;
  }
  return y + 6;
}

function pie(doc) {
  const paginas = doc.internal.getNumberOfPages();
  const ancho = doc.internal.pageSize.getWidth();
  const alto = doc.internal.pageSize.getHeight();
  for (let i = 1; i <= paginas; i++) {
    doc.setPage(i);
    doc.setDrawColor(226, 232, 240);
    doc.setLineWidth(0.2);
    doc.line(14, alto - 12, ancho - 14, alto - 12);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(...GRIS);
    doc.text(`Control Docente Integral · ${hoyLargo()}`, 14, alto - 8);
    doc.text(`Página ${i} de ${paginas}`, ancho - 14, alto - 8, { align: "right" });
  }
}

function nuevoDoc(orientacion = "p") {
  return new jsPDF({ orientation: orientacion, unit: "mm", format: "letter" });
}

/* ============================================================
   1. Reporte de tabla (asistencia, evaluación, ECOEMS, etc.)
   ============================================================ */

export function pdfTabla({ titulo, subtitulo, columnas, filas, config, ciclo, nota, resumen }) {
  const horizontal = columnas.length > 7;
  const doc = nuevoDoc(horizontal ? "l" : "p");
  let y = encabezado(doc, { titulo, subtitulo, config, ciclo });

  if (resumen?.length) {
    const ancho = doc.internal.pageSize.getWidth();
    const cajaAncho = (ancho - 28) / resumen.length;
    resumen.forEach((r, i) => {
      const x = 14 + i * cajaAncho;
      doc.setFillColor(...GRIS_CLARO);
      doc.roundedRect(x, y, cajaAncho - 2.5, 15, 1.5, 1.5, "F");
      doc.setFont("helvetica", "normal");
      doc.setFontSize(7.5);
      doc.setTextColor(...GRIS);
      doc.text(String(r.label), x + 3, y + 5.5, { maxWidth: cajaAncho - 8 });
      doc.setFont("helvetica", "bold");
      doc.setFontSize(12);
      doc.setTextColor(...TINTA);
      doc.text(String(r.valor), x + 3, y + 12);
    });
    y += 21;
  }

  autoTable(doc, {
    startY: y,
    head: [columnas],
    body: filas.length ? filas.map((f) => f.map((c) => (c === null || c === undefined ? "—" : String(c)))) : [[{ content: "Sin registros para estos filtros", colSpan: columnas.length, styles: { halign: "center", textColor: GRIS } }]],
    theme: "grid",
    headStyles: { fillColor: VERDE, textColor: 255, fontSize: 8, fontStyle: "bold", cellPadding: 2 },
    bodyStyles: { fontSize: 8, cellPadding: 1.8, textColor: TINTA },
    alternateRowStyles: { fillColor: [249, 250, 251] },
    styles: { lineColor: [226, 232, 240], lineWidth: 0.1, overflow: "linebreak" },
    margin: { left: 14, right: 14, bottom: 18 },
  });

  if (nota) {
    const yFinal = doc.lastAutoTable.finalY + 6;
    doc.setFont("helvetica", "italic");
    doc.setFontSize(7.5);
    doc.setTextColor(...GRIS);
    doc.text(nota, 14, yFinal, { maxWidth: doc.internal.pageSize.getWidth() - 28 });
  }

  pie(doc);
  doc.save(nombreArchivo(titulo));
  return true;
}

/* ============================================================
   2. Formato de reporte de incidencia, con espacios de firma
   ============================================================ */

export function pdfIncidencia({ incidencia, alumno, grupo, config, ciclo, avisoLegal }) {
  const doc = nuevoDoc("p");
  const ancho = doc.internal.pageSize.getWidth();
  const i = incidencia;

  doc.setFillColor(...VERDE);
  doc.rect(0, 0, ancho, 3, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(...TINTA);
  doc.text((config?.escuela || "Escuela").toUpperCase(), ancho / 2, 14, { align: "center", maxWidth: ancho - 28 });
  doc.setFontSize(10.5);
  doc.text("REPORTE / INCIDENCIA DE HECHOS", ancho / 2, 20.5, { align: "center" });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(...GRIS);
  doc.text(
    [`Folio ${i.folio}`, config?.cct ? `CCT ${config.cct}` : "", `Ciclo ${ciclo}`].filter(Boolean).join("   ·   "),
    ancho / 2,
    25.5,
    { align: "center" }
  );

  const par = (a, b, c, d) => [a, b, c, d];
  autoTable(doc, {
    startY: 30,
    body: [
      par("Fecha", fFecha(i.fecha), "Hora", i.hora || "—"),
      par("Alumna(o)", alumno, "Grado y grupo", grupo),
      par("Docente que reporta", i.docente || "—", "Lugar", i.lugar || "—"),
      par("Tipo de conducta", i.conducta || i.tipo || "—", "Clasificación", i.gravedad || "—"),
    ],
    theme: "grid",
    bodyStyles: { fontSize: 8.5, cellPadding: 2, textColor: TINTA },
    columnStyles: {
      0: { fillColor: GRIS_CLARO, fontStyle: "bold", cellWidth: 34 },
      1: { cellWidth: 55 },
      2: { fillColor: GRIS_CLARO, fontStyle: "bold", cellWidth: 34 },
      3: { cellWidth: "auto" },
    },
    styles: { lineColor: [203, 213, 225], lineWidth: 0.1 },
    margin: { left: 14, right: 14 },
  });

  let y = doc.lastAutoTable.finalY + 5;
  const bloque = (titulo, texto, minAlto = 14) => {
    if (y > doc.internal.pageSize.getHeight() - 50) {
      doc.addPage();
      y = 20;
    }
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.5);
    doc.setTextColor(...TINTA);
    doc.text(titulo.toUpperCase(), 14, y);
    y += 2;
    const contenido = texto && String(texto).trim() ? String(texto) : "—";
    const lineas = doc.splitTextToSize(contenido, ancho - 34);
    const alto = Math.max(minAlto, lineas.length * 4.2 + 5);
    doc.setDrawColor(203, 213, 225);
    doc.setLineWidth(0.2);
    doc.roundedRect(14, y, ancho - 28, alto, 1, 1, "S");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.5);
    doc.setTextColor(...TINTA);
    doc.text(lineas, 17, y + 5);
    y += alto + 5;
  };

  bloque("Descripción objetiva de los hechos", i.descripcion, 24);
  bloque("Personas involucradas", [i.involucrados, i.testigos ? `Testigos: ${i.testigos}` : ""].filter(Boolean).join("\n"));
  bloque("Evidencias o documentos relacionados", i.evidencias);
  bloque("Medidas de atención realizadas", [(i.medidas || []).join(" · "), i.accionInmediata].filter(Boolean).join("\n"));
  bloque("Acuerdos", i.acuerdos);
  bloque(
    "Seguimiento",
    [
      (i.seguimientos || []).map((s) => `${fFecha(s.fecha)}: ${s.nota}`).join("\n"),
      i.fechaSeguimiento ? `Fecha de seguimiento programada: ${fFecha(i.fechaSeguimiento)}` : "",
      `Estatus: ${i.estado}`,
    ]
      .filter(Boolean)
      .join("\n")
  );
  bloque("Observaciones", i.observaciones);

  // firmas
  if (y > doc.internal.pageSize.getHeight() - 60) {
    doc.addPage();
    y = 24;
  }
  y += 6;
  const firmas = ["Firma del docente", "Firma de prefectura", "Firma de orientación", "Firma de madre, padre o tutor", "Firma de autoridad escolar"];
  const anchoFirma = (ancho - 28) / 3;
  firmas.forEach((f, idx) => {
    const col = idx % 3;
    const fila = Math.floor(idx / 3);
    const x = 14 + col * anchoFirma;
    const yf = y + fila * 24;
    doc.setDrawColor(...GRIS);
    doc.setLineWidth(0.2);
    doc.line(x + 4, yf + 14, x + anchoFirma - 8, yf + 14);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(...GRIS);
    doc.text(f, x + anchoFirma / 2 - 2, yf + 18, { align: "center", maxWidth: anchoFirma - 6 });
  });
  y += 24 * Math.ceil(firmas.length / 3) + 4;

  if (avisoLegal) {
    doc.setFont("helvetica", "italic");
    doc.setFontSize(7);
    doc.setTextColor(...GRIS);
    doc.text(avisoLegal, 14, y, { maxWidth: ancho - 28 });
  }

  pie(doc);
  doc.save(nombreArchivo("incidencia", i.folio));
  return true;
}

/* ============================================================
   3. Reporte integral del alumno
   ============================================================ */

export function pdfAlumno({ alumno, grupo, config, ciclo, secciones, nota }) {
  const doc = nuevoDoc("p");
  let y = encabezado(doc, {
    titulo: "Reporte integral del estudiante",
    subtitulo: `${alumno.nombre}${grupo ? "  ·  " + grupo : ""}`,
    config,
    ciclo,
  });

  autoTable(doc, {
    startY: y,
    body: alumno.datos,
    theme: "grid",
    bodyStyles: { fontSize: 8.5, cellPadding: 2, textColor: TINTA },
    columnStyles: {
      0: { fillColor: GRIS_CLARO, fontStyle: "bold", cellWidth: 36 },
      1: { cellWidth: 52 },
      2: { fillColor: GRIS_CLARO, fontStyle: "bold", cellWidth: 36 },
      3: { cellWidth: "auto" },
    },
    styles: { lineColor: [203, 213, 225], lineWidth: 0.1 },
    margin: { left: 14, right: 14 },
  });
  y = doc.lastAutoTable.finalY + 7;

  secciones.forEach((sec) => {
    if (y > doc.internal.pageSize.getHeight() - 45) {
      doc.addPage();
      y = 20;
    }
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9.5);
    doc.setTextColor(...VERDE);
    doc.text(sec.titulo.toUpperCase(), 14, y);
    doc.setDrawColor(...VERDE);
    doc.setLineWidth(0.3);
    doc.line(14, y + 1.5, doc.internal.pageSize.getWidth() - 14, y + 1.5);

    autoTable(doc, {
      startY: y + 4,
      head: sec.columnas ? [sec.columnas] : undefined,
      body: sec.filas.length ? sec.filas : [[{ content: "Sin registros", colSpan: sec.columnas?.length || 1, styles: { halign: "center", textColor: GRIS } }]],
      theme: "grid",
      headStyles: { fillColor: [241, 245, 249], textColor: TINTA, fontSize: 7.5, fontStyle: "bold", cellPadding: 1.8 },
      bodyStyles: { fontSize: 8, cellPadding: 1.6, textColor: TINTA },
      styles: { lineColor: [226, 232, 240], lineWidth: 0.1, overflow: "linebreak" },
      margin: { left: 14, right: 14, bottom: 18 },
    });
    y = doc.lastAutoTable.finalY + 8;
  });

  if (nota) {
    if (y > doc.internal.pageSize.getHeight() - 30) {
      doc.addPage();
      y = 20;
    }
    doc.setFont("helvetica", "italic");
    doc.setFontSize(7);
    doc.setTextColor(...GRIS);
    doc.text(nota, 14, y, { maxWidth: doc.internal.pageSize.getWidth() - 28 });
  }

  pie(doc);
  doc.save(nombreArchivo("reporte_integral", alumno.archivo || ""));
  return true;
}

/* ============================================================
   4. Listas de permisos (autorizados y pendientes)
   ============================================================ */

export function pdfPermisos({ permiso, grupo, config, ciclo, autorizados, pendientes }) {
  const doc = nuevoDoc("p");
  let y = encabezado(doc, {
    titulo: permiso.nombre,
    subtitulo: `${fFecha(permiso.fecha)}  ·  ${grupo}  ·  ${permiso.lugar || "Sin lugar"}  ·  Responsable: ${permiso.responsable || "—"}`,
    config,
    ciclo,
  });

  const tabla = (titulo, filas, color) => {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9.5);
    doc.setTextColor(...TINTA);
    doc.text(`${titulo} (${filas.length})`, 14, y);
    autoTable(doc, {
      startY: y + 3,
      head: [["No.", "Nombre", "Entregó", "Autorizado", "Observaciones"]],
      body: filas.length ? filas : [[{ content: "Sin estudiantes en esta lista", colSpan: 5, styles: { halign: "center", textColor: GRIS } }]],
      theme: "grid",
      headStyles: { fillColor: color, textColor: 255, fontSize: 8, fontStyle: "bold", cellPadding: 2 },
      bodyStyles: { fontSize: 8, cellPadding: 1.8, textColor: TINTA },
      columnStyles: { 0: { cellWidth: 12 }, 2: { cellWidth: 20 }, 3: { cellWidth: 24 } },
      styles: { lineColor: [226, 232, 240], lineWidth: 0.1 },
      margin: { left: 14, right: 14, bottom: 18 },
    });
    y = doc.lastAutoTable.finalY + 9;
  };

  tabla("Autorizados", autorizados, VERDE);
  if (y > doc.internal.pageSize.getHeight() - 60) {
    doc.addPage();
    y = 20;
  }
  tabla("Sin permiso entregado", pendientes, [225, 29, 72]);

  pie(doc);
  doc.save(nombreArchivo(permiso.nombre));
  return true;
}
