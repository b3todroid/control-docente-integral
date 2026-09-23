import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";
import { Home, Users, GraduationCap, CalendarCheck, FileText, BookOpen, BarChart3, Target, AlertTriangle, ClipboardList, NotebookPen, TrendingUp, Calendar, Settings, Search, Plus, Trash2, Pencil, Download, Upload, Save, X, ChevronRight, ChevronLeft, Printer, Check, Menu, ArrowLeft, Database, RotateCcw, Bell, Info, ShieldCheck, Undo2, FileDown, WifiOff } from "lucide-react";
import { leer, guardar, pedirPermanencia, espacio, aPapelera, verPapelera, sacarDePapelera, borrarDePapelera, vaciarPapelera, respaldoAutomatico, verRespaldos, leerRespaldo } from "./almacen";
import { pdfTabla, pdfIncidencia, pdfAlumno, pdfPermisos } from "./pdf";

/* ============================================================
   CONTROL DOCENTE INTEGRAL
   Sistema de control docente, académico, asistencia,
   convivencia y seguimiento ECOEMS.
   ============================================================ */

/* ---------- utilidades ---------- */
const uid = (p = "id") => p + "_" + Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-4);
const hoy = () => new Date().toISOString().slice(0, 10);
const fFecha = (s) => { if (!s) return "—"; const [y, m, d] = String(s).split("-"); return `${d}/${m}/${y}`; };
const nomComp = (a) => (a ? `${a.apellidos || ""} ${a.nombre || ""}`.trim() : "—");
const round = (n, d = 1) => (isFinite(n) ? Number(Number(n).toFixed(d)) : 0);
const pct = (a, b) => (b > 0 ? round((a * 100) / b, 1) : 0);
const clone = (o) => JSON.parse(JSON.stringify(o));
const sortAl = (arr) => [...arr].sort((a, b) => (a.numLista || 999) - (b.numLista || 999) || nomComp(a).localeCompare(nomComp(b)));
const norm = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

/* ---------- catálogos base ---------- */
const EST = {
  A: { label: "Asistencia", corto: "A", full: "bg-emerald-600 text-white border-emerald-600", soft: "bg-emerald-50 text-emerald-700 border-emerald-200", dot: "#059669" },
  F: { label: "Falta", corto: "F", full: "bg-rose-600 text-white border-rose-600", soft: "bg-rose-50 text-rose-700 border-rose-200", dot: "#e11d48" },
  J: { label: "Justificada", corto: "J", full: "bg-sky-600 text-white border-sky-600", soft: "bg-sky-50 text-sky-700 border-sky-200", dot: "#0284c7" },
  R: { label: "Retardo", corto: "R", full: "bg-amber-500 text-white border-amber-500", soft: "bg-amber-50 text-amber-700 border-amber-200", dot: "#f59e0b" },
  P: { label: "Permiso", corto: "P", full: "bg-violet-600 text-white border-violet-600", soft: "bg-violet-50 text-violet-700 border-violet-200", dot: "#7c3aed" },
};
const ORDEN_EST = ["A", "F", "J", "R", "P"];

const NIVELES = {
  verde: { label: "Seguimiento docente", cls: "bg-emerald-50 text-emerald-800 border-emerald-300", punto: "bg-emerald-500" },
  amarillo: { label: "Seguimiento con tutoría u orientación", cls: "bg-amber-50 text-amber-800 border-amber-300", punto: "bg-amber-400" },
  naranja: { label: "Requiere intervención de autoridad escolar", cls: "bg-orange-50 text-orange-800 border-orange-300", punto: "bg-orange-500" },
  rojo: { label: "Requiere atención inmediata conforme a protocolos", cls: "bg-rose-50 text-rose-800 border-rose-300", punto: "bg-rose-500" },
};
const EST_INC = ["Abierta", "En seguimiento", "Cerrada"];
const EST_ACT = ["Pendiente", "Entregada", "Fuera de tiempo", "No entregada"];
const EST_BIT = ["Concluida", "Parcialmente concluida", "Reprogramada", "Suspendida"];
const EST_ALUMNO = ["Activo", "Baja", "Traslado"];
const CIRC_OPC = ["Sí", "No", "No se puede determinar", "No aplica"];
const CAMPOS_FORM = ["Lenguajes", "Saberes y Pensamiento Científico", "Ética, Naturaleza y Sociedades", "De lo Humano y lo Comunitario"];
const EJES = ["Inclusión", "Pensamiento crítico", "Interculturalidad crítica", "Igualdad de género", "Vida saludable", "Apropiación de las culturas a través de la lectura y la escritura", "Artes y experiencias estéticas"];
const AVISO_LEGAL = "Esta valoración es un registro administrativo y no sustituye la normativa, protocolos o procedimientos establecidos por la autoridad educativa.";
const AVISO_CONVIVENCIA = "La clasificación es un campo administrativo de seguimiento. No sustituye la valoración que corresponda conforme al Marco para la Convivencia Escolar y demás normativa aplicable, ni genera sanciones de manera automática.";
const AVISO_ECOEMS = "El semáforo es una referencia de trabajo definida por el docente. No es una predicción de ingreso a ninguna institución de educación media superior.";

const GRAVEDAD = {
  Leve: { cls: "bg-emerald-50 text-emerald-800 border-emerald-200", punto: "bg-emerald-500" },
  Moderada: { cls: "bg-amber-50 text-amber-800 border-amber-200", punto: "bg-amber-400" },
  Grave: { cls: "bg-orange-50 text-orange-800 border-orange-200", punto: "bg-orange-500" },
  "Alta prioridad de atención": { cls: "bg-rose-50 text-rose-800 border-rose-200", punto: "bg-rose-500" },
};
const MEDIDAS_ATENCION = [
  "Diálogo y orientación", "Acuerdo de convivencia", "Aviso al tutor del grupo",
  "Comunicación con madre, padre o tutor", "Citatorio", "Canalización a orientación educativa",
  "Intervención de prefectura", "Intervención de dirección", "Seguimiento individual",
  "Acciones restaurativas", "Otra medida establecida por la autoridad escolar",
];

const PERFILES = {
  Administrador: { modulos: null, desc: "Acceso completo al sistema" },
  Docente: { modulos: ["inicio", "grupos", "alumnos", "asistencia", "permisos", "actividades", "evaluacion", "ecoems", "incidencias", "valoracion", "bitacora", "estadisticas", "reportes", "calendario", "configuracion"], desc: "Sus grupos, evaluación, asistencia, actividades, bitácora e incidencias" },
  "Tutor de grupo": { modulos: ["inicio", "grupos", "alumnos", "asistencia", "evaluacion", "ecoems", "incidencias", "estadisticas", "reportes", "calendario"], desc: "Seguimiento integral de su grupo" },
  "Orientación y prefectura": { modulos: ["inicio", "alumnos", "asistencia", "incidencias", "valoracion", "estadisticas", "reportes", "calendario"], desc: "Incidencias y seguimiento de convivencia" },
};

const semaforoDe = (porcentaje, cfg) => {
  const s = cfg?.semaforo || { alto: 80, medio: 60 };
  if (porcentaje >= s.alto) return { nivel: "alto", label: "Desempeño alto", cls: "bg-emerald-50 text-emerald-800 border-emerald-200", punto: "bg-emerald-500", color: "#047857" };
  if (porcentaje >= s.medio) return { nivel: "medio", label: "Desempeño medio", cls: "bg-amber-50 text-amber-800 border-amber-200", punto: "bg-amber-400", color: "#f59e0b" };
  return { nivel: "bajo", label: "Requiere reforzamiento", cls: "bg-rose-50 text-rose-800 border-rose-200", punto: "bg-rose-500", color: "#e11d48" };
};

const catalogosBase = () => ({
  tiposActividad: ["Actividad", "Proyecto", "Examen", "Exposición", "Participación", "Tarea", "Producto final", "Trabajo colaborativo", "Otro"],
  materias: ["Español", "Matemáticas", "Biología", "Física", "Química", "Historia", "Geografía", "Formación Cívica y Ética", "Inglés", "Artes", "Educación Física", "Tecnología", "Tutoría"],
  conductas: [
    "Conductas que perturban el orden",
    "Conductas altamente perturbadoras",
    "Conductas que afectan la convivencia escolar",
    "Agresiones o violencia entre integrantes de la comunidad escolar",
    "Faltas de respeto",
    "Incumplimiento de normas escolares",
    "Daño a instalaciones, materiales o bienes",
    "Uso inadecuado de dispositivos electrónicos",
    "Conductas relacionadas con acoso escolar",
    "Conductas relacionadas con ciberacoso",
    "Otras conductas previstas por la normativa escolar",
  ],
  instrumentos: [
    { id: "ins_act", nombre: "Actividades", porcentaje: 40 },
    { id: "ins_part", nombre: "Participación", porcentaje: 10 },
    { id: "ins_proy", nombre: "Proyecto", porcentaje: 20 },
    { id: "ins_exa", nombre: "Examen", porcentaje: 30 },
  ],
  materiasEcoems: [
    { id: "e_esp", nombre: "Español", reactivos: 12 },
    { id: "e_hverb", nombre: "Habilidad verbal", reactivos: 16 },
    { id: "e_mat", nombre: "Matemáticas", reactivos: 12 },
    { id: "e_hmat", nombre: "Habilidad matemática", reactivos: 16 },
    { id: "e_bio", nombre: "Biología", reactivos: 12 },
    { id: "e_fis", nombre: "Física", reactivos: 12 },
    { id: "e_qui", nombre: "Química", reactivos: 12 },
    { id: "e_geo", nombre: "Geografía", reactivos: 12 },
    { id: "e_his", nombre: "Historia", reactivos: 12 },
    { id: "e_fce", nombre: "Formación Cívica y Ética", reactivos: 12 },
  ],
});

const configBase = (ciclo) => ({
  escuela: "Escuela Secundaria N.º 82 «Abraham Lincoln»", cct: "", docente: "", asignatura: "", turno: "Matutino", horario: "",
  ciclo, trimestre: 1, escalaMax: 10, escalaMin: 5,
  trimestres: [
    { n: 1, inicio: `${ciclo.slice(0, 4)}-08-25`, fin: `${ciclo.slice(0, 4)}-11-28` },
    { n: 2, inicio: `${ciclo.slice(0, 4)}-12-01`, fin: `${ciclo.slice(-4)}-03-20` },
    { n: 3, inicio: `${ciclo.slice(-4)}-03-23`, fin: `${ciclo.slice(-4)}-07-15` },
  ],
  semaforo: { alto: 80, medio: 60 },
  alertas: { faltasMax: 3, asistenciaMin: 80, promedioMin: 6, ecoemsMin: 60, actPendientes: 2, diasSeguimiento: 7 },
});

const dbVacia = (ciclo) => ({
  ciclo, config: configBase(ciclo), catalogos: catalogosBase(),
  grupos: [], alumnos: [], asistencias: [], actividades: [], entregas: [],
  permisos: [], incidencias: [], valoraciones: [], bitacoras: [], ecoems: [], eventos: [],
  folio: 0,
});

const trimestreDe = (fecha, cfg) => {
  const f = fecha || hoy();
  const t = (cfg?.trimestres || []).find((x) => x.inicio && x.fin && f >= x.inicio && f <= x.fin);
  if (t) return t.n;
  const m = Number(f.slice(5, 7));
  if (m >= 8 && m <= 11) return 1;
  if (m === 12 || m <= 3) return 2;
  return 3;
};

/* ---------- datos ficticios de prueba ---------- */
const NOMBRES_DEMO = [
  ["Aguilar Ríos", "Ana Sofía"], ["Bautista Mena", "Bruno"], ["Cordero Vela", "Camila"],
  ["Delgado Nava", "Diego"], ["Escobar Lira", "Elena"], ["Fuentes Ávila", "Fernando"],
  ["Guzmán Prado", "Gabriela"], ["Herrera Solís", "Hugo"], ["Ibarra Cano", "Itzel"],
  ["Jiménez Rosas", "Javier"],
];

function generarDemo(ciclo) {
  const db = dbVacia(ciclo);
  db.config.escuela = "Escuela Secundaria (datos de prueba)";
  db.config.cct = "00XXX0000X";
  db.config.docente = "Docente de prueba";
  db.config.asignatura = "Matemáticas";
  const g = { id: uid("grp"), grado: 3, grupo: "D", turno: "Matutino", asignatura: "Matemáticas", docente: "Docente de prueba", activo: true };
  db.grupos.push(g);
  db.alumnos = NOMBRES_DEMO.map((n, i) => ({
    id: uid("alu"), numLista: i + 1, apellidos: n[0], nombre: n[1], grupoId: g.id,
    curp: "XXXX000000XXXXXX00", turno: "Matutino", tutorGrupo: "Docente de prueba",
    tutorNombre: `Tutor(a) de ${n[1]}`, tutorTelefono: "55 0000 0000", tutorCorreo: "",
    estatus: "Activo", activo: true, observaciones: "",
  }));

  // asistencia: 20 sesiones
  const base = new Date(); base.setDate(base.getDate() - 40);
  for (let s = 0; s < 20; s++) {
    const d = new Date(base); d.setDate(d.getDate() + s * 2);
    if (d.getDay() === 0 || d.getDay() === 6) continue;
    const marcas = {};
    db.alumnos.forEach((a, i) => {
      const r = (i * 7 + s * 3) % 20;
      marcas[a.id] = r === 0 ? "F" : r === 5 ? "R" : r === 11 ? "J" : "A";
    });
    db.asistencias.push({ id: uid("asi"), grupoId: g.id, fecha: d.toISOString().slice(0, 10), sesion: 1, marcas });
  }

  // actividades + entregas
  const acts = [
    { nombre: "Operaciones con números enteros", ins: "ins_act" },
    { nombre: "Proyecto: estadística de mi comunidad", ins: "ins_proy" },
    { nombre: "Examen del primer trimestre", ins: "ins_exa" },
    { nombre: "Ecuaciones de primer grado", ins: "ins_act" },
  ];
  acts.forEach((a, i) => {
    const d = new Date(base); d.setDate(d.getDate() + i * 9);
    const act = {
      id: uid("act"), grupoId: g.id, nombre: a.nombre, instrumentoId: a.ins,
      fecha: d.toISOString().slice(0, 10), fechaEntrega: d.toISOString().slice(0, 10),
      trimestre: 1, campoFormativo: CAMPOS_FORM[1], contenido: "Contenido de prueba",
      pda: "Proceso de desarrollo de aprendizaje de prueba", eje: EJES[1], proyecto: "",
      proposito: "Propósito de prueba", descripcion: "Descripción de prueba",
      producto: "Producto esperado de prueba", valor: 10, materia: "Matemáticas", evidencia: "Cuaderno",
    };
    db.actividades.push(act);
    db.alumnos.forEach((al, j) => {
      const r = (i * 5 + j) % 9;
      db.entregas.push({
        id: uid("ent"), actividadId: act.id, alumnoId: al.id,
        estado: r === 0 ? "No entregada" : r === 1 ? "Pendiente" : "Entregada",
        calificacion: r === 0 ? 0 : round(6 + ((j * 3 + i * 2) % 5), 1),
      });
    });
  });

  // permiso
  const p = {
    id: uid("per"), nombre: "Visita al museo regional", fecha: hoy(), lugar: "Museo regional",
    responsable: "Docente de prueba", grupoId: g.id, horaSalida: "08:00", horaRegreso: "13:00",
    descripcion: "Actividad académica de prueba", registros: {},
  };
  db.alumnos.forEach((a, i) => { p.registros[a.id] = { entregado: i % 3 !== 0, fechaEntrega: i % 3 !== 0 ? hoy() : "", autorizado: i % 3 !== 0, obs: "" }; });
  db.permisos.push(p);

  // incidencia + valoración
  db.folio = 1;
  const inc = {
    id: uid("inc"), folio: `INC-${ciclo.slice(0, 4)}-0001`, fecha: hoy(), hora: "10:30",
    alumnoId: db.alumnos[3].id, grupoId: g.id, docente: "Docente de prueba", lugar: "Aula",
    conducta: "Uso inadecuado de dispositivos electrónicos", gravedad: "Leve",
    tipo: "Uso inadecuado de dispositivos electrónicos",
    descripcion: "Se registró el uso del teléfono celular durante la clase, después de la indicación de guardarlo.",
    involucrados: "", testigos: "", evidencias: "Registro en la agenda escolar",
    medidas: ["Diálogo y orientación", "Comunicación con madre, padre o tutor"],
    acuerdos: "El estudiante mantendrá el dispositivo guardado durante la sesión.",
    accionInmediata: "Se dialogó con el estudiante y se resguardó el dispositivo hasta el fin de la jornada.",
    canalizacion: "", fechaSeguimiento: hoy(), observaciones: "", estado: "En seguimiento",
    seguimientos: [{ id: uid("seg"), fecha: hoy(), nota: "Se informó a la familia por medio de la agenda escolar." }],
  };
  db.incidencias.push(inc);
  db.valoraciones.push({
    id: uid("val"), incidenciaId: inc.id, tipo: "Normativa escolar",
    circ: { afectacionPersonas: "No", afectacionMateriales: "No", intencional: "Sí", recurrente: "No", versionAlumno: "Sí", versionesCorrespondientes: "Sí" },
    nivel: "verde", medida: "Acuerdo de uso responsable del dispositivo dentro del aula.",
    acuerdo: "El estudiante mantendrá el dispositivo guardado durante la sesión.",
    comunicacion: "Se informó a la familia.", canalizacion: "", fechaSeguimiento: hoy(), resultado: "",
  });

  // bitácora
  db.bitacoras.push({
    id: uid("bit"), grupoId: g.id, fecha: hoy(), trimestre: 1, sesion: 1, horario: "08:00 - 08:50",
    campoFormativo: CAMPOS_FORM[1], contenido: "Ecuaciones de primer grado", pda: "Resuelve situaciones con ecuaciones lineales",
    eje: EJES[1], proposito: "Que el grupo resuelva ecuaciones de primer grado en contextos cotidianos",
    inicioAct: "Recuperación de conocimientos previos con ejercicios en el pizarrón",
    inicioPregunta: "¿Dónde encontramos una incógnita en la vida diaria?", inicioPrevios: "Operaciones con enteros",
    desarrollo: "Resolución guiada y después en parejas", estrategias: "Modelado, práctica guiada, trabajo colaborativo",
    individual: "Cinco ejercicios del cuaderno", colaborativo: "Resolución de un problema por equipo", recursos: "Pizarrón, cuaderno, hojas de ejercicios",
    cierreProducto: "Ejercicios resueltos", cierreEvidencia: "Cuaderno", cierreReflexion: "Puesta en común de procedimientos", cierreTarea: "Tres ejercicios",
    evalEvidencia: "Cuaderno", evalInstrumento: "Lista de cotejo", logros: "La mayoría identificó el procedimiento", dificultades: "Manejo de signos",
    apoyo: "Se dará acompañamiento en la siguiente sesión", situaciones: "", acuerdos: "Repasar signos al inicio de la próxima sesión",
    observaciones: "", reflexionLogro: "Sí",
    reflexionFunciono: "El trabajo en parejas", reflexionDificultades: "El tiempo fue ajustado", reflexionModificar: "Reducir el número de ejercicios",
    estado: "Concluida",
  });

  // ECOEMS: 2 simuladores
  const mats = db.catalogos.materiasEcoems;
  [1, 2].forEach((num, k) => {
    const d = new Date(base); d.setDate(d.getDate() + k * 20);
    db.alumnos.forEach((a, i) => {
      const res = {};
      mats.forEach((m, j) => {
        const bruto = m.reactivos * (0.45 + ((i * 3 + j * 2 + k * 4) % 10) / 22 + k * 0.06);
        res[m.id] = Math.min(m.reactivos, Math.max(0, Math.round(bruto)));
      });
      db.ecoems.push({ id: uid("eco"), alumnoId: a.id, grupoId: g.id, fecha: d.toISOString().slice(0, 10), numero: num, resultados: res });
    });
  });
  return db;
}

/* ---------- almacenamiento permanente ----------
   Los datos viven en IndexedDB dentro del dispositivo: se guardan sin
   internet, sobreviven al cerrar la app y solo se borran si se pide.  */
const K_META = "cdi:meta";
const K_CICLO = (id) => "cdi:ciclo:" + id;

const sGet = (k) => leer(k);
const sSet = async (k, v) => (await guardar(k, v)).ok;

/* ============================================================
   COMPONENTES DE INTERFAZ
   ============================================================ */
const Btn = ({ children, onClick, tipo = "primario", size = "md", icon: Icon, className = "", disabled, title }) => {
  const tipos = {
    primario: "bg-emerald-700 text-white hover:bg-emerald-800 border-emerald-700",
    secundario: "bg-white text-slate-700 hover:bg-slate-50 border-slate-300",
    peligro: "bg-rose-600 text-white hover:bg-rose-700 border-rose-600",
    suave: "bg-slate-100 text-slate-700 hover:bg-slate-200 border-slate-200",
    fantasma: "bg-transparent text-slate-600 hover:bg-slate-100 border-transparent",
  };
  const sizes = { sm: "px-2.5 py-1.5 text-xs gap-1", md: "px-3.5 py-2 text-sm gap-1.5", lg: "px-5 py-2.5 text-base gap-2" };
  return (
    <button type="button" title={title} disabled={disabled} onClick={onClick}
      className={`inline-flex items-center justify-center rounded-lg border font-medium transition-colors disabled:opacity-40 disabled:cursor-not-allowed focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:ring-offset-1 ${tipos[tipo]} ${sizes[size]} ${className}`}>
      {Icon && <Icon size={size === "sm" ? 14 : 16} />}{children}
    </button>
  );
};

const Card = ({ children, className = "", pad = true }) => (
  <div className={`bg-white border border-slate-200 rounded-xl ${pad ? "p-4" : ""} ${className}`}>{children}</div>
);

const Titulo = ({ children, sub, right }) => (
  <div className="flex items-start justify-between gap-3 mb-4 flex-wrap">
    <div>
      <h2 className="text-xl font-semibold text-slate-900">{children}</h2>
      {sub && <p className="text-sm text-slate-500 mt-0.5">{sub}</p>}
    </div>
    {right}
  </div>
);

const Campo = ({ label, children, hint, req }) => (
  <label className="block">
    <span className="block text-xs font-medium text-slate-600 mb-1">{label}{req && <span className="text-rose-500"> *</span>}</span>
    {children}
    {hint && <span className="block text-[11px] text-slate-400 mt-1">{hint}</span>}
  </label>
);

const inputCls = "w-full rounded-lg border border-slate-300 px-3 py-2 text-sm text-slate-900 bg-white focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 placeholder:text-slate-400";
const Inp = (p) => <input {...p} className={`${inputCls} ${p.className || ""}`} />;
const Sel = ({ children, ...p }) => <select {...p} className={`${inputCls} ${p.className || ""}`}>{children}</select>;
const Area = (p) => <textarea {...p} className={`${inputCls} min-h-[76px] ${p.className || ""}`} />;

const Modal = ({ open, onClose, title, children, ancho = "max-w-2xl", footer }) => {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-slate-900/50 p-0 sm:p-4" onClick={onClose}>
      <div className={`bg-white w-full ${ancho} rounded-t-2xl sm:rounded-2xl max-h-[92vh] flex flex-col`} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-4 py-3 border-b border-slate-200 shrink-0">
          <h3 className="font-semibold text-slate-900">{title}</h3>
          <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-500"><X size={18} /></button>
        </div>
        <div className="p-4 overflow-y-auto grow">{children}</div>
        {footer && <div className="px-4 py-3 border-t border-slate-200 flex justify-end gap-2 shrink-0">{footer}</div>}
      </div>
    </div>
  );
};

const Confirmar = ({ open, texto, onSi, onNo, textoSi = "Eliminar" }) => (
  <Modal open={open} onClose={onNo} title="Confirmar" ancho="max-w-md"
    footer={<><Btn tipo="secundario" onClick={onNo}>Cancelar</Btn><Btn tipo="peligro" onClick={onSi}>{textoSi}</Btn></>}>
    <p className="text-sm text-slate-700">{texto}</p>
    <p className="text-xs text-slate-500 mt-2">Lo que elimines se guarda en la papelera y puedes recuperarlo desde Configuración.</p>
  </Modal>
);

const Stat = ({ label, valor, sub, color = "text-slate-900", icon: Icon, onClick }) => (
  <div onClick={onClick} className={`bg-white border border-slate-200 rounded-xl p-3.5 ${onClick ? "cursor-pointer hover:border-emerald-400" : ""}`}>
    <div className="flex items-center justify-between">
      <span className="text-xs text-slate-500 font-medium">{label}</span>
      {Icon && <Icon size={15} className="text-slate-300" />}
    </div>
    <div className={`text-2xl font-semibold mt-1 ${color}`}>{valor}</div>
    {sub && <div className="text-[11px] text-slate-400 mt-0.5">{sub}</div>}
  </div>
);

const Pill = ({ children, cls = "bg-slate-100 text-slate-700 border-slate-200" }) => (
  <span className={`inline-flex items-center px-2 py-0.5 rounded-md text-[11px] font-medium border ${cls}`}>{children}</span>
);

const Vacio = ({ texto, accion }) => (
  <div className="text-center py-10 px-4">
    <p className="text-sm text-slate-500">{texto}</p>
    {accion && <div className="mt-3">{accion}</div>}
  </div>
);

const Aviso = ({ children }) => (
  <div className="flex gap-2 items-start bg-slate-50 border border-slate-200 rounded-lg p-3 text-[12px] text-slate-600">
    <Info size={15} className="shrink-0 mt-0.5 text-slate-400" /><span>{children}</span>
  </div>
);

const Tabs = ({ tabs, activa, set }) => (
  <div className="flex gap-1 overflow-x-auto pb-1 -mx-1 px-1">
    {tabs.map((t) => (
      <button key={t.id} onClick={() => set(t.id)}
        className={`px-3 py-1.5 rounded-lg text-sm whitespace-nowrap font-medium border transition-colors ${activa === t.id ? "bg-emerald-700 text-white border-emerald-700" : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"}`}>
        {t.label}
      </button>
    ))}
  </div>
);

const Tabla = ({ cols, children }) => (
  <div className="overflow-x-auto -mx-4 sm:mx-0">
    <table className="min-w-full text-sm">
      <thead>
        <tr className="border-b border-slate-200">
          {cols.map((c, i) => <th key={i} className={`text-left font-medium text-slate-500 text-xs py-2 px-3 whitespace-nowrap ${c.cls || ""}`}>{c.label ?? c}</th>)}
        </tr>
      </thead>
      <tbody className="divide-y divide-slate-100">{children}</tbody>
    </table>
  </div>
);

/* ============================================================
   CÁLCULOS AUTOMÁTICOS (no editables desde la interfaz)
   ============================================================ */
function statsAsistencia(db, alumnoId, { grupoId, trim, desde, hasta } = {}) {
  const cfg = db.config;
  let t = { total: 0, A: 0, F: 0, J: 0, R: 0, P: 0 };
  db.asistencias.forEach((s) => {
    if (grupoId && s.grupoId !== grupoId) return;
    if (desde && s.fecha < desde) return;
    if (hasta && s.fecha > hasta) return;
    if (trim && trimestreDe(s.fecha, cfg) !== trim) return;
    const m = s.marcas?.[alumnoId];
    if (!m) return;
    t.total++; t[m] = (t[m] || 0) + 1;
  });
  t.porcentaje = pct(t.A + t.J + t.P + t.R * 0.5, t.total);
  return t;
}

function statsGrupoAsistencia(db, grupoId, trim) {
  const alumnos = db.alumnos.filter((a) => a.grupoId === grupoId && a.activo);
  const acc = { total: 0, A: 0, F: 0, J: 0, R: 0, P: 0 };
  alumnos.forEach((a) => { const s = statsAsistencia(db, a.id, { grupoId, trim }); ORDEN_EST.forEach((k) => (acc[k] += s[k])); acc.total += s.total; });
  acc.porcentaje = pct(acc.A + acc.J + acc.P + acc.R * 0.5, acc.total);
  acc.sesiones = db.asistencias.filter((s) => s.grupoId === grupoId && (!trim || trimestreDe(s.fecha, db.config) === trim)).length;
  return acc;
}

function calificacionTrimestre(db, alumnoId, trim) {
  const al = db.alumnos.find((a) => a.id === alumnoId);
  if (!al) return { final: 0, detalle: [] };
  const ins = db.catalogos.instrumentos;
  const detalle = ins.map((i) => {
    const acts = db.actividades.filter((a) => a.grupoId === al.grupoId && a.instrumentoId === i.id && Number(a.trimestre) === Number(trim));
    const notas = [];
    acts.forEach((a) => {
      const e = db.entregas.find((x) => x.actividadId === a.id && x.alumnoId === alumnoId);
      if (e && e.estado !== "Pendiente" && e.estado !== "En proceso") notas.push(e.estado === "No entregada" ? 0 : Number(e.calificacion) || 0);
    });
    const prom = notas.length ? round(notas.reduce((s, n) => s + n, 0) / notas.length, 1) : null;
    return { id: i.id, nombre: i.nombre, porcentaje: i.porcentaje, actividades: acts.length, evaluadas: notas.length, promedio: prom };
  });
  const conDatos = detalle.filter((d) => d.promedio !== null);
  const pesoTotal = conDatos.reduce((s, d) => s + d.porcentaje, 0);
  const final = pesoTotal > 0 ? round(conDatos.reduce((s, d) => s + d.promedio * d.porcentaje, 0) / pesoTotal, 1) : 0;
  return { final, detalle };
}

function promedioAcumulado(db, alumnoId) {
  const cs = [1, 2, 3].map((t) => calificacionTrimestre(db, alumnoId, t).final).filter((x) => x > 0);
  return cs.length ? round(cs.reduce((s, n) => s + n, 0) / cs.length, 1) : 0;
}

function statsActividades(db, alumnoId) {
  const al = db.alumnos.find((a) => a.id === alumnoId);
  if (!al) return { asignadas: 0, entregadas: 0, noEntregadas: 0, pendientes: 0, fueraTiempo: 0, porcentaje: 0 };
  const acts = db.actividades.filter((a) => a.grupoId === al.grupoId);
  let r = { asignadas: acts.length, entregadas: 0, noEntregadas: 0, pendientes: 0, fueraTiempo: 0 };
  acts.forEach((a) => {
    const e = db.entregas.find((x) => x.actividadId === a.id && x.alumnoId === alumnoId);
    const est = e?.estado || "Pendiente";
    if (est === "Entregada" || est === "Fuera de tiempo") { r.entregadas++; if (est === "Fuera de tiempo") r.fueraTiempo++; }
    else if (est === "No entregada") r.noEntregadas++;
    else r.pendientes++;
  });
  r.porcentaje = pct(r.entregadas, r.asignadas);
  return r;
}

function resumenSim(sim, materias) {
  if (!sim) return { aciertos: 0, reactivos: 0, porcentaje: 0, porMateria: [] };
  const porMateria = materias.map((m) => {
    const ac = Number(sim.resultados?.[m.id] || 0);
    return { ...m, aciertos: ac, porcentaje: pct(ac, m.reactivos) };
  });
  const aciertos = porMateria.reduce((s, m) => s + m.aciertos, 0);
  const reactivos = materias.reduce((s, m) => s + Number(m.reactivos || 0), 0);
  return { aciertos, reactivos, porcentaje: pct(aciertos, reactivos), porMateria };
}

function historialEcoems(db, alumnoId) {
  const mats = db.catalogos.materiasEcoems;
  const sims = db.ecoems.filter((s) => s.alumnoId === alumnoId).sort((a, b) => a.numero - b.numero);
  const items = sims.map((s) => ({ sim: s, ...resumenSim(s, mats) }));
  const ini = items[0]?.porcentaje || 0;
  const act = items[items.length - 1]?.porcentaje || 0;
  return { items, inicial: ini, actual: act, diferencia: round(act - ini, 1), avance: ini > 0 ? round(((act - ini) / ini) * 100, 1) : 0 };
}

function alertas(db) {
  const out = [];
  const cfg = db.config.alertas;
  db.alumnos.filter((a) => a.activo).forEach((a) => {
    const s = statsAsistencia(db, a.id);
    if (s.total > 0 && s.F >= cfg.faltasMax) out.push({ tipo: "Asistencia", alumnoId: a.id, texto: `${nomComp(a)} acumula ${s.F} faltas`, nivel: "amarillo" });
    if (s.total >= 5 && s.porcentaje < cfg.asistenciaMin) out.push({ tipo: "Asistencia", alumnoId: a.id, texto: `${nomComp(a)} tiene ${s.porcentaje}% de asistencia`, nivel: "naranja" });
    const p = promedioAcumulado(db, a.id);
    if (p > 0 && p < cfg.promedioMin) out.push({ tipo: "Desempeño", alumnoId: a.id, texto: `${nomComp(a)} requiere apoyo académico (promedio ${p})`, nivel: "amarillo" });
    const ac = statsActividades(db, a.id);
    if (ac.noEntregadas >= cfg.actPendientes) out.push({ tipo: "Actividades", alumnoId: a.id, texto: `${nomComp(a)} tiene ${ac.noEntregadas} actividades sin entregar`, nivel: "amarillo" });
    const h = historialEcoems(db, a.id);
    if (h.items.length && semaforoDe(h.actual, db.config).nivel === "bajo") out.push({ tipo: "ECOEMS", alumnoId: a.id, texto: `${nomComp(a)} obtuvo ${h.actual}% en la última aplicación: conviene reforzar`, nivel: "verde" });
  });
  db.incidencias.filter((i) => i.estado !== "Cerrada").forEach((i) => {
    const a = db.alumnos.find((x) => x.id === i.alumnoId);
    const vencido = i.fechaSeguimiento && i.fechaSeguimiento < hoy();
    out.push({
      tipo: "Incidencias", alumnoId: i.alumnoId,
      texto: vencido ? `Folio ${i.folio} de ${nomComp(a)}: seguimiento programado para el ${fFecha(i.fechaSeguimiento)}` : `Folio ${i.folio} (${i.estado}) de ${nomComp(a)}`,
      nivel: i.gravedad === "Alta prioridad de atención" || i.gravedad === "Grave" ? "rojo" : vencido ? "naranja" : "amarillo",
    });
  });
  db.permisos.forEach((p) => {
    const pend = db.alumnos.filter((a) => a.grupoId === p.grupoId && a.activo && !p.registros?.[a.id]?.entregado).length;
    if (pend > 0) out.push({ tipo: "Permisos", texto: `${p.nombre}: ${pend} permisos sin entregar`, nivel: "verde" });
  });
  return out;
}

/* ---------- eliminación segura ----------
   Nada se destruye: lo eliminado viaja a la papelera y puede regresar. */
async function eliminarConRespaldo({ tipo, descripcion, ciclo, datos }) {
  return aPapelera({ tipo, descripcion, ciclo, datos });
}

/* ---------- exportación ---------- */
function descargar(nombre, contenido, tipo) {
  const blob = new Blob([contenido], { type: tipo });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = nombre; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
function aCSV(filas) {
  return "\uFEFF" + filas.map((f) => f.map((c) => `"${String(c ?? "").replace(/"/g, '""')}"`).join(",")).join("\r\n");
}
function exportarExcel(nombre, filas, hoja = "Datos") {
  const ws = XLSX.utils.aoa_to_sheet(filas);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, hoja);
  XLSX.writeFile(wb, nombre.endsWith(".xlsx") ? nombre : nombre + ".xlsx");
}
/* La generación de PDF vive en pdf.js y descarga el archivo directamente,
   sin abrir el cuadro de impresión del navegador. */
/* ============================================================
   MÓDULO: INICIO (tablero principal)
   ============================================================ */
function Inicio({ db, ir }) {
  const alumnos = db.alumnos.filter((a) => a.activo);
  const gruposAct = db.grupos.filter((g) => g.activo);
  const hoyStr = hoy();
  const sesionesHoy = db.asistencias.filter((s) => s.fecha === hoyStr);
  const marcasHoy = sesionesHoy.flatMap((s) => Object.values(s.marcas || {}));
  const asistHoy = marcasHoy.filter((m) => m === "A").length;
  const faltasHoy = marcasHoy.filter((m) => m === "F").length;
  const promGeneral = useMemo(() => {
    const ps = alumnos.map((a) => promedioAcumulado(db, a.id)).filter((x) => x > 0);
    return ps.length ? round(ps.reduce((s, n) => s + n, 0) / ps.length, 1) : 0;
  }, [db]);
  const actPend = db.actividades.filter((a) => a.fechaEntrega >= hoyStr).length;
  const incAbiertas = db.incidencias.filter((i) => i.estado === "Abierta" || i.estado === "En seguimiento").length;
  const permPend = db.permisos.reduce((s, p) => s + db.alumnos.filter((a) => a.grupoId === p.grupoId && a.activo && !p.registros?.[a.id]?.entregado).length, 0);
  const ultimoSim = useMemo(() => {
    const mats = db.catalogos.materiasEcoems;
    const nums = [...new Set(db.ecoems.map((s) => s.numero))].sort((a, b) => b - a);
    if (!nums.length) return null;
    const sims = db.ecoems.filter((s) => s.numero === nums[0]);
    const rr = sims.map((s) => resumenSim(s, mats));
    const porMateria = mats.map((m) => ({ nombre: m.nombre, promedio: round(sims.reduce((t, s) => t + pct(Number(s.resultados?.[m.id] || 0), m.reactivos), 0) / sims.length, 1) })).sort((a, b) => b.promedio - a.promedio);
    return {
      numero: nums[0], n: sims.length, max: mats.reduce((t, m) => t + Number(m.reactivos || 0), 0),
      aciertos: round(rr.reduce((a, b) => a + b.aciertos, 0) / rr.length, 1),
      promedio: round(rr.reduce((a, b) => a + b.porcentaje, 0) / rr.length, 1),
      mejor: porMateria[0], peor: porMateria[porMateria.length - 1],
    };
  }, [db]);
  const conductaFrecuente = useMemo(() => {
    const c = {};
    db.incidencias.forEach((i) => { const k = i.conducta || i.tipo || "Otra"; c[k] = (c[k] || 0) + 1; });
    const orden = Object.entries(c).sort((a, b) => b[1] - a[1]);
    return orden.length ? { nombre: orden[0][0], n: orden[0][1] } : null;
  }, [db]);
  const al = alertas(db);

  const asistPorGrupo = gruposAct.map((g) => ({ nombre: `${g.grado}° ${g.grupo}`, asistencia: statsGrupoAsistencia(db, g.id).porcentaje }));
  const trimData = [1, 2, 3].map((t) => {
    const ps = alumnos.map((a) => calificacionTrimestre(db, a.id, t).final).filter((x) => x > 0);
    return { nombre: `T${t}`, promedio: ps.length ? round(ps.reduce((s, n) => s + n, 0) / ps.length, 1) : 0 };
  });

  return (
    <div className="space-y-4">
      <Titulo sub={`${db.config.escuela || "Sin escuela configurada"} · Ciclo ${db.ciclo} · Trimestre ${db.config.trimestre}`}>Inicio</Titulo>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <Stat label="Grupos" valor={gruposAct.length} icon={Users} onClick={() => ir("grupos")} />
        <Stat label="Alumnas y alumnos" valor={alumnos.length} icon={GraduationCap} onClick={() => ir("alumnos")} />
        <Stat label="Asistencias de hoy" valor={asistHoy} sub={sesionesHoy.length ? `${sesionesHoy.length} sesión(es) registradas` : "Sin pase de lista hoy"} color="text-emerald-700" icon={CalendarCheck} onClick={() => ir("asistencia")} />
        <Stat label="Faltas de hoy" valor={faltasHoy} color={faltasHoy ? "text-rose-600" : "text-slate-900"} icon={AlertTriangle} />
        <Stat label="Promedio académico" valor={promGeneral || "—"} sub="Acumulado del ciclo" icon={BarChart3} onClick={() => ir("evaluacion")} />
        <Stat label="Actividades por entregar" valor={actPend} icon={BookOpen} onClick={() => ir("actividades")} />
        <Stat label="Incidencias abiertas" valor={incAbiertas} color={incAbiertas ? "text-orange-600" : "text-slate-900"} icon={AlertTriangle} onClick={() => ir("incidencias")} />
        <Stat label="Permisos sin entregar" valor={permPend} icon={FileText} onClick={() => ir("permisos")} />
      </div>

      <Card>
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
          <p className="text-sm font-medium text-slate-700">Convivencia escolar</p>
          <Btn size="sm" tipo="secundario" onClick={() => ir("incidencias")}>Ver incidencias</Btn>
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 text-center">
          <div className="bg-slate-50 border border-slate-200 rounded-lg py-2.5"><p className="text-xl font-semibold text-slate-800">{db.incidencias.length}</p><p className="text-[11px] text-slate-500">Registradas</p></div>
          <div className="bg-rose-50 border border-rose-200 rounded-lg py-2.5"><p className="text-xl font-semibold text-rose-700">{db.incidencias.filter((i) => i.estado === "Abierta").length}</p><p className="text-[11px] text-slate-600">Abiertas</p></div>
          <div className="bg-amber-50 border border-amber-200 rounded-lg py-2.5"><p className="text-xl font-semibold text-amber-700">{db.incidencias.filter((i) => i.estado === "En seguimiento").length}</p><p className="text-[11px] text-slate-600">En seguimiento</p></div>
          <div className="bg-emerald-50 border border-emerald-200 rounded-lg py-2.5"><p className="text-xl font-semibold text-emerald-700">{db.incidencias.filter((i) => i.estado === "Cerrada").length}</p><p className="text-[11px] text-slate-600">Cerradas</p></div>
        </div>
        {conductaFrecuente && <p className="text-xs text-slate-500 mt-3">Conducta más frecuente: <span className="text-slate-800 font-medium">{conductaFrecuente.nombre}</span> ({conductaFrecuente.n} registros)</p>}
      </Card>

      {ultimoSim && (
        <Card>
          <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
            <div>
              <p className="text-sm font-medium text-slate-700">Simulador ECOEMS · aplicación {ultimoSim.numero}</p>
              <p className="text-xs text-slate-400">{ultimoSim.n} registros</p>
            </div>
            <Btn size="sm" tipo="secundario" icon={Target} onClick={() => ir("ecoems")}>Ver ECOEMS</Btn>
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 text-center">
            <div className="bg-slate-50 border border-slate-200 rounded-lg py-2.5"><p className="text-xl font-semibold text-slate-800">{ultimoSim.aciertos}/{ultimoSim.max}</p><p className="text-[11px] text-slate-500">Aciertos promedio</p></div>
            <div className={`rounded-lg py-2.5 border ${semaforoDe(ultimoSim.promedio, db.config).cls}`}><p className="text-xl font-semibold">{ultimoSim.promedio}%</p><p className="text-[11px]">Promedio del grupo</p></div>
            <div className="bg-emerald-50 border border-emerald-200 rounded-lg py-2.5"><p className="text-sm font-semibold text-emerald-800 leading-tight pt-1">{ultimoSim.mejor?.nombre || "—"}</p><p className="text-[11px] text-slate-600">Mejor área ({ultimoSim.mejor?.promedio}%)</p></div>
            <div className="bg-rose-50 border border-rose-200 rounded-lg py-2.5"><p className="text-sm font-semibold text-rose-800 leading-tight pt-1">{ultimoSim.peor?.nombre || "—"}</p><p className="text-[11px] text-slate-600">Mayor reforzamiento ({ultimoSim.peor?.promedio}%)</p></div>
          </div>
        </Card>
      )}

      <div className="grid lg:grid-cols-2 gap-4">
        <Card>
          <p className="text-sm font-medium text-slate-700 mb-3">Asistencia por grupo</p>
          {asistPorGrupo.length ? (
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={asistPorGrupo}><CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={11} /><YAxis domain={[0, 100]} fontSize={11} /><Tooltip formatter={(v) => v + "%"} /><Bar dataKey="asistencia" fill="#047857" radius={[4, 4, 0, 0]} /></BarChart>
            </ResponsiveContainer>
          ) : <Vacio texto="Registra un grupo para ver la gráfica." />}
        </Card>
        <Card>
          <p className="text-sm font-medium text-slate-700 mb-3">Promedio por trimestre</p>
          <ResponsiveContainer width="100%" height={200}>
            <LineChart data={trimData}><CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={11} /><YAxis domain={[0, 10]} fontSize={11} /><Tooltip /><Line type="monotone" dataKey="promedio" stroke="#047857" strokeWidth={2} dot={{ r: 4 }} /></LineChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <Card>
        <div className="flex items-center gap-2 mb-3"><Bell size={16} className="text-slate-400" /><p className="text-sm font-medium text-slate-700">Avisos de seguimiento ({al.length})</p></div>
        {al.length === 0 ? <Vacio texto="No hay avisos por ahora." /> : (
          <div className="space-y-1.5 max-h-72 overflow-y-auto">
            {al.slice(0, 40).map((a, i) => (
              <div key={i} className={`flex items-center gap-2 text-sm px-3 py-2 rounded-lg border ${NIVELES[a.nivel].cls} ${a.alumnoId ? "cursor-pointer" : ""}`} onClick={() => a.alumnoId && ir("ficha", { alumnoId: a.alumnoId })}>
                <span className={`w-2 h-2 rounded-full shrink-0 ${NIVELES[a.nivel].punto}`} />
                <span className="text-[11px] font-medium opacity-70 w-20 shrink-0">{a.tipo}</span>
                <span className="grow">{a.texto}</span>
              </div>
            ))}
          </div>
        )}
        <p className="text-[11px] text-slate-400 mt-3">Los avisos son informativos y sirven para organizar el acompañamiento. No son etiquetas sobre las y los estudiantes.</p>
      </Card>
    </div>
  );
}

/* ============================================================
   MÓDULO: GRUPOS
   ============================================================ */
function Grupos({ db, upd, ir, toast }) {
  const [modal, setModal] = useState(null);
  const [conf, setConf] = useState(null);
  const vacio = { grado: 1, grupo: "A", turno: db.config.turno || "Matutino", asignatura: db.config.asignatura || "", docente: db.config.docente || "", activo: true };

  const guardar = () => {
    if (!modal.grupo) return toast("Escribe la letra del grupo", "error");
    const dup = db.grupos.find((g) => g.id !== modal.id && g.grado === Number(modal.grado) && g.grupo === modal.grupo && g.turno === modal.turno);
    if (dup) return toast("Ese grupo ya existe en este ciclo", "error");
    upd((d) => {
      if (modal.id) { const i = d.grupos.findIndex((g) => g.id === modal.id); d.grupos[i] = { ...modal, grado: Number(modal.grado) }; }
      else d.grupos.push({ ...modal, id: uid("grp"), grado: Number(modal.grado) });
    });
    setModal(null); toast("Grupo guardado");
  };

  const eliminar = (g) => {
    const n = db.alumnos.filter((a) => a.grupoId === g.id).length;
    setConf({
      texto: n ? `El grupo ${g.grado}° ${g.grupo} tiene ${n} estudiantes registrados. Al quitarlo, los alumnos se conservan pero quedan sin grupo asignado.` : `¿Enviar el grupo ${g.grado}° ${g.grupo} a la papelera?`,
      onSi: () => {
        eliminarConRespaldo({ tipo: "Grupo", descripcion: `${g.grado}° ${g.grupo}`, ciclo: db.ciclo, datos: { grupos: [g] } });
        upd((d) => { d.grupos = d.grupos.filter((x) => x.id !== g.id); });
        setConf(null); toast("Grupo enviado a la papelera");
      },
    });
  };

  return (
    <div className="space-y-4">
      <Titulo sub="Cada grupo agrupa alumnas, alumnos, asistencia, actividades y seguimiento." right={<Btn icon={Plus} onClick={() => setModal(vacio)}>Nuevo grupo</Btn>}>Grupos</Titulo>
      {db.grupos.length === 0 ? (
        <Card><Vacio texto="Todavía no hay grupos. Empieza creando uno." accion={<Btn icon={Plus} onClick={() => setModal(vacio)}>Crear el primer grupo</Btn>} /></Card>
      ) : (
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">
          {db.grupos.map((g) => {
            const n = db.alumnos.filter((a) => a.grupoId === g.id && a.activo).length;
            const s = statsGrupoAsistencia(db, g.id);
            return (
              <Card key={g.id} className={!g.activo ? "opacity-60" : ""}>
                <div className="flex items-start justify-between">
                  <div>
                    <p className="text-2xl font-semibold text-slate-900">{g.grado}° {g.grupo}</p>
                    <p className="text-xs text-slate-500">{g.turno} · {g.asignatura || "Sin asignatura"}</p>
                  </div>
                  <div className="flex gap-1">
                    <button onClick={() => setModal(g)} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-500"><Pencil size={15} /></button>
                    <button onClick={() => eliminar(g)} className="p-1.5 rounded-lg hover:bg-rose-50 text-rose-500"><Trash2 size={15} /></button>
                  </div>
                </div>
                <div className="grid grid-cols-3 gap-2 mt-3 text-center">
                  <div className="bg-slate-50 rounded-lg py-1.5"><p className="text-base font-semibold text-slate-800">{n}</p><p className="text-[10px] text-slate-500">Alumnos</p></div>
                  <div className="bg-slate-50 rounded-lg py-1.5"><p className="text-base font-semibold text-emerald-700">{s.porcentaje}%</p><p className="text-[10px] text-slate-500">Asistencia</p></div>
                  <div className="bg-slate-50 rounded-lg py-1.5"><p className="text-base font-semibold text-slate-800">{s.sesiones}</p><p className="text-[10px] text-slate-500">Sesiones</p></div>
                </div>
                <div className="flex gap-2 mt-3">
                  <Btn size="sm" tipo="secundario" className="grow" onClick={() => ir("alumnos", { grupoId: g.id })}>Ver alumnos</Btn>
                  <Btn size="sm" tipo="secundario" className="grow" onClick={() => ir("tableroGrupo", { grupoId: g.id })}>Tablero</Btn>
                </div>
                {!g.activo && <p className="text-[11px] text-slate-400 mt-2">Grupo inactivo</p>}
              </Card>
            );
          })}
        </div>
      )}

      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.id ? "Editar grupo" : "Nuevo grupo"} ancho="max-w-lg"
        footer={<><Btn tipo="secundario" onClick={() => setModal(null)}>Cancelar</Btn><Btn icon={Save} onClick={guardar}>Guardar grupo</Btn></>}>
        {modal && (
          <div className="grid grid-cols-2 gap-3">
            <Campo label="Grado" req><Sel value={modal.grado} onChange={(e) => setModal({ ...modal, grado: e.target.value })}>{[1, 2, 3].map((g) => <option key={g} value={g}>{g}°</option>)}</Sel></Campo>
            <Campo label="Grupo" req><Inp value={modal.grupo} maxLength={2} onChange={(e) => setModal({ ...modal, grupo: e.target.value.toUpperCase() })} placeholder="A" /></Campo>
            <Campo label="Turno"><Sel value={modal.turno} onChange={(e) => setModal({ ...modal, turno: e.target.value })}><option>Matutino</option><option>Vespertino</option></Sel></Campo>
            <Campo label="Asignatura"><Inp value={modal.asignatura} onChange={(e) => setModal({ ...modal, asignatura: e.target.value })} /></Campo>
            <Campo label="Docente"><Inp value={modal.docente} onChange={(e) => setModal({ ...modal, docente: e.target.value })} /></Campo>
            <Campo label="Estado"><Sel value={modal.activo ? "1" : "0"} onChange={(e) => setModal({ ...modal, activo: e.target.value === "1" })}><option value="1">Activo</option><option value="0">Inactivo</option></Sel></Campo>
          </div>
        )}
      </Modal>
      <Confirmar open={!!conf} texto={conf?.texto} onSi={conf?.onSi} onNo={() => setConf(null)} />
    </div>
  );
}

/* ============================================================
   MÓDULO: ALUMNAS Y ALUMNOS
   ============================================================ */
function Alumnos({ db, upd, ir, toast, params }) {
  const [grupoId, setGrupoId] = useState(params?.grupoId || db.grupos[0]?.id || "");
  const [busca, setBusca] = useState("");
  const [modal, setModal] = useState(null);
  const [conf, setConf] = useState(null);
  const [imp, setImp] = useState(null);
  const fileRef = useRef();

  const [fEstatus, setFEstatus] = useState("");
  const lista = useMemo(() => sortAl(db.alumnos.filter((a) =>
    (!grupoId || a.grupoId === grupoId) &&
    (!fEstatus || (a.estatus || (a.activo ? "Activo" : "Baja")) === fEstatus) &&
    (!busca || norm(`${nomComp(a)} ${a.curp || ""} ${a.numLista}`).includes(norm(busca))))), [db, grupoId, busca, fEstatus]);
  const grupo = db.grupos.find((g) => g.id === grupoId);

  const guardar = () => {
    if (!modal.apellidos || !modal.nombre) return toast("Escribe nombre y apellidos", "error");
    const dup = db.alumnos.find((a) => a.id !== modal.id && a.grupoId === modal.grupoId && norm(nomComp(a)) === norm(`${modal.apellidos} ${modal.nombre}`));
    if (dup) return toast("Ya existe un alumno con ese nombre en el grupo", "error");
    if (modal.curp && modal.curp.trim().length !== 18) return toast("La CURP debe tener 18 caracteres", "error");
    const dupCurp = modal.curp && db.alumnos.find((a) => a.id !== modal.id && norm(a.curp || "") === norm(modal.curp));
    if (dupCurp) return toast(`Esa CURP ya está registrada en ${nomComp(dupCurp)}`, "error");
    upd((d) => {
      const base = { ...modal, numLista: Number(modal.numLista) || 0, curp: (modal.curp || "").toUpperCase().trim(), activo: (modal.estatus || "Activo") === "Activo" };
      if (modal.id) { const i = d.alumnos.findIndex((a) => a.id === modal.id); d.alumnos[i] = base; }
      else d.alumnos.push({ ...base, id: uid("alu"), numLista: Number(modal.numLista) || d.alumnos.filter((a) => a.grupoId === modal.grupoId).length + 1 });
    });
    setModal(null); toast("Alumno guardado");
  };

  const baja = (a) => setConf({
    texto: `¿Dar de baja a ${nomComp(a)}? Se conservan todos sus registros históricos y dejará de aparecer en las listas activas.`,
    textoSi: "Dar de baja",
    onSi: () => { upd((d) => { const i = d.alumnos.findIndex((x) => x.id === a.id); d.alumnos[i].activo = false; }); setConf(null); toast("Alumno dado de baja"); },
  });

  const leerArchivo = (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    const r = new FileReader();
    r.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const filas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
        const datos = filas.slice(1).filter((f) => f && (f[1] || f[2])).map((f) => ({
          numLista: Number(f[0]) || 0, apellidos: String(f[1] || "").trim(), nombre: String(f[2] || "").trim(),
          grupoTexto: String(f[3] || "").trim(), curp: String(f[4] || "").trim().toUpperCase(),
        }));
        const conDup = datos.map((d) => ({
          ...d,
          dup: !!db.alumnos.find((a) => (a.grupoId === grupoId && norm(nomComp(a)) === norm(`${d.apellidos} ${d.nombre}`)) || (d.curp && norm(a.curp || "") === norm(d.curp))),
        }));
        setImp({ datos: conDup, incluir: conDup.map((d) => !d.dup) });
      } catch (err) { toast("No se pudo leer el archivo", "error"); }
      e.target.value = "";
    };
    r.readAsBinaryString(f);
  };

  const confirmarImport = () => {
    const nuevos = imp.datos.filter((_, i) => imp.incluir[i]);
    upd((d) => { nuevos.forEach((n, i) => d.alumnos.push({
      id: uid("alu"), numLista: n.numLista || d.alumnos.filter((a) => a.grupoId === grupoId).length + i + 1,
      apellidos: n.apellidos, nombre: n.nombre, grupoId, curp: n.curp || "", turno: d.config.turno || "Matutino",
      tutorGrupo: d.config.docente || "", tutorNombre: "", tutorTelefono: "", tutorCorreo: "",
      estatus: "Activo", activo: true, observaciones: "" })); });
    setImp(null); toast(`${nuevos.length} estudiantes importados`);
  };

  const exportar = () => exportarExcel(`alumnos_${grupo ? grupo.grado + grupo.grupo : "todos"}`,
    [["No.", "Apellidos", "Nombre", "CURP", "Grupo", "Turno", "Tutor de grupo", "Contacto tutor", "Teléfono", "Estatus"],
     ...lista.map((a) => { const g = db.grupos.find((x) => x.id === a.grupoId);
       return [a.numLista, a.apellidos, a.nombre, a.curp || "", g ? `${g.grado}° ${g.grupo}` : "", a.turno || "", a.tutorGrupo || "", a.tutorNombre || "", a.tutorTelefono || "", a.estatus || (a.activo ? "Activo" : "Baja")]; })]);

  const pdfPadron = () => {
    pdfTabla({
      titulo: "Padrón de alumnas y alumnos",
      subtitulo: [grupo ? `${grupo.grado}° ${grupo.grupo}` : "Todos los grupos", fEstatus || "Todos los estatus"].join("   ·   "),
      columnas: ["No.", "Nombre", "CURP", "Grupo", "Estatus", "Asistencia", "Promedio"],
      filas: lista.map((a) => { const g2 = db.grupos.find((x) => x.id === a.grupoId); const st = statsAsistencia(db, a.id);
        return [a.numLista, nomComp(a), a.curp || "—", g2 ? `${g2.grado}° ${g2.grupo}` : "—", a.estatus || (a.activo ? "Activo" : "Baja"), st.total ? st.porcentaje + "%" : "—", promedioAcumulado(db, a.id) || "—"]; }),
      config: db.config, ciclo: db.ciclo,
      resumen: [{ label: "Estudiantes", valor: lista.length }, { label: "Activos", valor: lista.filter((a) => (a.estatus || (a.activo ? "Activo" : "Baja")) === "Activo").length }],
      nota: "Documento con datos de menores de edad. Su manejo es confidencial y de uso exclusivamente escolar.",
    });
    toast("PDF descargado");
  };

  return (
    <div className="space-y-4">
      <Titulo sub="Cada estudiante se registra una sola vez y su información se usa en todos los módulos."
        right={<div className="flex gap-2 flex-wrap">
          <Btn tipo="secundario" size="sm" icon={Upload} onClick={() => fileRef.current?.click()} disabled={!grupoId}>Importar</Btn>
          <Btn tipo="secundario" size="sm" icon={FileDown} onClick={pdfPadron}>PDF</Btn>
          <Btn tipo="secundario" size="sm" icon={Download} onClick={exportar}>Excel</Btn>
          <Btn size="sm" icon={Plus} onClick={() => setModal({ numLista: "", apellidos: "", nombre: "", grupoId, curp: "", turno: db.config.turno || "Matutino", tutorGrupo: db.config.docente || "", tutorNombre: "", tutorTelefono: "", tutorCorreo: "", estatus: "Activo", activo: true, observaciones: "" })} disabled={!db.grupos.length}>Nuevo</Btn>
        </div>}>Alumnas y alumnos</Titulo>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={leerArchivo} className="hidden" />

      <div className="flex gap-2 flex-wrap">
        <Sel value={grupoId} onChange={(e) => setGrupoId(e.target.value)} className="max-w-[190px]">
          <option value="">Todos los grupos</option>
          {db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}
        </Sel>
        <Sel value={fEstatus} onChange={(e) => setFEstatus(e.target.value)} className="max-w-[150px]">
          <option value="">Todos los estatus</option>{EST_ALUMNO.map((s2) => <option key={s2}>{s2}</option>)}
        </Sel>
        <div className="relative grow max-w-sm">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={busca} onChange={(e) => setBusca(e.target.value)} placeholder="Buscar por nombre, CURP o número de lista" className={inputCls + " pl-9"} />
        </div>
      </div>

      <Card pad={false} className="p-4">
        {lista.length === 0 ? <Vacio texto="No hay estudiantes con estos filtros." /> : (
          <Tabla cols={["No.", "Nombre", "Grupo", "Estatus", "Asistencia", "Promedio", { label: "", cls: "text-right" }]}>
            {lista.map((a) => {
              const g = db.grupos.find((x) => x.id === a.grupoId);
              const s = statsAsistencia(db, a.id);
              const p = promedioAcumulado(db, a.id);
              return (
                <tr key={a.id} className={`hover:bg-slate-50 ${!a.activo ? "opacity-50" : ""}`}>
                  <td className="py-2 px-3 text-slate-500">{a.numLista}</td>
                  <td className="py-2 px-3">
                    <button onClick={() => ir("ficha", { alumnoId: a.id })} className="font-medium text-slate-800 hover:text-emerald-700 text-left">{nomComp(a)}</button>
                    {a.curp && <span className="block text-[10px] text-slate-400 font-mono">{a.curp}</span>}
                  </td>
                  <td className="py-2 px-3 text-slate-600">{g ? `${g.grado}° ${g.grupo}` : "—"}</td>
                  <td className="py-2 px-3"><Pill cls={(a.estatus || (a.activo ? "Activo" : "Baja")) === "Activo" ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-slate-100 text-slate-600 border-slate-200"}>{a.estatus || (a.activo ? "Activo" : "Baja")}</Pill></td>
                  <td className="py-2 px-3"><span className={s.porcentaje >= 80 ? "text-emerald-700" : "text-amber-600"}>{s.total ? s.porcentaje + "%" : "—"}</span></td>
                  <td className="py-2 px-3">{p || "—"}</td>
                  <td className="py-2 px-3 text-right whitespace-nowrap">
                    <button onClick={() => setModal(a)} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-500"><Pencil size={15} /></button>
                    {a.activo && <button onClick={() => baja(a)} className="p-1.5 rounded-lg hover:bg-rose-50 text-rose-500"><Trash2 size={15} /></button>}
                  </td>
                </tr>
              );
            })}
          </Tabla>
        )}
      </Card>

      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.id ? "Editar estudiante" : "Nuevo estudiante"} ancho="max-w-lg"
        footer={<><Btn tipo="secundario" onClick={() => setModal(null)}>Cancelar</Btn><Btn icon={Save} onClick={guardar}>Guardar</Btn></>}>
        {modal && (
          <div className="space-y-3">
            <div className="grid grid-cols-3 gap-3">
              <Campo label="No. de lista"><Inp type="number" value={modal.numLista} onChange={(e) => setModal({ ...modal, numLista: e.target.value })} /></Campo>
              <Campo label="Grupo"><Sel value={modal.grupoId} onChange={(e) => setModal({ ...modal, grupoId: e.target.value })}>{db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}</Sel></Campo>
              <Campo label="Estatus"><Sel value={modal.estatus || "Activo"} onChange={(e) => setModal({ ...modal, estatus: e.target.value })}>{EST_ALUMNO.map((s2) => <option key={s2}>{s2}</option>)}</Sel></Campo>
            </div>
            <Campo label="Apellidos" req><Inp value={modal.apellidos} onChange={(e) => setModal({ ...modal, apellidos: e.target.value })} /></Campo>
            <Campo label="Nombre" req><Inp value={modal.nombre} onChange={(e) => setModal({ ...modal, nombre: e.target.value })} /></Campo>
            <div className="grid sm:grid-cols-2 gap-3">
              <Campo label="CURP" hint="18 caracteres. El sistema avisa si ya está registrada."><Inp value={modal.curp || ""} maxLength={18} onChange={(e) => setModal({ ...modal, curp: e.target.value.toUpperCase() })} className="font-mono" /></Campo>
              <Campo label="Turno"><Sel value={modal.turno || "Matutino"} onChange={(e) => setModal({ ...modal, turno: e.target.value })}><option>Matutino</option><option>Vespertino</option></Sel></Campo>
            </div>
            <Campo label="Tutor(a) de grupo"><Inp value={modal.tutorGrupo || ""} onChange={(e) => setModal({ ...modal, tutorGrupo: e.target.value })} /></Campo>
            <div className="border-t border-slate-200 pt-3">
              <p className="text-xs font-medium text-slate-600 mb-2">Contacto de madre, padre o tutor</p>
              <div className="grid sm:grid-cols-3 gap-3">
                <Campo label="Nombre"><Inp value={modal.tutorNombre || ""} onChange={(e) => setModal({ ...modal, tutorNombre: e.target.value })} /></Campo>
                <Campo label="Teléfono"><Inp value={modal.tutorTelefono || ""} onChange={(e) => setModal({ ...modal, tutorTelefono: e.target.value })} /></Campo>
                <Campo label="Correo"><Inp value={modal.tutorCorreo || ""} onChange={(e) => setModal({ ...modal, tutorCorreo: e.target.value })} /></Campo>
              </div>
            </div>
            <Campo label="Observaciones" hint="Notas escolares breves. Evita registrar datos personales que no sean necesarios."><Area value={modal.observaciones} onChange={(e) => setModal({ ...modal, observaciones: e.target.value })} /></Campo>
          </div>
        )}
      </Modal>

      <Modal open={!!imp} onClose={() => setImp(null)} title="Revisar importación" ancho="max-w-2xl"
        footer={<><Btn tipo="secundario" onClick={() => setImp(null)}>Cancelar</Btn><Btn icon={Check} onClick={confirmarImport}>Importar seleccionados</Btn></>}>
        {imp && (
          <div className="space-y-3">
            <Aviso>Formato esperado: columna A No., B Apellidos, C Nombre, D Grupo y, si la tienes, E CURP. Los posibles duplicados aparecen desmarcados.</Aviso>
            <Tabla cols={["", "No.", "Apellidos", "Nombre", "CURP", "Aviso"]}>
              {imp.datos.map((d, i) => (
                <tr key={i} className={d.dup ? "bg-amber-50" : ""}>
                  <td className="py-1.5 px-3"><input type="checkbox" checked={imp.incluir[i]} onChange={() => { const c = [...imp.incluir]; c[i] = !c[i]; setImp({ ...imp, incluir: c }); }} /></td>
                  <td className="py-1.5 px-3">{d.numLista}</td><td className="py-1.5 px-3">{d.apellidos}</td><td className="py-1.5 px-3">{d.nombre}</td>
                  <td className="py-1.5 px-3 font-mono text-[11px]">{d.curp}</td>
                  <td className="py-1.5 px-3 text-xs text-amber-700">{d.dup ? "Ya existe" : ""}</td>
                </tr>
              ))}
            </Tabla>
          </div>
        )}
      </Modal>
      <Confirmar open={!!conf} texto={conf?.texto} textoSi={conf?.textoSi} onSi={conf?.onSi} onNo={() => setConf(null)} />
    </div>
  );
}

/* ============================================================
   FICHA INTEGRAL DEL ALUMNO
   ============================================================ */
function Ficha({ db, ir, params, toast }) {
  const a = db.alumnos.find((x) => x.id === params?.alumnoId);
  const [tab, setTab] = useState("academico");
  if (!a) return <Card><Vacio texto="Selecciona un estudiante desde el padrón." accion={<Btn onClick={() => ir("alumnos")}>Ir al padrón</Btn>} /></Card>;
  const g = db.grupos.find((x) => x.id === a.grupoId);
  const asis = statsAsistencia(db, a.id);
  const act = statsActividades(db, a.id);
  const eco = historialEcoems(db, a.id);
  const trims = [1, 2, 3].map((t) => ({ t, ...calificacionTrimestre(db, a.id, t) }));
  const acum = promedioAcumulado(db, a.id);
  const incs = db.incidencias.filter((i) => i.alumnoId === a.id);
  const perms = db.permisos.filter((p) => p.grupoId === a.grupoId);
  const tipoIncFrecuente = (() => { const c = {}; incs.forEach((i) => { const k = i.conducta || i.tipo || "Otra"; c[k] = (c[k] || 0) + 1; });
    const o = Object.entries(c).sort((x, y) => y[1] - x[1]); return o.length ? o[0][0] : null; })();

  const progreso = [
    ...trims.filter((t) => t.final > 0).map((t) => ({ nombre: `T${t.t}`, valor: t.final * 10 })),
    ...eco.items.map((s) => ({ nombre: `ECOEMS ${s.sim.numero}`, valor: s.porcentaje })),
  ];

  const descargarFicha = () => {
    pdfAlumno({
      config: db.config, ciclo: db.ciclo,
      grupo: g ? `${g.grado}° ${g.grupo}` : "Sin grupo",
      alumno: {
        nombre: nomComp(a),
        archivo: nomComp(a).split(" ")[0],
        datos: [
          ["Nombre", nomComp(a), "Grupo", g ? `${g.grado}° ${g.grupo}` : "—"],
          ["CURP", a.curp || "—", "Turno", a.turno || g?.turno || "—"],
          ["No. de lista", String(a.numLista), "Estatus", a.estatus || (a.activo ? "Activo" : "Baja")],
          ["Tutor(a) de grupo", a.tutorGrupo || "—", "Contacto familiar", [a.tutorNombre, a.tutorTelefono].filter(Boolean).join(" · ") || "—"],
        ],
      },
      secciones: [
        { titulo: "Asistencia",
          columnas: ["Clases", "Asistencias", "Faltas", "Justificadas", "Retardos", "Permisos", "Porcentaje"],
          filas: [[asis.total, asis.A, asis.F, asis.J, asis.R, asis.P, asis.porcentaje + "%"]] },
        { titulo: "Evaluación",
          columnas: ["Trimestre 1", "Trimestre 2", "Trimestre 3", "Promedio acumulado"],
          filas: [[trims[0].final || "—", trims[1].final || "—", trims[2].final || "—", acum || "—"]] },
        { titulo: "Actividades",
          columnas: ["Asignadas", "Entregadas", "Fuera de tiempo", "No entregadas", "Pendientes", "% entrega"],
          filas: [[act.asignadas, act.entregadas, act.fueraTiempo, act.noEntregadas, act.pendientes, act.porcentaje + "%"]] },
        { titulo: "Simulador ECOEMS",
          columnas: ["Aplicación", "Fecha", "Aciertos", "Máximo", "Porcentaje", "Semáforo"],
          filas: eco.items.map((x) => [x.sim.numero, fFecha(x.sim.fecha), x.aciertos, x.reactivos, x.porcentaje + "%", semaforoDe(x.porcentaje, db.config).label]) },
        { titulo: "Incidencias y seguimiento",
          columnas: ["Folio", "Fecha", "Conducta", "Clasificación", "Estatus", "Seguimientos"],
          filas: incs.map((i) => [i.folio, fFecha(i.fecha), i.conducta || i.tipo, i.gravedad || "—", i.estado, (i.seguimientos || []).length]) },
        { titulo: "Permisos",
          columnas: ["Actividad", "Fecha", "Entregado", "Autorizado"],
          filas: perms.map((pp) => { const r = pp.registros?.[a.id] || {}; return [pp.nombre, fFecha(pp.fecha), r.entregado ? "Sí" : "No", r.autorizado ? "Sí" : "No"]; }) },
      ],
      nota: "Este reporte es un registro administrativo de seguimiento docente. La información es confidencial y su uso se limita a fines educativos.",
    });
  };

  return (
    <div className="space-y-4">
      <button onClick={() => ir("alumnos")} className="inline-flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><ArrowLeft size={15} />Volver al padrón</button>
      <Card>
        <div className="flex items-start justify-between gap-3 flex-wrap">
          <div>
            <h2 className="text-xl font-semibold text-slate-900">{nomComp(a)}</h2>
            <p className="text-sm text-slate-500">{g ? `${g.grado}° ${g.grupo} · ${a.turno || g.turno}` : "Sin grupo"} · No. {a.numLista} · {g?.asignatura || db.config.asignatura || "—"}</p>
            <div className="flex flex-wrap gap-2 mt-2">
              <Pill cls={(a.estatus || (a.activo ? "Activo" : "Baja")) === "Activo" ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-slate-100 text-slate-600 border-slate-200"}>{a.estatus || (a.activo ? "Activo" : "Baja")}</Pill>
              {a.curp && <Pill cls="bg-slate-100 text-slate-600 border-slate-200">CURP {a.curp}</Pill>}
              {a.tutorGrupo && <Pill cls="bg-slate-100 text-slate-600 border-slate-200">Tutor(a) de grupo: {a.tutorGrupo}</Pill>}
            </div>
            {(a.tutorNombre || a.tutorTelefono || a.tutorCorreo) && (
              <p className="text-xs text-slate-500 mt-2">Contacto: {[a.tutorNombre, a.tutorTelefono, a.tutorCorreo].filter(Boolean).join(" · ")}</p>
            )}
            {a.observaciones && <p className="text-xs text-slate-500 mt-2 bg-slate-50 border border-slate-200 rounded-lg p-2">{a.observaciones}</p>}
          </div>
          <Btn tipo="secundario" size="sm" icon={FileDown} onClick={descargarFicha}>Descargar PDF</Btn>
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mt-4">
          <div className="bg-emerald-50 border border-emerald-100 rounded-lg p-2.5"><p className="text-[11px] text-emerald-800">Asistencia</p><p className="text-xl font-semibold text-emerald-800">{asis.total ? asis.porcentaje + "%" : "—"}</p></div>
          <div className="bg-slate-50 border border-slate-200 rounded-lg p-2.5"><p className="text-[11px] text-slate-600">Promedio</p><p className="text-xl font-semibold text-slate-800">{acum || "—"}</p></div>
          <div className="bg-slate-50 border border-slate-200 rounded-lg p-2.5"><p className="text-[11px] text-slate-600">Entregas</p><p className="text-xl font-semibold text-slate-800">{act.entregadas}/{act.asignadas}</p></div>
          <div className="bg-sky-50 border border-sky-100 rounded-lg p-2.5"><p className="text-[11px] text-sky-800">ECOEMS</p><p className="text-xl font-semibold text-sky-800">{eco.items.length ? eco.actual + "%" : "—"}</p></div>
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mt-2">
          <div className="bg-slate-50 border border-slate-200 rounded-lg p-2.5"><p className="text-[11px] text-slate-600">Actividades pendientes</p><p className="text-xl font-semibold text-slate-800">{act.pendientes}</p></div>
          <div className="bg-slate-50 border border-slate-200 rounded-lg p-2.5"><p className="text-[11px] text-slate-600">No entregadas</p><p className="text-xl font-semibold text-slate-800">{act.noEntregadas}</p></div>
          <div className="bg-slate-50 border border-slate-200 rounded-lg p-2.5"><p className="text-[11px] text-slate-600">Incidencias</p><p className="text-xl font-semibold text-slate-800">{incs.length}</p></div>
          <div className="bg-slate-50 border border-slate-200 rounded-lg p-2.5"><p className="text-[11px] text-slate-600">Tipo más frecuente</p><p className="text-xs font-medium text-slate-800 pt-1.5 leading-tight">{tipoIncFrecuente || "—"}</p></div>
        </div>
      </Card>

      {progreso.length > 1 && (
        <Card>
          <p className="text-sm font-medium text-slate-700 mb-3">Línea de progreso</p>
          <ResponsiveContainer width="100%" height={190}>
            <LineChart data={progreso}><CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={11} /><YAxis domain={[0, 100]} fontSize={11} /><Tooltip formatter={(v) => v + " / 100"} /><Line type="monotone" dataKey="valor" stroke="#0369a1" strokeWidth={2} dot={{ r: 4 }} /></LineChart>
          </ResponsiveContainer>
          <p className="text-[11px] text-slate-400 mt-1">Calificaciones y resultados ECOEMS en escala de 100 para poder compararlos.</p>
        </Card>
      )}

      <Tabs activa={tab} set={setTab} tabs={[
        { id: "academico", label: "Académico" }, { id: "asistencia", label: "Asistencia" }, { id: "actividades", label: "Actividades" },
        { id: "incidencias", label: "Incidencias" }, { id: "permisos", label: "Permisos" }, { id: "ecoems", label: "ECOEMS" },
      ]} />

      {tab === "academico" && (
        <Card>
          <Tabla cols={["Trimestre", "Instrumento", "Porcentaje", "Actividades", "Promedio"]}>
            {trims.flatMap((t) => t.detalle.map((d, i) => (
              <tr key={t.t + d.id}>
                <td className="py-2 px-3 text-slate-500">{i === 0 ? `T${t.t}` : ""}</td>
                <td className="py-2 px-3">{d.nombre}</td><td className="py-2 px-3">{d.porcentaje}%</td>
                <td className="py-2 px-3">{d.evaluadas}/{d.actividades}</td>
                <td className="py-2 px-3 font-medium">{d.promedio ?? "—"}</td>
              </tr>
            )).concat([
              <tr key={"tot" + t.t} className="bg-slate-50"><td className="py-2 px-3" /><td className="py-2 px-3 font-medium" colSpan={3}>Calificación del trimestre {t.t}</td><td className="py-2 px-3 font-semibold text-emerald-700">{t.final || "—"}</td></tr>
            ]))}
            <tr className="bg-emerald-50"><td className="py-2 px-3" colSpan={4}><span className="font-medium">Promedio acumulado</span></td><td className="py-2 px-3 font-semibold text-emerald-800">{acum || "—"}</td></tr>
          </Tabla>
        </Card>
      )}

      {tab === "asistencia" && (
        <Card>
          <div className="grid grid-cols-3 sm:grid-cols-6 gap-2 mb-4">
            <Stat label="Clases" valor={asis.total} />
            {ORDEN_EST.map((k) => <Stat key={k} label={EST[k].label} valor={asis[k]} />)}
          </div>
          <Tabla cols={["Fecha", "Sesión", "Estado"]}>
            {db.asistencias.filter((s) => s.marcas?.[a.id]).sort((x, y) => y.fecha.localeCompare(x.fecha)).slice(0, 60).map((s) => (
              <tr key={s.id}><td className="py-2 px-3">{fFecha(s.fecha)}</td><td className="py-2 px-3 text-slate-500">{s.sesion}</td>
                <td className="py-2 px-3"><Pill cls={EST[s.marcas[a.id]].soft}>{EST[s.marcas[a.id]].label}</Pill></td></tr>
            ))}
          </Tabla>
        </Card>
      )}

      {tab === "actividades" && (
        <Card>
          <Tabla cols={["Actividad", "Entrega", "Trim.", "Estado", "Calificación"]}>
            {db.actividades.filter((x) => x.grupoId === a.grupoId).map((x) => {
              const e = db.entregas.find((y) => y.actividadId === x.id && y.alumnoId === a.id);
              return <tr key={x.id}><td className="py-2 px-3">{x.nombre}</td><td className="py-2 px-3 text-slate-500">{fFecha(x.fechaEntrega)}</td>
                <td className="py-2 px-3 text-slate-500">T{x.trimestre}</td><td className="py-2 px-3">{e?.estado || "Pendiente"}</td>
                <td className="py-2 px-3 font-medium">{e && e.estado !== "Pendiente" ? e.calificacion : "—"}</td></tr>;
            })}
          </Tabla>
        </Card>
      )}

      {tab === "incidencias" && (
        <Card>
          {incs.length === 0 ? <Vacio texto="Sin incidencias registradas." /> : (
            <Tabla cols={["Folio", "Fecha", "Conducta", "Clasificación", "Estatus", "Seguimientos"]}>
              {incs.map((i) => <tr key={i.id} className="cursor-pointer hover:bg-slate-50" onClick={() => ir("incidencias", { incId: i.id })}>
                <td className="py-2 px-3 font-mono text-xs">{i.folio}</td><td className="py-2 px-3">{fFecha(i.fecha)}</td>
                <td className="py-2 px-3">{i.conducta || i.tipo}</td>
                <td className="py-2 px-3"><Pill cls={(GRAVEDAD[i.gravedad] || GRAVEDAD.Leve).cls}>{i.gravedad || "—"}</Pill></td>
                <td className="py-2 px-3"><Pill>{i.estado}</Pill></td><td className="py-2 px-3">{i.seguimientos?.length || 0}</td></tr>)}
            </Tabla>
          )}
        </Card>
      )}

      {tab === "permisos" && (
        <Card>
          {perms.length === 0 ? <Vacio texto="Sin permisos registrados." /> : (
            <Tabla cols={["Actividad", "Fecha", "Entregado", "Autorizado"]}>
              {perms.map((p) => { const r = p.registros?.[a.id] || {}; return <tr key={p.id}><td className="py-2 px-3">{p.nombre}</td><td className="py-2 px-3">{fFecha(p.fecha)}</td>
                <td className="py-2 px-3">{r.entregado ? `Sí · ${fFecha(r.fechaEntrega)}` : "No"}</td><td className="py-2 px-3">{r.autorizado ? "Sí" : "No"}</td></tr>; })}
            </Tabla>
          )}
        </Card>
      )}

      {tab === "ecoems" && (
        <Card>
          {eco.items.length === 0 ? <Vacio texto="Sin simuladores registrados." /> : (
            <>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-4">
                <Stat label="Resultado inicial" valor={eco.inicial + "%"} /><Stat label="Resultado actual" valor={eco.actual + "%"} />
                <Stat label="Diferencia" valor={(eco.diferencia > 0 ? "+" : "") + eco.diferencia} color={eco.diferencia >= 0 ? "text-emerald-700" : "text-amber-600"} />
                <Stat label="Avance" valor={(eco.avance > 0 ? "+" : "") + eco.avance + "%"} />
              </div>
              <ResponsiveContainer width="100%" height={190}>
                <LineChart data={eco.items.map((s) => ({ nombre: "S" + s.sim.numero, porcentaje: s.porcentaje }))}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={11} /><YAxis domain={[0, 100]} fontSize={11} /><Tooltip formatter={(v) => v + "%"} />
                  <Line type="monotone" dataKey="porcentaje" stroke="#0369a1" strokeWidth={2} dot={{ r: 4 }} />
                </LineChart>
              </ResponsiveContainer>
              <Tabla cols={["Materia", ...eco.items.map((s) => "S" + s.sim.numero)]}>
                {db.catalogos.materiasEcoems.map((m) => (
                  <tr key={m.id}><td className="py-2 px-3">{m.nombre}</td>
                    {eco.items.map((s) => { const pm = s.porMateria.find((x) => x.id === m.id); return <td key={s.sim.id} className="py-2 px-3">{pm.aciertos}/{m.reactivos} <span className="text-slate-400">({pm.porcentaje}%)</span></td>; })}
                  </tr>
                ))}
                <tr className="bg-slate-50 font-medium"><td className="py-2 px-3">Total</td>{eco.items.map((s) => <td key={s.sim.id} className="py-2 px-3">{s.aciertos}/{s.reactivos} ({s.porcentaje}%)</td>)}</tr>
              </Tabla>
            </>
          )}
        </Card>
      )}
    </div>
  );
}

/* ============================================================
   MÓDULO: ASISTENCIA
   ============================================================ */
function Asistencia({ db, upd, toast }) {
  const [grupoId, setGrupoId] = useState(db.grupos[0]?.id || "");
  const [fecha, setFecha] = useState(hoy());
  const [sesion, setSesion] = useState(1);
  const alumnos = useMemo(() => sortAl(db.alumnos.filter((a) => a.grupoId === grupoId && a.activo)), [db, grupoId]);
  const registro = db.asistencias.find((s) => s.grupoId === grupoId && s.fecha === fecha && Number(s.sesion) === Number(sesion));
  const [marcas, setMarcas] = useState({});
  const [notas, setNotas] = useState({});
  const [obsSesion, setObsSesion] = useState("");
  const [sucio, setSucio] = useState(false);
  const [abierto, setAbierto] = useState(null);

  useEffect(() => {
    setMarcas(registro?.marcas ? { ...registro.marcas } : {});
    setNotas(registro?.notas ? { ...registro.notas } : {});
    setObsSesion(registro?.observaciones || "");
    setSucio(false);
  }, [grupoId, fecha, sesion, registro?.id]);

  const marcar = (id, val) => { setMarcas((m) => ({ ...m, [id]: val })); setSucio(true); };
  const todos = (val) => { const m = {}; alumnos.forEach((a) => (m[a.id] = val)); setMarcas(m); setSucio(true); };

  const guardar = () => {
    if (!grupoId) return toast("Selecciona un grupo", "error");
    upd((d) => {
      const i = d.asistencias.findIndex((s) => s.grupoId === grupoId && s.fecha === fecha && Number(s.sesion) === Number(sesion));
      if (i >= 0) { d.asistencias[i].marcas = marcas; d.asistencias[i].notas = notas; d.asistencias[i].observaciones = obsSesion; }
      else d.asistencias.push({ id: uid("asi"), grupoId, fecha, sesion: Number(sesion), marcas, notas, observaciones: obsSesion });
    });
    setSucio(false); toast("Asistencia guardada");
  };

  const conteo = ORDEN_EST.reduce((o, k) => ({ ...o, [k]: Object.values(marcas).filter((v) => v === k).length }), {});
  const sinMarcar = alumnos.length - Object.keys(marcas).filter((k) => alumnos.find((a) => a.id === k)).length;

  return (
    <div className="space-y-4">
      <Titulo sub="Selecciona grupo y fecha, marca a todo el grupo y ajusta solo los casos particulares.">Asistencia</Titulo>
      <Card>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <Campo label="Grupo"><Sel value={grupoId} onChange={(e) => setGrupoId(e.target.value)}>{db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}</Sel></Campo>
          <Campo label="Fecha"><Inp type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Campo>
          <Campo label="Clase / sesión"><Sel value={sesion} onChange={(e) => setSesion(Number(e.target.value))}>{[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => <option key={n} value={n}>Clase {n}</option>)}</Sel></Campo>
          <Campo label="Trimestre"><div className="px-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg text-slate-600">T{trimestreDe(fecha, db.config)}</div></Campo>
        </div>
        {registro && <p className="text-xs text-emerald-700 mt-2">Ya existe un registro para esta fecha y clase. Al guardar se actualiza.</p>}
      </Card>

      {alumnos.length === 0 ? <Card><Vacio texto="Este grupo no tiene estudiantes activos." /></Card> : (
        <>
          <div className="flex gap-2 flex-wrap items-center">
            <Btn size="sm" icon={Check} onClick={() => todos("A")}>Marcar asistencia a todo el grupo</Btn>
            <Btn size="sm" tipo="secundario" onClick={() => setMarcas({})}>Limpiar</Btn>
            {sinMarcar > 0 && <span className="text-xs text-amber-600">{sinMarcar} sin marcar</span>}
          </div>

          <Card pad={false} className="p-2 sm:p-3">
            <div className="divide-y divide-slate-100">
              {alumnos.map((a) => (
                <div key={a.id} className="py-2 px-1">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="w-6 text-xs text-slate-400 shrink-0">{a.numLista}</span>
                    <span className="grow min-w-[120px] text-sm text-slate-800">{nomComp(a)}</span>
                    <div className="flex gap-1">
                      {ORDEN_EST.map((k) => (
                        <button key={k} onClick={() => marcar(a.id, k)} title={EST[k].label}
                          className={`w-9 h-9 rounded-lg border text-xs font-semibold transition-colors ${marcas[a.id] === k ? EST[k].full : "bg-white text-slate-400 border-slate-200 hover:border-slate-400"}`}>
                          {EST[k].corto}
                        </button>
                      ))}
                      <button onClick={() => setAbierto(abierto === a.id ? null : a.id)} title="Observación o justificación"
                        className={`w-9 h-9 rounded-lg border text-xs transition-colors ${notas[a.id] ? "bg-slate-800 text-white border-slate-800" : "bg-white text-slate-300 border-slate-200 hover:border-slate-400"}`}>
                        <NotebookPen size={14} className="mx-auto" />
                      </button>
                    </div>
                  </div>
                  {abierto === a.id && (
                    <input autoFocus value={notas[a.id] || ""} onChange={(e) => { setNotas({ ...notas, [a.id]: e.target.value }); setSucio(true); }}
                      placeholder="Justificación presentada u observación" className={inputCls + " mt-2 text-xs"} />
                  )}
                  {notas[a.id] && abierto !== a.id && <p className="text-[11px] text-slate-500 ml-8 mt-1">{notas[a.id]}</p>}
                </div>
              ))}
            </div>
          </Card>

          <Card>
            <Campo label="Observaciones de la sesión"><Inp value={obsSesion} onChange={(e) => { setObsSesion(e.target.value); setSucio(true); }} placeholder="Suspensión parcial, actividad especial, guardia…" /></Campo>
          </Card>

          <div className="sticky bottom-16 lg:bottom-4 z-20">
            <Card className="flex items-center justify-between gap-3 flex-wrap shadow-lg">
              <div className="flex gap-2 flex-wrap text-xs">
                {ORDEN_EST.map((k) => <Pill key={k} cls={EST[k].soft}>{EST[k].label}: {conteo[k]}</Pill>)}
              </div>
              <Btn icon={Save} onClick={guardar} disabled={!sucio && !!registro}>Guardar asistencia</Btn>
            </Card>
          </div>
        </>
      )}
    </div>
  );
}

/* ============================================================
   MÓDULO: PERMISOS Y AUTORIZACIONES
   ============================================================ */
function Permisos({ db, upd, toast }) {
  const [modal, setModal] = useState(null);
  const [abierto, setAbierto] = useState(null);
  const [conf, setConf] = useState(null);
  const vacio = { nombre: "", fecha: hoy(), lugar: "", responsable: db.config.docente || "", grupoId: db.grupos[0]?.id || "", horaSalida: "", horaRegreso: "", descripcion: "", registros: {}, archivo: "" };

  const guardar = () => {
    if (!modal.nombre || !modal.grupoId) return toast("Falta el nombre de la actividad o el grupo", "error");
    upd((d) => {
      if (modal.id) { const i = d.permisos.findIndex((p) => p.id === modal.id); d.permisos[i] = modal; }
      else d.permisos.push({ ...modal, id: uid("per") });
    });
    setModal(null); toast("Permiso guardado");
  };

  const setReg = (permisoId, alumnoId, campo, valor) => upd((d) => {
    const p = d.permisos.find((x) => x.id === permisoId);
    p.registros = p.registros || {};
    p.registros[alumnoId] = { entregado: false, fechaEntrega: "", autorizado: false, obs: "", ...(p.registros[alumnoId] || {}), [campo]: valor };
    if (campo === "entregado" && valor && !p.registros[alumnoId].fechaEntrega) p.registros[alumnoId].fechaEntrega = hoy();
  });

  const detalle = db.permisos.find((p) => p.id === abierto);
  const alumnosDet = detalle ? sortAl(db.alumnos.filter((a) => a.grupoId === detalle.grupoId && a.activo)) : [];

  const descargarListas = (p) => {
    const als = sortAl(db.alumnos.filter((a) => a.grupoId === p.grupoId && a.activo));
    const g = db.grupos.find((x) => x.id === p.grupoId);
    const fila = (a) => { const r = p.registros?.[a.id] || {}; return [a.numLista, nomComp(a), r.entregado ? "Sí" : "No", r.autorizado ? "Sí" : "No", r.obs || ""]; };
    pdfPermisos({
      permiso: p, config: db.config, ciclo: db.ciclo,
      grupo: g ? `${g.grado}° ${g.grupo}` : "Sin grupo",
      autorizados: als.filter((a) => p.registros?.[a.id]?.autorizado).map(fila),
      pendientes: als.filter((a) => !p.registros?.[a.id]?.entregado).map(fila),
    });
  };

  return (
    <div className="space-y-4">
      <Titulo sub="Control de permisos por actividad escolar." right={<Btn icon={Plus} onClick={() => setModal(vacio)} disabled={!db.grupos.length}>Nueva actividad</Btn>}>Permisos</Titulo>
      {db.permisos.length === 0 ? <Card><Vacio texto="Aún no hay actividades con permiso." /></Card> : (
        <div className="grid sm:grid-cols-2 gap-3">
          {db.permisos.map((p) => {
            const als = db.alumnos.filter((a) => a.grupoId === p.grupoId && a.activo);
            const ent = als.filter((a) => p.registros?.[a.id]?.entregado).length;
            const aut = als.filter((a) => p.registros?.[a.id]?.autorizado).length;
            const g = db.grupos.find((x) => x.id === p.grupoId);
            return (
              <Card key={p.id}>
                <div className="flex items-start justify-between">
                  <div><p className="font-medium text-slate-900">{p.nombre}</p>
                    <p className="text-xs text-slate-500">{fFecha(p.fecha)} · {g ? `${g.grado}° ${g.grupo}` : ""} · {p.lugar}</p></div>
                  <div className="flex gap-1">
                    <button onClick={() => setModal(p)} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-500"><Pencil size={15} /></button>
                    <button onClick={() => setConf({ texto: `¿Eliminar la actividad "${p.nombre}" y sus registros de permiso?`, onSi: () => {
                      eliminarConRespaldo({ tipo: "Permiso", descripcion: p.nombre, ciclo: db.ciclo, datos: { permisos: [p] } });
                      upd((d) => { d.permisos = d.permisos.filter((x) => x.id !== p.id); });
                      setConf(null); toast("Enviado a la papelera");
                    } })} className="p-1.5 rounded-lg hover:bg-rose-50 text-rose-500"><Trash2 size={15} /></button>
                  </div>
                </div>
                <div className="grid grid-cols-3 gap-2 mt-3 text-center text-xs">
                  <div className="bg-emerald-50 rounded-lg py-1.5"><p className="text-base font-semibold text-emerald-700">{aut}</p><p className="text-slate-500">Autorizados</p></div>
                  <div className="bg-slate-50 rounded-lg py-1.5"><p className="text-base font-semibold text-slate-700">{ent}</p><p className="text-slate-500">Entregaron</p></div>
                  <div className="bg-amber-50 rounded-lg py-1.5"><p className="text-base font-semibold text-amber-700">{als.length - ent}</p><p className="text-slate-500">Pendientes</p></div>
                </div>
                <div className="flex gap-2 mt-3">
                  <Btn size="sm" tipo="secundario" className="grow" onClick={() => setAbierto(p.id)}>Pasar lista de permisos</Btn>
                  <Btn size="sm" tipo="secundario" icon={FileDown} onClick={() => descargarListas(p)}>PDF</Btn>
                </div>
              </Card>
            );
          })}
        </div>
      )}

      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.id ? "Editar actividad" : "Nueva actividad con permiso"}
        footer={<><Btn tipo="secundario" onClick={() => setModal(null)}>Cancelar</Btn><Btn icon={Save} onClick={guardar}>Guardar</Btn></>}>
        {modal && (
          <div className="grid sm:grid-cols-2 gap-3">
            <Campo label="Nombre de la actividad" req><Inp value={modal.nombre} onChange={(e) => setModal({ ...modal, nombre: e.target.value })} /></Campo>
            <Campo label="Fecha"><Inp type="date" value={modal.fecha} onChange={(e) => setModal({ ...modal, fecha: e.target.value })} /></Campo>
            <Campo label="Lugar"><Inp value={modal.lugar} onChange={(e) => setModal({ ...modal, lugar: e.target.value })} /></Campo>
            <Campo label="Responsable"><Inp value={modal.responsable} onChange={(e) => setModal({ ...modal, responsable: e.target.value })} /></Campo>
            <Campo label="Grupo"><Sel value={modal.grupoId} onChange={(e) => setModal({ ...modal, grupoId: e.target.value })}>{db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}</Sel></Campo>
            <div className="grid grid-cols-2 gap-3">
              <Campo label="Hora de salida"><Inp type="time" value={modal.horaSalida} onChange={(e) => setModal({ ...modal, horaSalida: e.target.value })} /></Campo>
              <Campo label="Hora de regreso"><Inp type="time" value={modal.horaRegreso} onChange={(e) => setModal({ ...modal, horaRegreso: e.target.value })} /></Campo>
            </div>
            <div className="sm:col-span-2"><Campo label="Descripción"><Area value={modal.descripcion} onChange={(e) => setModal({ ...modal, descripcion: e.target.value })} /></Campo></div>
            <div className="sm:col-span-2"><Campo label="Referencia del archivo digital" hint="Escribe dónde se resguarda el permiso firmado (carpeta, folder físico o enlace).">
              <Inp value={modal.archivo} onChange={(e) => setModal({ ...modal, archivo: e.target.value })} /></Campo></div>
          </div>
        )}
      </Modal>

      <Modal open={!!detalle} onClose={() => setAbierto(null)} title={detalle?.nombre} ancho="max-w-3xl"
        footer={<Btn onClick={() => setAbierto(null)}>Listo</Btn>}>
        {detalle && (
          <Tabla cols={["No.", "Nombre", "Entregó", "Fecha", "Autorizado", "Observaciones"]}>
            {alumnosDet.map((a) => { const r = detalle.registros?.[a.id] || {};
              return (
                <tr key={a.id}>
                  <td className="py-2 px-3 text-slate-400">{a.numLista}</td>
                  <td className="py-2 px-3">{nomComp(a)}</td>
                  <td className="py-2 px-3"><input type="checkbox" checked={!!r.entregado} onChange={(e) => setReg(detalle.id, a.id, "entregado", e.target.checked)} /></td>
                  <td className="py-2 px-3 text-xs text-slate-500">{r.fechaEntrega ? fFecha(r.fechaEntrega) : "—"}</td>
                  <td className="py-2 px-3"><input type="checkbox" checked={!!r.autorizado} onChange={(e) => setReg(detalle.id, a.id, "autorizado", e.target.checked)} /></td>
                  <td className="py-2 px-3"><input value={r.obs || ""} onChange={(e) => setReg(detalle.id, a.id, "obs", e.target.value)} className="w-full border border-slate-200 rounded px-2 py-1 text-xs" /></td>
                </tr>
              ); })}
          </Tabla>
        )}
      </Modal>
      <Confirmar open={!!conf} texto={conf?.texto} onSi={conf?.onSi} onNo={() => setConf(null)} />
    </div>
  );
}

/* ============================================================
   MÓDULO: ACTIVIDADES
   ============================================================ */
function Actividades({ db, upd, toast }) {
  const [grupoId, setGrupoId] = useState(db.grupos[0]?.id || "");
  const [trim, setTrim] = useState(db.config.trimestre);
  const [modal, setModal] = useState(null);
  const [calif, setCalif] = useState(null);
  const [conf, setConf] = useState(null);
  const lista = db.actividades.filter((a) => a.grupoId === grupoId && Number(a.trimestre) === Number(trim));
  const vacio = { grupoId, nombre: "", fecha: hoy(), fechaEntrega: hoy(), trimestre: trim, instrumentoId: db.catalogos.instrumentos[0]?.id || "", materia: db.config.asignatura || "", evidencia: "", campoFormativo: "", contenido: "", pda: "", eje: "", proyecto: "", proposito: "", descripcion: "", producto: "", valor: 10 };

  const guardar = () => {
    if (!modal.nombre) return toast("Escribe el nombre de la actividad", "error");
    upd((d) => {
      if (modal.id) { const i = d.actividades.findIndex((a) => a.id === modal.id); d.actividades[i] = { ...modal, trimestre: Number(modal.trimestre) }; }
      else {
        const nueva = { ...modal, id: uid("act"), trimestre: Number(modal.trimestre) };
        d.actividades.push(nueva);
        d.alumnos.filter((a) => a.grupoId === nueva.grupoId && a.activo).forEach((a) => d.entregas.push({ id: uid("ent"), actividadId: nueva.id, alumnoId: a.id, estado: "Pendiente", calificacion: 0 }));
      }
    });
    setModal(null); toast("Actividad guardada");
  };

  const setEnt = (actId, alumnoId, campo, valor) => upd((d) => {
    let e = d.entregas.find((x) => x.actividadId === actId && x.alumnoId === alumnoId);
    if (!e) { e = { id: uid("ent"), actividadId: actId, alumnoId, estado: "Pendiente", calificacion: 0 }; d.entregas.push(e); }
    e[campo] = valor;
    if (campo === "calificacion" && e.estado === "Pendiente") e.estado = "Entregada";
  });

  const actCalif = db.actividades.find((a) => a.id === calif);
  const alumnosCalif = actCalif ? sortAl(db.alumnos.filter((a) => a.grupoId === actCalif.grupoId && a.activo)) : [];

  return (
    <div className="space-y-4">
      <Titulo sub="Planeación, seguimiento de entregas y calificación en un mismo lugar."
        right={<Btn icon={Plus} onClick={() => setModal({ ...vacio, grupoId, trimestre: trim })} disabled={!db.grupos.length}>Nueva actividad</Btn>}>Actividades</Titulo>
      <div className="flex gap-2 flex-wrap">
        <Sel value={grupoId} onChange={(e) => setGrupoId(e.target.value)} className="max-w-[170px]">{db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}</Sel>
        <Sel value={trim} onChange={(e) => setTrim(Number(e.target.value))} className="max-w-[150px]">{[1, 2, 3].map((t) => <option key={t} value={t}>Trimestre {t}</option>)}</Sel>
      </div>

      {lista.length === 0 ? <Card><Vacio texto="No hay actividades en este trimestre." /></Card> : (
        <div className="space-y-2">
          {lista.map((a) => {
            const ents = db.entregas.filter((e) => e.actividadId === a.id);
            const ent = ents.filter((e) => e.estado === "Entregada").length;
            const tarde = ents.filter((e) => e.estado === "Fuera de tiempo").length;
            const no = ents.filter((e) => e.estado === "No entregada").length;
            const ins = db.catalogos.instrumentos.find((i) => i.id === a.instrumentoId);
            const notas = ents.filter((e) => e.estado === "Entregada" || e.estado === "Fuera de tiempo").map((e) => Number(e.calificacion) || 0);
            const prom = notas.length ? round(notas.reduce((s, n) => s + n, 0) / notas.length, 1) : null;
            return (
              <Card key={a.id}>
                <div className="flex items-start justify-between gap-3 flex-wrap">
                  <div className="grow min-w-[200px]">
                    <div className="flex items-center gap-2 flex-wrap">
                      <p className="font-medium text-slate-900">{a.nombre}</p>
                      <Pill cls="bg-slate-100 text-slate-600 border-slate-200">{ins?.nombre || "Sin instrumento"}</Pill>
                    </div>
                    <p className="text-xs text-slate-500 mt-0.5">{a.materia ? a.materia + " · " : ""}Entrega {fFecha(a.fechaEntrega)} · {a.campoFormativo || "Sin campo formativo"}</p>
                    <div className="flex gap-2 mt-2 text-[11px] flex-wrap">
                      <Pill cls="bg-emerald-50 text-emerald-700 border-emerald-200">Entregadas {ent}</Pill>
                      {tarde > 0 && <Pill cls="bg-amber-50 text-amber-700 border-amber-200">Fuera de tiempo {tarde}</Pill>}
                      <Pill cls="bg-rose-50 text-rose-700 border-rose-200">No entregadas {no}</Pill>
                      <Pill cls="bg-slate-50 text-slate-600 border-slate-200">Pendientes {ents.length - ent - tarde - no}</Pill>
                      {prom !== null && <Pill cls="bg-sky-50 text-sky-700 border-sky-200">Promedio {prom}</Pill>}
                    </div>
                  </div>
                  <div className="flex gap-1 shrink-0">
                    <Btn size="sm" tipo="secundario" onClick={() => setCalif(a.id)}>Calificar</Btn>
                    <button onClick={() => setModal(a)} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-500"><Pencil size={15} /></button>
                    <button onClick={() => setConf({ texto: `¿Eliminar "${a.nombre}" y todas sus calificaciones?`, onSi: () => {
                      eliminarConRespaldo({ tipo: "Actividad", descripcion: a.nombre, ciclo: db.ciclo, datos: { actividades: [a], entregas: db.entregas.filter((x) => x.actividadId === a.id) } });
                      upd((d) => { d.actividades = d.actividades.filter((x) => x.id !== a.id); d.entregas = d.entregas.filter((x) => x.actividadId !== a.id); });
                      setConf(null); toast("Actividad enviada a la papelera");
                    } })} className="p-1.5 rounded-lg hover:bg-rose-50 text-rose-500"><Trash2 size={15} /></button>
                  </div>
                </div>
              </Card>
            );
          })}
        </div>
      )}

      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.id ? "Editar actividad" : "Nueva actividad"}
        footer={<><Btn tipo="secundario" onClick={() => setModal(null)}>Cancelar</Btn><Btn icon={Save} onClick={guardar}>Guardar</Btn></>}>
        {modal && (
          <div className="grid sm:grid-cols-2 gap-3">
            <div className="sm:col-span-2"><Campo label="Nombre" req><Inp value={modal.nombre} onChange={(e) => setModal({ ...modal, nombre: e.target.value })} /></Campo></div>
            <Campo label="Grupo"><Sel value={modal.grupoId} onChange={(e) => setModal({ ...modal, grupoId: e.target.value })}>{db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}</Sel></Campo>
            <Campo label="Instrumento de evaluación"><Sel value={modal.instrumentoId} onChange={(e) => setModal({ ...modal, instrumentoId: e.target.value })}>{db.catalogos.instrumentos.map((i) => <option key={i.id} value={i.id}>{i.nombre} ({i.porcentaje}%)</option>)}</Sel></Campo>
            <Campo label="Fecha"><Inp type="date" value={modal.fecha} onChange={(e) => setModal({ ...modal, fecha: e.target.value })} /></Campo>
            <Campo label="Fecha de entrega"><Inp type="date" value={modal.fechaEntrega} onChange={(e) => setModal({ ...modal, fechaEntrega: e.target.value })} /></Campo>
            <Campo label="Trimestre"><Sel value={modal.trimestre} onChange={(e) => setModal({ ...modal, trimestre: e.target.value })}>{[1, 2, 3].map((t) => <option key={t} value={t}>Trimestre {t}</option>)}</Sel></Campo>
            <Campo label="Valor (escala)"><Inp type="number" value={modal.valor} onChange={(e) => setModal({ ...modal, valor: e.target.value })} /></Campo>
            <Campo label="Materia"><Sel value={modal.materia || ""} onChange={(e) => setModal({ ...modal, materia: e.target.value })}><option value="">Sin especificar</option>{(db.catalogos.materias || []).map((m) => <option key={m}>{m}</option>)}</Sel></Campo>
            <Campo label="Evidencia"><Inp value={modal.evidencia || ""} onChange={(e) => setModal({ ...modal, evidencia: e.target.value })} placeholder="Cuaderno, fotografía, producto…" /></Campo>
            <Campo label="Campo formativo"><Sel value={modal.campoFormativo} onChange={(e) => setModal({ ...modal, campoFormativo: e.target.value })}><option value="">Sin especificar</option>{CAMPOS_FORM.map((c) => <option key={c}>{c}</option>)}</Sel></Campo>
            <Campo label="Eje articulador"><Sel value={modal.eje} onChange={(e) => setModal({ ...modal, eje: e.target.value })}><option value="">Sin especificar</option>{EJES.map((c) => <option key={c}>{c}</option>)}</Sel></Campo>
            <div className="sm:col-span-2"><Campo label="Contenido"><Inp value={modal.contenido} onChange={(e) => setModal({ ...modal, contenido: e.target.value })} /></Campo></div>
            <div className="sm:col-span-2"><Campo label="Proceso de desarrollo de aprendizaje (PDA)"><Area value={modal.pda} onChange={(e) => setModal({ ...modal, pda: e.target.value })} /></Campo></div>
            <Campo label="Proyecto"><Inp value={modal.proyecto} onChange={(e) => setModal({ ...modal, proyecto: e.target.value })} /></Campo>
            <Campo label="Producto esperado"><Inp value={modal.producto} onChange={(e) => setModal({ ...modal, producto: e.target.value })} /></Campo>
            <div className="sm:col-span-2"><Campo label="Propósito"><Area value={modal.proposito} onChange={(e) => setModal({ ...modal, proposito: e.target.value })} /></Campo></div>
            <div className="sm:col-span-2"><Campo label="Descripción"><Area value={modal.descripcion} onChange={(e) => setModal({ ...modal, descripcion: e.target.value })} /></Campo></div>
          </div>
        )}
      </Modal>

      <Modal open={!!actCalif} onClose={() => setCalif(null)} title={`Calificar: ${actCalif?.nombre || ""}`} ancho="max-w-2xl" footer={<Btn onClick={() => setCalif(null)}>Listo</Btn>}>
        {actCalif && (
          <>
            <div className="flex gap-2 mb-3 flex-wrap">
              <Btn size="sm" tipo="secundario" onClick={() => alumnosCalif.forEach((a) => setEnt(actCalif.id, a.id, "estado", "Entregada"))}>Marcar todas como entregadas</Btn>
            </div>
            <Tabla cols={["No.", "Nombre", "Estado", "Calificación"]}>
              {alumnosCalif.map((a) => { const e = db.entregas.find((x) => x.actividadId === actCalif.id && x.alumnoId === a.id) || {};
                return (
                  <tr key={a.id}>
                    <td className="py-2 px-3 text-slate-400">{a.numLista}</td>
                    <td className="py-2 px-3">{nomComp(a)}</td>
                    <td className="py-2 px-3"><select value={e.estado || "Pendiente"} onChange={(ev) => setEnt(actCalif.id, a.id, "estado", ev.target.value)} className="border border-slate-200 rounded px-2 py-1 text-xs">{EST_ACT.map((s) => <option key={s}>{s}</option>)}</select></td>
                    <td className="py-2 px-3"><input type="number" step="0.1" min="0" max={actCalif.valor} value={e.calificacion ?? ""} onChange={(ev) => setEnt(actCalif.id, a.id, "calificacion", Number(ev.target.value))} className="w-20 border border-slate-200 rounded px-2 py-1 text-xs" /></td>
                  </tr>
                ); })}
            </Tabla>
          </>
        )}
      </Modal>
      <Confirmar open={!!conf} texto={conf?.texto} onSi={conf?.onSi} onNo={() => setConf(null)} />
    </div>
  );
}

/* ============================================================
   MÓDULO: EVALUACIÓN
   ============================================================ */
function Evaluacion({ db, upd, ir, toast }) {
  const [grupoId, setGrupoId] = useState(db.grupos[0]?.id || "");
  const [trim, setTrim] = useState(db.config.trimestre);
  const [editIns, setEditIns] = useState(false);
  const [ins, setIns] = useState(db.catalogos.instrumentos);
  const alumnos = sortAl(db.alumnos.filter((a) => a.grupoId === grupoId && a.activo));
  const suma = ins.reduce((s, i) => s + Number(i.porcentaje || 0), 0);

  const guardarIns = () => {
    if (suma !== 100) return toast(`Los porcentajes suman ${suma}%. Deben sumar 100%.`, "error");
    upd((d) => { d.catalogos.instrumentos = ins.map((i) => ({ ...i, porcentaje: Number(i.porcentaje) })); });
    setEditIns(false); toast("Porcentajes actualizados. Los registros anteriores se conservan.");
  };

  const filas = alumnos.map((a) => ({ a, ...calificacionTrimestre(db, a.id, trim), acum: promedioAcumulado(db, a.id) }));
  const evaluados = filas.filter((f) => f.final > 0);
  const promGrupo = evaluados.length ? round(evaluados.reduce((s, f) => s + f.final, 0) / evaluados.length, 1) : 0;

  const exportar = () => exportarExcel(`evaluacion_T${trim}`, [
    ["No.", "Nombre", ...db.catalogos.instrumentos.map((i) => `${i.nombre} (${i.porcentaje}%)`), `Calificacion T${trim}`, "Promedio acumulado"],
    ...filas.map((f) => [f.a.numLista, nomComp(f.a), ...f.detalle.map((d) => d.promedio ?? ""), f.final, f.acum]),
  ]);

  const pdfEvaluacion = () => {
    const g2 = db.grupos.find((x) => x.id === grupoId);
    pdfTabla({
      titulo: `Reporte de evaluación · Trimestre ${trim}`,
      subtitulo: [g2 ? `${g2.grado}° ${g2.grupo}` : "", db.config.docente, db.config.asignatura].filter(Boolean).join("   ·   "),
      columnas: ["No.", "Nombre", ...db.catalogos.instrumentos.map((i) => `${i.nombre} ${i.porcentaje}%`), `T${trim}`, "Acumulado"],
      filas: filas.map((f) => [f.a.numLista, nomComp(f.a), ...f.detalle.map((d) => d.promedio ?? "—"), f.final || "—", f.acum || "—"]),
      config: db.config, ciclo: db.ciclo,
      resumen: [{ label: "Promedio del grupo", valor: promGrupo || "—" }, { label: "Evaluados", valor: `${evaluados.length}/${filas.length}` }],
      nota: "Calificaciones calculadas con los porcentajes configurados en el sistema.",
    });
    toast("PDF descargado");
  };

  return (
    <div className="space-y-4">
      <Titulo sub="La calificación se calcula sola con las actividades y los porcentajes configurados."
        right={<div className="flex gap-2"><Btn size="sm" tipo="secundario" icon={FileDown} onClick={pdfEvaluacion}>PDF</Btn><Btn size="sm" tipo="secundario" icon={Download} onClick={exportar}>Excel</Btn><Btn size="sm" tipo="secundario" icon={Settings} onClick={() => { setIns(clone(db.catalogos.instrumentos)); setEditIns(true); }}>Porcentajes</Btn></div>}>Evaluación</Titulo>
      <div className="flex gap-2 flex-wrap">
        <Sel value={grupoId} onChange={(e) => setGrupoId(e.target.value)} className="max-w-[170px]">{db.grupos.map((x) => <option key={x.id} value={x.id}>{x.grado}° {x.grupo}</option>)}</Sel>
        <Sel value={trim} onChange={(e) => setTrim(Number(e.target.value))} className="max-w-[150px]">{[1, 2, 3].map((t) => <option key={t} value={t}>Trimestre {t}</option>)}</Sel>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <Stat label="Promedio del grupo" valor={promGrupo || "—"} color="text-emerald-700" />
        <Stat label="Estudiantes evaluados" valor={evaluados.length + "/" + filas.length} />
        <Stat label="Actividades del trimestre" valor={db.actividades.filter((a) => a.grupoId === grupoId && Number(a.trimestre) === Number(trim)).length} />
        <Stat label="Suma de porcentajes" valor={db.catalogos.instrumentos.reduce((s, i) => s + i.porcentaje, 0) + "%"} />
      </div>

      <Card pad={false} className="p-4">
        {alumnos.length === 0 ? <Vacio texto="No hay estudiantes en este grupo." /> : (
          <Tabla cols={["No.", "Nombre", ...db.catalogos.instrumentos.map((i) => `${i.nombre} ${i.porcentaje}%`), `T${trim}`, "Acumulado"]}>
            {filas.map((f) => (
              <tr key={f.a.id} className="hover:bg-slate-50">
                <td className="py-2 px-3 text-slate-400">{f.a.numLista}</td>
                <td className="py-2 px-3"><button onClick={() => ir("ficha", { alumnoId: f.a.id })} className="hover:text-emerald-700 text-left">{nomComp(f.a)}</button></td>
                {f.detalle.map((d) => <td key={d.id} className="py-2 px-3 text-slate-600">{d.promedio ?? "—"}</td>)}
                <td className={`py-2 px-3 font-semibold ${f.final >= 6 ? "text-emerald-700" : f.final > 0 ? "text-amber-600" : "text-slate-400"}`}>{f.final || "—"}</td>
                <td className="py-2 px-3 font-medium">{f.acum || "—"}</td>
              </tr>
            ))}
          </Tabla>
        )}
      </Card>

      <Modal open={editIns} onClose={() => setEditIns(false)} title="Instrumentos y porcentajes" ancho="max-w-lg"
        footer={<><Btn tipo="secundario" onClick={() => setEditIns(false)}>Cancelar</Btn><Btn icon={Save} onClick={guardarIns}>Guardar</Btn></>}>
        <div className="space-y-2">
          <Aviso>Al cambiar los porcentajes, las calificaciones se recalculan pero no se modifica ningún registro de actividad ya capturado.</Aviso>
          {ins.map((i, k) => (
            <div key={i.id} className="flex gap-2 items-center">
              <input value={i.nombre} onChange={(e) => { const c = [...ins]; c[k] = { ...i, nombre: e.target.value }; setIns(c); }} className={inputCls} />
              <input type="number" value={i.porcentaje} onChange={(e) => { const c = [...ins]; c[k] = { ...i, porcentaje: e.target.value }; setIns(c); }} className={inputCls + " w-24"} />
              <button onClick={() => setIns(ins.filter((_, j) => j !== k))} className="p-2 text-rose-500 hover:bg-rose-50 rounded-lg"><Trash2 size={15} /></button>
            </div>
          ))}
          <div className="flex justify-between items-center pt-2">
            <Btn size="sm" tipo="secundario" icon={Plus} onClick={() => setIns([...ins, { id: uid("ins"), nombre: "", porcentaje: 0 }])}>Agregar instrumento</Btn>
            <span className={`text-sm font-medium ${suma === 100 ? "text-emerald-700" : "text-rose-600"}`}>Suma: {suma}%</span>
          </div>
        </div>
      </Modal>
    </div>
  );
}

/* ============================================================
   MÓDULO: ECOEMS
   ============================================================ */
function Ecoems({ db, upd, ir, toast }) {
  const [grupoId, setGrupoId] = useState(db.grupos[0]?.id || "");
  const [tab, setTab] = useState("captura");
  const [alumnoId, setAlumnoId] = useState("");
  const [numero, setNumero] = useState(1);
  const [fecha, setFecha] = useState(hoy());
  const [res, setRes] = useState({});
  const [obs, setObs] = useState("");
  const mats = db.catalogos.materiasEcoems;
  const alumnos = sortAl(db.alumnos.filter((a) => a.grupoId === grupoId && a.activo));
  const maxTotal = mats.reduce((t, m) => t + Number(m.reactivos || 0), 0);

  useEffect(() => { setAlumnoId(alumnos[0]?.id || ""); }, [grupoId]);
  const existente = db.ecoems.find((s) => s.alumnoId === alumnoId && Number(s.numero) === Number(numero));
  useEffect(() => {
    setRes(existente?.resultados ? { ...existente.resultados } : {});
    setObs(existente?.observaciones || "");
    if (existente?.fecha) setFecha(existente.fecha);
  }, [alumnoId, numero, existente?.id]);

  const totalAciertos = mats.reduce((t, m) => t + (Number(res[m.id]) || 0), 0);
  const porcentaje = pct(totalAciertos, maxTotal);
  const sem = semaforoDe(porcentaje, db.config);

  const guardar = () => {
    if (!alumnoId) return toast("Selecciona un estudiante", "error");
    upd((d) => {
      const i = d.ecoems.findIndex((s) => s.alumnoId === alumnoId && Number(s.numero) === Number(numero));
      const dato = { alumnoId, grupoId, fecha, numero: Number(numero), resultados: res, observaciones: obs };
      if (i >= 0) d.ecoems[i] = { ...d.ecoems[i], ...dato }; else d.ecoems.push({ id: uid("eco"), ...dato });
    });
    toast(`Aplicación ${numero} guardada · ${totalAciertos}/${maxTotal} (${porcentaje}%)`);
    const idx = alumnos.findIndex((a) => a.id === alumnoId);
    if (idx >= 0 && idx < alumnos.length - 1) setAlumnoId(alumnos[idx + 1].id);
  };

  const numeros = [...new Set(db.ecoems.filter((s) => s.grupoId === grupoId).map((s) => s.numero))].sort((a, b) => a - b);

  const analisis = useMemo(() => {
    if (!numeros.length) return null;
    const ultimo = numeros[numeros.length - 1];
    const sims = db.ecoems.filter((s) => s.grupoId === grupoId && s.numero === ultimo);
    const porMateria = mats.map((m) => {
      const ac = sims.map((s) => Number(s.resultados?.[m.id] || 0));
      const prom = ac.length ? round(ac.reduce((a, b) => a + b, 0) / ac.length, 1) : 0;
      return { nombre: m.nombre, corto: m.nombre.length > 14 ? m.nombre.slice(0, 13) + "…" : m.nombre, reactivos: m.reactivos, aciertos: prom, promedio: pct(prom, m.reactivos) };
    });
    const filas = sims.map((s) => { const r = resumenSim(s, mats); return { alumno: db.alumnos.find((a) => a.id === s.alumnoId), aciertos: r.aciertos, porcentaje: r.porcentaje }; })
      .sort((a, b) => b.aciertos - a.aciertos);
    const ps = filas.map((f) => f.porcentaje);
    const ord = [...porMateria].sort((a, b) => b.promedio - a.promedio);
    const evolucion = numeros.map((n) => {
      const ss = db.ecoems.filter((s) => s.grupoId === grupoId && s.numero === n);
      const rr = ss.map((s) => resumenSim(s, mats));
      return { nombre: "Ap. " + n, aciertos: rr.length ? round(rr.reduce((a, b) => a + b.aciertos, 0) / rr.length, 1) : 0, promedio: rr.length ? round(rr.reduce((a, b) => a + b.porcentaje, 0) / rr.length, 1) : 0 };
    });
    const semaforo = { alto: filas.filter((f) => semaforoDe(f.porcentaje, db.config).nivel === "alto").length, medio: filas.filter((f) => semaforoDe(f.porcentaje, db.config).nivel === "medio").length, bajo: filas.filter((f) => semaforoDe(f.porcentaje, db.config).nivel === "bajo").length };
    return { ultimo, porMateria, evolucion, filas, semaforo,
      promedioAciertos: filas.length ? round(filas.reduce((a, b) => a + b.aciertos, 0) / filas.length, 1) : 0,
      promedio: ps.length ? round(ps.reduce((a, b) => a + b, 0) / ps.length, 1) : 0,
      mayor: Math.max(...ps, 0), menor: ps.length ? Math.min(...ps) : 0,
      mejorMateria: ord[0], areaRefuerzo: ord[ord.length - 1] };
  }, [db, grupoId]);

  const comparativoGrupos = useMemo(() => db.grupos.map((g) => {
    const nums = [...new Set(db.ecoems.filter((s) => s.grupoId === g.id).map((s) => s.numero))];
    if (!nums.length) return null;
    const ult = Math.max(...nums);
    const rr = db.ecoems.filter((s) => s.grupoId === g.id && s.numero === ult).map((s) => resumenSim(s, mats));
    return { nombre: `${g.grado}°${g.grupo}`, aplicacion: ult, aciertos: round(rr.reduce((a, b) => a + b.aciertos, 0) / rr.length, 1), promedio: round(rr.reduce((a, b) => a + b.porcentaje, 0) / rr.length, 1) };
  }).filter(Boolean), [db]);

  const exportarConcentrado = () => {
    const g = db.grupos.find((x) => x.id === grupoId);
    exportarExcel(`ecoems_${g ? g.grado + g.grupo : ""}`,
      [["No.", "Nombre", ...mats.map((m) => `${m.nombre} (${m.reactivos})`), "Aciertos", "Máximo", "Porcentaje", "Semáforo"],
       ...alumnos.flatMap((a) => db.ecoems.filter((s) => s.alumnoId === a.id).sort((x, y) => x.numero - y.numero).map((s) => {
         const r = resumenSim(s, mats);
         return [a.numLista, `${nomComp(a)} (Ap. ${s.numero})`, ...mats.map((m) => Number(s.resultados?.[m.id] || 0)), r.aciertos, r.reactivos, r.porcentaje, semaforoDe(r.porcentaje, db.config).label];
       }))]);
    toast("Concentrado exportado");
  };

  const pdfConcentrado = () => {
    const g2 = db.grupos.find((x) => x.id === grupoId);
    pdfTabla({
      titulo: "Concentrado del simulador ECOEMS",
      subtitulo: [g2 ? `${g2.grado}° ${g2.grupo}` : "", `${maxTotal} aciertos máximos`, `${mats.length} áreas`].filter(Boolean).join("   ·   "),
      columnas: ["No.", "Nombre", ...numeros.map((n) => `Ap. ${n}`), "Avance", "Semáforo"],
      filas: alumnos.map((a) => { const h = historialEcoems(db, a.id);
        return [a.numLista, nomComp(a), ...numeros.map((n) => { const it = h.items.find((x) => x.sim.numero === n); return it ? `${it.aciertos}/${it.reactivos}` : "—"; }),
          h.items.length > 1 ? (h.diferencia > 0 ? "+" : "") + h.diferencia + "%" : "—",
          h.items.length ? semaforoDe(h.actual, db.config).label : "—"]; }),
      config: db.config, ciclo: db.ciclo,
      resumen: analisis ? [{ label: "Promedio de aciertos", valor: `${analisis.promedioAciertos}/${maxTotal}` }, { label: "Promedio del grupo", valor: analisis.promedio + "%" }, { label: "Mejor área", valor: analisis.mejorMateria?.nombre || "—" }, { label: "Mayor reforzamiento", valor: analisis.areaRefuerzo?.nombre || "—" }] : null,
      nota: AVISO_ECOEMS,
    });
    toast("PDF descargado");
  };

  return (
    <div className="space-y-4">
      <Titulo sub={`Simulador de ${maxTotal} aciertos en ${mats.length} áreas. Captura, comparación y semáforo académico.`}>Simulador ECOEMS</Titulo>
      <div className="flex gap-2 flex-wrap items-center">
        <Sel value={grupoId} onChange={(e) => setGrupoId(e.target.value)} className="max-w-[170px]">{db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}</Sel>
        {maxTotal !== 128 && <span className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-2 py-1">Las áreas suman {maxTotal} aciertos. Ajusta el catálogo en Configuración si esperabas 128.</span>}
      </div>
      <Tabs activa={tab} set={setTab} tabs={[{ id: "captura", label: "Captura" }, { id: "concentrado", label: "Concentrado" }, { id: "analisis", label: "Análisis" }, { id: "comparativo", label: "Comparativo" }]} />

      {tab === "captura" && (
        <Card>
          <div className="grid sm:grid-cols-3 gap-3 mb-4">
            <Campo label="Estudiante"><Sel value={alumnoId} onChange={(e) => setAlumnoId(e.target.value)}>{alumnos.map((a) => <option key={a.id} value={a.id}>{a.numLista}. {nomComp(a)}</option>)}</Sel></Campo>
            <Campo label="Número de aplicación"><Inp type="number" min="1" value={numero} onChange={(e) => setNumero(Number(e.target.value))} /></Campo>
            <Campo label="Fecha del simulador"><Inp type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Campo>
          </div>
          {existente && <p className="text-xs text-emerald-700 mb-3">Ya hay un registro de esta aplicación. Al guardar se actualiza.</p>}
          <div className="space-y-1.5">
            {mats.map((m) => {
              const ac = Number(res[m.id]) || 0;
              const p = pct(ac, m.reactivos);
              return (
                <div key={m.id} className="flex items-center gap-2">
                  <span className="grow text-sm text-slate-700 min-w-[110px]">{m.nombre}</span>
                  <input type="number" min="0" max={m.reactivos} value={res[m.id] ?? ""} placeholder="0"
                    onChange={(e) => setRes({ ...res, [m.id]: Math.min(m.reactivos, Math.max(0, Number(e.target.value))) })}
                    className="w-16 border border-slate-300 rounded-lg px-2 py-1.5 text-sm text-center" />
                  <span className="text-xs text-slate-400 w-10">/ {m.reactivos}</span>
                  <span className="text-xs w-11 text-right font-medium text-slate-600">{p}%</span>
                  <span className={`w-2.5 h-2.5 rounded-full shrink-0 ${semaforoDe(p, db.config).punto}`} />
                </div>
              );
            })}
          </div>
          <div className="mt-3"><Campo label="Observaciones"><Inp value={obs} onChange={(e) => setObs(e.target.value)} placeholder="Notas sobre la aplicación" /></Campo></div>
          <div className="flex items-center justify-between mt-4 pt-3 border-t border-slate-200 flex-wrap gap-3">
            <div className="flex items-center gap-3 flex-wrap">
              <div className="text-sm"><span className="text-slate-500">Total: </span><span className="font-semibold text-slate-900">{totalAciertos} / {maxTotal}</span>
                <span className="ml-2 font-semibold text-slate-900">{porcentaje}%</span></div>
              <Pill cls={sem.cls}>{sem.label}</Pill>
            </div>
            <Btn icon={Save} onClick={guardar}>Guardar y seguir</Btn>
          </div>
          <p className="text-[11px] text-slate-400 mt-3">{AVISO_ECOEMS}</p>
        </Card>
      )}

      {tab === "concentrado" && (
        <Card pad={false} className="p-4">
          <div className="flex justify-end gap-2 mb-2">
            <Btn size="sm" tipo="secundario" icon={FileDown} onClick={pdfConcentrado}>PDF</Btn>
            <Btn size="sm" tipo="secundario" icon={Download} onClick={exportarConcentrado}>Excel</Btn>
          </div>
          {numeros.length === 0 ? <Vacio texto="Aún no hay aplicaciones registradas en este grupo." /> : (
            <Tabla cols={["No.", "Nombre", ...numeros.map((n) => `Ap. ${n}`), "Avance", "Semáforo"]}>
              {alumnos.map((a) => {
                const h = historialEcoems(db, a.id);
                const sm = semaforoDe(h.actual, db.config);
                return (
                  <tr key={a.id} className="hover:bg-slate-50">
                    <td className="py-2 px-3 text-slate-400">{a.numLista}</td>
                    <td className="py-2 px-3"><button onClick={() => ir("ficha", { alumnoId: a.id })} className="hover:text-emerald-700 text-left">{nomComp(a)}</button></td>
                    {numeros.map((n) => { const it = h.items.find((x) => x.sim.numero === n);
                      return <td key={n} className="py-2 px-3 text-slate-600">{it ? `${it.aciertos}/${it.reactivos}` : "—"}</td>; })}
                    <td className={`py-2 px-3 font-medium ${h.diferencia >= 0 ? "text-emerald-700" : "text-amber-600"}`}>{h.items.length > 1 ? (h.diferencia > 0 ? "+" : "") + h.diferencia + "%" : "—"}</td>
                    <td className="py-2 px-3">{h.items.length ? <span className={`inline-flex items-center gap-1.5 text-xs`}><span className={`w-2.5 h-2.5 rounded-full ${sm.punto}`} />{h.actual}%</span> : "—"}</td>
                  </tr>
                );
              })}
            </Tabla>
          )}
        </Card>
      )}

      {tab === "analisis" && (analisis ? (
        <div className="space-y-4">
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <Stat label="Promedio de aciertos" valor={`${analisis.promedioAciertos}/${maxTotal}`} sub={`Aplicación ${analisis.ultimo}`} />
            <Stat label="Promedio porcentual" valor={analisis.promedio + "%"} color="text-emerald-700" />
            <Stat label="Mejor área" valor={analisis.mejorMateria?.promedio + "%"} sub={analisis.mejorMateria?.nombre} />
            <Stat label="Mayor reforzamiento" valor={analisis.areaRefuerzo?.promedio + "%"} sub={analisis.areaRefuerzo?.nombre} color="text-amber-600" />
          </div>

          <Card>
            <p className="text-sm font-medium text-slate-700 mb-3">Semáforo del grupo</p>
            <div className="grid grid-cols-3 gap-2 text-center">
              <div className="bg-emerald-50 border border-emerald-200 rounded-lg py-3"><p className="text-2xl font-semibold text-emerald-700">{analisis.semaforo.alto}</p><p className="text-[11px] text-slate-600">Desempeño alto</p></div>
              <div className="bg-amber-50 border border-amber-200 rounded-lg py-3"><p className="text-2xl font-semibold text-amber-700">{analisis.semaforo.medio}</p><p className="text-[11px] text-slate-600">Desempeño medio</p></div>
              <div className="bg-rose-50 border border-rose-200 rounded-lg py-3"><p className="text-2xl font-semibold text-rose-700">{analisis.semaforo.bajo}</p><p className="text-[11px] text-slate-600">Requiere reforzamiento</p></div>
            </div>
            <p className="text-[11px] text-slate-400 mt-3">Alto desde {db.config.semaforo.alto}%, medio desde {db.config.semaforo.medio}%. {AVISO_ECOEMS}</p>
          </Card>

          <Card>
            <p className="text-sm font-medium text-slate-700 mb-3">Aciertos promedio por área · aplicación {analisis.ultimo}</p>
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={analisis.porMateria} layout="vertical" margin={{ left: 55 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis type="number" fontSize={11} /><YAxis type="category" dataKey="corto" fontSize={9} width={100} />
                <Tooltip formatter={(v, n, o) => [`${v} de ${o.payload.reactivos} aciertos (${o.payload.promedio}%)`, o.payload.nombre]} />
                <Bar dataKey="aciertos" radius={[0, 4, 4, 0]}>
                  {analisis.porMateria.map((m, i) => <Cell key={i} fill={semaforoDe(m.promedio, db.config).color} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </Card>

          <Card>
            <p className="text-sm font-medium text-slate-700 mb-3">Evolución del grupo entre aplicaciones</p>
            <ResponsiveContainer width="100%" height={210}>
              <LineChart data={analisis.evolucion}><CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={11} /><YAxis domain={[0, 100]} fontSize={11} /><Tooltip formatter={(v) => v + "%"} />
                <Line type="monotone" dataKey="promedio" stroke="#0369a1" strokeWidth={2} dot={{ r: 4 }} name="Promedio" /></LineChart>
            </ResponsiveContainer>
          </Card>

          <Card pad={false} className="p-4">
            <p className="text-sm font-medium text-slate-700 mb-1">Orden interno de resultados · aplicación {analisis.ultimo}</p>
            <p className="text-[11px] text-slate-400 mb-3">Solo para seguimiento académico del docente. No es una lista para publicar ni para comparar públicamente al grupo.</p>
            <Tabla cols={["#", "Nombre", "Aciertos", "%", "Semáforo"]}>
              {analisis.filas.map((f, i) => { const sm = semaforoDe(f.porcentaje, db.config);
                return (
                  <tr key={i} className="hover:bg-slate-50 cursor-pointer" onClick={() => f.alumno && ir("ficha", { alumnoId: f.alumno.id })}>
                    <td className="py-2 px-3 text-slate-400">{i + 1}</td><td className="py-2 px-3">{nomComp(f.alumno)}</td>
                    <td className="py-2 px-3">{f.aciertos}/{maxTotal}</td><td className="py-2 px-3 font-medium">{f.porcentaje}%</td>
                    <td className="py-2 px-3"><Pill cls={sm.cls}>{sm.label}</Pill></td>
                  </tr>
                ); })}
            </Tabla>
          </Card>

          <Card>
            <p className="text-sm font-medium text-slate-700 mb-2">Áreas que requieren reforzamiento</p>
            <div className="space-y-1.5">
              {analisis.porMateria.filter((m) => semaforoDe(m.promedio, db.config).nivel !== "alto").sort((a, b) => a.promedio - b.promedio).map((m) => (
                <div key={m.nombre} className={`flex justify-between text-sm rounded-lg px-3 py-2 border ${semaforoDe(m.promedio, db.config).cls}`}>
                  <span>{m.nombre}</span><span className="font-medium">{m.aciertos}/{m.reactivos} · {m.promedio}%</span>
                </div>
              ))}
              {analisis.porMateria.every((m) => semaforoDe(m.promedio, db.config).nivel === "alto") && <p className="text-sm text-emerald-700">Todas las áreas están en desempeño alto.</p>}
            </div>
          </Card>
        </div>
      ) : <Card><Vacio texto="Registra al menos una aplicación para ver el análisis." /></Card>)}

      {tab === "comparativo" && (
        <div className="space-y-4">
          <Card>
            <p className="text-sm font-medium text-slate-700 mb-3">Comparativo entre grupos · última aplicación de cada uno</p>
            {comparativoGrupos.length === 0 ? <Vacio texto="Aún no hay resultados registrados." /> : (
              <>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={comparativoGrupos}><CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={11} /><YAxis domain={[0, 100]} fontSize={11} /><Tooltip formatter={(v) => v + "%"} />
                    <Bar dataKey="promedio" radius={[3, 3, 0, 0]}>{comparativoGrupos.map((c, i) => <Cell key={i} fill={semaforoDe(c.promedio, db.config).color} />)}</Bar></BarChart>
                </ResponsiveContainer>
                <Tabla cols={["Grupo", "Aplicación", "Aciertos promedio", "Porcentaje"]}>
                  {comparativoGrupos.map((c) => <tr key={c.nombre}><td className="py-2 px-3">{c.nombre}</td><td className="py-2 px-3">{c.aplicacion}</td>
                    <td className="py-2 px-3">{c.aciertos}/{maxTotal}</td><td className="py-2 px-3 font-medium">{c.promedio}%</td></tr>)}
                </Tabla>
              </>
            )}
          </Card>
          <Card>
            <p className="text-sm font-medium text-slate-700 mb-3">Comparativo entre aplicaciones del grupo</p>
            {analisis ? (
              <Tabla cols={["Aplicación", "Aciertos promedio", "Porcentaje", "Cambio"]}>
                {analisis.evolucion.map((e, i) => (
                  <tr key={e.nombre}><td className="py-2 px-3">{e.nombre}</td><td className="py-2 px-3">{e.aciertos}/{maxTotal}</td>
                    <td className="py-2 px-3 font-medium">{e.promedio}%</td>
                    <td className={`py-2 px-3 ${i === 0 ? "text-slate-400" : e.promedio - analisis.evolucion[i - 1].promedio >= 0 ? "text-emerald-700" : "text-amber-600"}`}>
                      {i === 0 ? "—" : (e.promedio - analisis.evolucion[i - 1].promedio > 0 ? "+" : "") + round(e.promedio - analisis.evolucion[i - 1].promedio, 1) + "%"}</td></tr>
                ))}
              </Tabla>
            ) : <Vacio texto="Sin datos para comparar." />}
          </Card>
        </div>
      )}
    </div>
  );
}

/* ============================================================
   MÓDULO: INCIDENCIAS
   ============================================================ */
function Incidencias({ db, upd, ir, toast, params }) {
  const [modal, setModal] = useState(null);
  const [detalle, setDetalle] = useState(params?.incId || null);
  const [conf, setConf] = useState(null);
  const [fEstado, setFEstado] = useState("");
  const [fConducta, setFConducta] = useState("");
  const [fGravedad, setFGravedad] = useState("");
  const [nota, setNota] = useState("");

  const nuevo = () => ({
    fecha: hoy(), hora: new Date().toTimeString().slice(0, 5), alumnoId: "", grupoId: db.grupos[0]?.id || "",
    docente: db.config.docente || "", lugar: "", conducta: db.catalogos.conductas[0], gravedad: "Leve",
    descripcion: "", involucrados: "", testigos: "", evidencias: "", medidas: [], acuerdos: "",
    accionInmediata: "", canalizacion: "", fechaSeguimiento: "", observaciones: "", estado: "Abierta", seguimientos: [],
  });

  const guardar = () => {
    if (!modal.alumnoId || !modal.descripcion) return toast("Falta el estudiante o la descripción de los hechos", "error");
    upd((d) => {
      const dato = { ...modal, tipo: modal.conducta };
      if (modal.id) { const i = d.incidencias.findIndex((x) => x.id === modal.id); d.incidencias[i] = dato; }
      else {
        d.folio = (d.folio || 0) + 1;
        d.incidencias.push({ ...dato, id: uid("inc"), folio: `INC-${d.ciclo.slice(0, 4)}-${String(d.folio).padStart(4, "0")}` });
      }
    });
    setModal(null); toast("Incidencia registrada");
  };

  const agregarSeg = (incId) => {
    if (!nota.trim()) return;
    upd((d) => { const i = d.incidencias.find((x) => x.id === incId); i.seguimientos = [...(i.seguimientos || []), { id: uid("seg"), fecha: hoy(), nota }]; if (i.estado === "Abierta") i.estado = "En seguimiento"; });
    setNota(""); toast("Seguimiento agregado");
  };
  const cambiarEstado = (incId, estado) => upd((d) => { const i = d.incidencias.find((x) => x.id === incId); i.estado = estado; });

  const toggleMedida = (m) => setModal((x) => ({ ...x, medidas: (x.medidas || []).includes(m) ? x.medidas.filter((y) => y !== m) : [...(x.medidas || []), m] }));

  const lista = db.incidencias
    .filter((i) => (!fEstado || i.estado === fEstado) && (!fConducta || (i.conducta || i.tipo) === fConducta) && (!fGravedad || i.gravedad === fGravedad))
    .sort((a, b) => b.fecha.localeCompare(a.fecha));
  const inc = db.incidencias.find((i) => i.id === detalle);

  const formatoImpreso = (i) => {
    const a = db.alumnos.find((x) => x.id === i.alumnoId);
    const g = db.grupos.find((x) => x.id === i.grupoId);
    pdfIncidencia({
      incidencia: i, config: db.config, ciclo: db.ciclo,
      alumno: nomComp(a),
      grupo: g ? `${g.grado}° ${g.grupo} · ${g.turno}` : "—",
      avisoLegal: AVISO_CONVIVENCIA,
    });
  };

  return (
    <div className="space-y-4">
      <Titulo sub="Marco para la convivencia escolar. Registro de hechos objetivos, sin calificativos sobre las personas."
        right={<Btn icon={Plus} onClick={() => setModal(nuevo())} disabled={!db.alumnos.length}>Nueva incidencia</Btn>}>Incidencias</Titulo>

      <div className="flex gap-2 flex-wrap">
        <Sel value={fEstado} onChange={(e) => setFEstado(e.target.value)} className="max-w-[170px]"><option value="">Todos los estatus</option>{EST_INC.map((s2) => <option key={s2}>{s2}</option>)}</Sel>
        <Sel value={fGravedad} onChange={(e) => setFGravedad(e.target.value)} className="max-w-[180px]"><option value="">Toda clasificación</option>{Object.keys(GRAVEDAD).map((s2) => <option key={s2}>{s2}</option>)}</Sel>
        <Sel value={fConducta} onChange={(e) => setFConducta(e.target.value)} className="max-w-[240px]"><option value="">Todas las conductas</option>{db.catalogos.conductas.map((s2) => <option key={s2}>{s2}</option>)}</Sel>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <Stat label="Total" valor={db.incidencias.length} />
        <Stat label="Abiertas" valor={db.incidencias.filter((i) => i.estado === "Abierta").length} color="text-rose-600" />
        <Stat label="En seguimiento" valor={db.incidencias.filter((i) => i.estado === "En seguimiento").length} color="text-amber-600" />
        <Stat label="Cerradas" valor={db.incidencias.filter((i) => i.estado === "Cerrada").length} color="text-emerald-700" />
      </div>

      {lista.length === 0 ? <Card><Vacio texto="No hay incidencias con estos filtros." /></Card> : (
        <div className="space-y-2">
          {lista.map((i) => {
            const a = db.alumnos.find((x) => x.id === i.alumnoId);
            const g = db.grupos.find((x) => x.id === i.grupoId);
            const gr = GRAVEDAD[i.gravedad] || GRAVEDAD.Leve;
            return (
              <Card key={i.id}>
                <div className="flex justify-between items-start gap-3 flex-wrap cursor-pointer" onClick={() => setDetalle(i.id)}>
                  <div className="grow min-w-[180px]">
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="font-mono text-xs text-slate-500">{i.folio}</span>
                      <Pill cls={i.estado === "Cerrada" ? "bg-slate-100 text-slate-600 border-slate-200" : i.estado === "En seguimiento" ? "bg-amber-50 text-amber-700 border-amber-200" : "bg-rose-50 text-rose-700 border-rose-200"}>{i.estado}</Pill>
                      <Pill cls={gr.cls}>{i.gravedad || "Leve"}</Pill>
                    </div>
                    <p className="font-medium text-slate-900 mt-1">{nomComp(a)} · {g ? `${g.grado}° ${g.grupo}` : ""}</p>
                    <p className="text-xs text-slate-500">{fFecha(i.fecha)} {i.hora} · {i.conducta || i.tipo} · {i.lugar}</p>
                    <p className="text-sm text-slate-600 mt-1">{i.descripcion}</p>
                    {(i.medidas || []).length > 0 && <p className="text-[11px] text-slate-400 mt-1">Medidas: {(i.medidas || []).join(" · ")}</p>}
                  </div>
                  <div className="flex gap-1 shrink-0">
                    <button onClick={(e) => { e.stopPropagation(); formatoImpreso(i); }} title="Descargar el formato en PDF" className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-500"><FileDown size={15} /></button>
                    <button onClick={(e) => { e.stopPropagation(); setModal(i); }} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-500"><Pencil size={15} /></button>
                    <button onClick={(e) => { e.stopPropagation(); setConf({ texto: `¿Eliminar la incidencia ${i.folio}? También se elimina su valoración.`, onSi: () => {
                      eliminarConRespaldo({ tipo: "Incidencia", descripcion: `${i.folio} · ${nomComp(db.alumnos.find((x) => x.id === i.alumnoId))}`, ciclo: db.ciclo, datos: { incidencias: [i], valoraciones: db.valoraciones.filter((x) => x.incidenciaId === i.id) } });
                      upd((d) => { d.incidencias = d.incidencias.filter((x) => x.id !== i.id); d.valoraciones = d.valoraciones.filter((x) => x.incidenciaId !== i.id); });
                      setConf(null); toast("Enviada a la papelera");
                    } }); }} className="p-1.5 rounded-lg hover:bg-rose-50 text-rose-500"><Trash2 size={15} /></button>
                  </div>
                </div>
              </Card>
            );
          })}
        </div>
      )}

      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.id ? `Editar ${modal.folio}` : "Nueva incidencia"}
        footer={<><Btn tipo="secundario" onClick={() => setModal(null)}>Cancelar</Btn><Btn icon={Save} onClick={guardar}>Guardar</Btn></>}>
        {modal && (
          <div className="space-y-4">
            <div className="grid sm:grid-cols-2 gap-3">
              <Campo label="Fecha"><Inp type="date" value={modal.fecha} onChange={(e) => setModal({ ...modal, fecha: e.target.value })} /></Campo>
              <Campo label="Hora"><Inp type="time" value={modal.hora} onChange={(e) => setModal({ ...modal, hora: e.target.value })} /></Campo>
              <Campo label="Grupo"><Sel value={modal.grupoId} onChange={(e) => setModal({ ...modal, grupoId: e.target.value, alumnoId: "" })}>{db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}</Sel></Campo>
              <Campo label="Alumna(o)" req><Sel value={modal.alumnoId} onChange={(e) => setModal({ ...modal, alumnoId: e.target.value })}><option value="">Selecciona</option>{sortAl(db.alumnos.filter((a) => a.grupoId === modal.grupoId && a.activo)).map((a) => <option key={a.id} value={a.id}>{a.numLista}. {nomComp(a)}</option>)}</Sel></Campo>
              <Campo label="Docente que reporta"><Inp value={modal.docente} onChange={(e) => setModal({ ...modal, docente: e.target.value })} /></Campo>
              <Campo label="Lugar"><Inp value={modal.lugar} onChange={(e) => setModal({ ...modal, lugar: e.target.value })} /></Campo>
            </div>

            <div className="grid sm:grid-cols-2 gap-3">
              <div className="sm:col-span-2"><Campo label="Tipo de conducta" hint="Catálogo configurable en Configuración.">
                <Sel value={modal.conducta} onChange={(e) => setModal({ ...modal, conducta: e.target.value })}>{db.catalogos.conductas.map((c) => <option key={c}>{c}</option>)}</Sel></Campo></div>
              <div className="sm:col-span-2">
                <p className="text-xs font-medium text-slate-600 mb-1">Clasificación</p>
                <div className="grid grid-cols-2 gap-2">
                  {Object.entries(GRAVEDAD).map(([k, v]) => (
                    <button key={k} onClick={() => setModal({ ...modal, gravedad: k })}
                      className={`flex items-center gap-2 px-3 py-2 rounded-lg border text-sm text-left ${modal.gravedad === k ? v.cls + " ring-2 ring-slate-300" : "bg-white text-slate-600 border-slate-200"}`}>
                      <span className={`w-2.5 h-2.5 rounded-full shrink-0 ${v.punto}`} />{k}
                    </button>
                  ))}
                </div>
              </div>
            </div>

            <Campo label="Descripción objetiva de los hechos" req hint="Qué pasó, cuándo y dónde. Sin adjetivos ni interpretaciones sobre la persona.">
              <Area value={modal.descripcion} onChange={(e) => setModal({ ...modal, descripcion: e.target.value })} /></Campo>
            <div className="grid sm:grid-cols-2 gap-3">
              <Campo label="Personas involucradas"><Inp value={modal.involucrados} onChange={(e) => setModal({ ...modal, involucrados: e.target.value })} /></Campo>
              <Campo label="Testigos, cuando corresponda"><Inp value={modal.testigos} onChange={(e) => setModal({ ...modal, testigos: e.target.value })} /></Campo>
            </div>
            <Campo label="Evidencias o documentos relacionados" hint="Indica dónde se resguardan (carpeta, expediente, fotografía en el archivo escolar).">
              <Inp value={modal.evidencias} onChange={(e) => setModal({ ...modal, evidencias: e.target.value })} /></Campo>

            <div>
              <p className="text-xs font-medium text-slate-600 mb-1">Medidas de atención realizadas</p>
              <div className="flex flex-wrap gap-1.5">
                {MEDIDAS_ATENCION.map((m) => (
                  <button key={m} onClick={() => toggleMedida(m)}
                    className={`px-2.5 py-1.5 rounded-lg text-xs border text-left ${(modal.medidas || []).includes(m) ? "bg-slate-800 text-white border-slate-800" : "bg-white text-slate-600 border-slate-200 hover:border-slate-400"}`}>{m}</button>
                ))}
              </div>
            </div>
            <Campo label="Detalle de la atención"><Area value={modal.accionInmediata} onChange={(e) => setModal({ ...modal, accionInmediata: e.target.value })} /></Campo>
            <Campo label="Acuerdos"><Area value={modal.acuerdos} onChange={(e) => setModal({ ...modal, acuerdos: e.target.value })} /></Campo>
            <div className="grid sm:grid-cols-3 gap-3">
              <Campo label="Canalización"><Inp value={modal.canalizacion} onChange={(e) => setModal({ ...modal, canalizacion: e.target.value })} placeholder="Orientación, prefectura…" /></Campo>
              <Campo label="Fecha de seguimiento"><Inp type="date" value={modal.fechaSeguimiento} onChange={(e) => setModal({ ...modal, fechaSeguimiento: e.target.value })} /></Campo>
              <Campo label="Estatus"><Sel value={modal.estado} onChange={(e) => setModal({ ...modal, estado: e.target.value })}>{EST_INC.map((s2) => <option key={s2}>{s2}</option>)}</Sel></Campo>
            </div>
            <Campo label="Observaciones"><Area value={modal.observaciones} onChange={(e) => setModal({ ...modal, observaciones: e.target.value })} /></Campo>
            <Aviso>{AVISO_CONVIVENCIA}</Aviso>
          </div>
        )}
      </Modal>

      <Modal open={!!inc} onClose={() => setDetalle(null)} title={inc?.folio} ancho="max-w-2xl"
        footer={<><Btn tipo="secundario" icon={FileDown} onClick={() => formatoImpreso(inc)}>Descargar formato</Btn>
          <Btn onClick={() => { setDetalle(null); ir("valoracion", { incId: inc.id }); }}>Valoración</Btn></>}>
        {inc && (
          <div className="space-y-3 text-sm">
            <div className="grid grid-cols-2 gap-2 text-slate-600">
              <p><span className="text-slate-400">Alumna(o): </span>{nomComp(db.alumnos.find((x) => x.id === inc.alumnoId))}</p>
              <p><span className="text-slate-400">Fecha: </span>{fFecha(inc.fecha)} {inc.hora}</p>
              <p><span className="text-slate-400">Conducta: </span>{inc.conducta || inc.tipo}</p>
              <p><span className="text-slate-400">Clasificación: </span>{inc.gravedad || "—"}</p>
              <p><span className="text-slate-400">Lugar: </span>{inc.lugar || "—"}</p>
              <p><span className="text-slate-400">Seguimiento: </span>{inc.fechaSeguimiento ? fFecha(inc.fechaSeguimiento) : "—"}</p>
            </div>
            <div><p className="text-xs text-slate-400 mb-1">Hechos</p><p className="bg-slate-50 border border-slate-200 rounded-lg p-3">{inc.descripcion}</p></div>
            {(inc.medidas || []).length > 0 && <div><p className="text-xs text-slate-400 mb-1">Medidas de atención</p>
              <div className="flex flex-wrap gap-1.5">{inc.medidas.map((m) => <Pill key={m}>{m}</Pill>)}</div></div>}
            {inc.acuerdos && <div><p className="text-xs text-slate-400 mb-1">Acuerdos</p><p className="bg-slate-50 border border-slate-200 rounded-lg p-3">{inc.acuerdos}</p></div>}
            <div>
              <p className="text-xs text-slate-400 mb-1">Seguimientos ({inc.seguimientos?.length || 0})</p>
              <div className="space-y-1.5 mb-2">
                {(inc.seguimientos || []).map((x) => <div key={x.id} className="text-sm border-l-2 border-emerald-300 pl-3 py-1"><span className="text-xs text-slate-400">{fFecha(x.fecha)}</span><p>{x.nota}</p></div>)}
              </div>
              <div className="flex gap-2">
                <input value={nota} onChange={(e) => setNota(e.target.value)} placeholder="Registrar acción realizada" className={inputCls} />
                <Btn size="sm" onClick={() => agregarSeg(inc.id)}>Agregar</Btn>
              </div>
            </div>
            <div className="flex gap-2 flex-wrap items-center pt-2 border-t border-slate-200">
              <span className="text-xs text-slate-500">Cambiar estatus:</span>
              {EST_INC.map((e2) => <Btn key={e2} size="sm" tipo={inc.estado === e2 ? "primario" : "secundario"} onClick={() => cambiarEstado(inc.id, e2)}>{e2}</Btn>)}
            </div>
            <Aviso>{AVISO_CONVIVENCIA}</Aviso>
          </div>
        )}
      </Modal>
      <Confirmar open={!!conf} texto={conf?.texto} onSi={conf?.onSi} onNo={() => setConf(null)} />
    </div>
  );
}

/* ============================================================
   MÓDULO: VALORACIÓN DEL TIPO DE FALTA
   ============================================================ */
const PREGUNTAS = [
  ["afectacionPersonas", "¿Hubo afectación a otras personas?"],
  ["afectacionMateriales", "¿Hubo afectación a materiales?"],
  ["intencional", "¿La situación fue intencional?"],
  ["recurrente", "¿Es recurrente?"],
  ["versionAlumno", "¿Se escuchó la versión del alumno?"],
  ["versionesCorrespondientes", "¿Se escucharon las versiones correspondientes?"],
];

function Valoracion({ db, upd, ir, toast, params }) {
  const [incId, setIncId] = useState(params?.incId || db.incidencias[0]?.id || "");
  const inc = db.incidencias.find((i) => i.id === incId);
  const [f, setF] = useState({ incidenciaId: incId, tipo: "Convivencia", circ: {}, nivel: "verde", medida: "", acuerdo: "", comunicacion: "", canalizacion: "", fechaSeguimiento: "", resultado: "" });
  useEffect(() => {
    const v = db.valoraciones.find((x) => x.incidenciaId === incId);
    setF(v ? clone(v) : { incidenciaId: incId, tipo: "Convivencia", circ: {}, nivel: "verde", medida: "", acuerdo: "", comunicacion: "", canalizacion: "", fechaSeguimiento: "", resultado: "" });
  }, [incId, db.valoraciones.length]);

  const guardar = () => {
    if (!incId) return toast("Selecciona una incidencia", "error");
    upd((d) => {
      const i = d.valoraciones.findIndex((v) => v.incidenciaId === incId);
      if (i >= 0) d.valoraciones[i] = { ...d.valoraciones[i], ...f };
      else d.valoraciones.push({ ...f, id: uid("val"), incidenciaId: incId });
    });
    toast("Valoración guardada");
  };

  const alumno = inc ? db.alumnos.find((a) => a.id === inc.alumnoId) : null;

  return (
    <div className="space-y-4">
      <Titulo sub="Registro administrativo vinculado al folio de la incidencia.">Valoración de la falta</Titulo>
      {db.incidencias.length === 0 ? <Card><Vacio texto="Primero registra una incidencia." accion={<Btn onClick={() => ir("incidencias")}>Ir a incidencias</Btn>} /></Card> : (
        <>
          <Card>
            <Campo label="Folio de incidencia">
              <Sel value={incId} onChange={(e) => setIncId(e.target.value)}>
                {db.incidencias.map((i) => { const a = db.alumnos.find((x) => x.id === i.alumnoId); return <option key={i.id} value={i.id}>{i.folio} · {nomComp(a)} · {i.tipo}</option>; })}
              </Sel>
            </Campo>
            {inc && (
              <div className="mt-3 bg-slate-50 border border-slate-200 rounded-lg p-3 text-sm">
                <p className="text-slate-800 font-medium">{nomComp(alumno)}</p>
                <p className="text-xs text-slate-500">{fFecha(inc.fecha)} {inc.hora} · {inc.tipo} · {inc.lugar}</p>
                <p className="text-slate-600 mt-1">{inc.descripcion}</p>
              </div>
            )}
          </Card>

          <Card>
            <Campo label="Tipo de situación">
              <Sel value={f.tipo} onChange={(e) => setF({ ...f, tipo: e.target.value })}>
                {["Académica", "Convivencia", "Normativa escolar", "Situación que requiere atención específica"].map((t) => <option key={t}>{t}</option>)}
              </Sel>
            </Campo>
            <div className="mt-4 space-y-2">
              <p className="text-xs font-medium text-slate-600">Circunstancias</p>
              {PREGUNTAS.map(([k, q]) => (
                <div key={k} className="flex items-center justify-between gap-3 flex-wrap py-1.5 border-b border-slate-100 last:border-0">
                  <span className="text-sm text-slate-700 grow min-w-[180px]">{q}</span>
                  <div className="flex gap-1 flex-wrap">
                    {CIRC_OPC.map((o) => (
                      <button key={o} onClick={() => setF({ ...f, circ: { ...f.circ, [k]: o } })}
                        className={`px-2.5 py-1 rounded-lg text-xs border ${f.circ?.[k] === o ? "bg-slate-800 text-white border-slate-800" : "bg-white text-slate-500 border-slate-200 hover:border-slate-400"}`}>{o}</button>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </Card>

          <Card>
            <p className="text-xs font-medium text-slate-600 mb-2">Nivel de seguimiento</p>
            <div className="grid sm:grid-cols-2 gap-2">
              {Object.entries(NIVELES).map(([k, n]) => (
                <button key={k} onClick={() => setF({ ...f, nivel: k })}
                  className={`flex items-center gap-2 text-left px-3 py-2.5 rounded-lg border text-sm ${f.nivel === k ? n.cls + " ring-2 ring-slate-300" : "bg-white text-slate-600 border-slate-200"}`}>
                  <span className={`w-2.5 h-2.5 rounded-full shrink-0 ${n.punto}`} />{n.label}
                </button>
              ))}
            </div>
            <div className="mt-3"><Aviso>{AVISO_LEGAL} El sistema no genera sanciones de manera automática.</Aviso></div>
          </Card>

          <Card>
            <div className="grid sm:grid-cols-2 gap-3">
              <div className="sm:col-span-2"><Campo label="Medida educativa"><Area value={f.medida} onChange={(e) => setF({ ...f, medida: e.target.value })} /></Campo></div>
              <div className="sm:col-span-2"><Campo label="Acuerdo"><Area value={f.acuerdo} onChange={(e) => setF({ ...f, acuerdo: e.target.value })} /></Campo></div>
              <Campo label="Comunicación con madre, padre o tutor"><Inp value={f.comunicacion} onChange={(e) => setF({ ...f, comunicacion: e.target.value })} /></Campo>
              <Campo label="Canalización"><Inp value={f.canalizacion} onChange={(e) => setF({ ...f, canalizacion: e.target.value })} /></Campo>
              <Campo label="Fecha de seguimiento"><Inp type="date" value={f.fechaSeguimiento} onChange={(e) => setF({ ...f, fechaSeguimiento: e.target.value })} /></Campo>
              <Campo label="Resultado"><Inp value={f.resultado} onChange={(e) => setF({ ...f, resultado: e.target.value })} /></Campo>
            </div>
            <div className="flex justify-end mt-4"><Btn icon={Save} onClick={guardar}>Guardar valoración</Btn></div>
          </Card>
        </>
      )}
    </div>
  );
}

/* ============================================================
   MÓDULO: BITÁCORA DIGITAL DE CLASE
   ============================================================ */
function Bitacora({ db, upd, ir, toast }) {
  const [modal, setModal] = useState(null);
  const [conf, setConf] = useState(null);
  const [grupoId, setGrupoId] = useState(db.grupos[0]?.id || "");
  const [seccion, setSeccion] = useState("datos");

  const nueva = () => ({
    grupoId, fecha: hoy(), trimestre: trimestreDe(hoy(), db.config), sesion: db.bitacoras.filter((b) => b.grupoId === grupoId).length + 1, horario: db.config.horario || "",
    campoFormativo: "", contenido: "", pda: "", eje: "", proposito: "",
    inicioAct: "", inicioPregunta: "", inicioPrevios: "", desarrollo: "", estrategias: "", individual: "", colaborativo: "", recursos: "",
    cierreProducto: "", cierreEvidencia: "", cierreReflexion: "", cierreTarea: "",
    evalEvidencia: "", evalInstrumento: "", logros: "", dificultades: "", apoyo: "",
    situaciones: "", acuerdos: "", observaciones: "", reflexionLogro: "Sí", reflexionFunciono: "", reflexionDificultades: "", reflexionModificar: "", estado: "Concluida",
  });

  const conteoSesion = (gid, f) => {
    const regs = db.asistencias.filter((x) => x.grupoId === gid && x.fecha === f).flatMap((x) => Object.entries(x.marcas || {}));
    if (!regs.length) return null;
    const vistos = {};
    regs.forEach(([al, m]) => { vistos[al] = m; });
    const vals = Object.values(vistos);
    return { presentes: vals.filter((m) => m === "A" || m === "R" || m === "P").length, faltas: vals.filter((m) => m === "F" || m === "J").length, retardos: vals.filter((m) => m === "R").length };
  };
  const requierenSeguimiento = (gid) => sortAl(db.alumnos.filter((a) => a.grupoId === gid && a.activo))
    .filter((a) => { const st = statsAsistencia(db, a.id, { grupoId: gid }); const pr = promedioAcumulado(db, a.id); const ac = statsActividades(db, a.id);
      return (st.total >= 5 && st.porcentaje < db.config.alertas.asistenciaMin) || (pr > 0 && pr < db.config.alertas.promedioMin) || ac.noEntregadas >= db.config.alertas.actPendientes; });

  const guardar = () => {
    upd((d) => {
      if (modal.id) { const i = d.bitacoras.findIndex((b) => b.id === modal.id); d.bitacoras[i] = { ...modal, trimestre: Number(modal.trimestre), sesion: Number(modal.sesion) }; }
      else d.bitacoras.push({ ...modal, id: uid("bit"), trimestre: Number(modal.trimestre), sesion: Number(modal.sesion) });
    });
    setModal(null); toast("Bitácora guardada");
  };

  const lista = db.bitacoras.filter((b) => b.grupoId === grupoId).sort((a, b) => b.fecha.localeCompare(a.fecha));
  const T = ({ id, children }) => <button onClick={() => setSeccion(id)} className={`px-3 py-1.5 rounded-lg text-xs font-medium border whitespace-nowrap ${seccion === id ? "bg-emerald-700 text-white border-emerald-700" : "bg-white text-slate-600 border-slate-200"}`}>{children}</button>;

  return (
    <div className="space-y-4">
      <Titulo sub="Registro de cada sesión: planeación, desarrollo, evaluación y reflexión docente."
        right={<Btn icon={Plus} onClick={() => { setModal(nueva()); setSeccion("datos"); }} disabled={!db.grupos.length}>Nueva sesión</Btn>}>Bitácora de clase</Titulo>
      <Sel value={grupoId} onChange={(e) => setGrupoId(e.target.value)} className="max-w-[170px]">{db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}</Sel>

      {lista.length === 0 ? <Card><Vacio texto="Aún no hay sesiones registradas para este grupo." /></Card> : (
        <div className="space-y-2">
          {lista.map((b) => (
            <Card key={b.id}>
              <div className="flex justify-between items-start gap-3 flex-wrap">
                <div className="grow min-w-[200px]">
                  <div className="flex items-center gap-2 flex-wrap">
                    <p className="font-medium text-slate-900">Sesión {b.sesion} · {fFecha(b.fecha)}</p>
                    <Pill cls={b.estado === "Concluida" ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-amber-50 text-amber-700 border-amber-200"}>{b.estado}</Pill>
                    <Pill>T{b.trimestre}</Pill>
                  </div>
                  <p className="text-sm text-slate-600 mt-1">{b.contenido || "Sin contenido registrado"}</p>
                  <p className="text-xs text-slate-400">{b.campoFormativo}</p>
                  {(() => { const c = conteoSesion(b.grupoId, b.fecha); return c ? <p className="text-[11px] text-slate-500 mt-1">Presentes {c.presentes} · faltas {c.faltas} · retardos {c.retardos}</p> : <p className="text-[11px] text-slate-400 mt-1">Sin pase de lista de ese día</p>; })()}
                  {b.situaciones && <p className="text-xs text-amber-700 mt-1.5 bg-amber-50 border border-amber-200 rounded px-2 py-1">Situación relevante: {b.situaciones}</p>}
                </div>
                <div className="flex gap-1 shrink-0">
                  {b.situaciones && <Btn size="sm" tipo="secundario" onClick={() => ir("incidencias", { desdeBitacora: b.situaciones })}>A incidencia</Btn>}
                  <button onClick={() => { setModal(b); setSeccion("datos"); }} className="p-1.5 rounded-lg hover:bg-slate-100 text-slate-500"><Pencil size={15} /></button>
                  <button onClick={() => setConf({ texto: `¿Eliminar la bitácora de la sesión ${b.sesion}?`, onSi: () => {
                    eliminarConRespaldo({ tipo: "Bitácora", descripcion: `Sesión ${b.sesion} · ${fFecha(b.fecha)}`, ciclo: db.ciclo, datos: { bitacoras: [b] } });
                    upd((d) => { d.bitacoras = d.bitacoras.filter((x) => x.id !== b.id); });
                    setConf(null); toast("Enviada a la papelera");
                  } })} className="p-1.5 rounded-lg hover:bg-rose-50 text-rose-500"><Trash2 size={15} /></button>
                </div>
              </div>
            </Card>
          ))}
        </div>
      )}

      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.id ? "Editar sesión" : "Nueva sesión"} ancho="max-w-2xl"
        footer={<><Btn tipo="secundario" onClick={() => setModal(null)}>Cancelar</Btn><Btn icon={Save} onClick={guardar}>Guardar sesión</Btn></>}>
        {modal && (
          <div className="space-y-4">
            <div className="flex gap-1 overflow-x-auto pb-1">
              <T id="datos">Datos</T><T id="planeacion">Planeación</T><T id="inicio">Inicio</T><T id="desarrollo">Desarrollo</T><T id="cierre">Cierre</T><T id="evaluacion">Evaluación</T><T id="reflexion">Reflexión</T>
            </div>
            {seccion === "datos" && (
              <div className="grid sm:grid-cols-2 gap-3">
                <Campo label="Grupo"><Sel value={modal.grupoId} onChange={(e) => setModal({ ...modal, grupoId: e.target.value })}>{db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}</Sel></Campo>
                <Campo label="Fecha"><Inp type="date" value={modal.fecha} onChange={(e) => setModal({ ...modal, fecha: e.target.value, trimestre: trimestreDe(e.target.value, db.config) })} /></Campo>
                <Campo label="Número de sesión"><Inp type="number" value={modal.sesion} onChange={(e) => setModal({ ...modal, sesion: e.target.value })} /></Campo>
                <Campo label="Horario"><Inp value={modal.horario} onChange={(e) => setModal({ ...modal, horario: e.target.value })} placeholder="08:00 - 08:50" /></Campo>
                <Campo label="Trimestre"><Sel value={modal.trimestre} onChange={(e) => setModal({ ...modal, trimestre: e.target.value })}>{[1, 2, 3].map((t) => <option key={t} value={t}>Trimestre {t}</option>)}</Sel></Campo>
                <Campo label="Estado de la sesión"><Sel value={modal.estado} onChange={(e) => setModal({ ...modal, estado: e.target.value })}>{EST_BIT.map((s) => <option key={s}>{s}</option>)}</Sel></Campo>
                <div className="sm:col-span-2">
                  {(() => { const c = conteoSesion(modal.grupoId, modal.fecha);
                    return (
                      <div className="bg-slate-50 border border-slate-200 rounded-lg p-3">
                        <p className="text-xs font-medium text-slate-600 mb-1">Asistencia de la sesión</p>
                        {c ? <p className="text-sm text-slate-700">Presentes: <strong>{c.presentes}</strong> · Faltas: <strong>{c.faltas}</strong> · Retardos: <strong>{c.retardos}</strong></p>
                           : <p className="text-sm text-slate-500">No hay pase de lista de esta fecha. Se toma del módulo de Asistencia en cuanto lo registres.</p>}
                      </div>
                    ); })()}
                </div>
                <div className="sm:col-span-2">
                  {(() => { const r = requierenSeguimiento(modal.grupoId);
                    return (
                      <div className="border border-slate-200 rounded-lg p-3">
                        <p className="text-xs font-medium text-slate-600 mb-1">Alumnos que requieren seguimiento ({r.length})</p>
                        {r.length === 0 ? <p className="text-sm text-slate-500">Ninguno según los criterios configurados.</p>
                          : <div className="flex flex-wrap gap-1.5">{r.map((a) => <Pill key={a.id}>{nomComp(a)}</Pill>)}</div>}
                      </div>
                    ); })()}
                </div>
                <div className="sm:col-span-2"><Campo label="Situaciones relevantes" hint="Incidencias, retardos, problemas técnicos, suspensión parcial o situaciones extraordinarias.">
                  <Area value={modal.situaciones} onChange={(e) => setModal({ ...modal, situaciones: e.target.value })} /></Campo></div>
                <div className="sm:col-span-2"><Campo label="Acuerdos"><Area value={modal.acuerdos || ""} onChange={(e) => setModal({ ...modal, acuerdos: e.target.value })} /></Campo></div>
                <div className="sm:col-span-2"><Campo label="Observaciones"><Area value={modal.observaciones || ""} onChange={(e) => setModal({ ...modal, observaciones: e.target.value })} /></Campo></div>
              </div>
            )}
            {seccion === "planeacion" && (
              <div className="grid sm:grid-cols-2 gap-3">
                <Campo label="Campo formativo"><Sel value={modal.campoFormativo} onChange={(e) => setModal({ ...modal, campoFormativo: e.target.value })}><option value="">Sin especificar</option>{CAMPOS_FORM.map((c) => <option key={c}>{c}</option>)}</Sel></Campo>
                <Campo label="Eje articulador"><Sel value={modal.eje} onChange={(e) => setModal({ ...modal, eje: e.target.value })}><option value="">Sin especificar</option>{EJES.map((c) => <option key={c}>{c}</option>)}</Sel></Campo>
                <div className="sm:col-span-2"><Campo label="Contenido"><Inp value={modal.contenido} onChange={(e) => setModal({ ...modal, contenido: e.target.value })} /></Campo></div>
                <div className="sm:col-span-2"><Campo label="PDA"><Area value={modal.pda} onChange={(e) => setModal({ ...modal, pda: e.target.value })} /></Campo></div>
                <div className="sm:col-span-2"><Campo label="Propósito"><Area value={modal.proposito} onChange={(e) => setModal({ ...modal, proposito: e.target.value })} /></Campo></div>
              </div>
            )}
            {seccion === "inicio" && (
              <div className="space-y-3">
                <Campo label="Actividad de inicio"><Area value={modal.inicioAct} onChange={(e) => setModal({ ...modal, inicioAct: e.target.value })} /></Campo>
                <Campo label="Pregunta detonadora"><Inp value={modal.inicioPregunta} onChange={(e) => setModal({ ...modal, inicioPregunta: e.target.value })} /></Campo>
                <Campo label="Recuperación de conocimientos previos"><Area value={modal.inicioPrevios} onChange={(e) => setModal({ ...modal, inicioPrevios: e.target.value })} /></Campo>
              </div>
            )}
            {seccion === "desarrollo" && (
              <div className="space-y-3">
                <Campo label="Actividades"><Area value={modal.desarrollo} onChange={(e) => setModal({ ...modal, desarrollo: e.target.value })} /></Campo>
                <Campo label="Estrategias"><Area value={modal.estrategias} onChange={(e) => setModal({ ...modal, estrategias: e.target.value })} /></Campo>
                <div className="grid sm:grid-cols-2 gap-3">
                  <Campo label="Trabajo individual"><Area value={modal.individual} onChange={(e) => setModal({ ...modal, individual: e.target.value })} /></Campo>
                  <Campo label="Trabajo colaborativo"><Area value={modal.colaborativo} onChange={(e) => setModal({ ...modal, colaborativo: e.target.value })} /></Campo>
                </div>
                <Campo label="Recursos"><Inp value={modal.recursos} onChange={(e) => setModal({ ...modal, recursos: e.target.value })} /></Campo>
              </div>
            )}
            {seccion === "cierre" && (
              <div className="grid sm:grid-cols-2 gap-3">
                <Campo label="Producto"><Inp value={modal.cierreProducto} onChange={(e) => setModal({ ...modal, cierreProducto: e.target.value })} /></Campo>
                <Campo label="Evidencia"><Inp value={modal.cierreEvidencia} onChange={(e) => setModal({ ...modal, cierreEvidencia: e.target.value })} /></Campo>
                <div className="sm:col-span-2"><Campo label="Reflexión del grupo"><Area value={modal.cierreReflexion} onChange={(e) => setModal({ ...modal, cierreReflexion: e.target.value })} /></Campo></div>
                <div className="sm:col-span-2"><Campo label="Tarea"><Area value={modal.cierreTarea} onChange={(e) => setModal({ ...modal, cierreTarea: e.target.value })} /></Campo></div>
              </div>
            )}
            {seccion === "evaluacion" && (
              <div className="grid sm:grid-cols-2 gap-3">
                <Campo label="Evidencia"><Inp value={modal.evalEvidencia} onChange={(e) => setModal({ ...modal, evalEvidencia: e.target.value })} /></Campo>
                <Campo label="Instrumento"><Inp value={modal.evalInstrumento} onChange={(e) => setModal({ ...modal, evalInstrumento: e.target.value })} /></Campo>
                <div className="sm:col-span-2"><Campo label="Logros"><Area value={modal.logros} onChange={(e) => setModal({ ...modal, logros: e.target.value })} /></Campo></div>
                <div className="sm:col-span-2"><Campo label="Dificultades"><Area value={modal.dificultades} onChange={(e) => setModal({ ...modal, dificultades: e.target.value })} /></Campo></div>
                <div className="sm:col-span-2"><Campo label="Estudiantes que requieren apoyo" hint="Registra el acompañamiento necesario, sin calificativos."><Area value={modal.apoyo} onChange={(e) => setModal({ ...modal, apoyo: e.target.value })} /></Campo></div>
              </div>
            )}
            {seccion === "reflexion" && (
              <div className="space-y-3">
                <Campo label="¿Se logró el propósito?"><Sel value={modal.reflexionLogro} onChange={(e) => setModal({ ...modal, reflexionLogro: e.target.value })}><option>Sí</option><option>Parcialmente</option><option>No</option></Sel></Campo>
                <Campo label="¿Qué funcionó?"><Area value={modal.reflexionFunciono} onChange={(e) => setModal({ ...modal, reflexionFunciono: e.target.value })} /></Campo>
                <Campo label="¿Qué dificultades hubo?"><Area value={modal.reflexionDificultades} onChange={(e) => setModal({ ...modal, reflexionDificultades: e.target.value })} /></Campo>
                <Campo label="¿Qué se debe modificar?"><Area value={modal.reflexionModificar} onChange={(e) => setModal({ ...modal, reflexionModificar: e.target.value })} /></Campo>
              </div>
            )}
          </div>
        )}
      </Modal>
      <Confirmar open={!!conf} texto={conf?.texto} onSi={conf?.onSi} onNo={() => setConf(null)} />
    </div>
  );
}

/* ============================================================
   TABLERO POR GRUPO
   ============================================================ */
function TableroGrupo({ db, ir, params }) {
  const [grupoId, setGrupoId] = useState(params?.grupoId || db.grupos[0]?.id || "");
  const g = db.grupos.find((x) => x.id === grupoId);
  const alumnos = sortAl(db.alumnos.filter((a) => a.grupoId === grupoId && a.activo));
  if (!g) return <Card><Vacio texto="Selecciona un grupo." accion={<Btn onClick={() => ir("grupos")}>Ir a grupos</Btn>} /></Card>;

  const asis = statsGrupoAsistencia(db, grupoId);
  const porTrim = [1, 2, 3].map((t) => {
    const s = statsGrupoAsistencia(db, grupoId, t);
    const ps = alumnos.map((a) => calificacionTrimestre(db, a.id, t).final).filter((x) => x > 0);
    return { nombre: `T${t}`, asistencia: s.porcentaje, promedio: ps.length ? round(ps.reduce((x, y) => x + y, 0) / ps.length, 1) : 0, faltas: s.F, retardos: s.R };
  });
  const acts = db.actividades.filter((a) => a.grupoId === grupoId);
  const ents = db.entregas.filter((e) => acts.some((a) => a.id === e.actividadId));
  const incs = db.incidencias.filter((i) => i.grupoId === grupoId);
  const tiposInc = [...new Set(incs.map((i) => i.tipo))].map((t) => ({ nombre: t, valor: incs.filter((i) => i.tipo === t).length }));
  const eco = (() => {
    const nums = [...new Set(db.ecoems.filter((s) => s.grupoId === grupoId).map((s) => s.numero))];
    if (!nums.length) return null;
    const ult = Math.max(...nums);
    const sims = db.ecoems.filter((s) => s.grupoId === grupoId && s.numero === ult);
    return db.catalogos.materiasEcoems.map((m) => ({ nombre: m.nombre.slice(0, 12), promedio: round(sims.reduce((s, x) => s + pct(Number(x.resultados?.[m.id] || 0), m.reactivos), 0) / sims.length, 1) }));
  })();
  const seguimiento = alumnos.map((a) => ({ a, asis: statsAsistencia(db, a.id, { grupoId }), prom: promedioAcumulado(db, a.id), act: statsActividades(db, a.id) }))
    .filter((x) => (x.asis.total >= 5 && x.asis.porcentaje < db.config.alertas.asistenciaMin) || (x.prom > 0 && x.prom < db.config.alertas.promedioMin) || x.act.noEntregadas >= db.config.alertas.actPendientes);

  return (
    <div className="space-y-4">
      <Titulo sub="Vista completa del grupo: asistencia, evaluación, entregas, convivencia y ECOEMS."
        right={<Sel value={grupoId} onChange={(e) => setGrupoId(e.target.value)} className="max-w-[150px]">{db.grupos.map((x) => <option key={x.id} value={x.id}>{x.grado}° {x.grupo}</option>)}</Sel>}>Tablero {g.grado}° {g.grupo}</Titulo>

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <Stat label="Estudiantes" valor={alumnos.length} />
        <Stat label="Asistencia" valor={asis.porcentaje + "%"} sub={`${asis.sesiones} sesiones`} color="text-emerald-700" />
        <Stat label="Faltas" valor={asis.F} color={asis.F ? "text-rose-600" : "text-slate-900"} />
        <Stat label="Retardos" valor={asis.R} />
      </div>

      <div className="grid lg:grid-cols-2 gap-4">
        <Card>
          <p className="text-sm font-medium text-slate-700 mb-3">Asistencia y faltas por trimestre</p>
          <ResponsiveContainer width="100%" height={210}>
            <BarChart data={porTrim}><CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={11} /><YAxis fontSize={11} /><Tooltip /><Legend fontSize={11} />
              <Bar dataKey="faltas" fill="#e11d48" name="Faltas" radius={[3, 3, 0, 0]} /><Bar dataKey="retardos" fill="#f59e0b" name="Retardos" radius={[3, 3, 0, 0]} /></BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <p className="text-sm font-medium text-slate-700 mb-3">Promedio del grupo por trimestre</p>
          <ResponsiveContainer width="100%" height={210}>
            <BarChart data={porTrim}><CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={11} /><YAxis domain={[0, 10]} fontSize={11} /><Tooltip /><Bar dataKey="promedio" fill="#047857" radius={[3, 3, 0, 0]} /></BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <p className="text-sm font-medium text-slate-700 mb-3">Entregas de actividades</p>
          <div className="grid grid-cols-3 gap-2 text-center">
            <div className="bg-emerald-50 border border-emerald-100 rounded-lg py-3"><p className="text-xl font-semibold text-emerald-700">{ents.filter((e) => e.estado === "Entregada" || e.estado === "Fuera de tiempo").length}</p><p className="text-[11px] text-slate-500">Entregadas</p></div>
            <div className="bg-rose-50 border border-rose-100 rounded-lg py-3"><p className="text-xl font-semibold text-rose-700">{ents.filter((e) => e.estado === "No entregada").length}</p><p className="text-[11px] text-slate-500">No entregadas</p></div>
            <div className="bg-slate-50 border border-slate-200 rounded-lg py-3"><p className="text-xl font-semibold text-slate-700">{ents.filter((e) => e.estado === "Pendiente" || e.estado === "En proceso").length}</p><p className="text-[11px] text-slate-500">Pendientes</p></div>
          </div>
          <p className="text-xs text-slate-400 mt-3">{acts.length} actividades registradas en el ciclo.</p>
        </Card>
        <Card>
          <p className="text-sm font-medium text-slate-700 mb-3">Incidencias por tipo</p>
          {tiposInc.length === 0 ? <Vacio texto="Sin incidencias registradas." /> : (
            <ResponsiveContainer width="100%" height={200}>
              <PieChart><Pie data={tiposInc} dataKey="valor" nameKey="nombre" cx="50%" cy="50%" outerRadius={70} label={(e) => e.nombre.slice(0, 10)} fontSize={10}>
                {tiposInc.map((_, i) => <Cell key={i} fill={["#047857", "#0284c7", "#f59e0b", "#e11d48", "#7c3aed", "#64748b"][i % 6]} />)}</Pie><Tooltip /></PieChart>
            </ResponsiveContainer>
          )}
        </Card>
      </div>

      {eco && (
        <Card>
          <p className="text-sm font-medium text-slate-700 mb-3">ECOEMS: promedio por materia (último simulador)</p>
          <ResponsiveContainer width="100%" height={230}>
            <BarChart data={eco}><CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={9} angle={-25} textAnchor="end" height={60} /><YAxis domain={[0, 100]} fontSize={11} /><Tooltip formatter={(v) => v + "%"} />
              <Bar dataKey="promedio" radius={[3, 3, 0, 0]}>{eco.map((m, i) => <Cell key={i} fill={m.promedio >= 60 ? "#047857" : "#f59e0b"} />)}</Bar></BarChart>
          </ResponsiveContainer>
        </Card>
      )}

      <Card>
        <p className="text-sm font-medium text-slate-700 mb-1">Estudiantes que requieren acompañamiento</p>
        <p className="text-[11px] text-slate-400 mb-3">Criterios objetivos: asistencia menor a {db.config.alertas.asistenciaMin}%, promedio menor a {db.config.alertas.promedioMin} o {db.config.alertas.actPendientes} o más actividades sin entregar.</p>
        {seguimiento.length === 0 ? <Vacio texto="Ningún estudiante cumple los criterios de acompañamiento." /> : (
          <Tabla cols={["Nombre", "Asistencia", "Promedio", "Sin entregar"]}>
            {seguimiento.map((x) => (
              <tr key={x.a.id} className="hover:bg-slate-50 cursor-pointer" onClick={() => ir("ficha", { alumnoId: x.a.id })}>
                <td className="py-2 px-3">{nomComp(x.a)}</td>
                <td className="py-2 px-3">{x.asis.porcentaje}%</td>
                <td className="py-2 px-3">{x.prom || "—"}</td>
                <td className="py-2 px-3">{x.act.noEntregadas}</td>
              </tr>
            ))}
          </Tabla>
        )}
      </Card>
    </div>
  );
}

/* ============================================================
   MÓDULO: ESTADÍSTICAS
   ============================================================ */
function Estadisticas({ db, ir }) {
  const [tab, setTab] = useState("asistencia");
  const [grupoId, setGrupoId] = useState("");
  const [trim, setTrim] = useState(0);
  const [rango, setRango] = useState({ desde: "", hasta: "" });
  const alumnos = sortAl(db.alumnos.filter((a) => a.activo && (!grupoId || a.grupoId === grupoId)));
  const grupos = db.grupos.filter((g) => !grupoId || g.id === grupoId);

  const filtros = (
    <div className="flex gap-2 flex-wrap">
      <Sel value={grupoId} onChange={(e) => setGrupoId(e.target.value)} className="max-w-[170px]"><option value="">Todos los grupos</option>{db.grupos.map((g) => <option key={g.id} value={g.id}>{g.grado}° {g.grupo}</option>)}</Sel>
      <Sel value={trim} onChange={(e) => setTrim(Number(e.target.value))} className="max-w-[160px]"><option value={0}>Todo el ciclo</option>{[1, 2, 3].map((t) => <option key={t} value={t}>Trimestre {t}</option>)}</Sel>
      <input type="date" value={rango.desde} onChange={(e) => setRango({ ...rango, desde: e.target.value })} className={inputCls + " max-w-[150px]"} title="Desde" />
      <input type="date" value={rango.hasta} onChange={(e) => setRango({ ...rango, hasta: e.target.value })} className={inputCls + " max-w-[150px]"} title="Hasta" />
      {(rango.desde || rango.hasta) && <Btn size="sm" tipo="fantasma" onClick={() => setRango({ desde: "", hasta: "" })}>Limpiar fechas</Btn>}
    </div>
  );

  const porMes = useMemo(() => {
    const meses = {};
    db.asistencias.filter((x) => (!grupoId || x.grupoId === grupoId)).forEach((x) => {
      const k = x.fecha.slice(0, 7);
      meses[k] = meses[k] || { nombre: `${MESES[Number(k.slice(5, 7)) - 1].slice(0, 3)} ${k.slice(2, 4)}`, A: 0, F: 0, J: 0, R: 0, P: 0, total: 0 };
      Object.values(x.marcas || {}).forEach((m) => { meses[k][m] = (meses[k][m] || 0) + 1; meses[k].total++; });
    });
    return Object.entries(meses).sort().map(([, v]) => ({ ...v, porcentaje: pct(v.A + v.J + v.P + v.R * 0.5, v.total) }));
  }, [db, grupoId]);

  const pdfEstadistica = () => {
    const g2 = db.grupos.find((x) => x.id === grupoId);
    const sub = [g2 ? `${g2.grado}° ${g2.grupo}` : "Todos los grupos", trim ? `Trimestre ${trim}` : "Todo el ciclo",
      rango.desde || rango.hasta ? `Del ${fFecha(rango.desde) } al ${fFecha(rango.hasta)}` : ""].filter(Boolean).join("   ·   ");
    const tablas = {
      asistencia: () => ({ titulo: "Estadística de asistencia", columnas: ["Nombre", "Clases", "Asistencias", "Faltas", "Justificadas", "Retardos", "Permisos", "%"],
        filas: alumnos.map((a) => { const st = statsAsistencia(db, a.id, { trim: trim || undefined, desde: rango.desde || undefined, hasta: rango.hasta || undefined });
          return [nomComp(a), st.total, st.A, st.F, st.J, st.R, st.P, st.total ? st.porcentaje + "%" : "—"]; }) }),
      evaluacion: () => ({ titulo: "Estadística de evaluación", columnas: ["Nombre", "T1", "T2", "T3", "Promedio anual"],
        filas: alumnos.map((a) => [nomComp(a), ...[1, 2, 3].map((t) => calificacionTrimestre(db, a.id, t).final || "—"), promedioAcumulado(db, a.id) || "—"]) }),
      actividades: () => ({ titulo: "Estadística de actividades", columnas: ["Nombre", "Asignadas", "Entregadas", "Fuera de tiempo", "No entregadas", "Pendientes", "% entrega"],
        filas: alumnos.map((a) => { const st = statsActividades(db, a.id); return [nomComp(a), st.asignadas, st.entregadas, st.fueraTiempo, st.noEntregadas, st.pendientes, st.porcentaje + "%"]; }) }),
      convivencia: () => ({ titulo: "Estadística de convivencia", columnas: ["Folio", "Fecha", "Estudiante", "Conducta", "Clasificación", "Estatus"],
        filas: db.incidencias.filter((i) => !grupoId || i.grupoId === grupoId).map((i) => [i.folio, fFecha(i.fecha), nomComp(db.alumnos.find((a) => a.id === i.alumnoId)), i.conducta || i.tipo, i.gravedad || "—", i.estado]) }),
      ecoems: () => ({ titulo: "Estadística del simulador ECOEMS", columnas: ["Nombre", "Aplicaciones", "Inicial", "Actual", "Diferencia"],
        filas: alumnos.map((a) => { const h = historialEcoems(db, a.id);
          return [nomComp(a), h.items.length, h.items.length ? h.inicial + "%" : "—", h.items.length ? h.actual + "%" : "—", h.items.length > 1 ? (h.diferencia > 0 ? "+" : "") + h.diferencia + "%" : "—"]; }) }),
    };
    const t = (tablas[tab] || tablas.asistencia)();
    pdfTabla({ ...t, subtitulo: sub, config: db.config, ciclo: db.ciclo,
      nota: "Las estadísticas sirven para organizar el acompañamiento docente. No se utilizan para exhibir ni etiquetar a las alumnas y los alumnos." });
  };

  return (
    <div className="space-y-4">
      <Titulo sub="Consulta y comparación por grupo, trimestre y estudiante."
        right={<Btn size="sm" tipo="secundario" icon={FileDown} onClick={pdfEstadistica}>Descargar PDF</Btn>}>Estadísticas</Titulo>
      {filtros}
      <Tabs activa={tab} set={setTab} tabs={[
        { id: "asistencia", label: "Asistencia" }, { id: "evaluacion", label: "Evaluación" },
        { id: "actividades", label: "Actividades" }, { id: "convivencia", label: "Incidencias y permisos" }, { id: "ecoems", label: "ECOEMS" }]} />

      {tab === "asistencia" && (
        <>
          <Card>
            <p className="text-sm font-medium text-slate-700 mb-3">Asistencia por mes</p>
            {porMes.length === 0 ? <Vacio texto="Aún no hay pases de lista registrados." /> : (
              <ResponsiveContainer width="100%" height={210}>
                <BarChart data={porMes}><CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={10} /><YAxis fontSize={11} /><Tooltip /><Legend />
                  <Bar dataKey="A" stackId="a" fill="#059669" name="Asistencias" /><Bar dataKey="R" stackId="a" fill="#f59e0b" name="Retardos" />
                  <Bar dataKey="J" stackId="a" fill="#0284c7" name="Justificadas" /><Bar dataKey="F" stackId="a" fill="#e11d48" name="Faltas" radius={[3, 3, 0, 0]} /></BarChart>
              </ResponsiveContainer>
            )}
          </Card>
          <Card>
            <p className="text-sm font-medium text-slate-700 mb-3">Comparativa trimestral por grupo</p>
            <ResponsiveContainer width="100%" height={230}>
              <BarChart data={grupos.map((g) => ({ nombre: `${g.grado}°${g.grupo}`, T1: statsGrupoAsistencia(db, g.id, 1).porcentaje, T2: statsGrupoAsistencia(db, g.id, 2).porcentaje, T3: statsGrupoAsistencia(db, g.id, 3).porcentaje }))}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" /><XAxis dataKey="nombre" fontSize={11} /><YAxis domain={[0, 100]} fontSize={11} /><Tooltip formatter={(v) => v + "%"} /><Legend />
                <Bar dataKey="T1" fill="#047857" radius={[3, 3, 0, 0]} /><Bar dataKey="T2" fill="#0284c7" radius={[3, 3, 0, 0]} /><Bar dataKey="T3" fill="#7c3aed" radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </Card>
          <Card pad={false} className="p-4">
            <Tabla cols={["Nombre", "Clases", "A", "F", "J", "R", "P", "%"]}>
              {alumnos.map((a) => { const s = statsAsistencia(db, a.id, { trim: trim || undefined, desde: rango.desde || undefined, hasta: rango.hasta || undefined });
                return (
                  <tr key={a.id} className="hover:bg-slate-50 cursor-pointer" onClick={() => ir("ficha", { alumnoId: a.id })}>
                    <td className="py-2 px-3">{nomComp(a)}</td><td className="py-2 px-3 text-slate-500">{s.total}</td>
                    {ORDEN_EST.map((k) => <td key={k} className="py-2 px-3">{s[k]}</td>)}
                    <td className={`py-2 px-3 font-medium ${s.porcentaje >= db.config.alertas.asistenciaMin ? "text-emerald-700" : "text-amber-600"}`}>{s.total ? s.porcentaje + "%" : "—"}</td>
                  </tr>
                ); })}
            </Tabla>
          </Card>
        </>
      )}

      {tab === "evaluacion" && (
        <Card pad={false} className="p-4">
          <Tabla cols={["Nombre", "T1", "T2", "T3", "Promedio anual"]}>
            {alumnos.map((a) => (
              <tr key={a.id} className="hover:bg-slate-50 cursor-pointer" onClick={() => ir("ficha", { alumnoId: a.id })}>
                <td className="py-2 px-3">{nomComp(a)}</td>
                {[1, 2, 3].map((t) => { const c = calificacionTrimestre(db, a.id, t).final; return <td key={t} className="py-2 px-3">{c || "—"}</td>; })}
                <td className="py-2 px-3 font-medium">{promedioAcumulado(db, a.id) || "—"}</td>
              </tr>
            ))}
          </Tabla>
        </Card>
      )}

      {tab === "actividades" && (
        <Card pad={false} className="p-4">
          <Tabla cols={["Nombre", "Asignadas", "Entregadas", "Fuera de tiempo", "No entregadas", "Pendientes", "% entrega"]}>
            {alumnos.map((a) => { const s = statsActividades(db, a.id);
              return <tr key={a.id} className="hover:bg-slate-50"><td className="py-2 px-3">{nomComp(a)}</td><td className="py-2 px-3">{s.asignadas}</td>
                <td className="py-2 px-3 text-emerald-700">{s.entregadas}</td><td className="py-2 px-3 text-amber-600">{s.fueraTiempo}</td>
                <td className="py-2 px-3 text-rose-600">{s.noEntregadas}</td><td className="py-2 px-3">{s.pendientes}</td>
                <td className="py-2 px-3 font-medium">{s.porcentaje}%</td></tr>; })}
          </Tabla>
        </Card>
      )}

      {tab === "convivencia" && (
        <div className="space-y-4">
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <Stat label="Incidencias" valor={db.incidencias.filter((i) => !grupoId || i.grupoId === grupoId).length} />
            <Stat label="Abiertas o en seguimiento" valor={db.incidencias.filter((i) => (!grupoId || i.grupoId === grupoId) && i.estado !== "Cerrada").length} color="text-amber-600" />
            <Stat label="Con valoración" valor={db.valoraciones.length} />
            <Stat label="Actividades con permiso" valor={db.permisos.filter((p) => !grupoId || p.grupoId === grupoId).length} />
          </div>
          <Card pad={false} className="p-4">
            <Tabla cols={["Folio", "Fecha", "Estudiante", "Conducta", "Clasificación", "Estatus"]}>
              {db.incidencias.filter((i) => !grupoId || i.grupoId === grupoId).map((i) => (
                <tr key={i.id} className="hover:bg-slate-50 cursor-pointer" onClick={() => ir("incidencias", { incId: i.id })}>
                  <td className="py-2 px-3 font-mono text-xs">{i.folio}</td><td className="py-2 px-3">{fFecha(i.fecha)}</td>
                  <td className="py-2 px-3">{nomComp(db.alumnos.find((a) => a.id === i.alumnoId))}</td><td className="py-2 px-3">{i.conducta || i.tipo}</td>
                  <td className="py-2 px-3"><Pill cls={(GRAVEDAD[i.gravedad] || GRAVEDAD.Leve).cls}>{i.gravedad || "—"}</Pill></td>
                  <td className="py-2 px-3">{i.estado}</td></tr>
              ))}
            </Tabla>
          </Card>
        </div>
      )}

      {tab === "ecoems" && (
        <Card pad={false} className="p-4">
          <Tabla cols={["Nombre", "Simuladores", "Inicial", "Actual", "Diferencia"]}>
            {alumnos.map((a) => { const h = historialEcoems(db, a.id);
              return <tr key={a.id} className="hover:bg-slate-50 cursor-pointer" onClick={() => ir("ficha", { alumnoId: a.id })}>
                <td className="py-2 px-3">{nomComp(a)}</td><td className="py-2 px-3">{h.items.length}</td>
                <td className="py-2 px-3">{h.items.length ? h.inicial + "%" : "—"}</td><td className="py-2 px-3">{h.items.length ? h.actual + "%" : "—"}</td>
                <td className={`py-2 px-3 font-medium ${h.diferencia >= 0 ? "text-emerald-700" : "text-amber-600"}`}>{h.items.length > 1 ? (h.diferencia > 0 ? "+" : "") + h.diferencia : "—"}</td></tr>; })}
          </Tabla>
        </Card>
      )}
    </div>
  );
}

/* ============================================================
   MÓDULO: REPORTES
   ============================================================ */
function Reportes({ db, toast }) {
  const [tipo, setTipo] = useState("asistencia");
  const [grupoId, setGrupoId] = useState(db.grupos[0]?.id || "");
  const [trim, setTrim] = useState(0);
  const [materia, setMateria] = useState("");
  const g = db.grupos.find((x) => x.id === grupoId);
  const alumnos = sortAl(db.alumnos.filter((a) => a.grupoId === grupoId && a.activo));
  const encabezadoTexto = [g ? `${g.grado}° ${g.grupo}` : "", db.config.docente || "", trim ? `Trimestre ${trim}` : "Todo el ciclo", tipo === "materia" && materia ? materia : ""].filter(Boolean).join("   ·   ");

  const armar = () => {
    if (tipo === "asistencia") {
      const filas = alumnos.map((a) => { const s = statsAsistencia(db, a.id, { grupoId, trim: trim || undefined }); return [a.numLista, nomComp(a), s.total, s.A, s.F, s.J, s.R, s.P, s.porcentaje + "%"]; });
      return { titulo: "Reporte de asistencia", cols: ["No.", "Nombre", "Clases", "Asist.", "Faltas", "Just.", "Ret.", "Perm.", "%"], filas };
    }
    if (tipo === "evaluacion") {
      const filas = alumnos.map((a) => [a.numLista, nomComp(a), ...[1, 2, 3].map((t) => calificacionTrimestre(db, a.id, t).final || "—"), promedioAcumulado(db, a.id) || "—"]);
      return { titulo: "Reporte de evaluación", cols: ["No.", "Nombre", "T1", "T2", "T3", "Acumulado"], filas };
    }
    if (tipo === "actividades") {
      const filas = alumnos.map((a) => { const s = statsActividades(db, a.id); return [a.numLista, nomComp(a), s.asignadas, s.entregadas, s.noEntregadas, s.pendientes]; });
      return { titulo: "Reporte de actividades", cols: ["No.", "Nombre", "Asignadas", "Entregadas", "No entregadas", "Pendientes"], filas };
    }
    if (tipo === "permisos") {
      const filas = [];
      db.permisos.filter((p) => p.grupoId === grupoId).forEach((p) => alumnos.forEach((a) => { const r = p.registros?.[a.id] || {}; filas.push([p.nombre, fFecha(p.fecha), nomComp(a), r.entregado ? "Sí" : "No", r.autorizado ? "Sí" : "No", r.obs || ""]); }));
      return { titulo: "Reporte de permisos", cols: ["Actividad", "Fecha", "Nombre", "Entregó", "Autorizado", "Observaciones"], filas };
    }
    if (tipo === "incidencias") {
      const filas = db.incidencias.filter((i) => i.grupoId === grupoId).map((i) => [i.folio, fFecha(i.fecha), i.hora, nomComp(db.alumnos.find((a) => a.id === i.alumnoId)), i.conducta || i.tipo, i.gravedad || "—", i.estado, (i.seguimientos || []).length]);
      return { titulo: "Reporte de incidencias", cols: ["Folio", "Fecha", "Hora", "Nombre", "Conducta", "Clasificación", "Estatus", "Seguimientos"], filas };
    }
    if (tipo === "convivencia") {
      const incs = db.incidencias.filter((i) => !grupoId || i.grupoId === grupoId);
      const porConducta = {};
      incs.forEach((i) => { const k = i.conducta || i.tipo || "Otra"; porConducta[k] = porConducta[k] || { n: 0, Leve: 0, Moderada: 0, Grave: 0, alta: 0, abiertas: 0 };
        porConducta[k].n++; if (i.gravedad === "Alta prioridad de atención") porConducta[k].alta++; else if (porConducta[k][i.gravedad] !== undefined) porConducta[k][i.gravedad]++;
        if (i.estado !== "Cerrada") porConducta[k].abiertas++; });
      const filas = Object.entries(porConducta).sort((a, b) => b[1].n - a[1].n).map(([k, v]) => [k, v.n, v.Leve, v.Moderada, v.Grave, v.alta, v.abiertas]);
      return { titulo: "Reporte de convivencia escolar", cols: ["Conducta", "Total", "Leve", "Moderada", "Grave", "Alta prioridad", "Sin cerrar"], filas };
    }
    if (tipo === "materia") {
      const acts = db.actividades.filter((a) => a.grupoId === grupoId && (!materia || a.materia === materia) && (!trim || Number(a.trimestre) === trim));
      const filas = acts.map((a) => {
        const es = db.entregas.filter((e) => e.actividadId === a.id);
        const notas = es.filter((e) => e.estado === "Entregada" || e.estado === "Fuera de tiempo").map((e) => Number(e.calificacion) || 0);
        return [a.materia || db.config.asignatura || "—", a.nombre, `T${a.trimestre}`, fFecha(a.fechaEntrega), es.filter((e) => e.estado === "Entregada").length,
          es.filter((e) => e.estado === "Fuera de tiempo").length, es.filter((e) => e.estado === "No entregada").length,
          notas.length ? round(notas.reduce((x, y) => x + y, 0) / notas.length, 1) : "—"];
      });
      return { titulo: "Reporte por materia", cols: ["Materia", "Actividad", "Trim.", "Entrega", "Entregadas", "Fuera de tiempo", "No entregadas", "Promedio"], filas };
    }
    if (tipo === "comparativoEcoems") {
      const mats = db.catalogos.materiasEcoems;
      const nums = [...new Set(db.ecoems.filter((s2) => s2.grupoId === grupoId).map((s2) => s2.numero))].sort((a, b) => a - b);
      const filas = mats.map((m) => [m.nombre, m.reactivos, ...nums.map((n) => {
        const ss = db.ecoems.filter((s2) => s2.grupoId === grupoId && s2.numero === n);
        if (!ss.length) return "—";
        const prom = round(ss.reduce((t, s2) => t + Number(s2.resultados?.[m.id] || 0), 0) / ss.length, 1);
        return `${prom} (${pct(prom, m.reactivos)}%)`;
      })]);
      return { titulo: "Comparativo entre aplicaciones ECOEMS", cols: ["Área", "Aciertos máximos", ...nums.map((n) => `Aplicación ${n}`)], filas };
    }
    if (tipo === "seguimiento") {
      const filas = [];
      db.incidencias.filter((i) => i.grupoId === grupoId && i.estado !== "Cerrada").forEach((i) => {
        const v = db.valoraciones.find((x) => x.incidenciaId === i.id);
        filas.push([i.folio, nomComp(db.alumnos.find((a) => a.id === i.alumnoId)), i.estado, v ? NIVELES[v.nivel].label : "Sin valorar", v?.medida || "", (i.seguimientos || []).map((s) => `${fFecha(s.fecha)}: ${s.nota}`).join(" | ")]);
      });
      return { titulo: "Reporte de seguimiento", cols: ["Folio", "Nombre", "Estado", "Nivel", "Medida educativa", "Acciones realizadas"], filas };
    }
    if (tipo === "ecoems") {
      const nums = [...new Set(db.ecoems.filter((s) => s.grupoId === grupoId).map((s) => s.numero))].sort((a, b) => a - b);
      const filas = alumnos.map((a) => { const h = historialEcoems(db, a.id); return [a.numLista, nomComp(a), ...nums.map((n) => { const it = h.items.find((x) => x.sim.numero === n); return it ? `${it.aciertos}/${it.reactivos} (${it.porcentaje}%)` : "—"; }), h.items.length > 1 ? (h.diferencia > 0 ? "+" : "") + h.diferencia + "%" : "—", h.items.length ? semaforoDe(h.actual, db.config).label : "—"]; });
      return { titulo: "Reporte ECOEMS", cols: ["No.", "Nombre", ...nums.map((n) => `Aplicación ${n}`), "Avance", "Semáforo"], filas };
    }
    if (tipo === "bitacora") {
      const filas = db.bitacoras.filter((b) => b.grupoId === grupoId && (!trim || Number(b.trimestre) === trim)).sort((a, b) => a.fecha.localeCompare(b.fecha))
        .map((b) => [b.sesion, fFecha(b.fecha), b.contenido || "—", b.campoFormativo || "—", b.estado, b.logros || "—"]);
      return { titulo: "Reporte de bitácora", cols: ["Sesión", "Fecha", "Contenido", "Campo formativo", "Estado", "Logros"], filas };
    }
    const filas = alumnos.map((a) => {
      const s = statsAsistencia(db, a.id, { grupoId }); const ac = statsActividades(db, a.id); const h = historialEcoems(db, a.id);
      return [a.numLista, nomComp(a), s.porcentaje + "%", promedioAcumulado(db, a.id) || "—", `${ac.entregadas}/${ac.asignadas}`, db.incidencias.filter((i) => i.alumnoId === a.id).length, h.items.length ? h.actual + "%" : "—"];
    });
    return { titulo: "Reporte integral del grupo", cols: ["No.", "Nombre", "Asistencia", "Promedio", "Entregas", "Incidencias", "ECOEMS"], filas };
  };

  const r = armar();
  const resumenPDF = () => {
    if (tipo === "asistencia") {
      const st = statsGrupoAsistencia(db, grupoId, trim || undefined);
      return [{ label: "Estudiantes", valor: alumnos.length }, { label: "Asistencia", valor: st.porcentaje + "%" }, { label: "Faltas", valor: st.F }, { label: "Retardos", valor: st.R }];
    }
    if (tipo === "evaluacion" || tipo === "integral") {
      const ps = alumnos.map((a) => promedioAcumulado(db, a.id)).filter((x) => x > 0);
      return [{ label: "Estudiantes", valor: alumnos.length },
        { label: "Promedio del grupo", valor: ps.length ? round(ps.reduce((x, y) => x + y, 0) / ps.length, 1) : "—" },
        { label: "Evaluados", valor: ps.length }];
    }
    if (tipo === "incidencias" || tipo === "convivencia" || tipo === "seguimiento") {
      const ii = db.incidencias.filter((x) => x.grupoId === grupoId);
      return [{ label: "Registradas", valor: ii.length },
        { label: "Abiertas", valor: ii.filter((x) => x.estado === "Abierta").length },
        { label: "En seguimiento", valor: ii.filter((x) => x.estado === "En seguimiento").length },
        { label: "Cerradas", valor: ii.filter((x) => x.estado === "Cerrada").length }];
    }
    if (tipo === "ecoems" || tipo === "comparativoEcoems") {
      const mats = db.catalogos.materiasEcoems;
      const maxTotal = mats.reduce((t, m) => t + Number(m.reactivos || 0), 0);
      const hs = alumnos.map((a) => historialEcoems(db, a.id)).filter((h) => h.items.length);
      const prom = hs.length ? round(hs.reduce((t, h) => t + h.actual, 0) / hs.length, 1) : 0;
      return [{ label: "Estudiantes con registro", valor: hs.length },
        { label: "Aciertos máximos", valor: maxTotal },
        { label: "Promedio del grupo", valor: prom + "%" }];
    }
    return null;
  };

  const aPDF = () => {
    pdfTabla({
      titulo: r.titulo, subtitulo: encabezadoTexto, columnas: r.cols, filas: r.filas,
      config: db.config, ciclo: db.ciclo, resumen: resumenPDF(),
      nota: "Documento generado para seguimiento docente. La información de las alumnas y los alumnos es confidencial y su uso se limita a fines educativos.",
    });
    toast("PDF descargado");
  };
  const aExcel = () => { exportarExcel(r.titulo.replace(/ /g, "_"), [r.cols, ...r.filas]); toast("Archivo de Excel generado"); };
  const aArchivoCSV = () => { descargar(r.titulo.replace(/ /g, "_") + ".csv", aCSV([r.cols, ...r.filas]), "text/csv;charset=utf-8"); toast("Archivo CSV generado"); };

  return (
    <div className="space-y-4">
      <Titulo sub="Elige el reporte, revisa la vista previa y descárgalo. El PDF se guarda directo en tu dispositivo.">Reportes</Titulo>
      <Card>
        <div className="grid sm:grid-cols-3 gap-3">
          <Campo label="Tipo de reporte">
            <Sel value={tipo} onChange={(e) => setTipo(e.target.value)}>
              <option value="asistencia">Asistencia</option><option value="evaluacion">Evaluación</option><option value="actividades">Actividades</option>
              <option value="materia">Por materia</option><option value="permisos">Permisos</option>
              <option value="incidencias">Incidencias</option><option value="convivencia">Convivencia escolar</option><option value="seguimiento">Seguimiento</option>
              <option value="ecoems">ECOEMS</option><option value="comparativoEcoems">Comparativo ECOEMS</option>
              <option value="bitacora">Bitácora</option><option value="integral">Integral del grupo</option>
            </Sel>
          </Campo>
          <Campo label="Grupo"><Sel value={grupoId} onChange={(e) => setGrupoId(e.target.value)}>{db.grupos.map((x) => <option key={x.id} value={x.id}>{x.grado}° {x.grupo}</option>)}</Sel></Campo>
          <Campo label="Periodo"><Sel value={trim} onChange={(e) => setTrim(Number(e.target.value))}><option value={0}>Todo el ciclo</option>{[1, 2, 3].map((t) => <option key={t} value={t}>Trimestre {t}</option>)}</Sel></Campo>
          {tipo === "materia" && <Campo label="Materia"><Sel value={materia} onChange={(e) => setMateria(e.target.value)}><option value="">Todas</option>{(db.catalogos.materias || []).map((m) => <option key={m}>{m}</option>)}</Sel></Campo>}
        </div>
        <div className="flex gap-2 mt-4 flex-wrap">
          <Btn icon={FileDown} onClick={aPDF}>Descargar PDF</Btn>
          <Btn tipo="secundario" icon={Download} onClick={aExcel}>Excel</Btn>
          <Btn tipo="secundario" icon={Download} onClick={aArchivoCSV}>CSV</Btn>
        </div>
      </Card>

      <Card pad={false} className="p-4">
        <p className="text-sm font-medium text-slate-700 mb-3">{r.titulo} · vista previa ({r.filas.length} filas)</p>
        {r.filas.length === 0 ? <Vacio texto="No hay datos para este reporte." /> : (
          <Tabla cols={r.cols}>
            {r.filas.slice(0, 60).map((f, i) => <tr key={i} className="hover:bg-slate-50">{f.map((c, j) => <td key={j} className="py-2 px-3">{c}</td>)}</tr>)}
          </Tabla>
        )}
      </Card>
    </div>
  );
}

/* ============================================================
   MÓDULO: CALENDARIO
   ============================================================ */
const MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
const DIAS = ["Dom", "Lun", "Mar", "Mié", "Jue", "Vie", "Sáb"];

function Calendario({ db, upd, ir, toast }) {
  const [vista, setVista] = useState("mes");
  const [ref, setRef] = useState(new Date());
  const [modal, setModal] = useState(null);

  const eventos = useMemo(() => {
    const e = [];
    db.actividades.forEach((a) => { const g = db.grupos.find((x) => x.id === a.grupoId); e.push({ fecha: a.fechaEntrega, titulo: a.nombre, tipo: "Entrega", grupo: g ? `${g.grado}°${g.grupo}` : "", color: "bg-sky-100 text-sky-800 border-sky-200", destino: "actividades" }); });
    db.permisos.forEach((p) => { const g = db.grupos.find((x) => x.id === p.grupoId); e.push({ fecha: p.fecha, titulo: p.nombre, tipo: "Actividad escolar", grupo: g ? `${g.grado}°${g.grupo}` : "", color: "bg-violet-100 text-violet-800 border-violet-200", destino: "permisos" }); });
    db.ecoems.forEach((s) => { if (!e.find((x) => x.fecha === s.fecha && x.tipo === "ECOEMS" && x.titulo.includes(s.numero))) e.push({ fecha: s.fecha, titulo: `Simulador ${s.numero}`, tipo: "ECOEMS", grupo: "", color: "bg-emerald-100 text-emerald-800 border-emerald-200", destino: "ecoems" }); });
    db.bitacoras.forEach((b) => { const g = db.grupos.find((x) => x.id === b.grupoId); e.push({ fecha: b.fecha, titulo: `Sesión ${b.sesion}: ${b.contenido || "clase"}`, tipo: "Bitácora", grupo: g ? `${g.grado}°${g.grupo}` : "", color: "bg-slate-100 text-slate-700 border-slate-200", destino: "bitacora" }); });
    db.valoraciones.filter((v) => v.fechaSeguimiento).forEach((v) => { const i = db.incidencias.find((x) => x.id === v.incidenciaId); e.push({ fecha: v.fechaSeguimiento, titulo: `Seguimiento ${i?.folio || ""}`, tipo: "Seguimiento", grupo: "", color: "bg-amber-100 text-amber-800 border-amber-200", destino: "valoracion" }); });
    (db.eventos || []).forEach((v) => e.push({ ...v, color: "bg-rose-100 text-rose-800 border-rose-200", tipo: v.tipo || "Evento", destino: "calendario", propio: true }));
    return e;
  }, [db]);

  const evDe = (f) => eventos.filter((e) => e.fecha === f);
  const guardarEv = () => {
    if (!modal.titulo) return toast("Escribe el nombre del evento", "error");
    upd((d) => { d.eventos = [...(d.eventos || []), { id: uid("ev"), titulo: modal.titulo, fecha: modal.fecha, tipo: modal.tipo }]; });
    setModal(null); toast("Evento agregado");
  };

  const y = ref.getFullYear(), m = ref.getMonth();
  const primero = new Date(y, m, 1).getDay();
  const dias = new Date(y, m + 1, 0).getDate();
  const celdas = [...Array(primero).fill(null), ...Array.from({ length: dias }, (_, i) => i + 1)];
  const iso = (d) => `${y}-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;

  const inicioSemana = new Date(ref); inicioSemana.setDate(ref.getDate() - ref.getDay());
  const semana = Array.from({ length: 7 }, (_, i) => { const d = new Date(inicioSemana); d.setDate(inicioSemana.getDate() + i); return d; });
  const mover = (n) => { const d = new Date(ref); if (vista === "mes") d.setMonth(d.getMonth() + n); else if (vista === "semana") d.setDate(d.getDate() + n * 7); else d.setDate(d.getDate() + n); setRef(d); };

  return (
    <div className="space-y-4">
      <Titulo sub="Clases, entregas, actividades escolares, simuladores y seguimientos."
        right={<Btn size="sm" icon={Plus} onClick={() => setModal({ titulo: "", fecha: hoy(), tipo: "Evento" })}>Agregar evento</Btn>}>Calendario</Titulo>
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-1">
          <Btn size="sm" tipo="secundario" onClick={() => mover(-1)}><ChevronLeft size={15} /></Btn>
          <span className="text-sm font-medium text-slate-800 px-2 min-w-[150px] text-center">
            {vista === "mes" ? `${MESES[m]} ${y}` : vista === "semana" ? `${fFecha(semana[0].toISOString().slice(0, 10))} al ${fFecha(semana[6].toISOString().slice(0, 10))}` : fFecha(ref.toISOString().slice(0, 10))}
          </span>
          <Btn size="sm" tipo="secundario" onClick={() => mover(1)}><ChevronRight size={15} /></Btn>
          <Btn size="sm" tipo="fantasma" onClick={() => setRef(new Date())}>Hoy</Btn>
        </div>
        <Tabs activa={vista} set={setVista} tabs={[{ id: "dia", label: "Día" }, { id: "semana", label: "Semana" }, { id: "mes", label: "Mes" }]} />
      </div>

      {vista === "mes" && (
        <Card pad={false} className="p-2 overflow-x-auto">
          <div className="grid grid-cols-7 gap-1 min-w-[560px]">
            {DIAS.map((d) => <div key={d} className="text-center text-[11px] font-medium text-slate-400 py-1">{d}</div>)}
            {celdas.map((d, i) => (
              <div key={i} className={`min-h-[76px] rounded-lg p-1 border ${d ? "border-slate-100" : "border-transparent"} ${d && iso(d) === hoy() ? "bg-emerald-50 border-emerald-200" : ""}`}>
                {d && <>
                  <p className="text-[11px] text-slate-500 mb-0.5">{d}</p>
                  {evDe(iso(d)).slice(0, 3).map((e, j) => (
                    <button key={j} onClick={() => ir(e.destino)} className={`block w-full text-left text-[10px] px-1 py-0.5 rounded mb-0.5 border truncate ${e.color}`}>{e.titulo}</button>
                  ))}
                  {evDe(iso(d)).length > 3 && <p className="text-[10px] text-slate-400">+{evDe(iso(d)).length - 3} más</p>}
                </>}
              </div>
            ))}
          </div>
        </Card>
      )}

      {vista === "semana" && (
        <div className="space-y-2">
          {semana.map((d, i) => { const f = d.toISOString().slice(0, 10);
            return (
              <Card key={i} className={f === hoy() ? "border-emerald-300" : ""}>
                <p className="text-sm font-medium text-slate-800">{DIAS[d.getDay()]} {d.getDate()} de {MESES[d.getMonth()].toLowerCase()}</p>
                <div className="mt-2 space-y-1">
                  {evDe(f).length === 0 ? <p className="text-xs text-slate-400">Sin registros</p> :
                    evDe(f).map((e, j) => <button key={j} onClick={() => ir(e.destino)} className={`block w-full text-left text-xs px-2 py-1.5 rounded-lg border ${e.color}`}><span className="font-medium">{e.tipo}</span> · {e.titulo} {e.grupo && `· ${e.grupo}`}</button>)}
                </div>
              </Card>
            ); })}
        </div>
      )}

      {vista === "dia" && (
        <Card>
          <p className="text-sm font-medium text-slate-800 mb-3">{fFecha(ref.toISOString().slice(0, 10))}</p>
          {evDe(ref.toISOString().slice(0, 10)).length === 0 ? <Vacio texto="No hay nada registrado en este día." /> : (
            <div className="space-y-1.5">
              {evDe(ref.toISOString().slice(0, 10)).map((e, j) => <button key={j} onClick={() => ir(e.destino)} className={`block w-full text-left text-sm px-3 py-2 rounded-lg border ${e.color}`}><span className="font-medium">{e.tipo}</span> · {e.titulo} {e.grupo && `· ${e.grupo}`}</button>)}
            </div>
          )}
        </Card>
      )}

      <Modal open={!!modal} onClose={() => setModal(null)} title="Agregar evento" ancho="max-w-md"
        footer={<><Btn tipo="secundario" onClick={() => setModal(null)}>Cancelar</Btn><Btn icon={Save} onClick={guardarEv}>Guardar</Btn></>}>
        {modal && (
          <div className="space-y-3">
            <Campo label="Nombre" req><Inp value={modal.titulo} onChange={(e) => setModal({ ...modal, titulo: e.target.value })} /></Campo>
            <Campo label="Fecha"><Inp type="date" value={modal.fecha} onChange={(e) => setModal({ ...modal, fecha: e.target.value })} /></Campo>
            <Campo label="Tipo"><Sel value={modal.tipo} onChange={(e) => setModal({ ...modal, tipo: e.target.value })}>{["Evento", "Examen", "Proyecto", "Junta", "Simulador", "Otro"].map((t) => <option key={t}>{t}</option>)}</Sel></Campo>
          </div>
        )}
      </Modal>
    </div>
  );
}

/* ============================================================
   MÓDULO: CONFIGURACIÓN, RESPALDO Y CICLOS
   ============================================================ */
function Configuracion({ db, upd, meta, setCiclo, nuevoCiclo, restaurar, toast, perfil, setPerfil }) {
  const [tab, setTab] = useState("escuela");
  const [c, setC] = useState(db.config);
  const [cat, setCat] = useState(db.catalogos);
  const [conf, setConf] = useState(null);
  const fileRef = useRef();
  const [almacenamiento, setAlmacenamiento] = useState(null);
  const [copias, setCopias] = useState([]);
  const [papelera, setPapelera] = useState([]);
  useEffect(() => { setC(clone(db.config)); setCat(clone(db.catalogos)); }, [db.ciclo]);

  const revisarAlmacenamiento = useCallback(async () => {
    const [perm, esp, cop, pap] = await Promise.all([pedirPermanencia(), espacio(), verRespaldos(), verPapelera()]);
    setAlmacenamiento({ ...perm, ...(esp || {}) });
    setCopias(cop); setPapelera(pap);
  }, []);
  useEffect(() => { if (tab === "respaldo" || tab === "papelera") revisarAlmacenamiento(); }, [tab, revisarAlmacenamiento]);

  const activarPermanencia = async () => {
    const r = await pedirPermanencia();
    await revisarAlmacenamiento();
    toast(r.permanente ? "Almacenamiento permanente activado" : "El navegador no concedió el permiso. Instala la app desde el menú del navegador y vuelve a intentarlo.", r.permanente ? "ok" : "error");
  };

  const restaurarCopia = (copia) => setConf({
    texto: `Se restaurará la copia del ${fFecha(copia.fecha)} del ciclo ${copia.ciclo}, con ${copia.alumnos} estudiantes y ${copia.registros} registros. La información actual de ese ciclo se guardará antes en la papelera, por si quieres volver.`,
    textoSi: "Restaurar",
    onSi: async () => {
      const datos = await leerRespaldo(copia.id);
      if (!datos) { setConf(null); return toast("No se pudo leer la copia", "error"); }
      await eliminarConRespaldo({ tipo: "Ciclo reemplazado", descripcion: `${db.ciclo} antes de restaurar la copia del ${fFecha(copia.fecha)}`, ciclo: db.ciclo, datos: clone(db) });
      restaurar(datos); setConf(null);
    },
  });

  const recuperar = async (item) => {
    const datos = await sacarDePapelera(item.id);
    if (!datos) return toast("No se pudo recuperar", "error");
    if (item.tipo === "Ciclo reemplazado") { restaurar(datos.datos); return; }
    upd((d) => {
      Object.entries(datos.datos || {}).forEach(([coleccion, registros]) => {
        if (!Array.isArray(d[coleccion]) || !Array.isArray(registros)) return;
        registros.forEach((r) => { if (!d[coleccion].some((x) => x.id === r.id)) d[coleccion].push(r); });
      });
    });
    await revisarAlmacenamiento();
    toast(`${item.tipo} recuperado`);
  };

  const guardarConfig = () => { upd((d) => { d.config = { ...c, trimestre: Number(c.trimestre) }; }); toast("Configuración guardada"); };
  const guardarCat = () => { upd((d) => { d.catalogos = cat; }); toast("Catálogos guardados"); };

  const respaldar = () => {
    descargar(`respaldo_${db.ciclo}_${hoy()}.json`, JSON.stringify({ aplicacion: "Control Docente Integral", version: 2, generado: new Date().toISOString(), meta, db }, null, 2), "application/json");
    toast("Respaldo descargado");
  };
  const abrirArchivoRespaldo = (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    const r = new FileReader();
    r.onload = (ev) => {
      try {
        const datos = JSON.parse(ev.target.result);
        if (!datos.db || !datos.db.ciclo) throw new Error("formato");
        setConf({ texto: `El respaldo corresponde al ciclo ${datos.db.ciclo} con ${datos.db.alumnos?.length || 0} estudiantes. Al restaurar se reemplaza la información actual de ese ciclo.`, textoSi: "Restaurar", onSi: () => { restaurar(datos.db); setConf(null); } });
      } catch (err) { toast("El archivo no es un respaldo válido", "error"); }
      e.target.value = "";
    };
    r.readAsText(f);
  };

  const setLista = (campo, i, val) => { const c2 = clone(cat); c2[campo][i] = val; setCat(c2); };

  return (
    <div className="space-y-4">
      <Titulo sub="Datos de la escuela, catálogos, ciclos escolares y respaldo de información.">Configuración</Titulo>
      <Tabs activa={tab} set={setTab} tabs={[{ id: "escuela", label: "Escuela" }, { id: "catalogos", label: "Catálogos" }, { id: "alertas", label: "Alertas y semáforo" }, { id: "perfil", label: "Perfil" }, { id: "ciclos", label: "Ciclos" }, { id: "respaldo", label: "Respaldo" }, { id: "papelera", label: "Papelera" }]} />

      {tab === "perfil" && (
        <Card>
          <p className="text-sm text-slate-600 mb-3">El perfil define qué módulos se muestran en el menú. Es un control de vista dentro de este dispositivo, no una cuenta con contraseña.</p>
          <div className="space-y-2">
            {Object.entries(PERFILES).map(([k, v]) => (
              <button key={k} onClick={() => setPerfil(k)}
                className={`w-full text-left px-3 py-2.5 rounded-lg border ${perfil === k ? "bg-emerald-50 border-emerald-300" : "bg-white border-slate-200 hover:border-slate-300"}`}>
                <p className="text-sm font-medium text-slate-800">{k}{perfil === k && <span className="ml-2 text-xs text-emerald-700">en uso</span>}</p>
                <p className="text-xs text-slate-500">{v.desc}</p>
                <p className="text-[11px] text-slate-400 mt-0.5">{v.modulos ? `${v.modulos.length} módulos visibles` : "Todos los módulos"}</p>
              </button>
            ))}
          </div>
          <div className="mt-3"><Aviso>Los datos de las alumnas y los alumnos se guardan solo en este dispositivo. Protege el acceso con el bloqueo de pantalla y evita compartir el respaldo por medios abiertos.</Aviso></div>
        </Card>
      )}

      {tab === "escuela" && (
        <Card>
          <div className="grid sm:grid-cols-2 gap-3">
            <Campo label="Nombre de la escuela"><Inp value={c.escuela} onChange={(e) => setC({ ...c, escuela: e.target.value })} /></Campo>
            <Campo label="CCT"><Inp value={c.cct} onChange={(e) => setC({ ...c, cct: e.target.value })} /></Campo>
            <Campo label="Docente"><Inp value={c.docente} onChange={(e) => setC({ ...c, docente: e.target.value })} /></Campo>
            <Campo label="Asignatura"><Inp value={c.asignatura} onChange={(e) => setC({ ...c, asignatura: e.target.value })} /></Campo>
            <Campo label="Turno"><Sel value={c.turno} onChange={(e) => setC({ ...c, turno: e.target.value })}><option>Matutino</option><option>Vespertino</option></Sel></Campo>
            <Campo label="Horario"><Inp value={c.horario} onChange={(e) => setC({ ...c, horario: e.target.value })} placeholder="Lunes a viernes 07:00 - 13:30" /></Campo>
            <Campo label="Trimestre actual"><Sel value={c.trimestre} onChange={(e) => setC({ ...c, trimestre: e.target.value })}>{[1, 2, 3].map((t) => <option key={t} value={t}>Trimestre {t}</option>)}</Sel></Campo>
            <Campo label="Ciclo escolar"><div className="px-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg text-slate-600">{db.ciclo}</div></Campo>
          </div>
          <p className="text-xs font-medium text-slate-600 mt-5 mb-2">Fechas de los trimestres</p>
          <div className="space-y-2">
            {c.trimestres.map((t, i) => (
              <div key={t.n} className="flex gap-2 items-center flex-wrap">
                <span className="text-sm text-slate-600 w-24">Trimestre {t.n}</span>
                <input type="date" value={t.inicio} onChange={(e) => { const ts = [...c.trimestres]; ts[i] = { ...t, inicio: e.target.value }; setC({ ...c, trimestres: ts }); }} className={inputCls + " max-w-[160px]"} />
                <input type="date" value={t.fin} onChange={(e) => { const ts = [...c.trimestres]; ts[i] = { ...t, fin: e.target.value }; setC({ ...c, trimestres: ts }); }} className={inputCls + " max-w-[160px]"} />
              </div>
            ))}
          </div>
          <div className="flex justify-end mt-4"><Btn icon={Save} onClick={guardarConfig}>Guardar configuración</Btn></div>
        </Card>
      )}

      {tab === "catalogos" && (
        <div className="space-y-4">
          <Card>
            <p className="text-sm font-medium text-slate-700 mb-2">Materias del simulador ECOEMS</p>
            <div className="space-y-2">
              {cat.materiasEcoems.map((m, i) => (
                <div key={m.id} className="flex gap-2 items-center">
                  <input value={m.nombre} onChange={(e) => setLista("materiasEcoems", i, { ...m, nombre: e.target.value })} className={inputCls} />
                  <input type="number" value={m.reactivos} onChange={(e) => setLista("materiasEcoems", i, { ...m, reactivos: Number(e.target.value) })} className={inputCls + " w-24"} />
                  <button onClick={() => setCat({ ...cat, materiasEcoems: cat.materiasEcoems.filter((_, j) => j !== i) })} className="p-2 text-rose-500 hover:bg-rose-50 rounded-lg"><Trash2 size={15} /></button>
                </div>
              ))}
            </div>
            <div className="flex justify-between items-center mt-3">
              <Btn size="sm" tipo="secundario" icon={Plus} onClick={() => setCat({ ...cat, materiasEcoems: [...cat.materiasEcoems, { id: uid("m"), nombre: "", reactivos: 10 }] })}>Agregar materia</Btn>
              <span className={`text-sm font-medium ${cat.materiasEcoems.reduce((s2, m) => s2 + Number(m.reactivos || 0), 0) === 128 ? "text-emerald-700" : "text-amber-600"}`}>
                Total de aciertos: {cat.materiasEcoems.reduce((s2, m) => s2 + Number(m.reactivos || 0), 0)} {cat.materiasEcoems.reduce((s2, m) => s2 + Number(m.reactivos || 0), 0) === 128 ? "· coincide con el simulador de 128" : "· el simulador estándar son 128"}
              </span>
            </div>
          </Card>
          <Card>
            <p className="text-sm font-medium text-slate-700 mb-2">Catálogo de conductas · Marco para la Convivencia Escolar</p>
            <p className="text-xs text-slate-500 mb-3">Actualízalo cuando cambie la normativa aplicable en tu entidad.</p>
            <div className="space-y-2">
              {cat.conductas.map((t, i) => (
                <div key={i} className="flex gap-2 items-center">
                  <input value={t} onChange={(e) => setLista("conductas", i, e.target.value)} className={inputCls} />
                  <button onClick={() => setCat({ ...cat, conductas: cat.conductas.filter((_, j) => j !== i) })} className="p-2 text-rose-500 hover:bg-rose-50 rounded-lg"><Trash2 size={15} /></button>
                </div>
              ))}
            </div>
            <div className="mt-3"><Btn size="sm" tipo="secundario" icon={Plus} onClick={() => setCat({ ...cat, conductas: [...cat.conductas, ""] })}>Agregar conducta</Btn></div>
          </Card>
          <Card>
            <p className="text-sm font-medium text-slate-700 mb-2">Materias</p>
            <div className="flex flex-wrap gap-2">
              {(cat.materias || []).map((t, i) => (
                <span key={i} className="inline-flex items-center gap-1 bg-slate-100 border border-slate-200 rounded-lg px-2 py-1 text-xs">
                  {t}<button onClick={() => setCat({ ...cat, materias: cat.materias.filter((_, j) => j !== i) })} className="text-slate-400 hover:text-rose-500"><X size={12} /></button>
                </span>
              ))}
            </div>
            <div className="flex gap-2 mt-3">
              <input placeholder="Nueva materia y Enter" className={inputCls} onKeyDown={(e) => { if (e.key === "Enter" && e.target.value) { setCat({ ...cat, materias: [...(cat.materias || []), e.target.value] }); e.target.value = ""; } }} />
            </div>
          </Card>
          <Card>
            <p className="text-sm font-medium text-slate-700 mb-2">Tipos de actividad</p>
            <div className="flex flex-wrap gap-2">
              {cat.tiposActividad.map((t, i) => (
                <span key={i} className="inline-flex items-center gap-1 bg-slate-100 border border-slate-200 rounded-lg px-2 py-1 text-xs">
                  {t}<button onClick={() => setCat({ ...cat, tiposActividad: cat.tiposActividad.filter((_, j) => j !== i) })} className="text-slate-400 hover:text-rose-500"><X size={12} /></button>
                </span>
              ))}
            </div>
            <div className="flex gap-2 mt-3">
              <input id="nuevoTipoAct" placeholder="Nuevo tipo" className={inputCls} onKeyDown={(e) => { if (e.key === "Enter" && e.target.value) { setCat({ ...cat, tiposActividad: [...cat.tiposActividad, e.target.value] }); e.target.value = ""; } }} />
            </div>
          </Card>
          <Card>
            <p className="text-sm font-medium text-slate-700 mb-2">Tipos de incidencia</p>
            <div className="flex flex-wrap gap-2">
              {cat.tiposIncidencia.map((t, i) => (
                <span key={i} className="inline-flex items-center gap-1 bg-slate-100 border border-slate-200 rounded-lg px-2 py-1 text-xs">
                  {t}<button onClick={() => setCat({ ...cat, tiposIncidencia: cat.tiposIncidencia.filter((_, j) => j !== i) })} className="text-slate-400 hover:text-rose-500"><X size={12} /></button>
                </span>
              ))}
            </div>
            <div className="flex gap-2 mt-3">
              <input placeholder="Nuevo tipo" className={inputCls} onKeyDown={(e) => { if (e.key === "Enter" && e.target.value) { setCat({ ...cat, tiposIncidencia: [...cat.tiposIncidencia, e.target.value] }); e.target.value = ""; } }} />
            </div>
          </Card>
          <div className="flex justify-end"><Btn icon={Save} onClick={guardarCat}>Guardar catálogos</Btn></div>
        </div>
      )}

      {tab === "alertas" && (
        <Card>
          <p className="text-sm text-slate-600 mb-4">Define a partir de qué momento el sistema muestra un aviso de acompañamiento.</p>
          <div className="grid sm:grid-cols-2 gap-3">
            <Campo label="Faltas para generar aviso"><Inp type="number" value={c.alertas.faltasMax} onChange={(e) => setC({ ...c, alertas: { ...c.alertas, faltasMax: Number(e.target.value) } })} /></Campo>
            <Campo label="Asistencia mínima esperada (%)"><Inp type="number" value={c.alertas.asistenciaMin} onChange={(e) => setC({ ...c, alertas: { ...c.alertas, asistenciaMin: Number(e.target.value) } })} /></Campo>
            <Campo label="Promedio mínimo esperado"><Inp type="number" step="0.1" value={c.alertas.promedioMin} onChange={(e) => setC({ ...c, alertas: { ...c.alertas, promedioMin: Number(e.target.value) } })} /></Campo>
            <Campo label="Porcentaje mínimo en ECOEMS"><Inp type="number" value={c.alertas.ecoemsMin} onChange={(e) => setC({ ...c, alertas: { ...c.alertas, ecoemsMin: Number(e.target.value) } })} /></Campo>
            <Campo label="Actividades sin entregar para aviso"><Inp type="number" value={c.alertas.actPendientes} onChange={(e) => setC({ ...c, alertas: { ...c.alertas, actPendientes: Number(e.target.value) } })} /></Campo>
            <Campo label="Días para avisar de un seguimiento vencido"><Inp type="number" value={c.alertas.diasSeguimiento ?? 7} onChange={(e) => setC({ ...c, alertas: { ...c.alertas, diasSeguimiento: Number(e.target.value) } })} /></Campo>
          </div>
          <div className="border-t border-slate-200 mt-5 pt-4">
            <p className="text-sm font-medium text-slate-700 mb-1">Semáforo académico ECOEMS</p>
            <p className="text-xs text-slate-500 mb-3">Define desde qué porcentaje se considera cada nivel.</p>
            <div className="grid sm:grid-cols-2 gap-3">
              <Campo label="Desempeño alto desde (%)"><Inp type="number" value={c.semaforo?.alto ?? 80} onChange={(e) => setC({ ...c, semaforo: { ...(c.semaforo || {}), alto: Number(e.target.value) } })} /></Campo>
              <Campo label="Desempeño medio desde (%)"><Inp type="number" value={c.semaforo?.medio ?? 60} onChange={(e) => setC({ ...c, semaforo: { ...(c.semaforo || {}), medio: Number(e.target.value) } })} /></Campo>
            </div>
            <div className="flex gap-2 mt-3 flex-wrap">
              {[95, 70, 40].map((v) => { const sm = semaforoDe(v, c); return <Pill key={v} cls={sm.cls}>{v}% · {sm.label}</Pill>; })}
            </div>
            <div className="mt-3"><Aviso>{AVISO_ECOEMS}</Aviso></div>
          </div>
          <div className="flex justify-end mt-4"><Btn icon={Save} onClick={guardarConfig}>Guardar alertas</Btn></div>
        </Card>
      )}

      {tab === "ciclos" && (
        <Card>
          <p className="text-sm text-slate-600 mb-3">Cada ciclo guarda por separado sus grupos, alumnos, asistencia, evaluaciones, actividades, incidencias, bitácoras y ECOEMS.</p>
          <div className="space-y-2">
            {meta.ciclos.map((ci) => (
              <div key={ci} className={`flex items-center justify-between px-3 py-2.5 rounded-lg border ${ci === meta.activo ? "bg-emerald-50 border-emerald-300" : "bg-white border-slate-200"}`}>
                <span className="text-sm font-medium text-slate-800">{ci}{ci === meta.activo && <span className="ml-2 text-xs text-emerald-700">en uso</span>}</span>
                {ci !== meta.activo && <Btn size="sm" tipo="secundario" onClick={() => setCiclo(ci)}>Abrir</Btn>}
              </div>
            ))}
          </div>
          <div className="mt-4"><Btn icon={Plus} tipo="secundario" onClick={() => { const n = prompt("Nuevo ciclo escolar (ejemplo: 2027-2028)"); if (n && /^\d{4}-\d{4}$/.test(n)) nuevoCiclo(n); else if (n) toast("Usa el formato 2027-2028", "error"); }}>Crear ciclo escolar</Btn></div>
        </Card>
      )}

      {tab === "respaldo" && (
        <div className="space-y-4">
          <Card>
            <div className="flex items-start gap-3">
              <ShieldCheck size={20} className={almacenamiento?.permanente ? "text-emerald-600 shrink-0 mt-0.5" : "text-amber-500 shrink-0 mt-0.5"} />
              <div className="grow">
                <p className="text-sm font-medium text-slate-800">
                  {almacenamiento === null ? "Revisando el almacenamiento…" : almacenamiento.permanente ? "Almacenamiento permanente activo" : "Almacenamiento permanente sin activar"}
                </p>
                <p className="text-sm text-slate-500 mt-0.5">
                  {almacenamiento?.permanente
                    ? "Tus datos están protegidos. El navegador no los borrará aunque el dispositivo se quede sin espacio."
                    : "Sin este permiso, el navegador podría liberar espacio y borrar la información cuando el dispositivo se llene."}
                </p>
                {almacenamiento && almacenamiento.usadoMB !== undefined && (
                  <p className="text-xs text-slate-400 mt-1">Ocupa {almacenamiento.usadoMB} MB de {almacenamiento.disponibleMB} MB disponibles.</p>
                )}
                {almacenamiento && !almacenamiento.permanente && (
                  <div className="mt-3"><Btn size="sm" icon={ShieldCheck} onClick={activarPermanencia}>Activar protección</Btn></div>
                )}
              </div>
            </div>
          </Card>

          <Card>
            <p className="text-sm font-medium text-slate-700 mb-1">Copias automáticas</p>
            <p className="text-sm text-slate-500 mb-3">La app guarda una copia cada día que la usas y conserva las diez más recientes. No necesitas hacer nada.</p>
            {copias.length === 0 ? <Vacio texto="Todavía no hay copias automáticas. Se creará una en cuanto registres información." /> : (
              <div className="space-y-1.5">
                {copias.map((c) => (
                  <div key={c.id} className="flex items-center justify-between gap-3 px-3 py-2 rounded-lg border border-slate-200 flex-wrap">
                    <div>
                      <p className="text-sm text-slate-800">{fFecha(c.fecha)} <span className="text-slate-400">· ciclo {c.ciclo}</span></p>
                      <p className="text-[11px] text-slate-400">{c.alumnos} estudiantes · {c.registros} registros</p>
                    </div>
                    <Btn size="sm" tipo="secundario" icon={RotateCcw} onClick={() => restaurarCopia(c)}>Restaurar</Btn>
                  </div>
                ))}
              </div>
            )}
          </Card>

          <Card>
            <p className="text-sm font-medium text-slate-700 mb-1">Respaldo en archivo</p>
            <p className="text-sm text-slate-500 mb-3">Descarga un archivo con toda la información del ciclo {db.ciclo}. Guárdalo en tu teléfono, en la nube o en una memoria USB. Sirve para cambiar de dispositivo.</p>
            <div className="flex gap-2 flex-wrap">
              <Btn icon={Database} onClick={respaldar}>Descargar respaldo</Btn>
              <input ref={fileRef} type="file" accept=".json" onChange={abrirArchivoRespaldo} className="hidden" />
              <Btn tipo="secundario" icon={Upload} onClick={() => fileRef.current?.click()}>Restaurar desde archivo</Btn>
            </div>
          </Card>

          <Aviso>La información vive en este dispositivo y funciona sin internet. Aun así, descarga un respaldo al final de cada trimestre y antes de cambiar de teléfono.</Aviso>
        </div>
      )}

      {tab === "papelera" && (
        <div className="space-y-4">
          <Card>
            <p className="text-sm font-medium text-slate-700 mb-1">Papelera</p>
            <p className="text-sm text-slate-500">Todo lo que eliminas llega aquí y puede regresar. Nada desaparece hasta que tú lo decides.</p>
          </Card>
          {papelera.length === 0 ? <Card><Vacio texto="La papelera está vacía." /></Card> : (
            <>
              <div className="space-y-1.5">
                {papelera.map((item) => (
                  <Card key={item.id}>
                    <div className="flex items-center justify-between gap-3 flex-wrap">
                      <div className="grow min-w-[160px]">
                        <div className="flex items-center gap-2 flex-wrap">
                          <Pill>{item.tipo}</Pill>
                          <span className="text-[11px] text-slate-400">{new Date(item.fecha).toLocaleString("es-MX")}</span>
                        </div>
                        <p className="text-sm text-slate-800 mt-1">{item.descripcion}</p>
                        <p className="text-[11px] text-slate-400">Ciclo {item.ciclo}</p>
                      </div>
                      <div className="flex gap-2 shrink-0">
                        <Btn size="sm" icon={Undo2} onClick={() => recuperar(item)}>Recuperar</Btn>
                        <Btn size="sm" tipo="secundario" onClick={() => setConf({
                          texto: `¿Eliminar definitivamente "${item.descripcion}"? Después de esto ya no se podrá recuperar.`,
                          textoSi: "Eliminar para siempre",
                          onSi: async () => { await borrarDePapelera(item.id); await revisarAlmacenamiento(); setConf(null); toast("Eliminado definitivamente"); },
                        })}>Eliminar</Btn>
                      </div>
                    </div>
                  </Card>
                ))}
              </div>
              <div className="flex justify-end">
                <Btn tipo="peligro" size="sm" icon={Trash2} onClick={() => setConf({
                  texto: `¿Vaciar la papelera por completo? Se eliminarán ${papelera.length} elementos y no se podrán recuperar.`,
                  textoSi: "Vaciar papelera",
                  onSi: async () => { await vaciarPapelera(); await revisarAlmacenamiento(); setConf(null); toast("Papelera vaciada"); },
                })}>Vaciar papelera</Btn>
              </div>
            </>
          )}
        </div>
      )}

      <Confirmar open={!!conf} texto={conf?.texto} textoSi={conf?.textoSi} onSi={conf?.onSi} onNo={() => setConf(null)} />
    </div>
  );
}

/* ============================================================
   BÚSQUEDA GLOBAL
   ============================================================ */
function Buscador({ db, ir, cerrar }) {
  const [q, setQ] = useState("");
  const res = useMemo(() => {
    if (q.trim().length < 2) return [];
    const n = norm(q); const r = [];
    db.alumnos.filter((a) => norm(`${nomComp(a)} ${a.curp || ""} ${a.numLista}`).includes(n)).slice(0, 8).forEach((a) => { const g = db.grupos.find((x) => x.id === a.grupoId); r.push({ tipo: "Alumno", texto: nomComp(a), sub: [g ? `${g.grado}° ${g.grupo}` : "", `No. ${a.numLista}`, a.curp].filter(Boolean).join(" · "), ir: () => ir("ficha", { alumnoId: a.id }) }); });
    db.grupos.filter((g) => norm(`${g.grado} ${g.grupo} ${g.asignatura}`).includes(n)).forEach((g) => r.push({ tipo: "Grupo", texto: `${g.grado}° ${g.grupo}`, sub: g.asignatura, ir: () => ir("tableroGrupo", { grupoId: g.id }) }));
    db.actividades.filter((a) => norm(`${a.nombre} ${a.materia || ""}`).includes(n)).slice(0, 6).forEach((a) => r.push({ tipo: "Actividad", texto: a.nombre, sub: `${a.materia ? a.materia + " · " : ""}Entrega ${fFecha(a.fechaEntrega)}`, ir: () => ir("actividades") }));
    db.incidencias.filter((i) => norm(`${i.folio} ${i.descripcion} ${i.conducta || i.tipo} ${i.gravedad || ""}`).includes(n)).slice(0, 6).forEach((i) => r.push({ tipo: "Incidencia", texto: i.folio, sub: `${i.conducta || i.tipo} · ${i.gravedad || ""}`, ir: () => ir("incidencias", { incId: i.id }) }));
    db.bitacoras.filter((b) => norm(`${b.contenido} ${b.pda}`).includes(n)).slice(0, 5).forEach((b) => r.push({ tipo: "Bitácora", texto: b.contenido || `Sesión ${b.sesion}`, sub: fFecha(b.fecha), ir: () => ir("bitacora") }));
    if (norm("simulador ecoems").includes(n) || n.includes("ecoems") || n.includes("simulador")) r.push({ tipo: "ECOEMS", texto: "Simuladores ECOEMS", sub: `${db.ecoems.length} registros`, ir: () => ir("ecoems") });
    return r.slice(0, 20);
  }, [q, db]);

  return (
    <Modal open onClose={cerrar} title="Buscar" ancho="max-w-xl">
      <input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="Nombre, CURP, número de lista, grupo, folio, conducta…" className={inputCls} />
      <div className="mt-3 space-y-1 max-h-80 overflow-y-auto">
        {q.length >= 2 && res.length === 0 && <p className="text-sm text-slate-400 py-4 text-center">Sin resultados para "{q}".</p>}
        {res.map((r, i) => (
          <button key={i} onClick={() => { r.ir(); cerrar(); }} className="w-full flex items-center gap-3 text-left px-3 py-2 rounded-lg hover:bg-slate-50 border border-transparent hover:border-slate-200">
            <span className="text-[10px] font-medium text-slate-400 w-16 shrink-0">{r.tipo}</span>
            <span className="grow"><span className="block text-sm text-slate-800">{r.texto}</span><span className="block text-xs text-slate-400">{r.sub}</span></span>
            <ChevronRight size={15} className="text-slate-300" />
          </button>
        ))}
      </div>
    </Modal>
  );
}

/* ============================================================
   APLICACIÓN
   ============================================================ */
const MENU = [
  { id: "inicio", label: "Inicio", icon: Home },
  { id: "grupos", label: "Grupos", icon: Users },
  { id: "alumnos", label: "Alumnos", icon: GraduationCap },
  { id: "asistencia", label: "Asistencia", icon: CalendarCheck },
  { id: "permisos", label: "Permisos", icon: FileText },
  { id: "actividades", label: "Actividades", icon: BookOpen },
  { id: "evaluacion", label: "Evaluación", icon: BarChart3 },
  { id: "ecoems", label: "ECOEMS", icon: Target },
  { id: "incidencias", label: "Incidencias", icon: AlertTriangle },
  { id: "valoracion", label: "Valoración", icon: ClipboardList },
  { id: "bitacora", label: "Bitácora", icon: NotebookPen },
  { id: "estadisticas", label: "Estadísticas", icon: TrendingUp },
  { id: "reportes", label: "Reportes", icon: Printer },
  { id: "calendario", label: "Calendario", icon: Calendar },
  { id: "configuracion", label: "Configuración", icon: Settings },
];
const MENU_MOVIL = ["inicio", "asistencia", "alumnos", "incidencias", "reportes"];

export default function App() {
  const [meta, setMeta] = useState(null);
  const [db, setDb] = useState(null);
  const [cargando, setCargando] = useState(true);
  const [vista, setVista] = useState("inicio");
  const [params, setParams] = useState({});
  const [menuAbierto, setMenuAbierto] = useState(false);
  const [perfil, setPerfilEstado] = useState("Docente");
  const [buscar, setBuscar] = useState(false);
  const [aviso, setAviso] = useState(null);
  const [guardando, setGuardando] = useState(false);
  const [enLinea, setEnLinea] = useState(typeof navigator === "undefined" ? true : navigator.onLine);
  const timer = useRef(null);
  const copiaHecha = useRef(false);

  const toast = useCallback((texto, tipo = "ok") => { setAviso({ texto, tipo }); setTimeout(() => setAviso(null), 2600); }, []);

  useEffect(() => {
    const cambio = () => setEnLinea(navigator.onLine);
    window.addEventListener("online", cambio);
    window.addEventListener("offline", cambio);
    return () => { window.removeEventListener("online", cambio); window.removeEventListener("offline", cambio); };
  }, []);

  useEffect(() => {
    (async () => {
      // Pide al navegador que no borre la información aunque falte espacio.
      pedirPermanencia().catch(() => {});

      let m = await sGet(K_META);
      if (!m) {
        const ciclo = "2026-2027";
        m = { ciclos: [ciclo], activo: ciclo };
        const demo = generarDemo(ciclo);
        await sSet(K_CICLO(ciclo), demo);
        await sSet(K_META, m);
        setMeta(m); setDb(demo); setCargando(false);
        return;
      }
      let d = await sGet(K_CICLO(m.activo));
      if (!d) d = dbVacia(m.activo);
      if (m.perfil && PERFILES[m.perfil]) setPerfilEstado(m.perfil);
      setMeta(m); setDb(d); setCargando(false);
    })();
  }, []);

  // Copia automática del día, una sola vez por sesión.
  useEffect(() => {
    if (!db || copiaHecha.current) return;
    copiaHecha.current = true;
    const t = setTimeout(() => respaldoAutomatico(db.ciclo, db).catch(() => {}), 4000);
    return () => clearTimeout(t);
  }, [db?.ciclo]);

  const persistir = useCallback((nuevo) => {  // guarda en el dispositivo, sin internet
    setGuardando(true);
    clearTimeout(timer.current);
    timer.current = setTimeout(async () => {
      const ok = await sSet(K_CICLO(nuevo.ciclo), nuevo);
      setGuardando(false);
      if (!ok) toast("No se pudo guardar en este dispositivo. Descarga un respaldo desde Configuración.", "error");
    }, 600);
  }, []);

  const upd = useCallback((fn) => {
    setDb((prev) => { const n = clone(prev); fn(n); persistir(n); return n; });
  }, [persistir]);

  const ir = useCallback((v, p = {}) => { setVista(v); setParams(p); setMenuAbierto(false); window.scrollTo(0, 0); }, []);

  const setPerfil = async (nuevo) => {
    setPerfilEstado(nuevo);
    const m = { ...meta, perfil: nuevo };
    await sSet(K_META, m); setMeta(m);
    const permitidos = PERFILES[nuevo].modulos;
    if (permitidos && !permitidos.includes(vista) && vista !== "configuracion") ir("inicio");
    toast(`Perfil ${nuevo}`);
  };
  const menuVisible = PERFILES[perfil]?.modulos ? MENU.filter((m) => PERFILES[perfil].modulos.includes(m.id) || m.id === "configuracion") : MENU;

  const setCiclo = async (ciclo) => {
    const d = (await sGet(K_CICLO(ciclo))) || dbVacia(ciclo);
    const m = { ...meta, activo: ciclo };
    await sSet(K_META, m);
    setMeta(m); setDb(d); ir("inicio"); toast(`Ciclo ${ciclo} abierto`);
  };
  const nuevoCiclo = async (ciclo) => {
    if (meta.ciclos.includes(ciclo)) return toast("Ese ciclo ya existe", "error");
    const d = dbVacia(ciclo);
    d.config = { ...clone(db.config), ciclo, trimestre: 1 };
    d.config.trimestres = configBase(ciclo).trimestres;
    d.catalogos = clone(db.catalogos);
    const m = { ciclos: [...meta.ciclos, ciclo], activo: ciclo };
    await sSet(K_CICLO(ciclo), d); await sSet(K_META, m);
    setMeta(m); setDb(d); ir("inicio"); toast(`Ciclo ${ciclo} creado`);
  };
  const restaurar = async (nuevo) => {
    const m = meta.ciclos.includes(nuevo.ciclo) ? { ...meta, activo: nuevo.ciclo } : { ciclos: [...meta.ciclos, nuevo.ciclo], activo: nuevo.ciclo };
    await sSet(K_CICLO(nuevo.ciclo), nuevo); await sSet(K_META, m);
    setMeta(m); setDb(nuevo); ir("inicio"); toast("Respaldo restaurado");
  };

  if (cargando || !db) return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center">
      <div className="text-center"><p className="text-slate-800 font-medium">Control Docente Integral</p><p className="text-sm text-slate-400 mt-1">Cargando información…</p></div>
    </div>
  );

  const props = { db, upd, ir, toast, params };
  const pantallas = {
    inicio: <Inicio db={db} ir={ir} />,
    grupos: <Grupos {...props} />,
    alumnos: <Alumnos {...props} />,
    ficha: <Ficha {...props} />,
    tableroGrupo: <TableroGrupo db={db} ir={ir} params={params} />,
    asistencia: <Asistencia db={db} upd={upd} toast={toast} />,
    permisos: <Permisos db={db} upd={upd} toast={toast} />,
    actividades: <Actividades db={db} upd={upd} toast={toast} />,
    evaluacion: <Evaluacion {...props} />,
    ecoems: <Ecoems {...props} />,
    incidencias: <Incidencias {...props} />,
    valoracion: <Valoracion {...props} />,
    bitacora: <Bitacora {...props} />,
    estadisticas: <Estadisticas db={db} ir={ir} />,
    reportes: <Reportes db={db} toast={toast} />,
    calendario: <Calendario db={db} upd={upd} ir={ir} toast={toast} />,
    configuracion: <Configuracion db={db} upd={upd} meta={meta} setCiclo={setCiclo} nuevoCiclo={nuevoCiclo} restaurar={restaurar} toast={toast} perfil={perfil} setPerfil={setPerfil} />,
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900" style={{ fontFamily: "ui-sans-serif, system-ui, -apple-system, sans-serif" }}>
      {/* Barra superior */}
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200">
        <div className="flex items-center gap-2 px-3 sm:px-5 h-14">
          <button onClick={() => setMenuAbierto(!menuAbierto)} className="lg:hidden p-2 rounded-lg hover:bg-slate-100 text-slate-600"><Menu size={20} /></button>
          <div className="grow min-w-0">
            <p className="font-semibold text-slate-900 leading-tight truncate">Control Docente Integral</p>
            <p className="text-[11px] text-slate-400 truncate">{db.config.escuela || "Configura tu escuela"} · Ciclo {db.ciclo}</p>
          </div>
          {!enLinea && (
            <span title="La app funciona sin internet" className="inline-flex items-center gap-1 text-[11px] text-slate-500 bg-slate-100 border border-slate-200 rounded-md px-1.5 py-1">
              <WifiOff size={12} />
              <span className="hidden sm:inline">Sin conexión</span>
            </span>
          )}
          <span className={`text-[11px] hidden sm:inline ${guardando ? "text-slate-400" : "text-emerald-600"}`}>
            {guardando ? "Guardando…" : "Guardado"}
          </span>
          <button onClick={() => setBuscar(true)} className="p-2 rounded-lg hover:bg-slate-100 text-slate-600"><Search size={19} /></button>
        </div>
      </header>

      <div className="flex">
        {/* Menú lateral */}
        <aside className={`${menuAbierto ? "block" : "hidden"} lg:block fixed lg:sticky top-14 left-0 z-20 w-60 h-[calc(100vh-3.5rem)] bg-white border-r border-slate-200 overflow-y-auto shrink-0`}>
          <div className="px-4 pt-3 pb-1">
            <p className="text-[10px] text-slate-400">Perfil</p>
            <p className="text-xs font-medium text-slate-700">{perfil}</p>
          </div>
          <nav className="p-2 space-y-0.5">
            {menuVisible.map((m) => (
              <button key={m.id} onClick={() => ir(m.id)}
                className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm transition-colors ${vista === m.id || (vista === "ficha" && m.id === "alumnos") || (vista === "tableroGrupo" && m.id === "grupos") ? "bg-emerald-50 text-emerald-800 font-medium" : "text-slate-600 hover:bg-slate-50"}`}>
                <m.icon size={17} className={vista === m.id ? "text-emerald-700" : "text-slate-400"} />{m.label}
              </button>
            ))}
          </nav>
        </aside>
        {menuAbierto && <div className="lg:hidden fixed inset-0 top-14 bg-slate-900/30 z-10" onClick={() => setMenuAbierto(false)} />}

        {/* Contenido */}
        <main className="grow min-w-0 p-3 sm:p-5 pb-24 lg:pb-8 max-w-6xl">{pantallas[vista] || pantallas.inicio}</main>
      </div>

      {/* Navegación inferior en celular */}
      <nav className="lg:hidden fixed bottom-0 inset-x-0 z-30 bg-white border-t border-slate-200 flex">
        {MENU_MOVIL.filter((id) => menuVisible.some((x) => x.id === id)).map((id) => {
          const m = MENU.find((x) => x.id === id);
          return (
            <button key={id} onClick={() => ir(id)} className={`grow flex flex-col items-center gap-0.5 py-2 ${vista === id ? "text-emerald-700" : "text-slate-400"}`}>
              <m.icon size={19} /><span className="text-[10px]">{m.label}</span>
            </button>
          );
        })}
      </nav>

      {buscar && <Buscador db={db} ir={ir} cerrar={() => setBuscar(false)} />}

      {aviso && (
        <div className="fixed bottom-20 lg:bottom-6 left-1/2 -translate-x-1/2 z-50">
          <div className={`px-4 py-2.5 rounded-xl text-sm text-white shadow-lg ${aviso.tipo === "error" ? "bg-rose-600" : "bg-slate-800"}`}>{aviso.texto}</div>
        </div>
      )}
    </div>
  );
}
