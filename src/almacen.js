/* ============================================================
   ALMACENAMIENTO PERMANENTE
   Los datos viven en IndexedDB dentro del dispositivo.
   - Se guardan aunque no haya internet.
   - No se borran al cerrar el navegador ni al limpiar el caché.
   - Solo se eliminan si la persona lo pide expresamente.
   - Cada eliminación pasa antes por la papelera.
   - Se conservan respaldos automáticos de los últimos días.
   ============================================================ */

const BD = "control-docente-integral";
const VERSION = 1;
const ALMACEN = "datos";       // ciclos escolares y configuración general
const PAPELERA = "papelera";   // registros eliminados, recuperables
const RESPALDOS = "respaldos"; // copias automáticas por día

let conexion = null;

function abrir() {
  if (conexion) return conexion;
  conexion = new Promise((resolve, reject) => {
    const req = indexedDB.open(BD, VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(ALMACEN)) db.createObjectStore(ALMACEN);
      if (!db.objectStoreNames.contains(PAPELERA)) db.createObjectStore(PAPELERA, { keyPath: "id" });
      if (!db.objectStoreNames.contains(RESPALDOS)) db.createObjectStore(RESPALDOS, { keyPath: "id" });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
    req.onblocked = () => reject(new Error("La base de datos está abierta en otra pestaña"));
  });
  return conexion;
}

function tx(almacen, modo, fn) {
  return abrir().then(
    (db) =>
      new Promise((resolve, reject) => {
        const t = db.transaction(almacen, modo);
        const pedido = fn(t.objectStore(almacen));
        t.oncomplete = () => resolve(pedido?.result);
        t.onerror = () => reject(t.error);
        t.onabort = () => reject(t.error);
      })
  );
}

/* ---------- lectura y escritura principales ---------- */

export async function leer(clave) {
  try {
    const v = await tx(ALMACEN, "readonly", (s) => s.get(clave));
    return v ?? null;
  } catch (e) {
    console.error("No se pudo leer", clave, e);
    return respaldoLocal(clave);
  }
}

export async function guardar(clave, valor) {
  try {
    await tx(ALMACEN, "readwrite", (s) => s.put(valor, clave));
    espejoLocal(clave, valor);
    return { ok: true };
  } catch (e) {
    console.error("No se pudo guardar", clave, e);
    const espejo = espejoLocal(clave, valor);
    return { ok: false, error: e, espejo };
  }
}

export async function claves() {
  try {
    return (await tx(ALMACEN, "readonly", (s) => s.getAllKeys())) || [];
  } catch {
    return [];
  }
}

/* ---------- copia de emergencia en el propio navegador ----------
   Si IndexedDB fallara en algún dispositivo, queda una copia reducida
   para no perder el trabajo del día. Es un salvavidas, no el almacén.  */

function espejoLocal(clave, valor) {
  try {
    localStorage.setItem("cdi:espejo:" + clave, JSON.stringify(valor));
    return true;
  } catch {
    return false; // el espejo puede no caber; el almacén principal es IndexedDB
  }
}
function respaldoLocal(clave) {
  try {
    const v = localStorage.getItem("cdi:espejo:" + clave);
    return v ? JSON.parse(v) : null;
  } catch {
    return null;
  }
}

/* ---------- permiso de almacenamiento permanente ----------
   Sin este permiso, el navegador podría liberar espacio y borrar datos
   cuando el teléfono se queda sin memoria. Con él, no los toca.        */

export async function pedirPermanencia() {
  try {
    if (!navigator.storage?.persist) return { permanente: false, soportado: false };
    let permanente = await navigator.storage.persisted();
    if (!permanente) permanente = await navigator.storage.persist();
    return { permanente, soportado: true };
  } catch {
    return { permanente: false, soportado: false };
  }
}

export async function espacio() {
  try {
    if (!navigator.storage?.estimate) return null;
    const { usage = 0, quota = 0 } = await navigator.storage.estimate();
    return {
      usadoMB: Math.round((usage / 1048576) * 10) / 10,
      disponibleMB: Math.round(quota / 1048576),
      porcentaje: quota ? Math.round((usage / quota) * 1000) / 10 : 0,
    };
  } catch {
    return null;
  }
}

/* ---------- papelera: nada se borra de golpe ---------- */

export async function aPapelera(registro) {
  const item = {
    id: "pap_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7),
    fecha: new Date().toISOString(),
    ...registro, // { tipo, descripcion, ciclo, datos }
  };
  try {
    await tx(PAPELERA, "readwrite", (s) => s.put(item));
    return item;
  } catch (e) {
    console.error("No se pudo enviar a la papelera", e);
    return null;
  }
}

export async function verPapelera() {
  try {
    const items = (await tx(PAPELERA, "readonly", (s) => s.getAll())) || [];
    return items.sort((a, b) => b.fecha.localeCompare(a.fecha));
  } catch {
    return [];
  }
}

export async function sacarDePapelera(id) {
  const items = await verPapelera();
  const item = items.find((x) => x.id === id);
  if (item) await tx(PAPELERA, "readwrite", (s) => s.delete(id));
  return item || null;
}

export async function borrarDePapelera(id) {
  try {
    await tx(PAPELERA, "readwrite", (s) => s.delete(id));
    return true;
  } catch {
    return false;
  }
}

export async function vaciarPapelera() {
  try {
    await tx(PAPELERA, "readwrite", (s) => s.clear());
    return true;
  } catch {
    return false;
  }
}

/* ---------- respaldos automáticos ----------
   Una copia por día, se conservan las últimas 10.                     */

const MAX_RESPALDOS = 10;

export async function respaldoAutomatico(ciclo, datos) {
  const dia = new Date().toISOString().slice(0, 10);
  const id = `${ciclo}__${dia}`;
  try {
    await tx(RESPALDOS, "readwrite", (s) =>
      s.put({ id, ciclo, fecha: dia, creado: new Date().toISOString(), datos })
    );
    const todos = (await tx(RESPALDOS, "readonly", (s) => s.getAll())) || [];
    const delCiclo = todos.filter((r) => r.ciclo === ciclo).sort((a, b) => b.fecha.localeCompare(a.fecha));
    for (const viejo of delCiclo.slice(MAX_RESPALDOS)) {
      await tx(RESPALDOS, "readwrite", (s) => s.delete(viejo.id));
    }
    return true;
  } catch (e) {
    console.error("No se pudo crear el respaldo automático", e);
    return false;
  }
}

export async function verRespaldos() {
  try {
    const todos = (await tx(RESPALDOS, "readonly", (s) => s.getAll())) || [];
    return todos
      .map((r) => ({
        id: r.id,
        ciclo: r.ciclo,
        fecha: r.fecha,
        creado: r.creado,
        alumnos: r.datos?.alumnos?.length || 0,
        registros:
          (r.datos?.asistencias?.length || 0) +
          (r.datos?.actividades?.length || 0) +
          (r.datos?.incidencias?.length || 0) +
          (r.datos?.ecoems?.length || 0) +
          (r.datos?.bitacoras?.length || 0),
      }))
      .sort((a, b) => b.fecha.localeCompare(a.fecha));
  } catch {
    return [];
  }
}

export async function leerRespaldo(id) {
  try {
    const r = await tx(RESPALDOS, "readonly", (s) => s.get(id));
    return r?.datos || null;
  } catch {
    return null;
  }
}

/* ---------- exportación completa ---------- */

export async function exportarTodo() {
  const ks = await claves();
  const datos = {};
  for (const k of ks) datos[k] = await leer(k);
  return {
    aplicacion: "Control Docente Integral",
    version: 2,
    generado: new Date().toISOString(),
    datos,
  };
}
