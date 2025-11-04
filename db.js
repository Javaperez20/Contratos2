// db.js - IndexedDB helper para almacenar último contrato (.docx blob) y settings como Ejecutivo
// Guardamos un objeto { blob, filename } bajo la clave 'ultimoContrato' para poder mantener el nombre del archivo.
// Mantenemos compatibilidad con versiones previas que guardaban únicamente el Blob.

const DB_NAME = 'ContratoDB';
const STORE_NAME = 'Contratos';

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = (ev) => {
      const db = ev.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

/**
 * saveContrato(blob, filename)
 * - Guarda en IndexedDB un objeto con blob y filename.
 * - Si se llama con un único argumento (blob), filename será 'Contrato.docx'.
 */
async function saveContrato(blob, filename) {
  const db = await openDB();
  const meta = {
    blob: blob,
    filename: String(filename || 'Contrato.docx')
  };
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const req = store.put(meta, 'ultimoContrato');
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error);
  });
}

/**
 * getContrato()
 * - Devuelve { blob, filename } o null si no existe.
 * - Si en la DB había solamente un Blob (versiones previas), lo normaliza a { blob, filename }.
 */
async function getContrato() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    const req = store.get('ultimoContrato');
    req.onsuccess = () => {
      const res = req.result;
      if (!res) return resolve(null);
      if (res instanceof Blob) {
        // compatibilidad con valor antiguo
        resolve({ blob: res, filename: 'Contrato.docx' });
      } else if (res && res.blob) {
        resolve({ blob: res.blob, filename: res.filename || 'Contrato.docx' });
      } else {
        resolve(null);
      }
    };
    req.onerror = () => reject(req.error);
  });
}

// --- Nuevas utilidades para Ejecutivo (persistencia y borrado) ---
async function saveEjecutivo(name) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const req = store.put(String(name || ''), 'ejecutivo');
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error);
  });
}

async function getEjecutivo() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    const req = store.get('ejecutivo');
    req.onsuccess = () => resolve(req.result || '');
    req.onerror = () => reject(req.error);
  });
}

async function deleteEjecutivo() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const req = store.delete('ejecutivo');
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error);
  });
}