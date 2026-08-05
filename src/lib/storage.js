/**
 * Reemplazo de window.storage (API de Claude) usando Google Apps Script
 * como backend, que a su vez lee/escribe en una Google Sheet.
 *
 * Mantiene la MISMA firma que window.storage, así que en
 * panel-adjudicaciones.jsx solo hace falta:
 *   1) agregar: import { storage } from "./lib/storage";
 *   2) reemplazar todas las apariciones de "window.storage" por "storage"
 *
 * Requiere la variable de entorno VITE_SHEETS_API_URL apuntando a la URL
 * del Web App de Apps Script (ver Code.gs).
 */

const API_URL = import.meta.env.VITE_SHEETS_API_URL;

async function get(key) {
  if (!API_URL) throw new Error("Falta configurar VITE_SHEETS_API_URL");
  const res = await fetch(`${API_URL}?action=get&key=${encodeURIComponent(key)}`);
  const data = await res.json();
  if (!data.ok) return null;
  return { key: data.key, value: data.value };
}

async function set(key, value) {
  if (!API_URL) throw new Error("Falta configurar VITE_SHEETS_API_URL");
  const res = await fetch(API_URL, {
    method: "POST",
    // text/plain evita el preflight CORS (OPTIONS) que Apps Script no maneja bien
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "set", key, value }),
  });
  const data = await res.json();
  return data.ok ? { key: data.key, value: data.value } : null;
}

async function del(key) {
  if (!API_URL) throw new Error("Falta configurar VITE_SHEETS_API_URL");
  const res = await fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "delete", key }),
  });
  const data = await res.json();
  return data.ok ? { key, deleted: true } : null;
}

async function list(prefix = "") {
  if (!API_URL) throw new Error("Falta configurar VITE_SHEETS_API_URL");
  const res = await fetch(`${API_URL}?action=list&prefix=${encodeURIComponent(prefix)}`);
  const data = await res.json();
  return data.ok ? { keys: data.keys } : null;
}

export const storage = { get, set, delete: del, list };