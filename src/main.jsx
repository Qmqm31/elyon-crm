/* ============ Adaptateur de stockage v3 — AUCUNE DONNÉE NE DOIT JAMAIS SE PERDRE ============
   1. Fusion automatique des modifications simultanées (3 essais, puis écriture forcée)
   2. File d'attente : si le réseau échoue, la donnée est mise en attente et renvoyée
      automatiquement toutes les 4 secondes jusqu'au succès
   3. Copie de secours locale : la file d'attente survit à la fermeture du navigateur
      et est rejouée au prochain lancement
   4. Avertissement avant de fermer l'onglet si un enregistrement est en attente        */

const __revs = {};
const __last = {};
const clone = (v) => v === undefined ? v : JSON.parse(JSON.stringify(v));
const signal = (state, detail) => window.dispatchEvent(new CustomEvent("crm-save", { detail: { state, ...detail } }));

/* ---- Fusion à 3 sources : on ré-applique uniquement NOS changements sur la version la plus récente ---- */
function merge3(base, ours, theirs) {
  const isIdArray = (v) => Array.isArray(v) && v.every((x) => x && typeof x === "object" && "id" in x);
  const J = JSON.stringify;
  if (isIdArray(ours) && isIdArray(theirs) && isIdArray(base || [])) {
    const baseMap = new Map((base || []).map((x) => [x.id, x]));
    const oursMap = new Map(ours.map((x) => [x.id, x]));
    let result = theirs.map((t) => {
      const o = oursMap.get(t.id), b = baseMap.get(t.id);
      if (o && (!b || J(o) !== J(b))) return o;
      return t;
    });
    const theirIds = new Set(theirs.map((x) => x.id));
    ours.forEach((o) => { if (!theirIds.has(o.id) && !baseMap.has(o.id)) result.push(o); });
    const oursIds = new Set(ours.map((x) => x.id));
    result = result.filter((x) => oursIds.has(x.id) || !baseMap.has(x.id));
    return result;
  }
  if (ours && theirs && typeof ours === "object" && typeof theirs === "object" && !Array.isArray(ours) && !Array.isArray(theirs)) {
    const b = base && typeof base === "object" ? base : {};
    const result = { ...theirs };
    Object.keys(ours).forEach((k) => { if (J(ours[k]) !== J(b[k])) result[k] = ours[k]; });
    Object.keys(b).forEach((k) => { if (!(k in ours) && k in result && J(result[k]) === J(b[k])) delete result[k]; });
    return result;
  }
  return ours;
}

async function rawPut(key, value, rev) {
  const headers = { "Content-Type": "application/json" };
  if (rev !== undefined && rev !== null) headers["X-Rev"] = rev;
  return fetch(`/api/storage/${encodeURIComponent(key)}`, { method: "PUT", headers, body: JSON.stringify({ value }) });
}

/* ---- Écriture robuste : fusion sur conflit, puis écriture inconditionnelle en dernier recours ---- */
async function doSet(key, value) {
  let toSave = value;
  let r = await rawPut(key, toSave, __revs[key]);
  let tries = 0;
  while (r.status === 409 && tries < 3) {
    tries++;
    const latest = await fetch(`/api/storage/${encodeURIComponent(key)}`);
    const latestRev = latest.headers.get("X-Rev");
    const theirs = latest.ok ? (await latest.json()).value : undefined;
    toSave = merge3(__last[key], value, theirs);
    r = await rawPut(key, toSave, latestRev);
  }
  /* dernier recours : écriture sans condition — la donnée DOIT être enregistrée */
  if (r.status === 409) r = await rawPut(key, toSave, undefined);
  if (!r.ok) throw new Error("HTTP " + r.status);
  const rev = r.headers.get("X-Rev");
  if (rev !== null) __revs[key] = rev;
  __last[key] = clone(toSave);
  return toSave;
}

/* ---- File d'attente persistante (survit à la fermeture du navigateur) ---- */
const pending = new Map();
const PENDING_KEY = "crm-attente-enregistrement";
function persistPending() {
  try {
    if (pending.size) localStorage.setItem(PENDING_KEY, JSON.stringify([...pending.entries()]));
    else localStorage.removeItem(PENDING_KEY);
  } catch { /* stockage local plein : la file reste en mémoire */ }
}
let flushing = false;
async function flushPending() {
  if (flushing || !pending.size) return;
  flushing = true;
  for (const [key, value] of [...pending.entries()]) {
    try {
      await doSet(key, value);
      /* ne retirer que si la valeur n'a pas changé entre-temps */
      if (pending.get(key) === value) pending.delete(key);
    } catch { break; /* réseau toujours indisponible : on réessaiera */ }
  }
  persistPending();
  flushing = false;
  if (!pending.size) signal("saved");
}
setInterval(flushPending, 4000);
window.addEventListener("online", flushPending);

/* Rejouer la file d'attente d'une session précédente (données saisies puis navigateur fermé) */
try {
  const saved = JSON.parse(localStorage.getItem(PENDING_KEY) || "[]");
  saved.forEach(([k, v]) => pending.set(k, v));
  if (pending.size) setTimeout(flushPending, 1500);
} catch { }

/* Avertir avant de fermer l'onglet si des enregistrements sont en attente */
window.addEventListener("beforeunload", (e) => {
  if (pending.size) {
    e.preventDefault();
    e.returnValue = "Des enregistrements sont en attente d'envoi — attendez quelques secondes.";
    return e.returnValue;
  }
});

window.storage = {
  async get(key) {
    const r = await fetch(`/api/storage/${encodeURIComponent(key)}`);
    if (!r.ok) throw new Error("key not found: " + key);
    const rev = r.headers.get("X-Rev");
    if (rev !== null) __revs[key] = rev;
    const data = await r.json();
    __last[key] = clone(data.value);
    /* une valeur en attente pour cette clé prime sur celle du serveur */
    if (pending.has(key)) return { key, value: pending.get(key) };
    return data;
  },
  async set(key, value) {
    signal("saving");
    try {
      const saved = await doSet(key, value);
      /* purger une éventuelle attente devenue obsolète */
      if (pending.has(key)) { pending.delete(key); persistPending(); }
      signal("saved");
      return { key, value: saved };
    } catch (e) {
      /* réseau indisponible : mise en file d'attente — la donnée N'EST PAS perdue */
      pending.set(key, value);
      persistPending();
      signal(navigator.onLine ? "queued" : "offline", { key, pending: pending.size });
      return { key, value, queued: true };
    }
  },
  async delete(key) {
    try {
      const r = await fetch(`/api/storage/${encodeURIComponent(key)}`, { method: "DELETE" });
      delete __revs[key]; delete __last[key];
      pending.delete(key); persistPending();
      return await r.json();
    } catch { return { key, deleted: false }; }
  },
  async list(prefix) {
    const r = await fetch(`/api/storage?prefix=${encodeURIComponent(prefix || "")}`);
    return await r.json();
  },
};

window.addEventListener("offline", () => signal("offline"));

import React from "react";
import { createRoot } from "react-dom/client";
import App from "./CRM.jsx";
createRoot(document.getElementById("root")).render(<App />);
