/* Adaptateur : API serveur + révisions + FUSION AUTOMATIQUE des modifications simultanées */
const __revs = {};
const __last = {};   // dernier état connu du serveur, pour fusionner intelligemment
const clone = (v) => v === undefined ? v : JSON.parse(JSON.stringify(v));
const signal = (state, detail) => window.dispatchEvent(new CustomEvent("crm-save", { detail: { state, ...detail } }));

/* Fusion à 3 sources : base (notre dernier état connu), ours (notre modification), theirs (l'état actuel du serveur).
   On ré-applique UNIQUEMENT ce que NOUS avons changé, par-dessus la version la plus récente. */
function merge3(base, ours, theirs) {
  const isIdArray = (v) => Array.isArray(v) && v.every((x) => x && typeof x === "object" && "id" in x);
  const J = JSON.stringify;
  if (isIdArray(ours) && isIdArray(theirs) && isIdArray(base || [])) {
    const baseMap = new Map((base || []).map((x) => [x.id, x]));
    const oursMap = new Map(ours.map((x) => [x.id, x]));
    let result = theirs.map((t) => {
      const o = oursMap.get(t.id), b = baseMap.get(t.id);
      /* nous l'avons modifié → notre version gagne pour CET élément */
      if (o && (!b || J(o) !== J(b))) return o;
      return t;
    });
    /* éléments que nous avons ajoutés */
    const theirIds = new Set(theirs.map((x) => x.id));
    ours.forEach((o) => { if (!theirIds.has(o.id) && !baseMap.has(o.id)) result.push(o); });
    /* éléments que nous avons supprimés */
    const oursIds = new Set(ours.map((x) => x.id));
    result = result.filter((x) => oursIds.has(x.id) || !baseMap.has(x.id));
    return result;
  }
  if (ours && theirs && typeof ours === "object" && typeof theirs === "object" && !Array.isArray(ours) && !Array.isArray(theirs)) {
    const b = base && typeof base === "object" ? base : {};
    const result = { ...theirs };
    /* clés que nous avons modifiées ou ajoutées */
    Object.keys(ours).forEach((k) => { if (J(ours[k]) !== J(b[k])) result[k] = ours[k]; });
    /* clés que nous avons supprimées */
    Object.keys(b).forEach((k) => { if (!(k in ours) && k in result && J(result[k]) === J(b[k])) delete result[k]; });
    return result;
  }
  return ours; /* types simples : notre valeur */
}

async function rawPut(key, value, rev) {
  const headers = { "Content-Type": "application/json" };
  if (rev !== undefined) headers["X-Rev"] = rev;
  return fetch(`/api/storage/${encodeURIComponent(key)}`, { method: "PUT", headers, body: JSON.stringify({ value }) });
}

window.storage = {
  async get(key) {
    const r = await fetch(`/api/storage/${encodeURIComponent(key)}`);
    if (!r.ok) throw new Error("key not found: " + key);
    const rev = r.headers.get("X-Rev");
    if (rev !== null) __revs[key] = rev;
    const data = await r.json();
    __last[key] = clone(data.value);
    return data;
  },
  async set(key, value) {
    signal("saving");
    try {
      let toSave = value;
      let r = await rawPut(key, toSave, __revs[key]);
      /* Conflit : quelqu'un a modifié cette donnée entre-temps → fusion automatique + nouvel essai */
      let tries = 0;
      while (r.status === 409 && tries < 3) {
        tries++;
        const latest = await fetch(`/api/storage/${encodeURIComponent(key)}`);
        const latestRev = latest.headers.get("X-Rev");
        const theirs = latest.ok ? (await latest.json()).value : undefined;
        toSave = merge3(__last[key], value, theirs);
        r = await rawPut(key, toSave, latestRev);
      }
      if (!r.ok) throw new Error("HTTP " + r.status);
      const rev = r.headers.get("X-Rev");
      if (rev !== null) __revs[key] = rev;
      __last[key] = clone(toSave);
      signal("saved");
      return { key, value: toSave };
    } catch (e) {
      signal("error", { key });
      throw e;
    }
  },
  async delete(key) {
    const r = await fetch(`/api/storage/${encodeURIComponent(key)}`, { method: "DELETE" });
    delete __revs[key]; delete __last[key];
    return await r.json();
  },
  async list(prefix) {
    const r = await fetch(`/api/storage?prefix=${encodeURIComponent(prefix || "")}`);
    return await r.json();
  },
};

window.addEventListener("offline", () => signal("offline"));
window.addEventListener("online", () => signal("online"));

import React from "react";
import { createRoot } from "react-dom/client";
import App from "./CRM.jsx";
createRoot(document.getElementById("root")).render(<App />);
