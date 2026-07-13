/* Adaptateur : API serveur + révisions anti-écrasement + indicateur d'enregistrement */
const __revs = {};
const signal = (state, detail) => window.dispatchEvent(new CustomEvent("crm-save", { detail: { state, ...detail } }));

window.storage = {
  async get(key) {
    const r = await fetch(`/api/storage/${encodeURIComponent(key)}`);
    if (!r.ok) throw new Error("key not found: " + key);
    const rev = r.headers.get("X-Rev");
    if (rev !== null) __revs[key] = rev;
    return await r.json();
  },
  async set(key, value) {
    signal("saving");
    try {
      const headers = { "Content-Type": "application/json" };
      if (__revs[key] !== undefined) headers["X-Rev"] = __revs[key];
      const r = await fetch(`/api/storage/${encodeURIComponent(key)}`, {
        method: "PUT", headers, body: JSON.stringify({ value }),
      });
      if (r.status === 409) {
        /* Quelqu'un d'autre a modifié cette donnée depuis notre lecture */
        signal("conflict", { key });
        throw new Error("conflict");
      }
      if (!r.ok) throw new Error("HTTP " + r.status);
      const rev = r.headers.get("X-Rev");
      if (rev !== null) __revs[key] = rev;
      signal("saved");
      return await r.json();
    } catch (e) {
      if (e.message !== "conflict") signal("error", { key });
      throw e;
    }
  },
  async delete(key) {
    const r = await fetch(`/api/storage/${encodeURIComponent(key)}`, { method: "DELETE" });
    delete __revs[key];
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
