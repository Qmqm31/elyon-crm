/* Adaptateur : remplace window.storage (Claude) par l'API du serveur */
window.storage = {
  async get(key) {
    const r = await fetch(`/api/storage/${encodeURIComponent(key)}`);
    if (!r.ok) throw new Error("key not found: " + key);
    return await r.json();
  },
  async set(key, value) {
    const r = await fetch(`/api/storage/${encodeURIComponent(key)}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ value }),
    });
    return await r.json();
  },
  async delete(key) {
    const r = await fetch(`/api/storage/${encodeURIComponent(key)}`, { method: "DELETE" });
    return await r.json();
  },
  async list(prefix) {
    const r = await fetch(`/api/storage?prefix=${encodeURIComponent(prefix || "")}`);
    return await r.json();
  },
};

import React from "react";
import { createRoot } from "react-dom/client";
import App from "./CRM.jsx";
createRoot(document.getElementById("root")).render(<App />);
