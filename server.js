/* ============ Serveur CRM ELYON & Associés ============ */
const express = require("express");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "data");
const STORE_FILE = path.join(DATA_DIR, "storage.json");
const ACCESS_CODE = process.env.ACCESS_CODE || "";

/* ---- Protection par code d'accès (HTTP Basic) ---- */
app.use((req, res, next) => {
  if (!ACCESS_CODE) return next(); // pas de code défini = accès libre (déconseillé)
  const hdr = req.headers.authorization || "";
  const [scheme, b64] = hdr.split(" ");
  if (scheme === "Basic" && b64) {
    const pwd = Buffer.from(b64, "base64").toString().split(":").slice(1).join(":");
    if (pwd === ACCESS_CODE) return next();
  }
  res.set("WWW-Authenticate", 'Basic realm="CRM ELYON"');
  res.status(401).send("Accès protégé — CRM ELYON & Associés");
});

app.use(express.json({ limit: "25mb" }));

/* ---- Stockage clé/valeur dans un fichier JSON persistant ---- */
fs.mkdirSync(DATA_DIR, { recursive: true });
let store = {};
try { store = JSON.parse(fs.readFileSync(STORE_FILE, "utf8")); } catch { store = {}; }
let writing = Promise.resolve();
function persist() {
  writing = writing.then(() => new Promise((resolve) => {
    const tmp = STORE_FILE + ".tmp";
    fs.writeFile(tmp, JSON.stringify(store), (err) => {
      if (!err) fs.rename(tmp, STORE_FILE, () => resolve());
      else resolve();
    });
  }));
}

app.get("/api/storage", (req, res) => {
  const prefix = req.query.prefix || "";
  res.json({ keys: Object.keys(store).filter((k) => k.startsWith(prefix)) });
});
app.get("/api/storage/:key", (req, res) => {
  const k = req.params.key;
  if (!(k in store)) return res.status(404).json({ error: "not found" });
  res.json({ key: k, value: store[k] });
});
app.put("/api/storage/:key", (req, res) => {
  store[req.params.key] = req.body.value;
  persist();
  res.json({ key: req.params.key, value: req.body.value });
});
app.delete("/api/storage/:key", (req, res) => {
  delete store[req.params.key];
  persist();
  res.json({ key: req.params.key, deleted: true });
});

/* ---- Fichiers statiques ---- */
app.use(express.static(path.join(__dirname, "dist")));
app.get("/favicon.png", (_, res) => res.sendFile(path.join(__dirname, "favicon.png")));
app.get("/", (_, res) => res.sendFile(path.join(__dirname, "index.html")));

app.listen(PORT, () => console.log(`CRM ELYON en ligne sur le port ${PORT}`));
