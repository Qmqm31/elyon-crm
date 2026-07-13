/* ============ Serveur CRM ELYON & Associés — v2 ============ */
const express = require("express");
const compression = require("compression");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "data");
const STORE_FILE = path.join(DATA_DIR, "storage.json");
const REVS_FILE = path.join(DATA_DIR, "revs.json");
const SNAP_DIR = path.join(DATA_DIR, "snapshots");
const ACCESS_CODE = process.env.ACCESS_CODE || "";

/* ---- Compression (réseau ~4x plus léger) ---- */
app.use(compression());

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
fs.mkdirSync(SNAP_DIR, { recursive: true });
let store = {};
try { store = JSON.parse(fs.readFileSync(STORE_FILE, "utf8")); } catch { store = {}; }
let revs = {};
try { revs = JSON.parse(fs.readFileSync(REVS_FILE, "utf8")); } catch { revs = {}; }

let writing = Promise.resolve();
function persist() {
  writing = writing.then(() => new Promise((resolve) => {
    const tmp = STORE_FILE + ".tmp";
    fs.writeFile(tmp, JSON.stringify(store), (err) => {
      if (!err) fs.rename(tmp, STORE_FILE, () => resolve());
      else resolve();
    });
  })).then(() => new Promise((resolve) => {
    fs.writeFile(REVS_FILE, JSON.stringify(revs), () => resolve());
  }));
}

/* ---- Instantanés quotidiens automatiques (30 jours conservés) ---- */
const todayISO = () => new Date().toISOString().slice(0, 10);
function makeSnapshot(label) {
  const name = label || todayISO();
  const file = path.join(SNAP_DIR, name + ".json");
  try { fs.writeFileSync(file, JSON.stringify(store)); } catch (e) { console.error("snapshot:", e.message); }
  try {
    const files = fs.readdirSync(SNAP_DIR).filter((f) => f.endsWith(".json")).sort();
    while (files.length > 35) fs.unlinkSync(path.join(SNAP_DIR, files.shift()));
  } catch { }
}
function ensureDailySnapshot() {
  const file = path.join(SNAP_DIR, todayISO() + ".json");
  if (!fs.existsSync(file) && Object.keys(store).length) makeSnapshot();
}
ensureDailySnapshot();
setInterval(ensureDailySnapshot, 30 * 60 * 1000);

app.get("/api/snapshots", (req, res) => {
  try {
    const files = fs.readdirSync(SNAP_DIR).filter((f) => f.endsWith(".json")).sort().reverse();
    res.json(files.map((f) => {
      const st = fs.statSync(path.join(SNAP_DIR, f));
      return { name: f.replace(".json", ""), size: st.size };
    }));
  } catch { res.json([]); }
});
app.post("/api/snapshots/restore", (req, res) => {
  const name = String(req.body.name || "").replace(/[^0-9A-Za-z_\-h]/g, "");
  const file = path.join(SNAP_DIR, name + ".json");
  if (!fs.existsSync(file)) return res.status(404).json({ error: "snapshot introuvable" });
  try {
    const now = new Date();
    makeSnapshot(`${todayISO()}_avant-restauration_${String(now.getHours()).padStart(2, "0")}h${String(now.getMinutes()).padStart(2, "0")}`);
    store = JSON.parse(fs.readFileSync(file, "utf8"));
    Object.keys(store).forEach((k) => { revs[k] = (revs[k] || 0) + 1; });
    persist();
    res.json({ ok: true, restored: name });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

/* ---- API stockage (avec numéro de révision anti-écrasement) ---- */
app.get("/api/storage", (req, res) => {
  const prefix = req.query.prefix || "";
  res.json({ keys: Object.keys(store).filter((k) => k.startsWith(prefix)) });
});
app.get("/api/storage/:key", (req, res) => {
  const k = req.params.key;
  if (!(k in store)) return res.status(404).json({ error: "not found" });
  res.set("X-Rev", String(revs[k] || 0));
  res.json({ key: k, value: store[k] });
});
app.put("/api/storage/:key", (req, res) => {
  const k = req.params.key;
  const clientRev = req.headers["x-rev"];
  if (clientRev !== undefined && clientRev !== "" && k in store && String(revs[k] || 0) !== String(clientRev)) {
    res.set("X-Rev", String(revs[k] || 0));
    return res.status(409).json({ error: "conflict", rev: revs[k] || 0 });
  }
  store[k] = req.body.value;
  revs[k] = (revs[k] || 0) + 1;
  persist();
  res.set("X-Rev", String(revs[k]));
  res.json({ key: k, value: req.body.value });
});
app.delete("/api/storage/:key", (req, res) => {
  delete store[req.params.key];
  revs[req.params.key] = (revs[req.params.key] || 0) + 1;
  persist();
  res.json({ key: req.params.key, deleted: true });
});

/* ---- Fichiers statiques (morceaux hachés en cache 1 an, le reste sans cache) ---- */
app.use(express.static(path.join(__dirname, "dist"), {
  setHeaders: (res, filePath) => {
    if (filePath.includes(path.sep + "chunks" + path.sep)) {
      res.set("Cache-Control", "public, max-age=31536000, immutable");
    } else {
      res.set("Cache-Control", "no-cache");
    }
  },
}));
app.get("/favicon.png", (_, res) => res.sendFile(path.join(__dirname, "favicon.png")));
app.get("/audit-elyon.pdf", (_, res) => res.sendFile(path.join(__dirname, "audit-elyon.pdf")));
app.get("/", (_, res) => res.sendFile(path.join(__dirname, "index.html")));

app.listen(PORT, () => console.log(`CRM ELYON v2 en ligne sur le port ${PORT}`));
