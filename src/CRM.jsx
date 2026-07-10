import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid,
} from "recharts";
import * as XLSX from "xlsx";

/* ================= CONSTANTES ================= */
const NAVY = "#0B2545";
const NAVY2 = "#13315C";
const GOLD = "#C9A24B";
const LIGHT = "#F5F7FA";

/* ---- Motifs de rendez-vous client ---- */
const MOTIFS_RDV = [
  "Faire un arbitrage",
  "Réaliser un bilan patrimonial",
  "Mettre à jour la situation personnelle et patrimoniale (EIC)",
  "Définir ou revoir les objectifs patrimoniaux",
  "Mettre en place un nouveau placement",
  "Ouvrir un contrat (assurance-vie, PER, PEA, CTO…)",
  "Effectuer un versement libre",
  "Réaliser un rachat partiel ou total",
  "Modifier la clause bénéficiaire d'une assurance-vie",
  "Optimiser la fiscalité",
  "Préparer la retraite",
  "Préparer la transmission du patrimoine",
  "Diversifier les investissements",
  "Faire un point sur la performance des placements",
  "Vérifier l'adéquation du portefeuille avec le profil de risque",
  "Réaliser un suivi annuel et ajuster la stratégie patrimoniale",
];


/* ---- Modèles d'e-mails par défaut (modifiables dans la Messagerie) ---- */
const DEFAULT_MAIL_TEMPLATES = [
  { id: "tpl-bilan", nom: "Invitation bilan annuel", sujet: "Votre bilan patrimonial annuel — ELYON & Associés", corps: "Bonjour,\n\nComme chaque année, je vous propose de faire un point ensemble sur votre situation patrimoniale : performance de vos contrats, adéquation avec vos objectifs et opportunités du moment.\n\nQuelles seraient vos disponibilités dans les deux prochaines semaines ?\n\nBien cordialement,\n\nELYON & Associés\n4 rue Largillière, 75016 Paris" },
  { id: "tpl-per", nom: "Rappel plafond PER avant fin d'année", sujet: "Optimisez votre fiscalité avant le 31 décembre", corps: "Bonjour,\n\nLa fin d'année approche : c'est le moment de vérifier que vous avez utilisé au mieux votre plafond de déduction PER pour réduire votre impôt sur le revenu.\n\nJe peux réaliser cette vérification pour vous en quelques minutes — souhaitez-vous que je vous appelle ?\n\nBien cordialement,\n\nELYON & Associés" },
  { id: "tpl-voeux", nom: "Vœux de nouvelle année", sujet: "Tous nos vœux pour cette nouvelle année", corps: "Bonjour,\n\nToute l'équipe d'ELYON & Associés vous présente ses meilleurs vœux pour cette nouvelle année : santé, réussite et sérénité dans vos projets.\n\nNous restons à vos côtés pour faire fructifier et protéger votre patrimoine.\n\nBien cordialement,\n\nELYON & Associés" },
  { id: "tpl-relance", nom: "Relance document manquant", sujet: "Votre dossier — pièce manquante", corps: "Bonjour,\n\nPour finaliser votre dossier, il nous manque encore un document. Pourriez-vous nous le transmettre en réponse à cet e-mail ?\n\nMerci d'avance,\n\nELYON & Associés" },
];

const COMPANIES = {
  "PER": ["Abeille", "Generali", "MMA", "Malakoff Humanis", "Swiss Life", "Optimum Vie"],
  "Assurance vie": ["MMA", "Swiss Life", "Abeille", "Generali"],
  "Prévoyance": ["Abeille", "Swiss Life", "April"],
  "Protection juridique": ["IAG Santé"],
  "Mutuelle": ["Abeille", "Generali", "MMA", "Malakoff Humanis", "Swiss Life", "April"],
  "Transfert": ["Abeille", "Generali", "MMA", "Malakoff Humanis", "Swiss Life", "April", "IAG Santé"],
};
const CONTRACT_TYPES = Object.keys(COMPANIES);
const SITUATIONS = ["Célibataire", "Marié(e)", "Pacsé(e)", "Concubinage"];
const ALERT_TYPES = [
  "Création de l'espace client",
  "Rappel pour documents",
  "Rappel pour signature de contrat",
  "Autre",
];
const BAREMES = ["Manager", "Commercial", "Téléprospecteur"];
const isTelepro = (u) => u && u.bareme === "Téléprospecteur";
/* Espaces interdits aux téléprospecteurs (onglets visibles, mais accès bloqué) */
const PAGES_BLOQUEES_TELEPRO = ["clients", "ventes", "messagerie"];
const STATUTS = ["En attente", "Payé", "Annulé", "Décommissionné"];

/* ---- Prospection ---- */
const PROSPECTION_STATUTS = [
  "1er appel — pas de réponse", "2e appel — pas de réponse", "3e appel — pas de réponse",
  "À rappeler", "Refus", "KO — ne plus rappeler",
  "RDV pris", "RDV honoré", "RDV annulé", "RDV reporté", "Proposition envoyée", "Signé", "Perdu",
];
const PROSPECTION_COLORS = {
  "1er appel — pas de réponse": "#b58900", "2e appel — pas de réponse": "#d97706", "3e appel — pas de réponse": "#b45309",
  "Pas de réponse": "#b58900", "À rappeler": "#b58900", "Refus": "#B3261E", "KO — ne plus rappeler": "#8f1d1d",
  "RDV pris": "#1b7a3d", "RDV honoré": "#0B2545",
  "RDV annulé": "#B3261E", "RDV reporté": "#b58900", "Proposition envoyée": "#7a5c17", "Signé": "#1b7a3d", "Perdu": "#B3261E",
};
/* Couleur de fond des LIGNES du tableau de prospection (texte foncé conservé = lisible) */
const PROSPECTION_ROW_BG = {
  "RDV pris": "#e3f5e9",
  "Signé": "#d7f0df",
  "RDV honoré": "#e3f5e9",
  "1er appel — pas de réponse": "#fdf6d8",
  "Pas de réponse": "#fdf6d8",
  "2e appel — pas de réponse": "#ffe9cc",
  "3e appel — pas de réponse": "#ffdcb3",
  "Refus": "#ffe3e0",
  "KO — ne plus rappeler": "#ffd6d1",
  "Perdu": "#ffe3e0",
};
const PROFESSIONS_SANTE = [
  "Infirmier(ère) libéral(e)", "Kinésithérapeute", "Dentiste", "Médecin généraliste",
  "Médecin spécialiste", "Sage-femme", "Ostéopathe", "Podologue", "Orthophoniste",
  "Pharmacien(ne)", "Vétérinaire", "Autre profession libérale", "Autre",
];
const NOTES_5 = ["1", "2", "3", "4", "5"];
const MONTH_NAMES = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];

const DEFAULT_USERS = [
  { id: "quentin", prenom: "Quentin", nom: "Marty", bareme: "Manager", isManager: true, password: null },
  { id: "arzou", prenom: "Arzou", nom: "Kaya", bareme: "Commercial", isManager: false },
  { id: "simon", prenom: "Simon", nom: "Mandel", bareme: "Commercial", isManager: false },
];

/* ================= HELPERS ================= */
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

const monthKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
const monthLabel = (key) => {
  const [y, m] = key.split("-");
  return `${MONTH_NAMES[parseInt(m, 10) - 1]} ${y}`;
};
const defaultMonths = () => {
  const out = [];
  const start = new Date(2025, 8, 1); // septembre 2025
  const now = new Date();
  const end = new Date(now.getFullYear(), now.getMonth(), 1);
  let d = start;
  while (d <= end) {
    out.push(monthKey(d));
    d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  }
  return out;
};
const nextMonthKey = (key) => {
  const [y, m] = key.split("-").map(Number);
  const d = new Date(y, m, 1); // mois suivant
  return monthKey(d);
};

const parseNum = (v) => {
  if (v === null || v === undefined) return 0;
  let s = String(v).replace(/\s|\u00a0/g, "").replace(/[^\d.,-]/g, "");
  if (!s) return 0;
  const lastC = s.lastIndexOf(","), lastD = s.lastIndexOf(".");
  if (lastC > -1 && lastD > -1) {
    /* les deux séparateurs présents : le dernier est la décimale */
    if (lastC > lastD) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (lastC > -1) s = s.replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};
const fmtEUR = (n) =>
  n.toLocaleString("fr-FR", { minimumFractionDigits: 0, maximumFractionDigits: 2 }) + " €";
const fmtDate = (iso) => {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
};
const todayISO = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};

/* Rappel automatique tous les 4 mois à partir de la date de signature */
const nextFollowUp = (signatureISO) => {
  if (!signatureISO) return null;
  const sig = new Date(signatureISO + "T00:00:00");
  if (isNaN(sig)) return null;
  const now = new Date(); now.setHours(0, 0, 0, 0);
  let d = new Date(sig);
  do { d = new Date(d.getFullYear(), d.getMonth() + 4, d.getDate()); } while (d < now);
  return d;
};
const daysUntil = (date) => {
  const now = new Date(); now.setHours(0, 0, 0, 0);
  return Math.round((date - now) / 86400000);
};

/* Prochain anniversaire d'un client (à partir de sa date de naissance) */
const nextBirthday = (dateNaissanceISO) => {
  if (!dateNaissanceISO) return null;
  const [, m, d] = dateNaissanceISO.split("-").map(Number);
  if (!m || !d) return null;
  const now = new Date(); now.setHours(0, 0, 0, 0);
  let bd = new Date(now.getFullYear(), m - 1, d);
  if (bd < now) bd = new Date(now.getFullYear() + 1, m - 1, d);
  return bd;
};
const ageAt = (dateNaissanceISO, at) => {
  const [y] = dateNaissanceISO.split("-").map(Number);
  return at.getFullYear() - y;
};

/* Export CSV compatible Excel FR (séparateur ; + BOM UTF-8) */
const downloadCSV = (filename, headers, rows) => {
  const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const csv = "\uFEFF" + [headers.map(esc).join(";"), ...rows.map((r) => r.map(esc).join(";"))].join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 3000);
};

/* Lien "Ajouter à Google Agenda" pré-rempli (heure locale, durée 1h par défaut) */
const gcalUrl = ({ titre, date, heure, duree = 60, details = "", lieu = "" }) => {
  if (!date) return null;
  const [h, m] = (heure || "09:00").split(":").map(Number);
  const pad = (n) => String(n).padStart(2, "0");
  const d0 = date.replace(/-/g, "");
  const endMin = h * 60 + m + duree;
  const start = `${d0}T${pad(h)}${pad(m)}00`;
  const end = `${d0}T${pad(Math.floor(endMin / 60) % 24)}${pad(endMin % 60)}00`;
  return `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(titre)}&dates=${start}/${end}&details=${encodeURIComponent(details)}&location=${encodeURIComponent(lieu)}`;
};

/* Format automatique en euros : "1200" → "1 200,00 €" (saisie libre conservée si non numérique) */
const autoEUR = (raw) => {
  const s = String(raw || "").trim();
  if (!s) return s;
  const n = parseNum(s);
  if (!n && n !== 0) return s;
  if (!/\d/.test(s)) return s;
  return n.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
};

/* ================= STOCKAGE ================= */
async function sGet(key) {
  try {
    const r = await window.storage.get(key, true);
    return r ? JSON.parse(r.value) : null;
  } catch { return null; }
}
async function sSet(key, val) {
  try { await window.storage.set(key, JSON.stringify(val), true); }
  catch (e) { console.error("Erreur de sauvegarde", key, e); }
}
async function sDel(key) {
  try { await window.storage.delete(key, true); } catch {}
}

/* Fichiers (base64) — limite 15 Mo par fichier */
const MAX_FILE = 15 * 1024 * 1024;
function readFileB64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("Lecture impossible"));
    r.readAsDataURL(file);
  });
}
async function storeFile(file) {
  if (file.size > MAX_FILE) throw new Error(`« ${file.name} » dépasse 15 Mo. Compressez le fichier avant de l'importer.`);
  const data = await readFileB64(file);
  const id = uid();
  await sSet(`crm-file-${id}`, { name: file.name, type: file.type || "application/octet-stream", data });
  return { id, name: file.name, size: file.size, date: todayISO() };
}
async function downloadFile(fileId, fallbackName) {
  const f = await sGet(`crm-file-${fileId}`);
  if (!f) { alert("Fichier introuvable."); return; }
  const bin = atob(f.data);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const blob = new Blob([bytes], { type: f.type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = f.name || fallbackName || "document";
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 3000);
}

const emptyRow = () => ({
  id: uid(), dateCreation: "", nom: "", type: "", compagnie: "", frais: "", ref: "",
  commentaire: "", apporteur: "", volume: "", remuneration: "", statut: "En attente",
});
const emptyMonthData = (users) => {
  const out = {};
  users.forEach((u) => { out[u.id] = { rows: Array.from({ length: 20 }, emptyRow), nonPayes: "" }; });
  return out;
};

/* ================= STYLES ================= */
const CSS = `
  .crm * { box-sizing: border-box; }
  .crm { font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif; color: ${NAVY}; min-height: 100vh; background: ${LIGHT}; display:flex; }
  .crm h1,.crm h2,.crm h3 { font-family: Georgia, 'Times New Roman', serif; margin: 0; }
  .side { width: 232px; min-height: 100vh; background: linear-gradient(180deg, ${NAVY} 0%, ${NAVY2} 100%); color:#fff; display:flex; flex-direction:column; flex-shrink:0; }
  .brand { padding: 26px 20px 18px; border-bottom: 1px solid rgba(255,255,255,.12); }
  .brand b { font-family: Georgia, serif; font-size: 19px; letter-spacing: .5px; display:block; }
  .brand span { color:${GOLD}; font-size: 11px; letter-spacing: 2.5px; text-transform: uppercase; }
  .nav { padding: 14px 10px; display:flex; flex-direction:column; gap:4px; flex:1; }
  .nav button { text-align:left; background:none; border:none; color:rgba(255,255,255,.82); padding:11px 14px; border-radius:8px; cursor:pointer; font-size:14px; display:flex; gap:10px; align-items:center; transition: background .15s; }
  .nav button:hover { background: rgba(255,255,255,.08); }
  .nav button.on { background: rgba(201,162,75,.18); color:#fff; border-left: 3px solid ${GOLD}; }
  .side .who { padding: 16px 20px; border-top: 1px solid rgba(255,255,255,.12); font-size: 13px; }
  .side .who b { color: ${GOLD}; }
  .main { flex:1; padding: 30px 34px; overflow:auto; min-width: 0; }
  .ph { display:flex; align-items:flex-end; justify-content:space-between; gap: 16px; margin-bottom: 22px; flex-wrap: wrap; }
  .ph h1 { font-size: 26px; }
  .ph .sub { color:#5b6b82; font-size: 13px; margin-top: 4px; }
  .card { background:#fff; border:1px solid #e3e8f0; border-radius: 12px; padding: 20px; box-shadow: 0 1px 3px rgba(11,37,69,.05); }
  .grid { display:grid; gap: 14px; }
  .kpis { display:grid; grid-template-columns: repeat(auto-fit,minmax(150px,1fr)); gap:12px; }
  .kpi { background:#fff; border:1px solid #e3e8f0; border-left: 4px solid ${GOLD}; border-radius: 10px; padding: 14px 16px; }
  .kpi .n { font-size: 26px; font-weight: 700; font-family: Georgia, serif; }
  .kpi .l { font-size: 12px; color:#5b6b82; text-transform: uppercase; letter-spacing: .8px; margin-top:2px; }
  .btn { background:${NAVY}; color:#fff; border:none; border-radius:8px; padding: 9px 16px; cursor:pointer; font-size: 14px; }
  .btn:hover { background:${NAVY2}; }
  .btn.gold { background:${GOLD}; color:${NAVY}; font-weight:600; }
  .btn.ghost { background:#fff; color:${NAVY}; border:1px solid #cdd6e2; }
  .btn.sm { padding: 5px 10px; font-size: 12.5px; border-radius:6px; }
  .btn.danger { background:#fff; color:#B3261E; border:1px solid #e6c9c7; }
  .in, .sel, .ta { width:100%; border:1px solid #cdd6e2; border-radius:8px; padding: 8px 10px; font-size:14px; color:${NAVY}; background:#fff; font-family: inherit; }
  textarea, input, select, button { font-family: inherit; }
  .in:focus,.sel:focus,.ta:focus { outline: 2px solid ${GOLD}55; border-color:${GOLD}; }
  .lbl { font-size: 12px; font-weight: 600; color:#41506a; display:block; margin-bottom:4px; letter-spacing:.3px; }
  .fgrid { display:grid; grid-template-columns: repeat(auto-fit,minmax(210px,1fr)); gap: 12px; }
  table.t { width:100%; border-collapse: collapse; font-size: 13px; }
  table.t th { background:${NAVY}; color:#fff; padding: 8px 7px; text-align:center; font-weight:600; font-size:11.5px; letter-spacing:.5px; text-transform:uppercase; white-space: nowrap; }
  table.t th:first-child { border-radius: 8px 0 0 0; } table.t th:last-child { border-radius: 0 8px 0 0; }
  table.t td { border-bottom:1px solid #edf1f6; padding: 3px 5px; text-align:center; font-size:12.5px; }
  table.salest td { height:34px; white-space:nowrap; vertical-align:middle; }
  table.salest tbody tr:nth-child(odd):not(.paye):not(.annule):not(.decom):not(.totrow):not(.nprow) td { background:#fbfcff; }
  table.nowrapt td { white-space:nowrap; }
  table.t tbody tr:not(.paye):not(.annule):not(.decom):not(.totrow):not(.nprow):hover td { background:#f6f8fb; }
  table.t input, table.t select { width:100%; border:1px solid transparent; background:transparent; padding: 4px 3px; font-size:12.5px; text-align:center; border-radius:6px; color:inherit; }
  table.t input:focus, table.t select:focus { background:#fff; border-color:${GOLD}; outline:none; }
  tr.paye td { background:#e4f3e6; } tr.annule td { background:#fbe4e2; } tr.decom td { background:#f8d0cc; color:#7a1410; font-weight:600; }
  .totrow td { background:${NAVY}; color:#fff; font-weight:700; padding: 10px 8px; }
  .nprow td { background:#fdf6e7; font-weight:600; }
  .modal-bg { position:fixed; inset:0; background:rgba(11,37,69,.55); display:flex; align-items:flex-start; justify-content:center; padding: 40px 16px; z-index:50; overflow:auto; }
  .modal { background:#fff; border-radius: 14px; padding: 26px; width: 100%; max-width: 760px; box-shadow: 0 20px 60px rgba(0,0,0,.3); }
  .modal h2 { font-size: 20px; margin-bottom: 16px; }
  .badge { display:inline-block; padding: 3px 10px; border-radius: 20px; font-size: 11.5px; font-weight:600; }
  .b-navy { background:${NAVY}; color:#fff; } .b-gold { background:${GOLD}22; color:#8a6a1f; border:1px solid ${GOLD}66; }
  .b-green { background:#e4f3e6; color:#1d6b2a; } .b-red { background:#fbe4e2; color:#a33028; } .b-grey { background:#eef1f5; color:#41506a; }
  .row { display:flex; gap:10px; align-items:center; flex-wrap:wrap; }
  .clientcard { display:flex; justify-content:space-between; align-items:center; padding: 14px 18px; background:#fff; border:1px solid #e3e8f0; border-radius:10px; cursor:pointer; transition: border-color .15s; }
  .clientcard:hover { border-color:${GOLD}; }
  .filelink { display:flex; justify-content:space-between; align-items:center; padding: 8px 12px; background:${LIGHT}; border:1px solid #e3e8f0; border-radius:8px; font-size:13px; margin-bottom:6px; }
  .alertline { display:flex; justify-content:space-between; align-items:center; gap:10px; padding:10px 14px; border-radius:8px; border:1px solid #e3e8f0; background:#fff; margin-bottom:8px; }
  .alertline.today { border-color:${GOLD}; background:#fdf6e7; }
  .tabs { display:flex; gap:6px; flex-wrap:wrap; margin-bottom: 16px; }
  .tabs button { border:1px solid #cdd6e2; background:#fff; border-radius:20px; padding: 7px 16px; cursor:pointer; font-size:13px; color:${NAVY}; }
  .tabs button.on { background:${NAVY}; color:#fff; border-color:${NAVY}; }
  .login { min-height:100vh; width:100%; display:flex; align-items:center; justify-content:center; background: linear-gradient(135deg, ${NAVY} 0%, ${NAVY2} 60%, #1d4066 100%); }
  .loginbox { background:#fff; border-radius: 16px; padding: 38px 36px; width: 100%; max-width: 420px; box-shadow: 0 30px 80px rgba(0,0,0,.4); }
  .loginbox h1 { font-size: 24px; text-align:center; }
  .loginbox .gold { color:${GOLD}; letter-spacing: 3px; font-size: 11px; text-transform: uppercase; text-align:center; display:block; margin-bottom: 26px; margin-top: 4px; }
  .userbtn { display:flex; justify-content:space-between; align-items:center; width:100%; padding: 13px 16px; border:1px solid #cdd6e2; border-radius:10px; background:#fff; cursor:pointer; font-size: 15px; margin-bottom: 10px; }
  .userbtn:hover { border-color:${GOLD}; background:#fdf9f0; }
  /* ============ OPTIMISATION MOBILE ============ */
  @media (max-width: 860px) {
    .crm { flex-direction: column; }
    /* Barre latérale → barre supérieure compacte avec menu horizontal défilant */
    .side { width:100%; min-height:0; position: sticky; top: 0; z-index: 80; }
    .side .nav { display:flex; flex-direction:row; overflow-x:auto; -webkit-overflow-scrolling: touch; padding: 4px 8px 8px; gap: 4px; }
    .side .nav button { white-space: nowrap; flex-shrink: 0; font-size: 12.5px; padding: 8px 12px; border-radius: 8px; }
    .main { padding: 14px 10px; }
    .ph { flex-direction: column; align-items: flex-start; gap: 10px; }
    .ph .row { flex-wrap: wrap; }
    h1 { font-size: 20px; }
    .kpis { grid-template-columns: repeat(2, 1fr); gap: 8px; }
    .kpi .n { font-size: 20px; }
    .fgrid { grid-template-columns: 1fr; }
    .grid { grid-template-columns: 1fr !important; }
    .modal { width: 94vw !important; max-width: 94vw !important; padding: 16px !important; max-height: 90vh; overflow-y: auto; }
    /* Tableaux : défilement horizontal fluide au doigt */
    .card { padding: 12px; overflow-x: auto; -webkit-overflow-scrolling: touch; }
    table.t { min-width: 640px; }
    .in, .sel, .ta, .btn { font-size: 16px; } /* évite le zoom automatique iPhone à la saisie */
    .btn { padding: 10px 14px; }
    .userbtn { padding: 14px 16px; }
  }
`;

/* ================= COMPOSANTS GÉNÉRIQUES ================= */
function Field({ label, children }) {
  return (
    <div>
      <span className="lbl">{label}</span>
      {children}
    </div>
  );
}

function FilePicker({ label, multiple, onFiles, busyText = "Import en cours…" }) {
  const ref = useRef(null);
  const [busy, setBusy] = useState(false);
  return (
    <span>
      <input
        type="file" multiple={multiple} ref={ref} style={{ display: "none" }}
        onChange={async (e) => {
          const files = Array.from(e.target.files || []);
          if (!files.length) return;
          setBusy(true);
          try {
            const stored = [];
            for (const f of files) stored.push(await storeFile(f));
            await onFiles(stored);
          } catch (err) { alert(err.message); }
          setBusy(false);
          if (ref.current) ref.current.value = "";
        }}
      />
      <button className="btn ghost sm" onClick={() => ref.current && ref.current.click()} disabled={busy}>
        {busy ? busyText : label}
      </button>
    </span>
  );
}

function FileList({ files, onDelete }) {
  if (!files || !files.length) return <div style={{ fontSize: 13, color: "#8593a8" }}>Aucun document importé.</div>;
  return (
    <div>
      {files.map((f) => (
        <div className="filelink" key={f.id}>
          <span>📄 {f.name} <span style={{ color: "#8593a8" }}>· {fmtDate(f.date)}</span></span>
          <span className="row">
            <button className="btn sm" onClick={() => downloadFile(f.id, f.name)}>Télécharger</button>
            {onDelete && <button className="btn danger sm" onClick={() => onDelete(f)}>Suppr.</button>}
          </span>
        </div>
      ))}
    </div>
  );
}

/* ================= APPLICATION ================= */
export default function App() {
  const [loading, setLoading] = useState(true);
  const [users, setUsers] = useState(DEFAULT_USERS);
  const [clients, setClients] = useState([]);
  const [sales, setSales] = useState({});          // { "2025-09": { userId: {rows, nonPayes} } }
  const [docs, setDocs] = useState([]);            // dossiers partagés
  const [bordereaux, setBordereaux] = useState({}); // { userId: { "2025-09": [files] } }

  const [me, setMe] = useState(null);       // utilisateur connecté
  const [viewAs, setViewAs] = useState(null); // espace consulté (Quentin peut voir les autres)
  const [page, setPage] = useState("dash");
  const [openClient, setOpenClient] = useState(null);
  const [prospection, setProspection] = useState([]);   // fiches d'appels / RDV prospection
  const [objectifs, setObjectifs] = useState({});       // { "2026-07": { userId: { contrats, volume } } }
  const [trash, setTrash] = useState([]);               // corbeille (fiches supprimées, 30 jours)
  const [rdvClients, setRdvClients] = useState([]);     // rendez-vous clients (motifs patrimoniaux)
  const [mailTpl, setMailTpl] = useState([]);           // modèles d'e-mails (messagerie)
  const [backupModal, setBackupModal] = useState(null); // fenêtre export/import par copier-coller

  /* ---- Chargement initial ---- */
  useEffect(() => {
    (async () => {
      const [u, c, s, d, b, p, o, t, rc, mt] = await Promise.all([
        sGet("crm-users"), sGet("crm-clients"), sGet("crm-sales"), sGet("crm-docs"), sGet("crm-bordereaux"),
        sGet("crm-prospection"), sGet("crm-objectifs"), sGet("crm-trash"),
        sGet("crm-rdv-clients"), sGet("crm-mailtpl"),
      ]);
      const loadedUsers = u && u.length ? u : DEFAULT_USERS;
      setUsers(loadedUsers);
      if (!u) await sSet("crm-users", DEFAULT_USERS);
      setClients(c || []);
      let salesData = s || {};
      let changed = false;
      defaultMonths().forEach((mk) => {
        if (!salesData[mk]) { salesData[mk] = emptyMonthData(loadedUsers); changed = true; }
      });
      setSales(salesData);
      if (!s || changed) await sSet("crm-sales", salesData);
      setDocs(d || []);
      setBordereaux(b || {});
      setProspection(p || []);
      setObjectifs(o || {});
      setRdvClients(rc || []);
      setMailTpl(mt && mt.length ? mt : DEFAULT_MAIL_TEMPLATES);
      if (!mt || !mt.length) await sSet("crm-mailtpl", DEFAULT_MAIL_TEMPLATES);
      /* Purge automatique de la corbeille après 30 jours */
      const freshTrash = (t || []).filter((x) => Date.now() - new Date(x.deletedAt).getTime() < 30 * 86400000);
      setTrash(freshTrash);
      if ((t || []).length !== freshTrash.length) await sSet("crm-trash", freshTrash);
      setLoading(false);
    })();
  }, []);

  const saveUsers = (v) => { setUsers(v); sSet("crm-users", v); };
  const saveClients = (v) => { setClients(v); sSet("crm-clients", v); };
  const saveSales = (v) => { setSales(v); sSet("crm-sales", v); };
  const saveDocs = (v) => { setDocs(v); sSet("crm-docs", v); };
  const saveBordereaux = (v) => { setBordereaux(v); sSet("crm-bordereaux", v); };
  const saveProspection = (v) => { setProspection(v); sSet("crm-prospection", v); };
  const saveObjectifs = (v) => { setObjectifs(v); sSet("crm-objectifs", v); };
  const saveTrash = (v) => { setTrash(v); sSet("crm-trash", v); };
  const saveRdvClients = (v) => { setRdvClients(v); sSet("crm-rdv-clients", v); };
  const saveMailTpl = (v) => { setMailTpl(v); sSet("crm-mailtpl", v); };
  const toTrash = (kind, data) => saveTrash([...trash, { id: uid(), kind, data, deletedAt: new Date().toISOString(), deletedBy: me ? me.id : "?" }]);

  if (loading) {
    return (
      <div className="crm" style={{ alignItems: "center", justifyContent: "center", background: NAVY }}>
        <style>{CSS}</style>
        <div style={{ color: "#fff", fontFamily: "Georgia, serif", fontSize: 20 }}>
          ELYON <span style={{ color: GOLD }}>& Associés</span> — chargement…
        </div>
      </div>
    );
  }

  if (!me) {
    return (
      <Login
        users={users}
        onLogin={(u) => { setMe(u); setViewAs(u); setPage("dash"); }}
        onSetPassword={(userId, pwd) => {
          saveUsers(users.map((u) => (u.id === userId ? { ...u, password: pwd } : u)));
        }}
      />
    );
  }

  const view = viewAs || me;
  const NAV = [
    ["dash", "📊 Tableau de bord"],
    ["clients", "👥 Clients"],
    ["prospection", "🎯 Prospection"],
    ["ventes", "📈 Ventes équipe"],
    ["messagerie", "✉️ Messagerie"],
    ["paye", "💶 Ma rémunération"],
    ["docs", "📁 Documents"],
    ...(me.isManager ? [["decom", "📉 Décommissions"], ["equipe", "🧑‍💼 Mon équipe"], ["corbeille", "🗑️ Corbeille"]] : []),
  ];

  return (
    <div className="crm">
      <style>{CSS}</style>
      <aside className="side">
        <div className="brand">
          <b>ELYON <span style={{ color: GOLD }}>&</span> Associés</b>
          <span>Gestion de patrimoine</span>
        </div>
        <nav className="nav">
          {NAV.map(([k, l]) => (
            <button key={k} className={page === k ? "on" : ""} onClick={() => { setPage(k); setOpenClient(null); }}>
              {l}
            </button>
          ))}
        </nav>
        <GlobalSearch
          clients={clients} prospection={prospection} me={me} users={users}
          goClient={(id) => { setPage("clients"); setOpenClient(id); }}
          goProspection={() => { setPage("prospection"); setOpenClient(null); }}
        />
        <div className="who">
          Connecté : <b>{me.prenom} {me.nom}</b>
          <div style={{ fontSize: 11.5, opacity: 0.75 }}>Barème {me.bareme}</div>
          {me.isManager && (
            <div style={{ marginTop: 10 }}>
              <span style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: 1, opacity: 0.7 }}>Espace consulté</span>
              <select
                className="sel" style={{ marginTop: 4, background: "rgba(255,255,255,.1)", color: "#fff", borderColor: "rgba(255,255,255,.3)" }}
                value={view.id}
                onChange={(e) => setViewAs(users.find((u) => u.id === e.target.value))}
              >
                {users.map((u) => <option key={u.id} value={u.id} style={{ color: NAVY }}>{u.prenom} {u.nom}</option>)}
              </select>
            </div>
          )}
          {me.isManager && (
            <div style={{ marginTop: 10 }}>
              <button
                className="btn ghost sm" style={{ width: "100%", marginBottom: 6 }}
                onClick={() => {
                  const payload = { version: 5, date: todayISO(), users, clients, sales, docs, bordereaux, prospection, objectifs, trash, rdvClients, mailTpl };
                  const text = JSON.stringify(payload);
                  /* Tentative de téléchargement (peut être bloqué selon l'environnement) */
                  try {
                    const blob = new Blob([text], { type: "application/json" });
                    const a = document.createElement("a");
                    a.href = URL.createObjectURL(blob);
                    a.download = `ELYON_CRM_sauvegarde_${todayISO()}.json`;
                    a.click();
                    URL.revokeObjectURL(a.href);
                  } catch {}
                  /* Et dans tous les cas : fenêtre de copie (fonctionne partout) */
                  setBackupModal({ mode: "export", text });
                }}
              >
                💾 Exporter les données
              </button>
              <button
                className="btn ghost sm" style={{ width: "100%" }}
                onClick={() => setBackupModal({ mode: "import", text: "" })}
              >
                📥 Importer une sauvegarde
              </button>
            </div>
          )}
          <button className="btn ghost sm" style={{ marginTop: 12, width: "100%" }} onClick={() => { setMe(null); setViewAs(null); }}>
            Se déconnecter
          </button>
        </div>
      </aside>

      <main className="main">
        <RdvReminder prospection={prospection} rdvClients={rdvClients} clients={clients} me={me} />
        {isTelepro(me) && PAGES_BLOQUEES_TELEPRO.includes(page) ? (
          <AccessDenied goBack={() => { setPage("prospection"); setOpenClient(null); }} />
        ) : (
        <>
        {page === "dash" && <Dashboard clients={clients} users={users} view={view} me={me} sales={sales} rdvClients={rdvClients} prospection={prospection} saveClients={saveClients} saveUsers={saveUsers} goClient={(c) => { setOpenClient(c.id); setPage("clients"); }} goProspection={() => setPage("prospection")} />}
        {page === "messagerie" && (
          <MessageriePage clients={clients} users={users} me={me} mailTpl={mailTpl} saveMailTpl={saveMailTpl} />
        )}
        {page === "prospection" && (
          <ProspectionPage
            prospection={prospection} saveProspection={saveProspection} me={me} users={users} toTrash={toTrash}
            clients={clients} saveClients={saveClients}
            goClient={(id) => { setPage("clients"); setOpenClient(id); }}
          />
        )}
        {page === "clients" && !openClient && (
          <ClientsPage clients={clients} saveClients={saveClients} me={me} users={users} openClient={(id) => setOpenClient(id)} />
        )}
        {page === "clients" && openClient && (
          <ClientDetail
            client={clients.find((c) => c.id === openClient)}
            me={me}
            users={users}
            rdvClients={rdvClients}
            saveRdvClients={saveRdvClients}
            back={() => setOpenClient(null)}
            update={(next) => saveClients(clients.map((c) => (c.id === next.id ? next : c)))}
            remove={() => { if (confirm("Mettre cette fiche client à la corbeille ? (restaurable pendant 30 jours)")) { toTrash("client", clients.find((c) => c.id === openClient)); saveClients(clients.filter((c) => c.id !== openClient)); setOpenClient(null); } }}
          />
        )}
        {page === "ventes" && <SalesPage sales={sales} saveSales={saveSales} users={users} objectifs={objectifs} saveObjectifs={saveObjectifs} me={me} clients={clients} saveClients={saveClients} />}
        {page === "decom" && me.isManager && <DecomPage sales={sales} users={users} />}
        {page === "paye" && <PayePage view={view} sales={sales} bordereaux={bordereaux} saveBordereaux={saveBordereaux} />}
        {page === "docs" && <DocsPage docs={docs} saveDocs={saveDocs} toTrash={toTrash} />}
        {page === "equipe" && me.isManager && <TeamPage users={users} saveUsers={saveUsers} sales={sales} saveSales={saveSales} me={me} />}
        {page === "corbeille" && me.isManager && (
          <TrashPage
            trash={trash} saveTrash={saveTrash} users={users}
            restoreClient={(item) => { saveClients([...clients, item.data]); saveTrash(trash.filter((x) => x.id !== item.id)); }}
            restoreProspect={(item) => { saveProspection([...prospection, item.data]); saveTrash(trash.filter((x) => x.id !== item.id)); }}
            restoreDoc={(item) => { saveDocs([...docs, item.data]); saveTrash(trash.filter((x) => x.id !== item.id)); }}
            restoreFichier={(item) => {
              const target = docs.find((x) => x.id === item.data.folderId);
              if (target) saveDocs(docs.map((x) => (x.id === target.id ? { ...x, files: [...x.files, item.data.file] } : x)));
              else saveDocs([...docs, { id: uid(), name: item.data.folderName || "Documents restaurés", files: [item.data.file] }]);
              saveTrash(trash.filter((x) => x.id !== item.id));
            }}
          />
        )}
        </>
        )}
        {backupModal && (
          <BackupModal
            modal={backupModal}
            close={() => setBackupModal(null)}
            doImport={(text) => {
              try {
                const p = JSON.parse(text);
                if (!p.clients && !p.users) { alert("Ce texte n'est pas une sauvegarde ELYON valide."); return; }
                if (!confirm("Importer cette sauvegarde ? Les données actuelles seront remplacées.")) return;
                if (p.users) saveUsers(p.users);
                if (p.clients) saveClients(p.clients);
                if (p.sales) saveSales(p.sales);
                if (p.docs) saveDocs(p.docs);
                if (p.bordereaux) saveBordereaux(p.bordereaux);
                if (p.prospection) saveProspection(p.prospection);
                if (p.objectifs) saveObjectifs(p.objectifs);
                if (p.trash) saveTrash(p.trash);
                if (p.rdvClients) saveRdvClients(p.rdvClients);
                if (p.mailTpl) saveMailTpl(p.mailTpl);
                alert("Import terminé ✓ Vos données sont restaurées.");
                setBackupModal(null);
              } catch { alert("Impossible de lire cette sauvegarde. Vérifiez que le texte est copié en entier."); }
            }}
          />
        )}
      </main>
    </div>
  );
}

/* ================= CONNEXION ================= */
function Login({ users, onLogin, onSetPassword }) {
  const [selected, setSelected] = useState(null);
  const [pwd, setPwd] = useState("");
  const [pwd2, setPwd2] = useState("");
  const [err, setErr] = useState("");

  const needsPassword = !!selected;
  const firstTime = selected && !selected.password;

  const submit = () => {
    setErr("");
    if (firstTime) {
      if (pwd.length < 4) { setErr("Choisissez un mot de passe d'au moins 4 caractères."); return; }
      if (pwd !== pwd2) { setErr("Les deux mots de passe ne correspondent pas."); return; }
      onSetPassword(selected.id, pwd);
      onLogin({ ...selected, password: pwd });
    } else if (needsPassword) {
      if (pwd === selected.password) onLogin(selected);
      else setErr("Mot de passe incorrect.");
    }
  };

  return (
    <div className="crm">
      <style>{CSS}</style>
      <div className="login">
        <div className="loginbox">
          <h1>ELYON <span style={{ color: GOLD }}>&</span> Associés</h1>
          <span className="gold">CRM · Gestion de patrimoine</span>
          {!selected && (
            <>
              <p style={{ fontSize: 13.5, color: "#5b6b82", marginBottom: 14 }}>Sélectionnez votre espace :</p>
              {users.map((u) => (
                <button key={u.id} className="userbtn" onClick={() => { setSelected(u); setPwd(""); setPwd2(""); setErr(""); }}>
                  <span>{u.prenom} {u.nom}</span>
                  <span className={"badge " + (u.isManager ? "b-gold" : "b-grey")}>{u.bareme} 🔒</span>
                </button>
              ))}
            </>
          )}
          {selected && needsPassword && (
            <>
              <p style={{ fontSize: 14, marginBottom: 12 }}>
                Espace de <b>{selected.prenom} {selected.nom}</b>
                {firstTime && <span style={{ color: "#5b6b82" }}> — première connexion, créez votre mot de passe.</span>}
              </p>
              <Field label={firstTime ? "Nouveau mot de passe" : "Mot de passe"}>
                <input className="in" type="password" value={pwd} onChange={(e) => setPwd(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} autoFocus />
              </Field>
              {firstTime && (
                <div style={{ marginTop: 10 }}>
                  <Field label="Confirmez le mot de passe">
                    <input className="in" type="password" value={pwd2} onChange={(e) => setPwd2(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} />
                  </Field>
                </div>
              )}
              {err && <div style={{ color: "#B3261E", fontSize: 13, marginTop: 8 }}>{err}</div>}
              <div style={{ marginTop: 16 }}>
                <button
                  type="button"
                  onClick={submit}
                  style={{ display: "block", width: "100%", background: GOLD, color: NAVY, border: "none", borderRadius: 8, padding: "13px 16px", fontSize: 15, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", lineHeight: 1.3 }}
                >
                  {firstTime ? "Créer mon mot de passe et me connecter" : "Se connecter"}
                </button>
                <button
                  type="button"
                  onClick={() => setSelected(null)}
                  style={{ display: "block", width: "100%", marginTop: 8, background: "#fff", color: NAVY, border: "1px solid #cdd6e2", borderRadius: 8, padding: "10px 16px", fontSize: 14, cursor: "pointer", fontFamily: "inherit" }}
                >
                  ← Retour
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

/* ================= TABLEAU DE BORD ================= */
function Dashboard({ clients: allClients, users, view, me, sales, rdvClients, prospection, saveClients, saveUsers, goClient, goProspection }) {
  const relancesActives = !((users.find((u) => u.id === me.id) || {}).relancesOff);
  const [showContrats, setShowContrats] = useState(null);
  /* Cloisonnement : un commercial ne voit que ses clients.
     Le manager voit tout depuis son espace, ou le portefeuille du conseiller consulté. */
  const clients = useMemo(() => {
    if (!me.isManager) return allClients.filter((c) => (c.createdBy || "quentin") === me.id);
    if (view.id !== me.id) return allClients.filter((c) => (c.createdBy || "quentin") === view.id);
    return allClients;
  }, [allClients, me, view]);

  const stats = useMemo(() => {
    const s = { clients: clients.length, contrats: 0, PER: 0, "Assurance vie": 0, "Prévoyance": 0, "Protection juridique": 0, "Mutuelle": 0, "Transfert": 0 };
    clients.forEach((c) => (c.contrats || []).forEach((k) => { s.contrats++; if (s[k.type] !== undefined) s[k.type]++; }));
    return s;
  }, [clients]);

  /* Alertes du jour / en retard + rappels 4 mois */
  const today = todayISO();
  const alerts = [];
  clients.forEach((c) => {
    (c.alertes || []).filter((a) => !a.done).forEach((a) => {
      if (a.date <= today) alerts.push({ kind: "alerte", client: c, alerte: a, date: a.date });
    });
    (c.contrats || []).forEach((k) => {
      const nf = nextFollowUp(k.dateSignature);
      if (nf) {
        const d = daysUntil(nf);
        if (d <= 7) alerts.push({ kind: "suivi", client: c, contrat: k, date: nf.toISOString().slice(0, 10), days: d });
      }
    });
    /* 🎂 Anniversaire client dans les 7 prochains jours */
    const bd = nextBirthday(c.dateNaissance);
    if (bd) {
      const d = daysUntil(bd);
      if (d >= 0 && d <= 7) alerts.push({ kind: "anniv", client: c, date: bd.toISOString().slice(0, 10), days: d, age: ageAt(c.dateNaissance, bd) });
    }
    /* 📋 Fiche patrimoniale non complétée 1 mois après la création de la fiche client */
    if (!(c.audit && c.audit.doneAt) && !(c.fichePatrimoniale && c.fichePatrimoniale.doneAt) && c.createdAt) {
      const age = Math.floor((Date.now() - new Date(c.createdAt).getTime()) / 86400000);
      if (age >= 30) alerts.push({ kind: "fiche", client: c, date: today, days: age });
    }
  });
  /* 🔁 Relances intelligentes : prospects "à rappeler" sans activité depuis 7 jours */
  const scopeOwner = (ownerId) => {
    if (!me.isManager) return ownerId === me.id;
    if (view.id !== me.id) return ownerId === view.id;
    return true;
  };
  const RELANCE_STATUTS = ["À rappeler", "1er appel — pas de réponse", "2e appel — pas de réponse", "Pas de réponse"];
  const dormants = !relancesActives ? [] : (prospection || [])
    .filter((p) => scopeOwner(p.ownerId) && RELANCE_STATUTS.includes(p.statut))
    .map((p) => ({ p, jours: Math.floor((Date.now() - new Date(p.updatedAt || p.dateAppel || p.createdAt || today).getTime()) / 86400000) }))
    .filter((x) => x.jours >= 7)
    .sort((a, b) => b.jours - a.jours);
  dormants.slice(0, 5).forEach(({ p, jours }) => alerts.push({ kind: "relanceProspect", prospect: p, days: jours, date: today }));
  if (dormants.length > 5) alerts.push({ kind: "relancePlus", count: dormants.length - 5, date: today });

  /* 😴 Clients sans aucune interaction depuis 6 mois */
  if (relancesActives) clients.forEach((c) => {
    const dates = [c.createdAt || "", ...((c.historique || []).map((h) => h.date || "")),
      ...((rdvClients || []).filter((r) => r.clientId === c.id).map((r) => r.date || ""))].filter(Boolean).sort();
    const last = dates[dates.length - 1];
    if (last) {
      const jours = Math.floor((Date.now() - new Date(last).getTime()) / 86400000);
      if (jours >= 180) alerts.push({ kind: "clientDormant", client: c, days: jours, date: today });
    }
  });

  /* 📅 RDV clients dans les 7 prochains jours (les miens, ou ceux de l'espace consulté) */
  (rdvClients || []).forEach((r) => {
    if (r.done || !r.date) return;
    if (!me.isManager && r.ownerId !== me.id) return;
    if (me.isManager && view.id !== me.id && r.ownerId !== view.id) return;
    const d = daysUntil(new Date(r.date + "T00:00:00"));
    if (d >= 0 && d <= 7) {
      const c = allClients.find((x) => x.id === r.clientId);
      if (c) alerts.push({ kind: "rdvclient", client: c, rdv: r, date: r.date, days: d });
    }
  });
  alerts.sort((a, b) => a.date.localeCompare(b.date));

  const markDone = (client, alerte) => {
    const next = { ...client, alertes: client.alertes.map((a) => (a.id === alerte.id ? { ...a, done: true } : a)) };
    saveClients(clientsReplace(client, next));
  };
  const clientsReplace = (oldC, newC) => clients.map((c) => (c.id === oldC.id ? newC : c));

  return (
    <div>
      <div className="ph">
        <div>
          <h1>Tableau de bord</h1>
          <div className="sub">Vue d'ensemble du cabinet · espace de {view.prenom} {view.nom}</div>
        </div>
        <div className="row">
          <button
            className="btn ghost"
            title="Active ou désactive les alertes de relance (prospects inactifs 7 j, clients sans contact 6 mois)"
            style={{ borderColor: relancesActives ? "#1b7a3d" : "#B3261E", color: relancesActives ? "#1b7a3d" : "#B3261E", fontWeight: 600 }}
            onClick={() => saveUsers(users.map((u) => (u.id === me.id ? { ...u, relancesOff: relancesActives } : u)))}
          >
            🔁 Relances : {relancesActives ? "ACTIVÉES" : "DÉSACTIVÉES"}
          </button>
          <button className="btn gold" onClick={() => setShowContrats("all")}>📄 Contrats signés du mois</button>
        </div>
      </div>

      <div className="kpis" style={{ marginBottom: 20 }}>
        {[
          [stats.clients, "Clients actifs", "all"], [stats.contrats, "Contrats", "all"], [stats.PER, "PER", "PER"],
          [stats["Assurance vie"], "Assurances vie", "Assurance vie"], [stats["Prévoyance"], "Prévoyances", "Prévoyance"],
          [stats["Protection juridique"], "Protections juridiques", "Protection juridique"], [stats["Mutuelle"], "Mutuelles", "Mutuelle"], [stats["Transfert"], "Transferts", "Transfert"],
        ].map(([n, l, t]) => (
          <div key={l} className="kpi" onClick={() => setShowContrats(t)} style={{ cursor: "pointer" }} title={`Voir les contrats ${t === "all" ? "réalisés" : t}`}>
            <div className="n">{n}</div><div className="l">{l}</div>
          </div>
        ))}
      </div>

      <div className="card">
        <h2 style={{ fontSize: 18, marginBottom: 14 }}>🔔 Alertes & rappels du jour</h2>
        {alerts.length === 0 && <div style={{ color: "#8593a8", fontSize: 14 }}>Aucune alerte aujourd'hui. Tout est à jour.</div>}
        {alerts.map((a, i) => (
          <div className={"alertline" + (a.date <= today ? " today" : "")} key={i}>
            <div>
              {a.kind === "alerte" ? (
                <>
                  <b>{a.alerte.type}</b> — {a.client.prenom} {a.client.nom}
                  {a.alerte.note && <span style={{ color: "#5b6b82" }}> · {a.alerte.note}</span>}
                  <div style={{ fontSize: 12, color: "#8593a8" }}>Prévu le {fmtDate(a.date)}{a.date < today ? " (en retard)" : ""}</div>
                </>
              ) : a.kind === "anniv" ? (
                <>
                  <b>🎂 Anniversaire</b> — {a.client.prenom} {a.client.nom} fêtera ses {a.age} ans
                  <div style={{ fontSize: 12, color: "#8593a8" }}>
                    Le {fmtDate(a.date)}{a.days === 0 ? " (aujourd'hui ! 🎉)" : ` (dans ${a.days} j)`} — une bonne occasion de prendre des nouvelles.
                  </div>
                </>
              ) : a.kind === "fiche" ? (
                <>
                  <b>🩺 Audit patrimonial à réaliser</b> — {a.client.nom.toUpperCase()} {a.client.prenom}
                  <div style={{ fontSize: 12, color: "#8593a8" }}>
                    Fiche client créée il y a {a.days} jours sans audit patrimonial. Ouvrez la fiche client pour le réaliser.
                  </div>
                </>
              ) : a.kind === "relanceProspect" ? (
                <>
                  <b>🔁 Relance prospection</b> — {(a.prospect.nom || "").toUpperCase()} {a.prospect.prenom || ""} · {a.prospect.statut}
                  <div style={{ fontSize: 12, color: "#8593a8" }}>
                    Aucune activité depuis <b>{a.days} jours</b>{a.prospect.telephone ? ` · 📞 ${a.prospect.telephone}` : ""} — à rappeler en priorité.
                  </div>
                </>
              ) : a.kind === "relancePlus" ? (
                <>
                  <b>🔁 … et {a.count} autre(s) prospect(s)</b> en attente de relance
                  <div style={{ fontSize: 12, color: "#8593a8" }}>Retrouvez-les dans l'espace Prospection, filtre « À rappeler ».</div>
                </>
              ) : a.kind === "clientDormant" ? (
                <>
                  <b>😴 Client sans contact</b> — {a.client.nom.toUpperCase()} {a.client.prenom}
                  <div style={{ fontSize: 12, color: "#8593a8" }}>
                    Aucune interaction depuis <b>{Math.floor(a.days / 30)} mois</b> — une bonne occasion de proposer un point patrimonial.
                  </div>
                </>
              ) : a.kind === "rdvclient" ? (
                <>
                  <b>📅 RDV client</b> — {a.client.nom.toUpperCase()} {a.client.prenom} · {a.rdv.motif}
                  <div style={{ fontSize: 12, color: "#8593a8" }}>
                    {fmtDate(a.date)}{a.rdv.heure && ` à ${a.rdv.heure}`}{a.days === 0 ? " (aujourd'hui)" : ` (dans ${a.days} j)`}
                  </div>
                </>
              ) : (
                <>
                  <b>Rappel client (4 mois)</b> — {a.client.prenom} {a.client.nom}
                  <div style={{ fontSize: 12, color: "#8593a8" }}>
                    Contrat {a.contrat.type} {a.contrat.compagnie} signé le {fmtDate(a.contrat.dateSignature)} · à rappeler le {fmtDate(a.date)}
                    {a.days === 0 ? " (aujourd'hui)" : a.days > 0 ? ` (dans ${a.days} j)` : ""}
                  </div>
                </>
              )}
            </div>
            <div className="row">
              <button className="btn sm" onClick={() => goClient(a.client)}>Voir la fiche</button>
              {a.kind === "alerte" && <button className="btn ghost sm" onClick={() => markDone(a.client, a.alerte)}>✓ Fait</button>}
            </div>
          </div>
        ))}
      </div>

      <Leaderboard sales={sales} users={users} />

      {showContrats && <ContratsModal sales={sales} users={users} me={me} initialType={showContrats === "all" ? "all" : showContrats} onClose={() => setShowContrats(null)} />}
    </div>
  );
}

/* ================= LISTE DES CONTRATS SIGNÉS (modal) ================= */
function ContratsModal({ sales, users, me, onClose, initialType }) {
  const months = Object.keys(sales).sort();
  const [month, setMonth] = useState(months[months.length - 1]);
  const [typeF, setTypeF] = useState(initialType || "all");
  const md = sales[month] || {};
  const visibleUsers = me.isManager ? users : users.filter((u) => u.id === me.id);
  const rows = [];
  visibleUsers.forEach((u) => {
    (((md[u.id] || {}).rows) || []).forEach((r) => {
      if ((r.nom || "").trim() && r.statut !== "Annulé" && !r.mirrorOf && (typeF === "all" || r.type === typeF)) rows.push({ ...r, commercial: u.prenom });
    });
  });
  rows.sort((a, b) => (b.dateCreation || "").localeCompare(a.dateCreation || ""));
  const totV = rows.reduce((s, r) => s + parseNum(r.volume), 0);

  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 760, maxHeight: "85vh", overflowY: "auto" }}>
        <div className="row" style={{ justifyContent: "space-between", marginBottom: 12 }}>
          <h2>📄 Contrats {typeF !== "all" ? typeF : "réalisés"} — {rows.length} contrat(s) · {fmtEUR(totV)}</h2>
          <div className="row">
            <select className="sel" style={{ width: 170 }} value={typeF} onChange={(e) => setTypeF(e.target.value)}>
              <option value="all">Tous les types</option>
              {CONTRACT_TYPES.map((t) => <option key={t}>{t}</option>)}
            </select>
            <select className="sel" style={{ width: 170 }} value={month} onChange={(e) => setMonth(e.target.value)}>
              {months.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
            </select>
          </div>
        </div>
        <table className="t">
          <thead>
            <tr>{me.isManager && <th>Commercial</th>}<th>Date</th><th>Client</th><th>Type</th><th>Compagnie</th><th>Volume</th><th>Statut</th></tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id} className={r.statut === "Payé" ? "paye" : ""}>
                {me.isManager && <td style={{ fontSize: 12.5 }}>{r.commercial}</td>}
                <td style={{ fontSize: 12.5 }}>{fmtDate(r.dateCreation)}</td>
                <td><b>{r.nom}</b></td>
                <td style={{ fontSize: 12.5 }}>{r.type}</td>
                <td style={{ fontSize: 12.5 }}>{r.compagnie}</td>
                <td style={{ fontSize: 12.5 }}>{r.volume}</td>
                <td><span className={"badge " + (r.statut === "Payé" ? "b-navy" : "b-grey")}>{r.statut}</span></td>
              </tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={7} style={{ color: "#8593a8", padding: 16 }}>Aucun contrat sur {monthLabel(month)}.</td></tr>}
          </tbody>
        </table>
        <div className="row" style={{ marginTop: 14, justifyContent: "flex-end" }}>
          <button className="btn ghost" onClick={onClose}>Fermer</button>
        </div>
      </div>
    </div>
  );
}

/* ================= CLASSEMENT ÉQUIPE ================= */
function Leaderboard({ sales, users: allUsers }) {
  const users = allUsers.filter((u) => !isTelepro(u));
  const months = Object.keys(sales).sort();
  const [month, setMonth] = useState(months[months.length - 1]);
  const md = sales[month] || {};

  const ranking = users
    .map((u) => {
      const rows = ((md[u.id] || {}).rows || []).filter((r) => (r.nom || "").trim() && r.statut !== "Annulé");
      return {
        user: u,
        contrats: rows.length,
        volume: rows.reduce((s, r) => s + parseNum(r.volume), 0),
        remuneration: rows.reduce((s, r) => s + parseNum(r.remuneration), 0),
      };
    })
    .sort((a, b) => b.volume - a.volume || b.contrats - a.contrats);

  const medals = ["🥇", "🥈", "🥉"];
  const maxVol = Math.max(1, ...ranking.map((r) => r.volume));

  return (
    <div className="card" style={{ marginTop: 20 }}>
      <div className="row" style={{ justifyContent: "space-between", marginBottom: 14 }}>
        <h2 style={{ fontSize: 18 }}>🏆 Classement de l'équipe</h2>
        <select className="sel" style={{ width: 180 }} value={month} onChange={(e) => setMonth(e.target.value)}>
          {months.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
        </select>
      </div>
      {ranking.every((r) => r.contrats === 0) && (
        <div style={{ color: "#8593a8", fontSize: 13.5 }}>Aucun contrat saisi sur {monthLabel(month)} pour l'instant.</div>
      )}
      {ranking.map((r, i) => (
        <div key={r.user.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 4px", borderBottom: i < ranking.length - 1 ? "1px solid #eef1f6" : "none" }}>
          <span style={{ fontSize: 22, width: 30 }}>{medals[i] || `${i + 1}.`}</span>
          <div style={{ flex: 1 }}>
            <b style={{ fontSize: 14.5 }}>{r.user.prenom} {r.user.nom}</b>
            <div style={{ fontSize: 12, color: "#5b6b82" }}>{r.contrats} contrat(s) · volume {fmtEUR(r.volume)}</div>
            <div style={{ height: 7, background: "#eef1f6", borderRadius: 4, marginTop: 5, overflow: "hidden" }}>
              <div style={{ height: "100%", width: `${Math.round((r.volume / maxVol) * 100)}%`, background: i === 0 ? GOLD : NAVY2, borderRadius: 4, transition: "width .4s" }} />
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ================= CLIENTS ================= */
function ClientsPage({ clients, saveClients, me, users, openClient }) {
  const [showForm, setShowForm] = useState(false);
  const [search, setSearch] = useState("");
  const [ownerFilter, setOwnerFilter] = useState("all");
  const [sortMode, setSortMode] = useState("az");

  /* Cloisonnement : un commercial ne voit QUE son portefeuille. Le manager voit tout. */
  const ownerOf = (c) => c.createdBy || "quentin";
  const myPortfolio = me.isManager ? clients : clients.filter((c) => ownerOf(c) === me.id);
  const scoped = me.isManager && ownerFilter !== "all"
    ? myPortfolio.filter((c) => ownerOf(c) === ownerFilter)
    : myPortfolio;
  const filtered = scoped
    .filter((c) => (c.nom + " " + c.prenom + " " + (c.profession || "")).toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => {
      if (sortMode === "recent") return (b.createdAt || "").localeCompare(a.createdAt || "");
      if (sortMode === "profession") return (a.profession || "zzz").localeCompare(b.profession || "zzz", "fr");
      return (a.nom || "").localeCompare(b.nom || "", "fr"); /* A → Z par défaut */
    });
  const ownerName = (c) => {
    const u = users.find((x) => x.id === ownerOf(c));
    return u ? `${u.prenom} ${u.nom}` : "—";
  };

  return (
    <div>
      <div className="ph">
        <div>
          <h1>Clients</h1>
          <div className="sub">
            {me.isManager
              ? `${scoped.length} fiche(s) — vue manager : tous les portefeuilles`
              : `${myPortfolio.length} fiche(s) — votre portefeuille personnel`}
          </div>
        </div>
        <div className="row">
          <select className="sel" value={sortMode} onChange={(e) => setSortMode(e.target.value)}>
            <option value="az">Tri : Nom A → Z</option>
            <option value="recent">Tri : Plus récents</option>
            <option value="profession">Tri : Profession</option>
          </select>
          {me.isManager && (
            <select className="sel" value={ownerFilter} onChange={(e) => setOwnerFilter(e.target.value)}>
              <option value="all">Tous les conseillers</option>
              {users.map((u) => <option key={u.id} value={u.id}>{u.prenom} {u.nom}</option>)}
            </select>
          )}
          <input className="in" style={{ width: 220 }} placeholder="Rechercher un client…" value={search} onChange={(e) => setSearch(e.target.value)} />
          <button className="btn gold" onClick={() => setShowForm(true)}>+ Nouvelle fiche client</button>
        </div>
      </div>

      <div className="grid">
        {filtered.map((c) => (
          <div className="clientcard" key={c.id} onClick={() => openClient(c.id)}>
            <div>
              <b style={{ fontSize: 15 }}>{c.nom.toUpperCase()} {c.prenom}</b>
              <div style={{ fontSize: 12.5, color: "#5b6b82" }}>
                {c.profession || "Profession non renseignée"} · {c.telephone || "—"} · {c.email || "—"}
              </div>
            </div>
            <div className="row">
              {me.isManager && <span className="badge b-grey">👤 {ownerName(c)}</span>}
              {(c.contrats || []).length > 0 && <span className="badge b-navy">{c.contrats.length} contrat(s)</span>}
              {(c.alertes || []).some((a) => !a.done && a.date <= todayISO()) && <span className="badge b-gold">🔔 alerte</span>}
              <span style={{ color: GOLD, fontSize: 18 }}>›</span>
            </div>
          </div>
        ))}
        {filtered.length === 0 && <div className="card" style={{ color: "#8593a8" }}>Aucun client. Créez la première fiche pour démarrer.</div>}
      </div>

      {showForm && (
        <ClientForm
          onClose={() => setShowForm(false)}
          onSave={(c) => { const nc = { ...c, id: uid(), createdBy: me.id, createdAt: todayISO(), contrats: [], alertes: [] }; saveClients([...clients, nc]); setShowForm(false); openClient(nc.id); }}
        />
      )}
    </div>
  );
}

function ClientForm({ initial, onSave, onClose }) {
  const [f, setF] = useState(initial || {
    nom: "", prenom: "", dateNaissance: "", telephone: "", email: "",
    profession: "", revenus: "", situation: "Célibataire",
  });
  const set = (k, v) => setF({ ...f, [k]: v });
  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{initial ? "Modifier la fiche client" : "Nouvelle fiche client"}</h2>
        <div className="fgrid">
          <Field label="Nom *"><input className="in" value={f.nom} onChange={(e) => set("nom", e.target.value)} /></Field>
          <Field label="Prénom *"><input className="in" value={f.prenom} onChange={(e) => set("prenom", e.target.value)} /></Field>
          <Field label="Date de naissance"><input className="in" type="date" value={f.dateNaissance} onChange={(e) => set("dateNaissance", e.target.value)} /></Field>
          <Field label="Téléphone"><input className="in" value={f.telephone} onChange={(e) => set("telephone", e.target.value.replace(/\s/g, ""))} /></Field>
          <Field label="E-mail"><input className="in" value={f.email} onChange={(e) => set("email", e.target.value)} /></Field>
          <Field label="Profession"><input className="in" value={f.profession} onChange={(e) => set("profession", e.target.value.toUpperCase())} style={{ textTransform: "uppercase" }} /></Field>
          <Field label="Revenus imposables (€)"><input className="in" value={f.revenus} onChange={(e) => set("revenus", e.target.value)} placeholder="ex : 48 000" /></Field>
          <Field label="Enfants"><input className="in" value={f.enfants || ""} onChange={(e) => set("enfants", e.target.value)} placeholder="ex : 2 (2015, 2018)" /></Field>
          <Field label="Ville"><input className="in" value={f.ville || ""} onChange={(e) => set("ville", e.target.value)} /></Field>
          <Field label="Situation matrimoniale">
            <select className="sel" value={f.situation} onChange={(e) => set("situation", e.target.value)}>
              {SITUATIONS.map((s) => <option key={s}>{s}</option>)}
            </select>
          </Field>
        </div>
        <div style={{ marginTop: 12 }}>
          <Field label="Commentaires">
            <textarea className="ta" rows={3} value={f.commentaires || ""} onChange={(e) => set("commentaires", e.target.value)} placeholder="Notes libres sur le client…" />
          </Field>
        </div>
        <div className="row" style={{ marginTop: 18, justifyContent: "flex-end" }}>
          <button className="btn ghost" onClick={onClose}>Annuler</button>
          <button className="btn gold" onClick={() => { if (!f.nom || !f.prenom) { alert("Nom et prénom obligatoires."); return; } onSave(f); }}>Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

function ClientDetail({ client, me, users, rdvClients, saveRdvClients, back, update, remove }) {
  const [showRdv, setShowRdv] = useState(false);
  const [showFiche, setShowFiche] = useState(false);
  const [showCourrier, setShowCourrier] = useState(false);
  const [showEdit, setShowEdit] = useState(false);
  const [showContract, setShowContract] = useState(false);
  const [editContract, setEditContract] = useState(null);
  const [showAlert, setShowAlert] = useState(false);
  if (!client) return null;

  const addAlert = (a) => update({ ...client, alertes: [...(client.alertes || []), { ...a, id: uid(), done: false }] });
  const toggleAlert = (id) => update({ ...client, alertes: client.alertes.map((a) => (a.id === id ? { ...a, done: !a.done } : a)) });
  const delAlert = (id) => update({ ...client, alertes: client.alertes.filter((a) => a.id !== id) });

  return (
    <div>
      <div className="ph">
        <div>
          <button className="btn ghost sm" onClick={back}>← Retour aux clients</button>
          <h1 style={{ marginTop: 10 }}>{client.nom.toUpperCase()} {client.prenom}</h1>
          <div className="sub">
            Fiche créée le {fmtDate(client.createdAt)}
            {" · Conseiller : "}
            {me && me.isManager ? (
              <select
                className="sel"
                style={{ fontSize: 12.5, padding: "2px 6px", marginLeft: 4 }}
                value={client.createdBy || "quentin"}
                onChange={(e) => update({ ...client, createdBy: e.target.value })}
              >
                {(users || []).map((u) => <option key={u.id} value={u.id}>{u.prenom} {u.nom}</option>)}
              </select>
            ) : (
              <b>{((users || []).find((u) => u.id === (client.createdBy || "quentin")) || {}).prenom || "—"} {((users || []).find((u) => u.id === (client.createdBy || "quentin")) || {}).nom || ""}</b>
            )}
          </div>
        </div>
        <div className="row">
          <button className="btn ghost" onClick={() => setShowEdit(true)}>✏️ Modifier</button>
          <button className="btn danger" onClick={remove}>Supprimer</button>
        </div>
      </div>

      <div className="grid" style={{ gridTemplateColumns: "1fr 1fr", alignItems: "start" }}>
        <div className="card">
          <h2 style={{ fontSize: 17, marginBottom: 12 }}>Informations personnelles</h2>
          <div style={{ fontSize: 14, lineHeight: 2 }}>
            <div><b>Date de naissance :</b> {fmtDate(client.dateNaissance)}</div>
            <div><b>Téléphone :</b> {client.telephone || "—"}</div>
            <div><b>E-mail :</b> {client.email || "—"}</div>
            <div><b>Profession :</b> {client.profession || "—"}</div>
            <div><b>Revenus imposables :</b> {client.revenus ? fmtEUR(parseNum(client.revenus)) : "—"}</div>
            <div><b>Situation matrimoniale :</b> {client.situation || "—"}</div>
            <div><b>Enfants :</b> {client.enfants || "—"}</div>
            <div><b>Ville :</b> {client.ville || "—"}</div>
          </div>
          {client.commentaires && (
            <div style={{ marginTop: 10, background: "#f8f9fc", borderRadius: 8, padding: "8px 12px", fontSize: 13, borderLeft: `3px solid ${GOLD}` }}>
              <b>💬 Commentaires :</b> <span style={{ whiteSpace: "pre-wrap" }}>{client.commentaires}</span>
            </div>
          )}
        </div>

        <div className="card">
          <div className="row" style={{ justifyContent: "space-between", marginBottom: 12 }}>
            <h2 style={{ fontSize: 17 }}>🔔 Alertes</h2>
            <button className="btn sm gold" onClick={() => setShowAlert(true)}>+ Créer une alerte</button>
          </div>
          {(client.alertes || []).length === 0 && <div style={{ color: "#8593a8", fontSize: 13.5 }}>Aucune alerte programmée.</div>}
          {(client.alertes || []).slice().sort((a, b) => a.date.localeCompare(b.date)).map((a) => (
            <div className={"alertline" + (!a.done && a.date <= todayISO() ? " today" : "")} key={a.id} style={a.done ? { opacity: 0.5 } : {}}>
              <div>
                <b>{a.type}</b>{a.note && <span style={{ color: "#5b6b82" }}> · {a.note}</span>}
                <div style={{ fontSize: 12, color: "#8593a8" }}>Rappel le {fmtDate(a.date)} {a.done && "· ✓ fait"}</div>
              </div>
              <div className="row">
                <button className="btn ghost sm" onClick={() => toggleAlert(a.id)}>{a.done ? "Réactiver" : "✓ Fait"}</button>
                <button className="btn danger sm" onClick={() => delAlert(a.id)}>✕</button>
              </div>
            </div>
          ))}
        </div>
      </div>

      <div className="card" style={{ marginTop: 16 }}>
        <div className="row" style={{ justifyContent: "space-between", marginBottom: 14 }}>
          <h2 style={{ fontSize: 17 }}>Contrats ({(client.contrats || []).length})</h2>
          <button className="btn gold" onClick={() => { setEditContract(null); setShowContract(true); }}>+ Ajouter un contrat</button>
        </div>
        {(client.contrats || []).map((k) => (
          <ContractCard
            key={k.id} contract={k}
            onEdit={() => { setEditContract(k); setShowContract(true); }}
            onDelete={() => { if (confirm("Supprimer ce contrat ?")) update({ ...client, contrats: client.contrats.filter((x) => x.id !== k.id) }); }}
            onFiles={(files) => update({ ...client, contrats: client.contrats.map((x) => (x.id === k.id ? { ...x, fichiers: [...(x.fichiers || []), ...files] } : x)) })}
            onQuick={(patch) => update({ ...client, contrats: client.contrats.map((x) => (x.id === k.id ? { ...x, ...patch } : x)) })}
            onFileDelete={(f) => { sDel(`crm-file-${f.id}`); update({ ...client, contrats: client.contrats.map((x) => (x.id === k.id ? { ...x, fichiers: x.fichiers.filter((y) => y.id !== f.id) } : x)) }); }}
          />
        ))}
        {(client.contrats || []).length === 0 && <div style={{ color: "#8593a8", fontSize: 13.5 }}>Aucun contrat. Ajoutez le premier contrat de ce client.</div>}
      </div>

      <HistoriqueCard client={client} update={update} me={me} users={users} />

      {/* ---- Rendez-vous client ---- */}
      <div className="card" style={{ marginTop: 16 }}>
        <div className="row" style={{ justifyContent: "space-between", marginBottom: 12 }}>
          <h2 style={{ fontSize: 17 }}>📅 Rendez-vous ({(rdvClients || []).filter((r) => r.clientId === client.id && !r.done).length} à venir)</h2>
          <div className="row">
            <button className="btn ghost sm" onClick={() => setShowCourrier(true)}>📄 Générer un courrier</button>
            <button className="btn gold sm" onClick={() => setShowRdv(true)}>+ Planifier un RDV</button>
          </div>
        </div>
        {(rdvClients || []).filter((r) => r.clientId === client.id).sort((a, b) => (a.date + a.heure).localeCompare(b.date + b.heure)).map((r) => (
          <div key={r.id} className="row" style={{ justifyContent: "space-between", padding: "8px 4px", borderBottom: "1px solid #eef1f6", opacity: r.done ? 0.55 : 1 }}>
            <div style={{ fontSize: 13.5 }}>
              <b>{fmtDate(r.date)}{r.heure && ` à ${r.heure}`}</b> — {r.motif}
              {r.note && <div style={{ fontSize: 12, color: "#5b6b82" }}>💬 {r.note}</div>}
              <div style={{ fontSize: 11.5, color: "#8593a8" }}>Avec {((users || []).find((u) => u.id === r.ownerId) || {}).prenom || "—"}{r.done && " · ✓ effectué"}</div>
            </div>
            <div className="row">
              <button
                className="btn ghost sm" title="Ajouter à Google Agenda"
                onClick={() => window.open(gcalUrl({
                  titre: `RDV client — ${client.nom.toUpperCase()} ${client.prenom}`,
                  date: r.date, heure: r.heure,
                  details: `${r.motif}${r.note ? "\n" + r.note : ""}\n(CRM ELYON & Associés)`,
                }), "_blank")}
              >📆</button>
              {!r.done && <button className="btn ghost sm" onClick={() => saveRdvClients(rdvClients.map((x) => (x.id === r.id ? { ...x, done: true } : x)))}>✓ Fait</button>}
              <button className="btn danger sm" onClick={() => confirm("Supprimer ce rendez-vous ?") && saveRdvClients(rdvClients.filter((x) => x.id !== r.id))}>✕</button>
            </div>
          </div>
        ))}
        {(rdvClients || []).filter((r) => r.clientId === client.id).length === 0 && (
          <div style={{ color: "#8593a8", fontSize: 13.5 }}>Aucun rendez-vous planifié avec ce client.</div>
        )}
      </div>

      {/* ---- Audit patrimonial ---- */}
      <div className="card" style={{ marginTop: 16, border: client.audit ? undefined : `2px solid ${GOLD}` }}>
        <div className="row" style={{ justifyContent: "space-between", marginBottom: 8 }}>
          <h2 style={{ fontSize: 17 }}>
            🩺 Audit patrimonial {client.audit
              ? <span className="badge b-navy" style={{ marginLeft: 6 }}>✓ réalisé le {fmtDate(client.audit.doneAt)}</span>
              : <span className="badge b-gold" style={{ marginLeft: 6 }}>À réaliser</span>}
          </h2>
          <div className="row">
            {client.audit && <button className="btn ghost sm" onClick={() => auditPDF(client)}>⬇️ Télécharger le PDF</button>}
            <button className="btn gold sm" onClick={() => setShowFiche(true)}>
              {client.audit ? "Voir / modifier" : "🩺 Réaliser l'audit"}
            </button>
          </div>
        </div>
        {client.audit ? (
          <AuditSynthese a={client.audit.f} />
        ) : (
          <div style={{ color: "#8593a8", fontSize: 13.5 }}>
            L'essentiel de l'audit patrimonial ELYON : état civil, revenus et fiscalité, patrimoine, objectifs et profil investisseur.
            Une fois complété, le PDF « Audit patrimonial Elyon » est téléchargeable.
          </div>
        )}
      </div>

      {showRdv && (
        <RdvClientForm
          client={client} me={me}
          onClose={() => setShowRdv(false)}
          onSave={(rdv) => { saveRdvClients([...(rdvClients || []), rdv]); setShowRdv(false); }}
        />
      )}
      {showFiche && (
        <AuditModal
          client={client}
          onClose={() => setShowFiche(false)}
          onSave={(f) => { update({ ...client, audit: { f, doneAt: (client.audit || {}).doneAt || todayISO() } }); setShowFiche(false); }}
        />
      )}
          onSave={(answers) => { update({ ...client, fichePatrimoniale: { answers, doneAt: (client.fichePatrimoniale || {}).doneAt || todayISO() } }); setShowFiche(false); }}
        />
      )}

      {showEdit && <ClientForm initial={client} onClose={() => setShowEdit(false)} onSave={(f) => { update({ ...client, ...f }); setShowEdit(false); }} />}
      {showContract && (
        <ContractForm
          initial={editContract}
          onClose={() => setShowContract(false)}
          onSave={(k) => {
            if (editContract) update({ ...client, contrats: client.contrats.map((x) => (x.id === editContract.id ? { ...x, ...k } : x)) });
            else update({ ...client, contrats: [...(client.contrats || []), { ...k, id: uid(), fichiers: [] }] });
            setShowContract(false);
          }}
        />
      )}
      {showAlert && <AlertForm onClose={() => setShowAlert(false)} onSave={(a) => { addAlert(a); setShowAlert(false); }} />}
    </div>
  );
}

function ContractCard({ contract: k, onEdit, onDelete, onFiles, onFileDelete, onQuick }) {
  const nf = nextFollowUp(k.dateSignature);
  const cycleTransfert = () => {
    const next = { "": "En cours", "En cours": "Fait", "Fait": "" }[k.transfertEtat || ""];
    onQuick({ transfertEtat: next });
  };
  const cycleEspace = () => {
    const next = { "": "Oui", "Oui": "Non", "Non": "" }[k.espaceClient || ""];
    onQuick({ espaceClient: next });
  };
  const tColor = k.transfertEtat === "Fait" ? "#1b7a3d" : k.transfertEtat === "En cours" ? "#d97706" : "#8593a8";
  const eColor = k.espaceClient === "Oui" ? "#1b7a3d" : k.espaceClient === "Non" ? "#B3261E" : "#8593a8";
  return (
    <div style={{ border: "1px solid #e3e8f0", borderRadius: 10, padding: 16, marginBottom: 12, borderLeft: `4px solid ${GOLD}` }}>
      <div className="row" style={{ justifyContent: "space-between" }}>
        <div>
          <b style={{ fontSize: 15 }}>{k.type} — {k.compagnie}</b>
          <span className="badge b-grey" style={{ marginLeft: 8 }}>N° {k.numero || "—"}</span>
        </div>
        <div className="row">
          <button
            className="btn ghost sm" onClick={cycleTransfert} title="Cliquer pour changer"
            style={{ borderColor: tColor, color: tColor, fontWeight: 600 }}
          >
            🔁 Transfert : {k.transfertEtat || "—"}
          </button>
          <button
            className="btn ghost sm" onClick={cycleEspace} title="Cliquer pour changer"
            style={{ borderColor: eColor, color: eColor, fontWeight: 600 }}
          >
            🖥️ Espace client : {k.espaceClient || "—"}
          </button>
          <button className="btn ghost sm" onClick={onEdit}>Modifier</button>
          <button className="btn danger sm" onClick={onDelete}>Supprimer</button>
        </div>
      </div>
      <div style={{ fontSize: 13.5, marginTop: 8, lineHeight: 1.9, color: "#33415c" }}>
        <b>Montant :</b> {k.montant ? fmtEUR(parseNum(k.montant)) : "—"} · <b>Frais :</b> {k.frais !== "" && k.frais !== undefined ? k.frais + " %" : "—"} ·{" "}
        <b>Signature :</b> {fmtDate(k.dateSignature)} · <b>1er prélèvement :</b> {fmtDate(k.datePrelevement)}
        {k.type === "PER" && (
          <> · <b>Transfert interne :</b> {k.transfertInterne === "oui" ? `Oui${k.fraisTransfert === "oui" ? " (avec frais)" : " (sans frais)"}` : "Non"}</>
        )}
        {nf && <> · <b>Prochain rappel (4 mois) :</b> <span style={{ color: GOLD, fontWeight: 600 }}>{nf.toLocaleDateString("fr-FR")}</span></>}
        {k.commentaire && <div style={{ marginTop: 4, fontStyle: "italic" }}>💬 {k.commentaire}</div>}
      </div>
      <div style={{ marginTop: 10 }}>
        <div className="row" style={{ justifyContent: "space-between", marginBottom: 8 }}>
          <span className="lbl" style={{ margin: 0 }}>Contrat & pièces justificatives {k.type === "Transfert" && "(y compris fiche de transfert signée)"}</span>
          <FilePicker label="+ Importer des documents" multiple onFiles={onFiles} />
        </div>
        <FileList files={k.fichiers} onDelete={onFileDelete} />
      </div>
    </div>
  );
}

function ContractForm({ initial, onSave, onClose }) {
  const [f, setF] = useState(initial || {
    type: "PER", compagnie: COMPANIES["PER"][0], numero: "", montant: "", frais: "",
    commentaire: "", dateSignature: todayISO(), datePrelevement: "", transfertInterne: "non", fraisTransfert: "non",
  });
  const set = (k, v) => setF({ ...f, [k]: v });
  const setType = (t) => setF({ ...f, type: t, compagnie: COMPANIES[t][0] });
  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{initial ? "Modifier le contrat" : "Ajouter un contrat"}</h2>
        <div className="fgrid">
          <Field label="Type de contrat">
            <select className="sel" value={f.type} onChange={(e) => setType(e.target.value)}>
              {CONTRACT_TYPES.map((t) => <option key={t}>{t}</option>)}
            </select>
          </Field>
          <Field label="Compagnie">
            <select className="sel" value={f.compagnie} onChange={(e) => set("compagnie", e.target.value)}>
              {COMPANIES[f.type].map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Field label="Numéro de contrat"><input className="in" value={f.numero} onChange={(e) => set("numero", e.target.value)} /></Field>
          <Field label="Montant du contrat (€)"><input className="in" value={f.montant} onChange={(e) => set("montant", e.target.value)} placeholder="ex : 15 000" /></Field>
          <Field label="Frais (0 à 5 %)">
            <input className="in" type="number" min="0" max="5" step="0.1" value={f.frais}
              onChange={(e) => { const v = e.target.value; if (v === "" || (parseFloat(v) >= 0 && parseFloat(v) <= 5)) set("frais", v); }} />
          </Field>
          <Field label="Date de signature"><input className="in" type="date" value={f.dateSignature} onChange={(e) => set("dateSignature", e.target.value)} /></Field>
          <Field label="Date du 1er prélèvement"><input className="in" type="date" value={f.datePrelevement} onChange={(e) => set("datePrelevement", e.target.value)} /></Field>
          {f.type === "PER" && (
            <>
              <Field label="Transfert interne à effectuer ?">
                <select className="sel" value={f.transfertInterne} onChange={(e) => set("transfertInterne", e.target.value)}>
                  <option value="non">Non</option><option value="oui">Oui</option>
                </select>
              </Field>
              {f.transfertInterne === "oui" && (
                <Field label="Frais appliqués au transfert ?">
                  <select className="sel" value={f.fraisTransfert} onChange={(e) => set("fraisTransfert", e.target.value)}>
                    <option value="non">Non</option><option value="oui">Oui</option>
                  </select>
                </Field>
              )}
            </>
          )}
        </div>
        <div style={{ marginTop: 12 }}>
          <Field label="Commentaire">
            <textarea className="ta" rows={3} value={f.commentaire} onChange={(e) => set("commentaire", e.target.value)} />
          </Field>
        </div>
        <div className="row" style={{ marginTop: 18, justifyContent: "flex-end" }}>
          <button className="btn ghost" onClick={onClose}>Annuler</button>
          <button className="btn gold" onClick={() => onSave(f)}>Enregistrer le contrat</button>
        </div>
      </div>
    </div>
  );
}

function AlertForm({ onSave, onClose }) {
  const [f, setF] = useState({ type: ALERT_TYPES[0], date: todayISO(), note: "" });
  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 480 }}>
        <h2>Créer une alerte</h2>
        <div className="grid">
          <Field label="Type d'alerte">
            <select className="sel" value={f.type} onChange={(e) => setF({ ...f, type: e.target.value })}>
              {ALERT_TYPES.map((t) => <option key={t}>{t}</option>)}
            </select>
          </Field>
          <Field label="Date du rappel">
            <input className="in" type="date" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} />
          </Field>
          <Field label="Note (optionnel)">
            <input className="in" value={f.note} onChange={(e) => setF({ ...f, note: e.target.value })} placeholder="ex : relancer pour le RIB" />
          </Field>
        </div>
        <div className="row" style={{ marginTop: 18, justifyContent: "flex-end" }}>
          <button className="btn ghost" onClick={onClose}>Annuler</button>
          <button className="btn gold" onClick={() => onSave(f)}>Programmer l'alerte</button>
        </div>
      </div>
    </div>
  );
}

/* ================= VENTES ÉQUIPE ================= */
function SalesPage({ sales, saveSales, users: allUsers, objectifs, saveObjectifs, me, clients, saveClients }) {
  const users = allUsers.filter((u) => !isTelepro(u)); /* pas de tableau de ventes pour les téléprospecteurs */
  const months = Object.keys(sales).sort();
  const [month, setMonth] = useState(months[months.length - 1]);
  const [showObj, setShowObj] = useState(false);
  const [vue, setVue] = useState("mois"); // mois | annee

  /* Créer une fiche client + le contrat associé depuis une ligne du tableau */
  const createFromRow = (u, r) => {
    /* Ligne recopiée : on attribue la fiche au commercial d'origine */
    let ownerU = u;
    if (r.mirrorOf) {
      for (const cu of users) {
        if (cu.id !== u.id && ((((sales[month] || {})[cu.id] || {}).rows) || []).some((x) => x.id === r.mirrorOf)) { ownerU = cu; break; }
      }
    }
    const words = (r.nom || "").trim().split(/\s+/);
    const nomP = words[0] || "";
    const prenomP = words.slice(1).join(" ");
    const doublon = clients.find((c) => (c.nom || "").trim().toLowerCase() === nomP.toLowerCase() && (c.prenom || "").trim().toLowerCase() === prenomP.toLowerCase());
    if (doublon) {
      if (!confirm(`Un client « ${doublon.nom.toUpperCase()} ${doublon.prenom} » existe déjà. Lui AJOUTER ce contrat ?`)) return;
      const contrat = { id: uid(), type: r.type || "", compagnie: r.compagnie || "", numero: r.ref || "", montant: "", frais: (r.frais || "").replace("%", "").trim(), commentaire: r.commentaire || "", dateSignature: r.dateCreation || todayISO(), datePrelevement: "", transfertInterne: "non", fraisTransfert: "non", fichiers: [] };
      saveClients(clients.map((c) => (c.id === doublon.id ? { ...c, contrats: [...(c.contrats || []), contrat] } : c)));
      alert("Contrat ajouté à la fiche existante ✓");
      return;
    }
    if (!confirm(`Créer la fiche client « ${nomP.toUpperCase()} ${prenomP} » avec son contrat ${r.type || ""} ${r.compagnie || ""} ?`)) return;
    const nc = {
      id: uid(), nom: nomP, prenom: prenomP, profession: "", telephone: "", email: "",
      dateNaissance: "", revenus: "", situation: "Célibataire",
      createdBy: ownerU.id, createdAt: todayISO(), alertes: [],
      contrats: [{ id: uid(), type: r.type || "", compagnie: r.compagnie || "", numero: r.ref || "", montant: "", frais: (r.frais || "").replace("%", "").trim(), commentaire: r.commentaire || "", dateSignature: r.dateCreation || todayISO(), datePrelevement: "", transfertInterne: "non", fraisTransfert: "non", fichiers: [] }],
    };
    saveClients([...clients, nc]);
    alert(`Fiche client créée ✓ (attribuée à ${ownerU.prenom}) — retrouvez-la dans l'espace Clients.`);
  };

  const addMonth = () => {
    const last = months[months.length - 1];
    const nm = nextMonthKey(last);
    const next = { ...sales, [nm]: emptyMonthData(users) };
    saveSales(next);
    setMonth(nm);
  };

  const exportCSV = () => {
    const md0 = sales[month] || {};
    const rows = [];
    users.forEach((u) => {
      (((md0[u.id] || {}).rows) || []).forEach((r) => {
        if (!(r.nom || "").trim()) return;
        rows.push([`${u.prenom} ${u.nom}`, fmtDate(r.dateCreation), r.nom, r.type, r.compagnie, r.frais, r.ref, r.commentaire, r.apporteur, r.volume, r.remuneration, r.statut]);
      });
    });
    downloadCSV(
      `ELYON_ventes_${month}.csv`,
      ["Commercial", "Date de création", "Client", "Type de contrat", "Compagnie", "Frais", "Réf. contrat", "Commentaires", "Apporteur", "Volume", "Rémunération", "Statut"],
      rows
    );
  };

  const updateCell = (userId, rowId, field, value) => {
    const md = sales[month];
    let monthData = {
      ...md,
      [userId]: {
        ...md[userId],
        rows: md[userId].rows.map((r) => (r.id === rowId ? { ...r, [field]: value } : r)),
      },
    };

    /* ---- Recopie automatique vers le tableau du manager ----
       Quand un commercial saisit/modifie une ligne, elle est dupliquée
       dans le tableau du manager, rémunération laissée VIDE (barème différent). */
    const manager = users.find((u) => u.isManager);
    if (manager && userId !== manager.id) {
      const src = monthData[userId].rows.find((r) => r.id === rowId);
      if (src && (src.nom || "").trim()) {
        const commercial = users.find((u) => u.id === userId);
        const mirrored = {
          dateCreation: src.dateCreation || "", nom: src.nom, type: src.type,
          compagnie: src.compagnie, frais: src.frais, ref: src.ref,
          commentaire: src.commentaire,
          apporteur: (src.apporteur || "").trim() || (commercial ? `${commercial.prenom} ${commercial.nom}` : ""),
          volume: src.volume, statut: src.statut,
        };
        const mgrData = monthData[manager.id] || { rows: [], nonPayes: "" };
        let rows = [...mgrData.rows];
        const idx = rows.findIndex((r) => r.mirrorOf === rowId);
        if (idx === -1) {
          /* première recopie : on prend la première ligne vide, sinon on en ajoute une */
          const freeIdx = rows.findIndex((r) => !(r.nom || "").trim() && !r.mirrorOf);
          const newRow = { ...emptyRow(), ...mirrored, mirrorOf: rowId, remuneration: "" };
          if (freeIdx === -1) rows.push(newRow); else rows[freeIdx] = newRow;
        } else {
          /* mise à jour : tout est resynchronisé SAUF la rémunération saisie par le manager */
          rows[idx] = { ...rows[idx], ...mirrored };
        }
        monthData = { ...monthData, [manager.id]: { ...mgrData, rows } };
      }
    }

    saveSales({ ...sales, [month]: monthData });
  };
  const addRows = (userId) => {
    const md = sales[month];
    saveSales({ ...sales, [month]: { ...md, [userId]: { ...md[userId], rows: [...md[userId].rows, ...Array.from({ length: 5 }, emptyRow)] } } });
  };

  const md = sales[month] || {};

  return (
    <div>
      <div className="ph">
        <div>
          <h1>Ventes de l'équipe</h1>
          <div className="sub">Vue d'ensemble mois par mois — un tableau par commercial · volume et rémunération saisis à la main</div>
        </div>
        <div className="row">
          <div className="row" style={{ gap: 0, border: "1px solid #cdd6e2", borderRadius: 8, overflow: "hidden" }}>
            {[["mois", "📅 Mensuel"], ["annee", "📊 Annuel"]].map(([k, l]) => (
              <button key={k} onClick={() => setVue(k)}
                style={{ border: "none", padding: "8px 14px", fontSize: 13, cursor: "pointer", background: vue === k ? NAVY : "#fff", color: vue === k ? "#fff" : NAVY, fontFamily: "inherit" }}>
                {l}
              </button>
            ))}
          </div>
          {vue === "mois" && <select className="sel" style={{ width: 200 }} value={month} onChange={(e) => setMonth(e.target.value)}>
            {months.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
          </select>}
          {vue === "mois" && <button className="btn ghost" onClick={exportCSV}>⬇️ Exporter (Excel)</button>}
          {vue === "mois" && me.isManager && <button className="btn ghost" onClick={() => setShowObj(true)}>🎯 Définir les objectifs</button>}
          {vue === "mois" && <button className="btn gold" onClick={addMonth}>+ Ajouter le mois suivant</button>}
        </div>
      </div>

      {vue === "annee" && <AnnualView sales={sales} users={users} me={me} />}

      {vue === "mois" && (
      <>

      {/* ---- Objectifs du mois ---- */}
      <div className="card" style={{ marginBottom: 18 }}>
        <h2 style={{ fontSize: 16, marginBottom: 12 }}>🎯 Objectifs — {monthLabel(month)}</h2>
        {!objectifs[month] && <div style={{ color: "#8593a8", fontSize: 13 }}>Aucun objectif défini pour ce mois{me.isManager ? " — cliquez sur « Définir les objectifs »." : "."}</div>}
        {objectifs[month] && (() => {
          /* Les lignes recopiées automatiquement (mirrorOf) ne comptent PAS dans les objectifs :
             seuls les contrats réellement saisis par chacun sont comptabilisés. */
          const ownRows = (uid2) => (((sales[month] || {})[uid2] || {}).rows || [])
            .filter((r) => (r.nom || "").trim() && r.statut !== "Annulé" && !r.mirrorOf);
          const Bar = ({ label, got, obj, money }) => {
            const pct = obj ? Math.min(100, Math.round((got / obj) * 100)) : null;
            if (pct === null) return null;
            return (
              <div style={{ marginTop: 6 }}>
                <div style={{ fontSize: 12, color: "#5b6b82" }}>{label} : {money ? fmtEUR(got) : got} / {money ? fmtEUR(obj) : obj} {pct >= 100 && "✅"}</div>
                <div style={{ height: 7, background: "#eef1f6", borderRadius: 4, marginTop: 3, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${pct}%`, background: pct >= 100 ? "#1b7a3d" : money ? NAVY2 : GOLD, borderRadius: 4 }} />
                </div>
              </div>
            );
          };

          const objEq = (objectifs[month] || {}).equipe || {};
          const objEqC = parseNum(objEq.contrats), objEqV = parseNum(objEq.volume);
          const totC = users.reduce((s, u) => s + ownRows(u.id).length, 0);
          const totV = users.reduce((s, u) => s + ownRows(u.id).reduce((x, r) => x + parseNum(r.volume), 0), 0);

          return (
            <>
              {(objEqC > 0 || objEqV > 0) && (
                <div style={{ background: "#fdf9f0", border: `1px solid ${GOLD}`, borderRadius: 10, padding: "10px 14px", marginBottom: 14 }}>
                  <b style={{ fontSize: 14, color: "#7a5c17" }}>🏆 Objectif général de l'équipe</b>
                  {objEqC > 0 && <Bar label={`Contrats équipe`} got={totC} obj={objEqC} />}
                  {objEqV > 0 && <Bar label={`Volume équipe`} got={totV} obj={objEqV} money />}
                </div>
              )}
              <div className="grid" style={{ gridTemplateColumns: `repeat(${Math.min(users.length, 3)}, 1fr)` }}>
                {users.map((u) => {
                  const obj = (objectifs[month] || {})[u.id] || {};
                  const objC = parseNum(obj.contrats), objV = parseNum(obj.volume);
                  if (!objC && !objV) return null;
                  const rows = ownRows(u.id);
                  const gotC = rows.length;
                  const gotV = rows.reduce((s, r) => s + parseNum(r.volume), 0);
                  return (
                    <div key={u.id} style={{ padding: "4px 2px" }}>
                      <b style={{ fontSize: 14 }}>{u.prenom} {u.nom}</b>
                      {objC > 0 && <Bar label="Contrats" got={gotC} obj={objC} />}
                      {objV > 0 && <Bar label="Volume" got={gotV} obj={objV} money />}
                    </div>
                  );
                })}
              </div>
              <div style={{ fontSize: 11.5, color: "#8593a8", marginTop: 10 }}>
                ℹ️ Les lignes recopiées automatiquement depuis les tableaux des commerciaux ne comptent pas dans les objectifs — seuls les contrats saisis par chacun sont comptabilisés.
              </div>
            </>
          );
        })()}
      </div>

      </>
      )}

      {showObj && (
        <ObjectifsForm
          users={users} month={month}
          initial={objectifs[month] || {}}
          onClose={() => setShowObj(false)}
          onSave={(v) => { saveObjectifs({ ...objectifs, [month]: v }); setShowObj(false); }}
        />
      )}

      <div className="row" style={{ marginBottom: 14, fontSize: 12.5, color: "#5b6b82" }}>
        <span className="badge b-green">Payé</span>
        <span className="badge b-red">Annulé</span>
        <span className="badge b-grey">En attente (blanc)</span>
        <span>— le statut colore automatiquement la ligne lorsque vous comparez avec vos bordereaux.</span>
      </div>

      {users.map((u) => {
        const data = md[u.id] || { rows: [], nonPayes: "" };
        const totVol = data.rows.reduce((s, r) => s + parseNum(r.volume), 0);
        const totRem = data.rows.reduce((s, r) => s + parseNum(r.remuneration), 0);
        return (
          <div className="card" key={u.id} style={{ marginBottom: 22, overflowX: "auto" }}>
            <div className="row" style={{ justifyContent: "space-between", marginBottom: 12 }}>
              <h2 style={{ fontSize: 17 }}>
                {u.prenom} {u.nom} <span className="badge b-gold" style={{ marginLeft: 6 }}>Barème {u.bareme}</span>
              </h2>
              <button className="btn ghost sm" onClick={() => addRows(u.id)}>+ 5 lignes</button>
            </div>
            <table className="t salest">
              <thead>
                <tr>
                  <th style={{ width: "9%" }}>Date création</th>
                  <th style={{ width: "13%" }}>Nom / Prénom client</th>
                  <th>Type de contrat</th>
                  <th>Compagnie</th>
                  <th style={{ width: "5%" }}>Frais</th>
                  <th>Réf. contrat</th>
                  <th style={{ width: "14%" }}>Commentaires</th>
                  <th>Apporteur</th>
                  <th>Volume</th>
                  <th>Rémunération</th>
                  <th>Statut</th>
                  <th style={{ width: 34 }}></th>
                </tr>
              </thead>
              <tbody>
                {data.rows.map((r) => (
                  <tr key={r.id} className={r.statut === "Payé" ? "paye" : r.statut === "Annulé" ? "annule" : r.statut === "Décommissionné" ? "decom" : ""} style={r.mirrorOf ? { background: "#fdf9f0" } : undefined} title={r.mirrorOf ? "Ligne recopiée automatiquement depuis le tableau d'un commercial — saisissez votre rémunération" : undefined}>
                    <td><input type="date" value={r.dateCreation || ""} onChange={(e) => updateCell(u.id, r.id, "dateCreation", e.target.value)} style={{ fontSize: 12 }} /></td>
                    <td><input value={r.nom} onChange={(e) => updateCell(u.id, r.id, "nom", e.target.value)} style={{ fontWeight: 700 }} /></td>
                    <td>
                      <select value={r.type} onChange={(e) => updateCell(u.id, r.id, "type", e.target.value)}>
                        <option value=""></option>
                        {CONTRACT_TYPES.map((t) => <option key={t}>{t}</option>)}
                      </select>
                    </td>
                    <td>
                      <select value={r.compagnie} onChange={(e) => updateCell(u.id, r.id, "compagnie", e.target.value)}>
                        <option value=""></option>
                        {(r.type && COMPANIES[r.type] ? COMPANIES[r.type] : COMPANIES["Transfert"]).map((c) => <option key={c}>{c}</option>)}
                      </select>
                    </td>
                    <td><input value={r.frais} onChange={(e) => updateCell(u.id, r.id, "frais", e.target.value)} placeholder="%" /></td>
                    <td><input value={r.ref} onChange={(e) => updateCell(u.id, r.id, "ref", e.target.value)} /></td>
                    <td><input value={r.commentaire} onChange={(e) => updateCell(u.id, r.id, "commentaire", e.target.value)} /></td>
                    <td><input value={r.apporteur} onChange={(e) => updateCell(u.id, r.id, "apporteur", e.target.value)} /></td>
                    <td><input value={r.volume} onChange={(e) => updateCell(u.id, r.id, "volume", e.target.value)} onBlur={(e) => { const f = autoEUR(e.target.value); if (f !== e.target.value) updateCell(u.id, r.id, "volume", f); }} placeholder="€" /></td>
                    <td><input value={r.remuneration} onChange={(e) => updateCell(u.id, r.id, "remuneration", e.target.value)} onBlur={(e) => { const f = autoEUR(e.target.value); if (f !== e.target.value) updateCell(u.id, r.id, "remuneration", f); }} placeholder="€" /></td>
                    <td>
                      <select value={r.statut} onChange={(e) => updateCell(u.id, r.id, "statut", e.target.value)}>
                        {STATUTS.map((s) => <option key={s}>{s}</option>)}
                      </select>
                    </td>
                    <td>
                      {(r.nom || "").trim() && (
                        <button
                          onClick={() => createFromRow(u, r)}
                          title="Créer la fiche client + le contrat associé"
                          style={{ border: "none", background: "transparent", cursor: "pointer", fontSize: 15 }}
                        >👤➕</button>
                      )}
                    </td>
                  </tr>
                ))}
                <tr className="totrow">
                  <td colSpan={8} style={{ textAlign: "right" }}>TOTAL {monthLabel(month).toUpperCase()}</td>
                  <td>{fmtEUR(totVol)}</td>
                  <td>{fmtEUR(totRem)}</td>
                  <td colSpan={2}></td>
                </tr>
                <tr className="nprow">
                  <td colSpan={2} style={{ textAlign: "right", paddingRight: 10 }}>AFFAIRES NON PAYÉES</td>
                  <td colSpan={10} style={{ fontSize: 13 }}>
                    {(() => {
                      const enAttente = data.rows.filter((r) => (r.nom || "").trim() && r.statut === "En attente");
                      if (enAttente.length === 0) return <span style={{ color: "#8593a8" }}>0 — tout est payé ou annulé ✓</span>;
                      const totalAttente = enAttente.reduce((s, r) => s + parseNum(r.remuneration), 0);
                      return (
                        <span>
                          <b>{enAttente.length}</b> affaire(s) en attente
                          {totalAttente > 0 && <> · <b>{fmtEUR(totalAttente)}</b> de rémunération</>}
                          {" — "}
                          <span style={{ color: "#5b6b82" }}>{enAttente.map((r) => r.nom).join(", ")}</span>
                        </span>
                      );
                    })()}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        );
      })}
    </div>
  );
}

/* ================= RÉMUNÉRATION & BORDEREAUX ================= */
function PayePage({ view, sales, bordereaux, saveBordereaux }) {
  const months = Object.keys(sales).sort();
  const [detailMonth, setDetailMonth] = useState(null); // clic sur une colonne → détail des affaires
  const data = months.map((m) => ({
    mois: monthLabel(m).replace(" 20", " '"),
    key: m,
    paye: (sales[m][view.id]?.rows || []).reduce((s, r) => s + parseNum(r.remuneration), 0),
  }));
  const total = data.reduce((s, d) => s + d.paye, 0);
  const userB = bordereaux[view.id] || {};

  const addBordereau = (mk, files) => {
    const next = { ...bordereaux, [view.id]: { ...userB, [mk]: [...(userB[mk] || []), ...files] } };
    saveBordereaux(next);
  };
  const delBordereau = (mk, f) => {
    sDel(`crm-file-${f.id}`);
    const next = { ...bordereaux, [view.id]: { ...userB, [mk]: (userB[mk] || []).filter((x) => x.id !== f.id) } };
    saveBordereaux(next);
  };

  return (
    <div>
      <div className="ph">
        <div>
          <h1>Rémunération — {view.prenom} {view.nom}</h1>
          <div className="sub">Évolution des payes mois par mois depuis septembre 2025 (calculée depuis le tableau des ventes)</div>
        </div>
        <div className="kpi" style={{ minWidth: 200 }}>
          <div className="n">{fmtEUR(total)}</div>
          <div className="l">Total cumulé</div>
        </div>
      </div>

      <div className="card" style={{ marginBottom: 20 }}>
        <h2 style={{ fontSize: 17, marginBottom: 14 }}>📈 Évolution des payes <span style={{ fontSize: 12, color: "#8593a8", fontWeight: 400 }}>— cliquez sur une colonne pour voir les affaires du mois</span></h2>
        <div style={{ width: "100%", height: 300 }}>
          <ResponsiveContainer>
            <BarChart data={data} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e3e8f0" />
              <XAxis dataKey="mois" tick={{ fontSize: 11, fill: "#5b6b82" }} />
              <YAxis tick={{ fontSize: 11, fill: "#5b6b82" }} tickFormatter={(v) => v.toLocaleString("fr-FR")} />
              <Tooltip formatter={(v) => [fmtEUR(v), "Rémunération"]} />
              <Bar dataKey="paye" fill={NAVY} radius={[6, 6, 0, 0]} cursor="pointer" onClick={(d) => d && (d.key || (d.payload && d.payload.key)) && setDetailMonth(d.key || d.payload.key)}>
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        {detailMonth && (() => {
          const rows = ((sales[detailMonth] || {})[view.id]?.rows || []).filter((r) => (r.nom || "").trim());
          const tot = rows.reduce((s, r) => s + parseNum(r.remuneration), 0);
          return (
            <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && setDetailMonth(null)}>
              <div className="modal" style={{ maxWidth: 960, width: "min(960px, 94vw)", maxHeight: "88vh", overflowY: "auto" }}>
                <h2>💶 Affaires de {monthLabel(detailMonth)} — {rows.length} ligne(s) · {fmtEUR(tot)}</h2>
                <div style={{ overflowX: "auto" }}>
                <table className="t nowrapt" style={{ marginTop: 12 }}>
                  <thead><tr><th>Date</th><th>Client</th><th>Type</th><th>Compagnie</th><th>Volume</th><th>Rémunération</th><th>Statut</th></tr></thead>
                  <tbody>
                    {rows.map((r) => (
                      <tr key={r.id} className={r.statut === "Payé" ? "paye" : r.statut === "Annulé" ? "annule" : r.statut === "Décommissionné" ? "decom" : ""}>
                        <td style={{ fontSize: 12 }}>{fmtDate(r.dateCreation)}</td>
                        <td><b>{r.nom}</b></td>
                        <td style={{ fontSize: 12 }}>{r.type}</td>
                        <td style={{ fontSize: 12 }}>{r.compagnie}</td>
                        <td style={{ fontSize: 12 }}>{r.volume}</td>
                        <td style={{ fontSize: 12 }}><b>{r.remuneration}</b></td>
                        <td style={{ fontSize: 12 }}>{r.statut}</td>
                      </tr>
                    ))}
                    {rows.length === 0 && <tr><td colSpan={7} style={{ color: "#8593a8", padding: 14 }}>Aucune affaire ce mois-là.</td></tr>}
                  </tbody>
                </table>
                </div>
                <div className="row" style={{ marginTop: 14, justifyContent: "flex-end" }}>
                  <button className="btn ghost" onClick={() => setDetailMonth(null)}>Fermer</button>
                </div>
              </div>
            </div>
          );
        })()}
      </div>

      <div className="card">
        <h2 style={{ fontSize: 17, marginBottom: 14 }}>🧾 Bordereaux de paiement</h2>
        <div className="grid">
          {months.slice().reverse().map((mk) => (
            <div key={mk} style={{ border: "1px solid #e3e8f0", borderRadius: 10, padding: 14 }}>
              <div className="row" style={{ justifyContent: "space-between", marginBottom: 8 }}>
                <b>{monthLabel(mk)}</b>
                <FilePicker label="+ Importer un bordereau" multiple onFiles={(files) => addBordereau(mk, files)} />
              </div>
              <FileList files={userB[mk]} onDelete={(f) => delBordereau(mk, f)} />
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ================= DOCUMENTS ================= */
function DocsPage({ docs, saveDocs, toTrash }) {
  const [newName, setNewName] = useState("");
  const [search, setSearch] = useState("");
  const addFolder = () => {
    if (!newName.trim()) return;
    saveDocs([...docs, { id: uid(), name: newName.trim(), files: [] }]);
    setNewName("");
  };
  /* Recherche par mots-clés : dans le nom du dossier ET les noms de fichiers */
  const q = search.trim().toLowerCase();
  const visibleDocs = !q ? docs : docs.filter((d) =>
    d.name.toLowerCase().includes(q) || (d.files || []).some((f) => (f.name || "").toLowerCase().includes(q))
  );
  return (
    <div>
      <div className="ph">
        <div>
          <h1>Documents partagés</h1>
          <div className="sub">Modèles de lettres, documents compagnies (MMA, Abeille, Swiss Life, Malakoff Humanis, Generali…) — téléchargeables par toute l'équipe</div>
        </div>
        <div className="row">
          <input className="in" style={{ width: 200 }} placeholder="🔍 Rechercher un dossier ou fichier…" value={search} onChange={(e) => setSearch(e.target.value)} />
          <input className="in" style={{ width: 200 }} placeholder="Nom du dossier (ex : Modèles MMA)" value={newName} onChange={(e) => setNewName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addFolder()} />
          <button className="btn gold" onClick={addFolder}>+ Créer un dossier</button>
        </div>
      </div>
      {q && <div style={{ fontSize: 12.5, color: "#8593a8", marginBottom: 10 }}>{visibleDocs.length} dossier(s) trouvé(s) pour « {search} »</div>}
      <div className="grid">
        {visibleDocs.map((d) => (
          <div className="card" key={d.id}>
            <div className="row" style={{ justifyContent: "space-between", marginBottom: 10 }}>
              <h2 style={{ fontSize: 16 }}>📁 {d.name} <span style={{ color: "#8593a8", fontSize: 13, fontWeight: 400 }}>· {d.files.length} fichier(s)</span></h2>
              <div className="row">
                <FilePicker label="+ Importer des fichiers" multiple onFiles={(files) => saveDocs(docs.map((x) => (x.id === d.id ? { ...x, files: [...x.files, ...files] } : x)))} />
                <button className="btn danger sm" onClick={() => { if (confirm("Mettre ce dossier et ses fichiers à la corbeille ? (restaurable pendant 30 jours)")) { toTrash("doc", d); saveDocs(docs.filter((x) => x.id !== d.id)); } }}>Supprimer</button>
              </div>
            </div>
            <FileList
              files={d.files}
              onDelete={(f) => { if (!confirm(`Mettre « ${f.name} » à la corbeille ?`)) return; toTrash("fichier", { file: f, folderId: d.id, folderName: d.name }); saveDocs(docs.map((x) => (x.id === d.id ? { ...x, files: x.files.filter((y) => y.id !== f.id) } : x))); }}
            />
          </div>
        ))}
        {docs.length === 0 && <div className="card" style={{ color: "#8593a8" }}>Aucun dossier. Créez par exemple « Modèles de lettres », « Documents MMA », « Documents Swiss Life »…</div>}
      </div>
    </div>
  );
}

/* ================= ÉQUIPE (Quentin uniquement) ================= */
function TeamPage({ users, saveUsers, sales, saveSales, me }) {
  const [f, setF] = useState({ prenom: "", nom: "", bareme: "Commercial" });
  const addUser = () => {
    if (!f.prenom.trim() || !f.nom.trim()) { alert("Renseignez le prénom et le nom."); return; }
    const nu = { id: uid(), prenom: f.prenom.trim(), nom: f.nom.trim(), bareme: f.bareme, isManager: f.bareme === "Manager" };
    const nextUsers = [...users, nu];
    saveUsers(nextUsers);
    // Créer son tableau dans tous les mois existants
    const nextSales = { ...sales };
    Object.keys(nextSales).forEach((mk) => {
      nextSales[mk] = { ...nextSales[mk], [nu.id]: { rows: Array.from({ length: 20 }, emptyRow), nonPayes: "" } };
    });
    saveSales(nextSales);
    setF({ prenom: "", nom: "", bareme: "Commercial" });
  };
  const removeUser = (u) => {
    if (u.id === me.id) { alert("Vous ne pouvez pas supprimer votre propre espace."); return; }
    if (!confirm(`Supprimer l'espace de ${u.prenom} ${u.nom} ? Ses tableaux de ventes seront conservés dans l'historique.`)) return;
    saveUsers(users.filter((x) => x.id !== u.id));
  };
  return (
    <div>
      <div className="ph">
        <div>
          <h1>Mon équipe</h1>
          <div className="sub">Créez un espace pour chaque commercial (tableau de ventes automatique) ou téléprospecteur (accès prospection uniquement — pas de ventes ni de clients)</div>
        </div>
      </div>
      <div className="card" style={{ marginBottom: 18 }}>
        <h2 style={{ fontSize: 17, marginBottom: 12 }}>+ Nouveau membre de l'équipe</h2>
        <div className="fgrid">
          <Field label="Prénom"><input className="in" value={f.prenom} onChange={(e) => setF({ ...f, prenom: e.target.value })} /></Field>
          <Field label="Nom"><input className="in" value={f.nom} onChange={(e) => setF({ ...f, nom: e.target.value })} /></Field>
          <Field label="Barème">
            <select className="sel" value={f.bareme} onChange={(e) => setF({ ...f, bareme: e.target.value })}>
              {BAREMES.map((b) => <option key={b}>{b}</option>)}
            </select>
          </Field>
          <div style={{ display: "flex", alignItems: "flex-end" }}>
            <button className="btn gold" onClick={addUser}>Créer l'espace</button>
          </div>
        </div>
      </div>
      <div className="grid">
        {users.map((u) => (
          <div className="clientcard" key={u.id} style={{ cursor: "default" }}>
            <div>
              <b>{u.prenom} {u.nom}</b>
              <div style={{ fontSize: 12.5, color: "#5b6b82" }}>{u.isManager ? "Manager · accès à tous les espaces · protégé par mot de passe" : "Commercial"}</div>
            </div>
            <div className="row">
              <span className={"badge " + (u.isManager ? "b-gold" : isTelepro(u) ? "b-navy" : "b-grey")}>{isTelepro(u) ? "📞 Téléprospecteur" : "Barème " + u.bareme}</span>
              {!u.isManager && (
                <button
                  className="btn ghost sm"
                  onClick={() => {
                    if (confirm(`Réinitialiser le mot de passe de ${u.prenom} ? Il/elle en créera un nouveau à sa prochaine connexion.`)) {
                      saveUsers(users.map((x) => (x.id === u.id ? { ...x, password: null } : x)));
                    }
                  }}
                >
                  🔑 Réinitialiser mdp
                </button>
              )}
              {!u.isManager && <button className="btn danger sm" onClick={() => removeUser(u)}>Supprimer</button>}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ================= OBJECTIFS (formulaire manager) ================= */
function ObjectifsForm({ users, month, initial, onSave, onClose }) {
  const [f, setF] = useState(() => {
    const base = {};
    users.forEach((u) => { base[u.id] = { contrats: (initial[u.id] || {}).contrats || "", volume: (initial[u.id] || {}).volume || "" }; });
    base.equipe = { contrats: (initial.equipe || {}).contrats || "", volume: (initial.equipe || {}).volume || "" };
    return base;
  });
  const set = (uid2, k, v) => setF({ ...f, [uid2]: { ...f[uid2], [k]: v } });
  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 560 }}>
        <h2>🎯 Objectifs — {monthLabel(month)}</h2>
        <p style={{ fontSize: 13, color: "#5b6b82", marginBottom: 14 }}>
          Laissez vide pour ne pas fixer d'objectif. Les jauges se remplissent automatiquement avec les contrats saisis dans le tableau des ventes.
        </p>
        <div className="row" style={{ marginBottom: 14, alignItems: "flex-end", background: "#fdf9f0", borderRadius: 8, padding: "8px 10px" }}>
          <div style={{ width: 150, fontSize: 14, fontWeight: 700, paddingBottom: 8, color: "#7a5c17" }}>🏆 ÉQUIPE (global)</div>
          <Field label="Contrats">
            <input className="in" style={{ width: 110 }} value={(f.equipe || {}).contrats || ""} onChange={(e) => setF({ ...f, equipe: { ...(f.equipe || {}), contrats: e.target.value } })} placeholder="ex : 40" />
          </Field>
          <Field label="Volume (€)">
            <input className="in" style={{ width: 140 }} value={(f.equipe || {}).volume || ""} onChange={(e) => setF({ ...f, equipe: { ...(f.equipe || {}), volume: e.target.value } })} placeholder="ex : 100 000" />
          </Field>
        </div>
        {users.map((u) => (
          <div key={u.id} className="row" style={{ marginBottom: 12, alignItems: "flex-end" }}>
            <div style={{ width: 150, fontSize: 14, fontWeight: 600, paddingBottom: 8 }}>{u.prenom} {u.nom}</div>
            <Field label="Contrats">
              <input className="in" style={{ width: 110 }} value={f[u.id].contrats} onChange={(e) => set(u.id, "contrats", e.target.value)} placeholder="ex : 8" />
            </Field>
            <Field label="Volume (€)">
              <input className="in" style={{ width: 140 }} value={f[u.id].volume} onChange={(e) => set(u.id, "volume", e.target.value)} placeholder="ex : 50 000" />
            </Field>
          </div>
        ))}
        <div className="row" style={{ marginTop: 18, justifyContent: "flex-end" }}>
          <button className="btn ghost" onClick={onClose}>Annuler</button>
          <button className="btn gold" onClick={() => onSave(f)}>Enregistrer les objectifs</button>
        </div>
      </div>
    </div>
  );
}

/* ================= PROSPECTION ================= */
const emptyProspect = (ownerId) => ({
  id: uid(), ownerId,
  nom: "", prenom: "", profession: "", telephone: "", ville: "", email: "", dateNaissance: "",
  per: "", prevoyance: "",
  dateAppel: todayISO(), repondu: "Oui",
  statut: "RDV pris",
  dateRdv: "", heureRdv: "",
  qualitePrise: "", noteRdv: "",
  commentaire: "",
  createdAt: todayISO(),
});

function Stars({ value }) {
  const n = parseInt(value, 10);
  if (!n) return <span style={{ color: "#c3ccd8" }}>—</span>;
  return <span style={{ color: GOLD, letterSpacing: 1 }}>{"★".repeat(n)}<span style={{ color: "#dfe4ec" }}>{"★".repeat(5 - n)}</span></span>;
}

function ProspectionPage({ prospection, saveProspection, me, users, toTrash, clients, saveClients, goClient }) {
  const [showForm, setShowForm] = useState(false);
  const [editEntry, setEditEntry] = useState(null);
  const [search, setSearch] = useState("");
  const [ownerFilter, setOwnerFilter] = useState("all");
  const [statutFilter, setStatutFilter] = useState("all");
  const [monthFilter, setMonthFilter] = useState("all");
  const [viewMode, setViewMode] = useState("table"); // table | jour | semaine | mois
  const [showImport, setShowImport] = useState(false);
  const [slotPrefill, setSlotPrefill] = useState(null); // créneau cliqué dans l'agenda

  /* Cloisonnement identique aux clients */
  const mine = me.isManager ? prospection : prospection.filter((p) => p.ownerId === me.id);

  const monthsAvailable = [...new Set(mine.map((p) => (p.dateAppel || "").slice(0, 7)).filter(Boolean))].sort().reverse();

  const scoped = mine.filter((p) => {
    if (me.isManager && ownerFilter !== "all" && p.ownerId !== ownerFilter) return false;
    if (statutFilter !== "all" && p.statut !== statutFilter) return false;
    if (monthFilter !== "all" && (p.dateAppel || "").slice(0, 7) !== monthFilter) return false;
    if (search && !(`${p.nom} ${p.prenom} ${p.profession} ${p.telephone} ${p.ville} ${p.commentaire}`.toLowerCase().includes(search.toLowerCase()))) return false;
    return true;
  }).sort((a, b) => (b.dateAppel || "").localeCompare(a.dateAppel || ""));

  /* ---- Statistiques détaillées ---- */
  const NON_REPONSE = ["1er appel — pas de réponse", "2e appel — pas de réponse", "3e appel — pas de réponse", "Pas de réponse"];
  const stat = (list) => {
    const appels = list.length;
    const repondus = list.filter((p) => !NON_REPONSE.includes(p.statut)).length;
    const rdvPris = list.filter((p) => ["RDV pris", "RDV honoré", "RDV reporté", "Proposition envoyée", "Signé"].includes(p.statut)).length;
    const rdvHonores = list.filter((p) => ["RDV honoré", "Proposition envoyée", "Signé"].includes(p.statut)).length;
    const signes = list.filter((p) => p.statut === "Signé").length;
    const qualites = list.map((p) => parseInt(p.qualitePrise, 10)).filter((n) => n >= 1);
    const notes = list.map((p) => parseInt(p.noteRdv, 10)).filter((n) => n >= 1);
    const pct = (a, b) => (b ? Math.round((a / b) * 100) : 0);
    return {
      appels, repondus, rdvPris, rdvHonores, signes,
      txReponse: pct(repondus, appels),
      txPrise: pct(rdvPris, repondus),
      txHonore: pct(rdvHonores, rdvPris),
      txTransfo: pct(signes, rdvPris),
      txGlobal: pct(signes, appels),
      qualiteMoy: qualites.length ? (qualites.reduce((s, n) => s + n, 0) / qualites.length).toFixed(1) : "—",
      noteMoy: notes.length ? (notes.reduce((s, n) => s + n, 0) / notes.length).toFixed(1) : "—",
    };
  };
  const S = stat(scoped);

  /* Répartition par profession */
  const parProfession = {};
  scoped.forEach((p) => { const k = p.profession || "Non renseignée"; parProfession[k] = (parProfession[k] || 0) + 1; });
  const professionsTri = Object.entries(parProfession).sort((a, b) => b[1] - a[1]).slice(0, 6);

  const ownerName = (p) => {
    const u = users.find((x) => x.id === p.ownerId);
    return u ? u.prenom : "—";
  };

  const save = (entry) => {
    entry = { ...entry, updatedAt: todayISO() }; /* trace de la dernière activité (relances intelligentes) */
    if (editEntry) saveProspection(prospection.map((p) => (p.id === entry.id ? entry : p)));
    else saveProspection([...prospection, entry]);
    setShowForm(false); setEditEntry(null);
  };
  const remove = (entry) => {
    if (confirm("Mettre cette fiche de prospection à la corbeille ? (restaurable pendant 30 jours)")) {
      toTrash("prospect", entry);
      saveProspection(prospection.filter((p) => p.id !== entry.id));
      setShowForm(false); setEditEntry(null);
    }
  };

  /* Conversion d'un prospect en fiche client (pré-remplie, attribuée au bon commercial) */
  const convert = (entry) => {
    const norm = (s) => (s || "").trim().toLowerCase();
    const doublon = clients.find((c) => norm(c.nom) === norm(entry.nom) && norm(c.prenom) === norm(entry.prenom));
    if (doublon && !confirm(`Un client « ${doublon.nom.toUpperCase()} ${doublon.prenom} » existe déjà. Créer quand même une nouvelle fiche ?`)) return;
    const newClient = {
      id: uid(),
      nom: entry.nom || "", prenom: entry.prenom || "",
      profession: entry.profession || "", telephone: entry.telephone || "",
      email: entry.email || "", dateNaissance: entry.dateNaissance || "", revenus: "", situation: "Célibataire",
      createdBy: entry.ownerId || me.id, createdAt: todayISO(),
      contrats: [], alertes: [],
      historique: [{
        id: uid(), type: "📝 Note interne", date: todayISO(),
        byId: me.id, createdAt: new Date().toISOString(),
        note: `Converti depuis la prospection (statut : ${entry.statut}).`
          + (entry.ville ? ` Ville : ${entry.ville}.` : "")
          + (entry.dateRdv ? ` RDV du ${fmtDate(entry.dateRdv)}${entry.heureRdv ? " à " + entry.heureRdv : ""}.` : "")
          + (entry.commentaire ? ` Commentaire prospection : ${entry.commentaire}` : ""),
      }],
    };
    saveClients([...clients, newClient]);
    saveProspection(prospection.map((p) => (p.id === entry.id ? { ...p, convertedClientId: newClient.id } : p)));
    setShowForm(false); setEditEntry(null);
    if (confirm("Fiche client créée ✓ Voulez-vous l'ouvrir maintenant ?")) goClient(newClient.id);
  };

  /* Annulation d'une conversion : le client repart en corbeille, le prospect redevient convertible */
  const unconvert = (entry) => {
    if (!confirm("Annuler la conversion ? La fiche client créée sera mise à la corbeille (restaurable 30 jours) et le prospect restera dans la prospection.")) return;
    const c = clients.find((x) => x.id === entry.convertedClientId);
    if (c) {
      toTrash("client", c);
      saveClients(clients.filter((x) => x.id !== c.id));
    }
    saveProspection(prospection.map((p) => (p.id === entry.id ? { ...p, convertedClientId: null } : p)));
    setShowForm(false); setEditEntry(null);
  };

  const exportCSV = () => {
    downloadCSV(
      `ELYON_prospection_${todayISO()}.csv`,
      ["Commercial", "Nom", "Prénom", "Profession", "Téléphone", "Ville", "Date de l'appel", "Répondu", "Statut", "Date du RDV", "Heure", "Qualité prise de RDV (/5)", "Note du RDV (/5)", "Commentaire"],
      scoped.map((p) => [ownerName(p), p.nom, p.prenom, p.profession, p.telephone, p.ville, fmtDate(p.dateAppel), p.repondu, p.statut, fmtDate(p.dateRdv), p.heureRdv, p.qualitePrise, p.noteRdv, p.commentaire])
    );
  };

  return (
    <div>
      <div className="ph">
        <div>
          <h1>🎯 Prospection</h1>
          <div className="sub">
            {me.isManager ? "Vue manager : toute l'équipe" : "Votre espace de prospection personnel"} — {scoped.length} fiche(s) affichée(s)
          </div>
        </div>
        <div className="row">
          {me.isManager && <button className="btn ghost" onClick={() => setShowImport(true)}>📥 Importer une liste</button>}
          <button className="btn ghost" onClick={exportCSV}>⬇️ Exporter (Excel)</button>
          <button className="btn gold" onClick={() => { setEditEntry(null); setShowForm(true); }}>+ Nouvel appel / RDV</button>
        </div>
      </div>

      {/* ---- KPI (cliquables : un clic filtre le tableau) ---- */}
      {(() => {
        const goto = (statut) => { setStatutFilter(statut); setViewMode("table"); };
        const K = ({ n, l, statut }) => (
          <div className="kpi" onClick={() => goto(statut)} style={{ cursor: "pointer" }} title="Cliquer pour afficher le détail">
            <div className="n">{n}</div><div className="l">{l}</div>
          </div>
        );
        return (
          <div className="kpis" style={{ marginBottom: 14 }}>
            <K n={S.appels} l="Appels enregistrés" statut="all" />
            <K n={`${S.txReponse}%`} l={`Taux de réponse (${S.repondus}/${S.appels})`} statut="all" />
            <K n={S.rdvPris} l="RDV pris" statut="RDV pris" />
            <K n={`${S.txPrise}%`} l="Taux de prise de RDV*" statut="RDV pris" />
            <K n={`${S.txHonore}%`} l={`RDV honorés (${S.rdvHonores}/${S.rdvPris})`} statut="RDV honoré" />
            <K n={S.signes} l="Signés" statut="Signé" />
            <K n={`${S.txTransfo}%`} l="Transformation RDV → signé" statut="Signé" />
            <K n={`${S.txGlobal}%`} l="Transformation globale" statut="all" />
          </div>
        );
      })()}
      <div className="row" style={{ marginBottom: 14, fontSize: 12, color: "#8593a8" }}>
        <span>* RDV pris / appels répondus</span>
        <span>· Qualité moyenne de prise de RDV : <b style={{ color: NAVY }}>{S.qualiteMoy}/5</b></span>
        <span>· Note moyenne des RDV : <b style={{ color: NAVY }}>{S.noteMoy}/5</b></span>
      </div>

      {/* ---- Listes importées (téléchargeables) ---- */}
      {me.isManager && (() => {
        const batches = {};
        prospection.filter((p) => p.batchId).forEach((p) => {
          (batches[p.batchId] = batches[p.batchId] || { nom: p.batchNom, date: p.batchDate, entries: [] }).entries.push(p);
        });
        const list = Object.entries(batches).sort((a, b) => (b[1].date || "").localeCompare(a[1].date || ""));
        if (!list.length) return null;
        const dlBatch = (b) => downloadCSV(
          `ELYON_liste_${(b.nom || "import").replace(/\.[^.]+$/, "")}.csv`,
          ["Commercial", "Nom", "Prénom", "Profession", "Téléphone", "Ville", "E-mail", "Statut", "Date appel", "RDV le", "PER", "Prévoyance", "Commentaire"],
          b.entries.map((p) => [ownerName(p), p.nom, p.prenom, p.profession, (p.telephone || "").replace(/\s/g, ""), p.ville, p.email, p.statut, fmtDate(p.dateAppel), fmtDate(p.dateRdv), p.per, p.prevoyance, p.commentaire])
        );
        return (
          <div className="card" style={{ marginBottom: 16 }}>
            <h2 style={{ fontSize: 15, marginBottom: 10 }}>📚 Listes importées</h2>
            {list.map(([id, b]) => {
              const signes = b.entries.filter((p) => p.statut === "Signé").length;
              const traites = b.entries.filter((p) => !["À rappeler"].includes(p.statut)).length;
              return (
                <div key={id} className="row" style={{ justifyContent: "space-between", padding: "7px 0", borderBottom: "1px solid #eef1f6" }}>
                  <div style={{ fontSize: 13.5 }}>
                    <b>📄 {b.nom || "Liste"}</b>
                    <span style={{ color: "#8593a8", fontSize: 12 }}> — importée le {fmtDate(b.date)} · {b.entries.length} prospect(s) · {traites} travaillé(s) · {signes} signé(s)</span>
                  </div>
                  <button className="btn ghost sm" onClick={() => dlBatch(b)}>⬇️ Télécharger</button>
                </div>
              );
            })}
          </div>
        );
      })()}

      {/* ---- Répartition par profession ---- */}
      {professionsTri.length > 0 && (
        <div className="card" style={{ marginBottom: 16 }}>
          <h2 style={{ fontSize: 15, marginBottom: 10 }}>Répartition par profession</h2>
          <div className="row" style={{ flexWrap: "wrap" }}>
            {professionsTri.map(([prof, n]) => (
              <span key={prof} className="badge b-navy" style={{ fontSize: 12.5 }}>{prof} · {n}</span>
            ))}
          </div>
        </div>
      )}

      {/* ---- Choix de la vue ---- */}
      <div className="row" style={{ marginBottom: 14, flexWrap: "wrap" }}>
        <div className="row" style={{ gap: 0, border: "1px solid #cdd6e2", borderRadius: 8, overflow: "hidden" }}>
          {[["table", "📋 Tableau"], ["jour", "📆 Jour"], ["semaine", "📅 Semaine"], ["mois", "🗓️ Mois"]].map(([k, l]) => (
            <button
              key={k}
              onClick={() => setViewMode(k)}
              style={{ border: "none", padding: "8px 14px", fontSize: 13, cursor: "pointer", background: viewMode === k ? NAVY : "#fff", color: viewMode === k ? "#fff" : NAVY, fontFamily: "inherit" }}
            >
              {l}
            </button>
          ))}
        </div>
        {me.isManager && (
          <select className="sel" value={ownerFilter} onChange={(e) => setOwnerFilter(e.target.value)}>
            <option value="all">Tous les commerciaux</option>
            {users.map((u) => <option key={u.id} value={u.id}>{u.prenom} {u.nom}</option>)}
          </select>
        )}
        {viewMode === "table" && (
          <>
            <select className="sel" value={statutFilter} onChange={(e) => setStatutFilter(e.target.value)}>
              <option value="all">Tous les statuts</option>
              {PROSPECTION_STATUTS.map((s) => <option key={s}>{s}</option>)}
            </select>
            <select className="sel" value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}>
              <option value="all">Tous les mois</option>
              {monthsAvailable.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
            </select>
          </>
        )}
        <input className="in" style={{ width: 240 }} placeholder="Rechercher (nom, profession, ville…)" value={search} onChange={(e) => setSearch(e.target.value)} />
      </div>

      {viewMode !== "table" && (
        <RdvCalendar
          entries={(me.isManager && ownerFilter !== "all" ? mine.filter((p) => p.ownerId === ownerFilter) : mine).filter((p) => p.dateRdv)}
          mode={viewMode}
          users={users} me={me}
          onOpen={(p) => { setEditEntry(p); setShowForm(true); }}
          onDayOpen={() => setViewMode("jour")}
          onCreateSlot={(dateIso, heure) => {
            setEditEntry(null);
            setSlotPrefill({ dateRdv: dateIso, heureRdv: heure, statut: "RDV pris" });
            setShowForm(true);
          }}
        />
      )}

      {viewMode === "table" && (
      <div className="card" style={{ overflowX: "auto" }}>
        <table className="t">
          <thead>
            <tr>
              {me.isManager && <th>Commercial</th>}
              <th>Prospect</th>
              <th>Profession</th>
              <th>Téléphone</th>
              <th>Né(e) le</th>
              <th>Appel le</th>
              <th>Statut</th>
              <th>RDV le</th>
              <th>PER</th>
              <th>Prév.</th>
              <th>Qualité prise</th>
              <th>Note RDV</th>
              <th style={{ width: "16%" }}>Commentaire</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {scoped.map((p) => (
              <tr key={p.id} style={PROSPECTION_ROW_BG[p.statut] ? { background: PROSPECTION_ROW_BG[p.statut] } : undefined}>
                {me.isManager && <td style={{ fontSize: 12.5 }}>{ownerName(p)}</td>}
                <td><b>{(p.nom || "").toUpperCase()} {p.prenom}</b>
                  {p.ville && <div style={{ fontSize: 11.5, color: "#8593a8" }}>{p.ville}</div>}
                  {p.convertedClientId && (
                    <div>
                      <span className="badge b-gold" style={{ fontSize: 10.5, cursor: "pointer" }} onClick={(e) => { e.stopPropagation(); goClient(p.convertedClientId); }}>👥 Client ✓</span>
                    </div>
                  )}
                </td>
                <td style={{ fontSize: 12.5 }}>{p.profession || "—"}</td>
                <td style={{ fontSize: 12.5 }}>{(p.telephone || "").replace(/\s/g, "") || "—"}</td>
                <td style={{ fontSize: 12.5 }}>{p.dateNaissance ? fmtDate(p.dateNaissance) : "—"}</td>
                <td style={{ fontSize: 12.5 }}>{fmtDate(p.dateAppel)}</td>
                <td><span className="badge" style={{ background: "#fff", border: `1px solid ${PROSPECTION_COLORS[p.statut] || "#8593a8"}`, color: PROSPECTION_COLORS[p.statut] || "#8593a8" }}>{p.statut}</span></td>
                <td style={{ fontSize: 12.5 }}>{p.dateRdv ? <>{fmtDate(p.dateRdv)}{p.heureRdv && <div style={{ fontSize: 11, color: "#8593a8" }}>{p.heureRdv}</div>}</> : "—"}</td>
                <td style={{ fontSize: 12.5 }}>{p.per || "—"}</td>
                <td style={{ fontSize: 12.5 }}>{p.prevoyance || "—"}</td>
                <td><Stars value={p.qualitePrise} /></td>
                <td><Stars value={p.noteRdv} /></td>
                <td style={{ fontSize: 12, color: "#5b6b82" }}>{p.commentaire || "—"}</td>
                <td><button className="btn ghost sm" onClick={() => { setEditEntry(p); setShowForm(true); }}>✏️</button></td>
              </tr>
            ))}
            {scoped.length === 0 && (
              <tr><td colSpan={me.isManager ? 14 : 13} style={{ color: "#8593a8", fontSize: 13.5, padding: 18 }}>Aucune fiche. Cliquez sur « + Nouvel appel / RDV » pour enregistrer votre premier appel.</td></tr>
            )}
          </tbody>
        </table>
      </div>
      )}

      {showImport && (
        <ImportListeModal
          users={users}
          onClose={() => setShowImport(false)}
          onImport={(entries) => { saveProspection([...prospection, ...entries]); setShowImport(false); alert(`${entries.length} prospect(s) importé(s) ✓`); }}
        />
      )}

      {showForm && (
        <ProspectForm
          initial={editEntry}
          prefill={slotPrefill}
          me={me} users={users}
          onClose={() => { setShowForm(false); setEditEntry(null); setSlotPrefill(null); }}
          onSave={save}
          onDelete={editEntry ? () => remove(editEntry) : null}
          onConvert={editEntry && !editEntry.convertedClientId ? () => convert(editEntry) : null}
          onUnconvert={editEntry && editEntry.convertedClientId ? () => unconvert(editEntry) : null}
        />
      )}
    </div>
  );
}

function ProspectForm({ initial, prefill, me, users, onSave, onClose, onDelete, onConvert, onUnconvert }) {
  const [f, setF] = useState(initial || { ...emptyProspect(me.id), ...(prefill || {}) });
  const set = (k, v) => setF({ ...f, [k]: v });
  const showRdvFields = ["RDV pris", "RDV honoré", "RDV annulé", "RDV reporté", "Proposition envoyée", "Signé", "Perdu"].includes(f.statut);
  const showNoteRdv = ["RDV honoré", "Proposition envoyée", "Signé", "Perdu"].includes(f.statut);
  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 640 }}>
        <h2>{initial ? "Modifier la fiche prospection" : "Nouvel appel / RDV"}</h2>
        <div className="fgrid">
          <Field label="Nom *"><input className="in" value={f.nom} onChange={(e) => set("nom", e.target.value)} /></Field>
          <Field label="Prénom"><input className="in" value={f.prenom} onChange={(e) => set("prenom", e.target.value)} /></Field>
          <Field label="Profession">
            <input className="in" value={f.profession} onChange={(e) => set("profession", e.target.value.toUpperCase())} placeholder="ex : INFIRMIÈRE LIBÉRALE" style={{ textTransform: "uppercase" }} />
          </Field>
          <Field label="Téléphone"><input className="in" value={f.telephone} onChange={(e) => set("telephone", e.target.value.replace(/\s/g, ""))} /></Field>
          <Field label="Date de naissance"><input className="in" type="date" value={f.dateNaissance || ""} onChange={(e) => set("dateNaissance", e.target.value)} /></Field>
          <Field label="Ville"><input className="in" value={f.ville} onChange={(e) => set("ville", e.target.value)} /></Field>
          <Field label="E-mail"><input className="in" value={f.email || ""} onChange={(e) => set("email", e.target.value)} /></Field>
          <Field label="PER"><input className="in" value={f.per || ""} onChange={(e) => set("per", e.target.value)} placeholder="" /></Field>
          <Field label="Prévoyance"><input className="in" value={f.prevoyance || ""} onChange={(e) => set("prevoyance", e.target.value)} placeholder="" /></Field>
          {me.isManager && (
            <Field label="Commercial">
              <select className="sel" value={f.ownerId} onChange={(e) => set("ownerId", e.target.value)}>
                {users.map((u) => <option key={u.id} value={u.id}>{u.prenom} {u.nom}</option>)}
              </select>
            </Field>
          )}
          <Field label="Date de l'appel (prise de RDV)"><input className="in" type="date" value={f.dateAppel} onChange={(e) => set("dateAppel", e.target.value)} /></Field>
          <Field label="Statut">
            <select className="sel" value={f.statut} onChange={(e) => set("statut", e.target.value)}>
              {PROSPECTION_STATUTS.map((s) => <option key={s}>{s}</option>)}
            </select>
          </Field>
          {showRdvFields && (
            <>
              <Field label="Date du RDV"><input className="in" type="date" value={f.dateRdv} onChange={(e) => set("dateRdv", e.target.value)} /></Field>
              <Field label="Heure du RDV"><input className="in" type="time" value={f.heureRdv} onChange={(e) => set("heureRdv", e.target.value)} /></Field>
              <Field label="Qualité de la prise de RDV (/5)">
                <select className="sel" value={f.qualitePrise} onChange={(e) => set("qualitePrise", e.target.value)}>
                  <option value="">—</option>
                  {NOTES_5.map((n) => <option key={n} value={n}>{"★".repeat(parseInt(n, 10))} ({n}/5)</option>)}
                </select>
              </Field>
            </>
          )}
          {showNoteRdv && (
            <Field label="Note du rendez-vous (/5)">
              <select className="sel" value={f.noteRdv} onChange={(e) => set("noteRdv", e.target.value)}>
                <option value="">—</option>
                {NOTES_5.map((n) => <option key={n} value={n}>{"★".repeat(parseInt(n, 10))} ({n}/5)</option>)}
              </select>
            </Field>
          )}
        </div>
        <div style={{ marginTop: 12 }}>
          <Field label="Commentaire">
            <textarea className="ta" rows={3} value={f.commentaire} onChange={(e) => set("commentaire", e.target.value)} placeholder="ex : très intéressé par un PER, rappeler après ses congés…" />
          </Field>
        </div>
        <div className="row" style={{ marginTop: 18, justifyContent: "space-between" }}>
          <div className="row">
            {onDelete && <button className="btn danger" onClick={onDelete}>Supprimer</button>}
            {f.dateRdv && (
              <button
                className="btn ghost sm" style={{ alignSelf: "center" }}
                onClick={() => window.open(gcalUrl({
                  titre: `RDV prospection — ${(f.nom || "").toUpperCase()} ${f.prenom || ""}`,
                  date: f.dateRdv, heure: f.heureRdv,
                  details: `${f.profession || ""}${f.telephone ? " · Tél : " + f.telephone : ""}${f.commentaire ? "\n" + f.commentaire : ""}\n(CRM ELYON & Associés)`,
                  lieu: f.ville || "",
                }), "_blank")}
              >
                📆 Ajouter à Google Agenda
              </button>
            )}
            {onConvert && <button className="btn ghost" style={{ borderColor: GOLD, color: "#7a5c17" }} onClick={onConvert}>👥 Convertir en client</button>}
            {initial && initial.convertedClientId && <span className="badge b-gold" style={{ alignSelf: "center" }}>✓ Déjà converti en client</span>}
            {onUnconvert && <button className="btn ghost sm" style={{ borderColor: "#B3261E", color: "#B3261E", alignSelf: "center" }} onClick={onUnconvert}>↩️ Annuler la conversion</button>}
          </div>
          <div className="row">
            <button className="btn ghost" onClick={onClose}>Annuler</button>
            <button className="btn gold" onClick={() => { if (!f.nom.trim()) { alert("Le nom du prospect est obligatoire."); return; } onSave(f); }}>Enregistrer</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ================= RECHERCHE GLOBALE ================= */
function GlobalSearch({ clients, prospection, me, users, goClient, goProspection }) {
  const [q, setQ] = useState("");

  const results = useMemo(() => {
    const query = q.trim().toLowerCase();
    if (query.length < 2) return null;
    const out = [];
    const visibleClients = me.isManager ? clients : clients.filter((c) => (c.createdBy || "quentin") === me.id);
    visibleClients.forEach((c) => {
      const hay = `${c.nom} ${c.prenom} ${c.profession || ""} ${c.telephone || ""} ${c.email || ""}`.toLowerCase();
      const contractHit = (c.contrats || []).some((k) => `${k.type || ""} ${k.compagnie || ""} ${k.ref || ""}`.toLowerCase().includes(query));
      if (hay.includes(query) || contractHit) out.push({ kind: "client", label: `${c.nom.toUpperCase()} ${c.prenom}`, sub: c.profession || "Client", id: c.id });
    });
    const visibleProspects = me.isManager ? prospection : prospection.filter((p) => p.ownerId === me.id);
    visibleProspects.forEach((p) => {
      const hay = `${p.nom} ${p.prenom} ${p.profession || ""} ${p.telephone || ""} ${p.ville || ""}`.toLowerCase();
      if (hay.includes(query)) out.push({ kind: "prospect", label: `${(p.nom || "").toUpperCase()} ${p.prenom || ""}`, sub: `Prospection · ${p.statut}`, id: p.id });
    });
    return out.slice(0, 8);
  }, [q, clients, prospection, me]);

  return (
    <div style={{ padding: "0 14px 10px", position: "relative" }}>
      <input
        className="in"
        style={{ width: "100%", background: "rgba(255,255,255,.08)", border: "1px solid rgba(255,255,255,.2)", color: "#fff", fontSize: 13 }}
        placeholder="🔍 Recherche globale…"
        value={q}
        onChange={(e) => setQ(e.target.value)}
      />
      {results && (
        <div style={{ position: "absolute", top: "100%", left: 14, right: 14, background: "#fff", borderRadius: 10, boxShadow: "0 14px 40px rgba(0,0,0,.35)", zIndex: 60, overflow: "hidden" }}>
          {results.length === 0 && <div style={{ padding: 12, fontSize: 13, color: "#8593a8" }}>Aucun résultat.</div>}
          {results.map((r) => (
            <div
              key={r.kind + r.id}
              onClick={() => { setQ(""); r.kind === "client" ? goClient(r.id) : goProspection(); }}
              style={{ padding: "9px 12px", cursor: "pointer", borderBottom: "1px solid #f0f2f6" }}
              onMouseEnter={(e) => (e.currentTarget.style.background = "#fdf9f0")}
              onMouseLeave={(e) => (e.currentTarget.style.background = "#fff")}
            >
              <b style={{ fontSize: 13.5, color: NAVY }}>{r.kind === "client" ? "👥" : "🎯"} {r.label}</b>
              <div style={{ fontSize: 11.5, color: "#8593a8" }}>{r.sub}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ================= HISTORIQUE DES INTERACTIONS ================= */
const INTERACTION_TYPES = ["📞 Appel", "🤝 Rendez-vous", "✉️ E-mail", "🔁 Relance", "📮 Courrier", "📝 Note interne"];

function HistoriqueCard({ client, update, me, users }) {
  const [showAdd, setShowAdd] = useState(false);
  const [f, setF] = useState({ type: INTERACTION_TYPES[0], date: todayISO(), note: "" });

  const historique = (client.historique || []).slice().sort((a, b) =>
    (b.date + (b.createdAt || "")).localeCompare(a.date + (a.createdAt || ""))
  );

  const authorName = (byId) => {
    const u = (users || []).find((x) => x.id === byId);
    return u ? `${u.prenom} ${u.nom}` : "—";
  };

  const add = () => {
    if (!f.note.trim()) { alert("Écrivez une note avant d'enregistrer."); return; }
    const entry = { ...f, id: uid(), byId: me.id, createdAt: new Date().toISOString() };
    update({ ...client, historique: [...(client.historique || []), entry] });
    setF({ type: INTERACTION_TYPES[0], date: todayISO(), note: "" });
    setShowAdd(false);
  };
  const remove = (id) => {
    if (confirm("Supprimer cette entrée de l'historique ?")) {
      update({ ...client, historique: (client.historique || []).filter((h) => h.id !== id) });
    }
  };

  return (
    <div className="card" style={{ marginTop: 16 }}>
      <div className="row" style={{ justifyContent: "space-between", marginBottom: 14 }}>
        <h2 style={{ fontSize: 17 }}>🕐 Historique des interactions ({historique.length})</h2>
        <button className="btn gold sm" onClick={() => setShowAdd(!showAdd)}>{showAdd ? "Fermer" : "+ Ajouter une interaction"}</button>
      </div>

      {showAdd && (
        <div style={{ background: "#f8f9fc", border: "1px solid #e3e8f0", borderRadius: 10, padding: 14, marginBottom: 14 }}>
          <div className="row" style={{ flexWrap: "wrap", alignItems: "flex-end" }}>
            <Field label="Type">
              <select className="sel" value={f.type} onChange={(e) => setF({ ...f, type: e.target.value })}>
                {INTERACTION_TYPES.map((t) => <option key={t}>{t}</option>)}
              </select>
            </Field>
            <Field label="Date">
              <input className="in" type="date" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} />
            </Field>
          </div>
          <div style={{ marginTop: 10 }}>
            <Field label="Note">
              <textarea
                className="ta" rows={2} value={f.note}
                onChange={(e) => setF({ ...f, note: e.target.value })}
                placeholder="ex : appel de suivi, souhaite augmenter son VP à 300 €/mois, rappeler en septembre…"
                autoFocus
              />
            </Field>
          </div>
          <div className="row" style={{ marginTop: 10, justifyContent: "flex-end" }}>
            <button className="btn gold sm" onClick={add}>Enregistrer</button>
          </div>
        </div>
      )}

      {historique.length === 0 && !showAdd && (
        <div style={{ color: "#8593a8", fontSize: 13.5 }}>
          Aucune interaction enregistrée. Notez ici chaque appel, RDV ou relance pour garder le fil du dossier.
        </div>
      )}

      <div style={{ position: "relative" }}>
        {historique.map((h, i) => (
          <div key={h.id} style={{ display: "flex", gap: 12, paddingBottom: i < historique.length - 1 ? 14 : 0, position: "relative" }}>
            {/* ligne de temps */}
            {i < historique.length - 1 && (
              <div style={{ position: "absolute", left: 9, top: 22, bottom: -2, width: 2, background: "#e3e8f0" }} />
            )}
            <div style={{ width: 20, height: 20, borderRadius: "50%", background: "#fff", border: `2px solid ${GOLD}`, flexShrink: 0, zIndex: 1, marginTop: 2 }} />
            <div style={{ flex: 1, minWidth: 0 }}>
              <div className="row" style={{ justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <b style={{ fontSize: 13.5 }}>{h.type}</b>
                  <span style={{ fontSize: 12, color: "#8593a8", marginLeft: 8 }}>
                    {fmtDate(h.date)} · par {authorName(h.byId)}
                  </span>
                </div>
                {(me.isManager || h.byId === me.id) && (
                  <button className="btn danger sm" style={{ padding: "2px 8px" }} onClick={() => remove(h.id)}>✕</button>
                )}
              </div>
              <div style={{ fontSize: 13.5, color: "#33415c", marginTop: 3, whiteSpace: "pre-wrap" }}>{h.note}</div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ================= CALENDRIER DES RDV PROSPECTION (façon Google Agenda) ================= */
const RDV_OWNER_STYLE = {
  arzou: { bg: "#d4ede6", border: "#1b6e5a", text: "#0d4a3c" },
  simon: { bg: "#e2f6df", border: "#3f9e4d", text: "#1d5c26" },
  quentin: { bg: "#dceafa", border: "#1d6fb8", text: "#0d3f6e" },
};
const RDV_SIGNED_STYLE = { bg: "#f7e9c4", border: "#C9A24B", text: "#6b4e0e" };

function RdvCalendar({ entries, mode, users, me, onOpen, onDayOpen, onCreateSlot }) {
  const [ref, setRef] = useState(() => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; });
  const [now, setNow] = useState(() => new Date());
  useEffect(() => { const t = setInterval(() => setNow(new Date()), 60000); return () => clearInterval(t); }, []);

  const H_START = 8, H_END = 20, HPX = 52; /* grille horaire 8h → 20h */
  const isoOf = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  const startOfWeek = (d) => { const x = new Date(d); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); x.setHours(0, 0, 0, 0); return x; };
  const today = isoOf(new Date());
  const DAYS = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"];
  const DAYS_FULL = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"];

  const byDate = {};
  entries.forEach((p) => { (byDate[p.dateRdv] = byDate[p.dateRdv] || []).push(p); });
  Object.values(byDate).forEach((list) => list.sort((a, b) => (a.heureRdv || "").localeCompare(b.heureRdv || "")));

  const ownerName = (p) => (users.find((x) => x.id === p.ownerId) || {}).prenom || "?";
  const styleOf = (p) => (p.statut === "Signé" ? RDV_SIGNED_STYLE : (RDV_OWNER_STYLE[p.ownerId] || { bg: "#eef1f6", border: NAVY2, text: NAVY }));
  const minutesOf = (heure) => { const [h, m] = (heure || "9:00").split(":").map(Number); return h * 60 + (m || 0); };

  const nav = (dir) => {
    const d = new Date(ref);
    if (mode === "jour") d.setDate(d.getDate() + dir);
    else if (mode === "semaine") d.setDate(d.getDate() + dir * 7);
    else d.setMonth(d.getMonth() + dir);
    setRef(d);
  };

  /* ---- Bloc d'événement positionné sur la grille horaire ---- */
  const TimedEvent = ({ p, wide }) => {
    const st = styleOf(p);
    const top = ((minutesOf(p.heureRdv) - H_START * 60) / 60) * HPX;
    return (
      <div
        onClick={(e) => { e.stopPropagation(); onOpen(p); }}
        title={`${p.heureRdv} · ${(p.nom || "").toUpperCase()} ${p.prenom || ""} · ${p.statut} · ${ownerName(p)}${p.commentaire ? "\n" + p.commentaire : ""}`}
        style={{
          position: "absolute", left: 3, right: 3, top: Math.max(0, top), height: HPX - 6,
          background: st.bg, borderLeft: `4px solid ${st.border}`, borderRadius: 8,
          padding: wide ? "5px 10px" : "3px 7px", cursor: "pointer", overflow: "hidden",
          boxShadow: "0 1px 3px rgba(11,37,69,.12)", zIndex: 2,
        }}
      >
        <div style={{ fontSize: wide ? 13 : 11.5, fontWeight: 700, color: st.text, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
          {p.heureRdv} · {(p.nom || "").toUpperCase()} {p.prenom || ""} {p.statut === "Signé" && "🏆"}
        </div>
        <div style={{ fontSize: wide ? 11.5 : 10.5, color: st.text, opacity: 0.85, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
          {ownerName(p)}{p.profession ? " · " + p.profession : ""}{wide && p.telephone ? " · 📞 " + p.telephone : ""}
        </div>
      </div>
    );
  };

  /* ---- Colonne d'un jour avec grille horaire cliquable ---- */
  const DayColumn = ({ d, wide }) => {
    const dIso = isoOf(d);
    const list = (byDate[dIso] || []).filter((p) => p.heureRdv);
    const noHour = (byDate[dIso] || []).filter((p) => !p.heureRdv);
    const isToday = dIso === today;
    const nowTop = isToday ? ((now.getHours() * 60 + now.getMinutes() - H_START * 60) / 60) * HPX : null;
    return (
      <div style={{ position: "relative", flex: 1, borderLeft: "1px solid #eef1f6", minWidth: 0 }}>
        {/* lignes des heures + créneaux cliquables */}
        {Array.from({ length: H_END - H_START }, (_, i) => (
          <div
            key={i}
            onClick={() => onCreateSlot && onCreateSlot(dIso, `${String(H_START + i).padStart(2, "0")}:00`)}
            title="Cliquer pour créer un RDV à ce créneau"
            style={{ height: HPX, borderBottom: "1px solid #f0f3f8", cursor: onCreateSlot ? "pointer" : "default" }}
            onMouseEnter={(e) => (e.currentTarget.style.background = "#fafcff")}
            onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
          />
        ))}
        {list.map((p) => <TimedEvent key={p.id} p={p} wide={wide} />)}
        {noHour.length > 0 && (
          <div style={{ position: "absolute", top: 0, left: 3, right: 3, zIndex: 3 }}>
            {noHour.map((p) => {
              const st = styleOf(p);
              return (
                <div key={p.id} onClick={(e) => { e.stopPropagation(); onOpen(p); }}
                  style={{ background: st.bg, borderLeft: `4px solid ${st.border}`, borderRadius: 6, fontSize: 10.5, padding: "2px 6px", marginBottom: 2, cursor: "pointer", color: st.text, fontWeight: 700 }}>
                  {(p.nom || "").toUpperCase()} (heure ?)
                </div>
              );
            })}
          </div>
        )}
        {nowTop !== null && nowTop >= 0 && nowTop <= (H_END - H_START) * HPX && (
          <div style={{ position: "absolute", left: 0, right: 0, top: nowTop, zIndex: 4, pointerEvents: "none" }}>
            <div style={{ height: 2, background: "#ea4335" }} />
            <div style={{ width: 9, height: 9, borderRadius: "50%", background: "#ea4335", marginTop: -6 }} />
          </div>
        )}
      </div>
    );
  };

  const TimeGutter = () => (
    <div style={{ width: 46, flexShrink: 0 }}>
      {Array.from({ length: H_END - H_START }, (_, i) => (
        <div key={i} style={{ height: HPX, fontSize: 10.5, color: "#8593a8", textAlign: "right", paddingRight: 7, transform: "translateY(-6px)" }}>
          {H_START + i}:00
        </div>
      ))}
    </div>
  );

  /* ---- En-tête ---- */
  let label;
  if (mode === "jour") label = `${DAYS_FULL[(ref.getDay() + 6) % 7]} ${ref.getDate()} ${MONTH_NAMES[ref.getMonth()].toLowerCase()} ${ref.getFullYear()}`;
  else if (mode === "semaine") {
    const s = startOfWeek(ref); const e = new Date(s); e.setDate(e.getDate() + 6);
    label = `${s.getDate()} ${MONTH_NAMES[s.getMonth()].toLowerCase().slice(0, 4)}. – ${e.getDate()} ${MONTH_NAMES[e.getMonth()].toLowerCase().slice(0, 4)}. ${e.getFullYear()}`;
  } else label = `${MONTH_NAMES[ref.getMonth()]} ${ref.getFullYear()}`;

  const legend = (
    <div className="row" style={{ marginTop: 12, flexWrap: "wrap", fontSize: 11.5 }}>
      {users.map((u) => {
        const st = RDV_OWNER_STYLE[u.id] || { bg: "#eef1f6", border: NAVY2 };
        return <span key={u.id} style={{ background: st.bg, borderLeft: `3px solid ${st.border}`, padding: "2px 9px", borderRadius: 4, fontWeight: 600 }}>{u.prenom}</span>;
      })}
      <span style={{ background: RDV_SIGNED_STYLE.bg, borderLeft: `3px solid ${RDV_SIGNED_STYLE.border}`, padding: "2px 9px", borderRadius: 4, fontWeight: 600 }}>🏆 Signé</span>
      <span style={{ color: "#8593a8" }}>· Cliquez sur un créneau vide pour créer un RDV</span>
    </div>
  );

  return (
    <div className="card" style={{ padding: 16 }}>
      <div className="row" style={{ justifyContent: "space-between", marginBottom: 14 }}>
        <h2 style={{ fontSize: 17, textTransform: "capitalize" }}>{label}</h2>
        <div className="row">
          <button className="btn ghost sm" onClick={() => nav(-1)}>‹</button>
          <button className="btn ghost sm" onClick={() => { const d = new Date(); d.setHours(0, 0, 0, 0); setRef(d); }}>Aujourd'hui</button>
          <button className="btn ghost sm" onClick={() => nav(1)}>›</button>
        </div>
      </div>

      {/* ============ VUE JOUR : grille horaire pleine largeur ============ */}
      {mode === "jour" && (
        <div style={{ display: "flex" }}>
          <TimeGutter />
          <DayColumn d={ref} wide />
        </div>
      )}

      {/* ============ VUE SEMAINE : grille horaire 7 jours ============ */}
      {mode === "semaine" && (
        <div>
          <div style={{ display: "flex", marginBottom: 6 }}>
            <div style={{ width: 46, flexShrink: 0 }} />
            {Array.from({ length: 7 }, (_, i) => {
              const d = startOfWeek(ref); d.setDate(d.getDate() + i);
              const isToday = isoOf(d) === today;
              return (
                <div key={i} style={{ flex: 1, textAlign: "center", minWidth: 0 }}>
                  <div style={{ fontSize: 11, color: "#8593a8", textTransform: "uppercase" }}>{DAYS[i]}</div>
                  <div style={{
                    fontSize: 17, fontWeight: 700, width: 34, height: 34, lineHeight: "34px", margin: "2px auto 0",
                    borderRadius: "50%", background: isToday ? NAVY : "transparent", color: isToday ? "#fff" : NAVY,
                  }}>{d.getDate()}</div>
                </div>
              );
            })}
          </div>
          <div style={{ display: "flex", border: "1px solid #eef1f6", borderRadius: 10, overflow: "hidden" }}>
            <TimeGutter />
            {Array.from({ length: 7 }, (_, i) => {
              const d = startOfWeek(ref); d.setDate(d.getDate() + i);
              return <DayColumn key={i} d={new Date(d)} />;
            })}
          </div>
        </div>
      )}

      {/* ============ VUE MOIS : grille type Google ============ */}
      {mode === "mois" && (() => {
        const first = new Date(ref.getFullYear(), ref.getMonth(), 1);
        const offset = (first.getDay() + 6) % 7;
        const start = new Date(first); start.setDate(start.getDate() - offset);
        const cells = Array.from({ length: 42 }, (_, i) => { const d = new Date(start); d.setDate(d.getDate() + i); return d; });
        return (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", marginBottom: 4 }}>
              {DAYS.map((d) => <div key={d} style={{ fontSize: 11, fontWeight: 700, color: "#8593a8", textAlign: "center", textTransform: "uppercase", letterSpacing: 1 }}>{d}</div>)}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", border: "1px solid #eef1f6", borderRadius: 10, overflow: "hidden" }}>
              {cells.map((d, i) => {
                const dIso = isoOf(d);
                const inMonth = d.getMonth() === ref.getMonth();
                const isToday = dIso === today;
                const list = byDate[dIso] || [];
                return (
                  <div key={i}
                    onClick={() => { setRef(new Date(d)); onDayOpen && onDayOpen(); }}
                    title="Cliquer pour ouvrir la journée"
                    style={{
                      minHeight: 112, padding: "5px 5px 3px", cursor: "pointer",
                      borderRight: (i % 7) < 6 ? "1px solid #f0f3f8" : "none",
                      borderBottom: i < 35 ? "1px solid #f0f3f8" : "none",
                      background: inMonth ? "#fff" : "#f8fafc",
                    }}
                    onMouseEnter={(e) => (e.currentTarget.style.background = "#fafcff")}
                    onMouseLeave={(e) => (e.currentTarget.style.background = inMonth ? "#fff" : "#f8fafc")}
                  >
                    <div style={{
                      fontSize: 12, fontWeight: 700, width: 24, height: 24, lineHeight: "24px", textAlign: "center",
                      borderRadius: "50%", margin: "0 auto 3px",
                      background: isToday ? NAVY : "transparent", color: isToday ? "#fff" : inMonth ? NAVY : "#b4becc",
                    }}>{d.getDate()}</div>
                    {list.slice(0, 3).map((p) => {
                      const st = styleOf(p);
                      return (
                        <div key={p.id}
                          onClick={(e) => { e.stopPropagation(); onOpen(p); }}
                          title={`${p.heureRdv || ""} ${(p.nom || "").toUpperCase()} · ${p.statut} · ${ownerName(p)}`}
                          style={{ background: st.bg, borderRadius: 5, fontSize: 10.5, fontWeight: 600, color: st.text, padding: "1px 5px", marginBottom: 2, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                          <span style={{ display: "inline-block", width: 6, height: 6, borderRadius: "50%", background: st.border, marginRight: 4, verticalAlign: "middle" }} />
                          {p.heureRdv && <b>{p.heureRdv}</b>} {(p.nom || "").toUpperCase()}
                        </div>
                      );
                    })}
                    {list.length > 3 && <div style={{ fontSize: 10, color: "#5b6b82", fontWeight: 600, paddingLeft: 3 }}>+ {list.length - 3} autres</div>}
                  </div>
                );
              })}
            </div>
          </div>
        );
      })()}

      {entries.length === 0 && (
        <div style={{ color: "#8593a8", fontSize: 13.5, marginTop: 10 }}>
          Aucun RDV planifié : renseignez la « date du RDV » sur vos fiches, ou cliquez sur un créneau pour en créer un.
        </div>
      )}
      {legend}
    </div>
  );
}

/* ================= CORBEILLE (manager) ================= */
function TrashPage({ trash, saveTrash, users, restoreClient, restoreProspect, restoreDoc, restoreFichier }) {
  const items = trash.slice().sort((a, b) => b.deletedAt.localeCompare(a.deletedAt));
  const who = (id) => {
    const u = users.find((x) => x.id === id);
    return u ? `${u.prenom} ${u.nom}` : "—";
  };
  const daysLeft = (deletedAt) => Math.max(0, 30 - Math.floor((Date.now() - new Date(deletedAt).getTime()) / 86400000));

  const purgeFiles = (item) => {
    if (item.kind === "client") (item.data.contrats || []).forEach((k) => (k.fichiers || []).forEach((f) => sDel(`crm-file-${f.id}`)));
    if (item.kind === "doc") (item.data.files || []).forEach((f) => sDel(`crm-file-${f.id}`));
    if (item.kind === "fichier") sDel(`crm-file-${item.data.file.id}`);
  };
  const purgeItem = (item) => {
    if (!confirm("Supprimer DÉFINITIVEMENT ? Cette action est irréversible.")) return;
    purgeFiles(item);
    saveTrash(trash.filter((x) => x.id !== item.id));
  };
  const restore = (item) => {
    if (item.kind === "client") restoreClient(item);
    else if (item.kind === "prospect") restoreProspect(item);
    else if (item.kind === "doc") restoreDoc(item);
    else if (item.kind === "fichier") restoreFichier(item);
  };
  const iconOf = { client: "👥", prospect: "🎯", doc: "📁", fichier: "📄" };
  const labelOf = (item) => {
    if (item.kind === "client") return `${(item.data.nom || "").toUpperCase()} ${item.data.prenom || ""} — fiche client · ${(item.data.contrats || []).length} contrat(s)`;
    if (item.kind === "prospect") return `${(item.data.nom || "").toUpperCase()} ${item.data.prenom || ""} — prospection · ${item.data.statut}`;
    if (item.kind === "doc") return `${item.data.name} — dossier · ${(item.data.files || []).length} fichier(s)`;
    return `${item.data.file.name} — fichier (dossier « ${item.data.folderName} »)`;
  };

  return (
    <div>
      <div className="ph">
        <div>
          <h1>🗑️ Corbeille</h1>
          <div className="sub">{items.length} élément(s) — suppression automatique et définitive après 30 jours</div>
        </div>
        {items.length > 0 && (
          <button className="btn danger" onClick={() => { if (confirm("Vider toute la corbeille définitivement ? Cette action est irréversible.")) { items.forEach(purgeFiles); saveTrash([]); } }}>
            🗑️ Vider la corbeille
          </button>
        )}
      </div>

      {items.length === 0 && <div className="card" style={{ color: "#8593a8" }}>La corbeille est vide. Les fiches clients, fiches de prospection, dossiers et fichiers supprimés arriveront ici.</div>}

      <div className="grid">
        {items.map((item) => (
          <div className="clientcard" key={item.id} style={{ cursor: "default" }}>
            <div>
              <b style={{ fontSize: 14.5 }}>{iconOf[item.kind] || "🗑️"} {labelOf(item)}</b>
              <div style={{ fontSize: 12, color: "#5b6b82" }}>
                Supprimé le {fmtDate(item.deletedAt.slice(0, 10))} par {who(item.deletedBy)}
              </div>
              <div style={{ fontSize: 11.5, color: daysLeft(item.deletedAt) <= 7 ? "#B3261E" : "#8593a8" }}>
                ⏳ {daysLeft(item.deletedAt)} jour(s) avant suppression définitive
              </div>
            </div>
            <div className="row">
              <button className="btn gold sm" onClick={() => restore(item)}>↩️ Restaurer</button>
              <button className="btn danger sm" onClick={() => purgeItem(item)}>✕</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ================= SAUVEGARDE / RESTAURATION (copier-coller) ================= */
function BackupModal({ modal, close, doImport }) {
  const [text, setText] = useState(modal.text || "");
  const [copied, setCopied] = useState(false);
  const taRef = useRef(null);

  const copyAll = async () => {
    try {
      await navigator.clipboard.writeText(modal.text);
      setCopied(true);
    } catch {
      /* solution de repli si le presse-papiers est bloqué */
      if (taRef.current) {
        taRef.current.select();
        try { document.execCommand("copy"); setCopied(true); } catch {}
      }
    }
  };

  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && close()}>
      <div className="modal" style={{ maxWidth: 640 }}>
        {modal.mode === "export" ? (
          <>
            <h2>💾 Votre sauvegarde</h2>
            <p style={{ fontSize: 13, color: "#5b6b82", margin: "8px 0 12px" }}>
              Si aucun fichier ne s'est téléchargé, utilisez le bouton ci-dessous : copiez le texte,
              puis collez-le dans « 📥 Importer » de votre autre CRM, ou gardez-le dans un document en lieu sûr.
            </p>
            <button className="btn gold" style={{ width: "100%", marginBottom: 10 }} onClick={copyAll}>
              {copied ? "✓ Copié dans le presse-papiers !" : "📋 Tout copier"}
            </button>
            <textarea
              ref={taRef} className="ta" rows={8} readOnly value={modal.text}
              onClick={(e) => e.target.select()}
              style={{ fontSize: 11, fontFamily: "monospace", wordBreak: "break-all" }}
            />
          </>
        ) : (
          <>
            <h2>📥 Restaurer une sauvegarde</h2>
            <p style={{ fontSize: 13, color: "#5b6b82", margin: "8px 0 12px" }}>
              Collez ici le texte de sauvegarde (obtenu via « 💾 Exporter » → 📋 Tout copier),
              puis cliquez sur Importer. Les données actuelles seront remplacées.
            </p>
            <textarea
              className="ta" rows={8} value={text} autoFocus
              onChange={(e) => setText(e.target.value)}
              placeholder='Collez la sauvegarde ici (commence par {"version":...)'
              style={{ fontSize: 11, fontFamily: "monospace", wordBreak: "break-all" }}
            />
            <button className="btn gold" style={{ width: "100%", marginTop: 10 }} onClick={() => doImport(text)} disabled={!text.trim()}>
              Importer cette sauvegarde
            </button>
          </>
        )}
        <div className="row" style={{ marginTop: 12, justifyContent: "flex-end" }}>
          <button className="btn ghost" onClick={close}>Fermer</button>
        </div>
      </div>
    </div>
  );
}

/* ================= ALERTE RDV (10 minutes avant) ================= */
function RdvReminder({ prospection, rdvClients, clients, me }) {
  const [now, setNow] = useState(() => new Date());
  const [dismissed, setDismissed] = useState([]);

  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 30000); // vérification toutes les 30 s
    return () => clearInterval(t);
  }, []);

  const todayIso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
  const isSoon = (dateISO, heure) => {
    if (dateISO !== todayIso || !heure) return false;
    const [h, m] = heure.split(":").map(Number);
    const rdv = new Date(now); rdv.setHours(h, m, 0, 0);
    const diffMin = (rdv - now) / 60000;
    return diffMin <= 10 && diffMin > -15; // de J-10 min jusqu'à 15 min après le début
  };

  const soon = [
    ...prospection
      .filter((p) => p.ownerId === me.id && !dismissed.includes(p.id) && isSoon(p.dateRdv, p.heureRdv))
      .map((p) => ({ id: p.id, heure: p.heureRdv, titre: `${(p.nom || "").toUpperCase()} ${p.prenom || ""}`, sous: `Prospection${p.profession ? " · " + p.profession : ""}${p.telephone ? " · 📞 " + p.telephone : ""}` })),
    ...(rdvClients || [])
      .filter((r) => r.ownerId === me.id && !r.done && !dismissed.includes(r.id) && isSoon(r.date, r.heure))
      .map((r) => {
        const c = (clients || []).find((x) => x.id === r.clientId) || {};
        return { id: r.id, heure: r.heure, titre: `${(c.nom || "").toUpperCase()} ${c.prenom || ""}`, sous: `Client · ${r.motif}${c.telephone ? " · 📞 " + c.telephone : ""}` };
      }),
  ];

  if (soon.length === 0) return null;
  return (
    <div style={{ position: "sticky", top: 0, zIndex: 90, marginBottom: 16 }}>
      {soon.map((p) => (
        <div key={p.id} style={{
          background: "#fdf3d7", border: `2px solid ${GOLD}`, borderRadius: 12,
          padding: "12px 16px", marginBottom: 8, display: "flex", justifyContent: "space-between",
          alignItems: "center", boxShadow: "0 6px 20px rgba(201,162,75,.25)",
        }}>
          <div>
            <b style={{ color: "#7a5c17", fontSize: 15 }}>⏰ RDV imminent — {p.heure}</b>
            <div style={{ fontSize: 13.5, color: NAVY, marginTop: 2 }}>{p.titre} · {p.sous}</div>
          </div>
          <button className="btn ghost sm" onClick={() => setDismissed([...dismissed, p.id])}>✓ Vu</button>
        </div>
      ))}
    </div>
  );
}

/* ================= RDV CLIENT (motifs patrimoniaux) ================= */
function RdvClientForm({ client, me, onClose, onSave }) {
  const [f, setF] = useState({ motif: MOTIFS_RDV[0], date: todayISO(), heure: "10:00", note: "" });
  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 520 }}>
        <h2>📅 Planifier un RDV — {client.nom.toUpperCase()} {client.prenom}</h2>
        <div style={{ marginTop: 12 }}>
          <Field label="Motif du rendez-vous">
            <select className="sel" style={{ width: "100%" }} value={f.motif} onChange={(e) => setF({ ...f, motif: e.target.value })}>
              {MOTIFS_RDV.map((m) => <option key={m}>{m}</option>)}
            </select>
          </Field>
        </div>
        <div className="row" style={{ marginTop: 10 }}>
          <Field label="Date"><input className="in" type="date" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} /></Field>
          <Field label="Heure"><input className="in" type="time" value={f.heure} onChange={(e) => setF({ ...f, heure: e.target.value })} /></Field>
        </div>
        <div style={{ marginTop: 10 }}>
          <Field label="Note (facultatif)">
            <textarea className="ta" rows={2} value={f.note} onChange={(e) => setF({ ...f, note: e.target.value })} placeholder="ex : apporter le dernier relevé d'assurance vie…" />
          </Field>
        </div>
        <div className="row" style={{ marginTop: 16, justifyContent: "flex-end" }}>
          <button className="btn ghost" onClick={onClose}>Annuler</button>
          <button className="btn gold" onClick={() => {
            onSave({ id: uid(), clientId: client.id, ownerId: me.id, ...f, done: false, createdAt: todayISO() });
            if (confirm("RDV planifié ✓ (rappel 10 min avant dans le CRM)\n\nL'ajouter aussi à votre Google Agenda ?")) {
              window.open(gcalUrl({
                titre: `RDV client — ${client.nom.toUpperCase()} ${client.prenom}`,
                date: f.date, heure: f.heure,
                details: `${f.motif}${f.note ? "\n" + f.note : ""}\n(CRM ELYON & Associés)`,
              }), "_blank");
            }
          }}>
            Planifier le rendez-vous
          </button>
        </div>
      </div>
    </div>
  );
}



/* ================= MESSAGERIE ================= */
function MessageriePage({ clients, users, me, mailTpl, saveMailTpl }) {
  /* Cloisonnement identique au reste */
  const mine = me.isManager ? clients : clients.filter((c) => (c.createdBy || "quentin") === me.id);
  const withEmail = mine.filter((c) => (c.email || "").includes("@"));

  const [selected, setSelected] = useState([]);
  const [profFilter, setProfFilter] = useState("all");
  const [tplId, setTplId] = useState(mailTpl[0] ? mailTpl[0].id : "");
  const [sujet, setSujet] = useState(mailTpl[0] ? mailTpl[0].sujet : "");
  const [corps, setCorps] = useState(mailTpl[0] ? mailTpl[0].corps : "");
  const [editTpl, setEditTpl] = useState(false);

  const professions = [...new Set(mine.map((c) => (c.profession || "").trim()).filter(Boolean))].sort();
  const visible = profFilter === "all" ? withEmail : withEmail.filter((c) => (c.profession || "").trim() === profFilter);

  const toggle = (id) => setSelected(selected.includes(id) ? selected.filter((x) => x !== id) : [...selected, id]);
  const allVisibleSelected = visible.length > 0 && visible.every((c) => selected.includes(c.id));

  const pickTpl = (id) => {
    const t = mailTpl.find((x) => x.id === id);
    setTplId(id);
    if (t) { setSujet(t.sujet); setCorps(t.corps); }
  };

  const emails = mine.filter((c) => selected.includes(c.id)).map((c) => c.email.trim());
  const openMail = () => {
    if (!emails.length) { alert("Sélectionnez au moins un client."); return; }
    /* Destinataires en Cci : chaque client ne voit pas les adresses des autres (RGPD) */
    const url = `mailto:?bcc=${encodeURIComponent(emails.join(","))}&subject=${encodeURIComponent(sujet)}&body=${encodeURIComponent(corps)}`;
    if (url.length > 7500) {
      alert("Trop de destinataires pour un seul e-mail (limite technique des boîtes mail). Réduisez la sélection (max ~50 clients) ou copiez les adresses avec l'autre bouton.");
      return;
    }
    window.location.href = url;
  };
  const copyEmails = async () => {
    if (!emails.length) { alert("Sélectionnez au moins un client."); return; }
    try { await navigator.clipboard.writeText(emails.join(", ")); alert(`${emails.length} adresse(s) copiée(s) ✓ Collez-les en Cci dans votre boîte mail.`); }
    catch { prompt("Copiez les adresses ci-dessous :", emails.join(", ")); }
  };

  const saveCurrentAsTpl = () => {
    const nom = prompt("Nom du modèle :", "Nouveau modèle");
    if (!nom) return;
    saveMailTpl([...mailTpl, { id: uid(), nom, sujet, corps }]);
  };

  return (
    <div>
      <div className="ph">
        <div>
          <h1>✉️ Messagerie</h1>
          <div className="sub">Campagnes d'e-mails clients — {withEmail.length} client(s) avec adresse e-mail{me.isManager ? " (tous portefeuilles)" : " (votre portefeuille)"}</div>
        </div>
        <button className="btn ghost" onClick={() => setEditTpl(!editTpl)}>{editTpl ? "Fermer les modèles" : "⚙️ Gérer les modèles"}</button>
      </div>

      {editTpl && (
        <div className="card" style={{ marginBottom: 16 }}>
          <h2 style={{ fontSize: 15, marginBottom: 10 }}>Modèles enregistrés</h2>
          {mailTpl.map((t) => (
            <div key={t.id} className="row" style={{ justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid #eef1f6" }}>
              <div style={{ fontSize: 13.5 }}><b>{t.nom}</b> <span style={{ color: "#8593a8" }}>— {t.sujet}</span></div>
              <div className="row">
                <button className="btn ghost sm" onClick={() => pickTpl(t.id)}>Utiliser</button>
                <button className="btn danger sm" onClick={() => confirm(`Supprimer le modèle « ${t.nom} » ?`) && saveMailTpl(mailTpl.filter((x) => x.id !== t.id))}>✕</button>
              </div>
            </div>
          ))}
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "5fr 7fr", gap: 16 }}>
        {/* ---- Destinataires ---- */}
        <div className="card">
          <h2 style={{ fontSize: 15, marginBottom: 10 }}>1️⃣ Destinataires ({selected.length} sélectionné(s))</h2>
          <div className="row" style={{ marginBottom: 10 }}>
            <select className="sel" value={profFilter} onChange={(e) => setProfFilter(e.target.value)}>
              <option value="all">Toutes les professions</option>
              {professions.map((p) => <option key={p}>{p}</option>)}
            </select>
            <button className="btn ghost sm" onClick={() => setSelected(allVisibleSelected ? selected.filter((id) => !visible.some((c) => c.id === id)) : [...new Set([...selected, ...visible.map((c) => c.id)])])}>
              {allVisibleSelected ? "Tout désélectionner" : "Tout sélectionner"}
            </button>
          </div>
          <div style={{ maxHeight: 420, overflowY: "auto" }}>
            {visible.map((c) => (
              <label key={c.id} style={{ display: "flex", gap: 8, alignItems: "center", padding: "6px 4px", borderBottom: "1px solid #f0f2f6", cursor: "pointer", fontSize: 13.5 }}>
                <input type="checkbox" checked={selected.includes(c.id)} onChange={() => toggle(c.id)} />
                <span><b>{c.nom.toUpperCase()}</b> {c.prenom} <span style={{ color: "#8593a8", fontSize: 12 }}>· {c.profession || "—"} · {c.email}</span></span>
              </label>
            ))}
            {visible.length === 0 && <div style={{ color: "#8593a8", fontSize: 13 }}>Aucun client avec e-mail pour ce filtre.</div>}
          </div>
        </div>

        {/* ---- Message ---- */}
        <div className="card">
          <h2 style={{ fontSize: 15, marginBottom: 10 }}>2️⃣ Message</h2>
          <Field label="Modèle">
            <select className="sel" style={{ width: "100%" }} value={tplId} onChange={(e) => pickTpl(e.target.value)}>
              {mailTpl.map((t) => <option key={t.id} value={t.id}>{t.nom}</option>)}
            </select>
          </Field>
          <div style={{ marginTop: 10 }}>
            <Field label="Objet"><input className="in" style={{ width: "100%" }} value={sujet} onChange={(e) => setSujet(e.target.value)} /></Field>
          </div>
          <div style={{ marginTop: 10 }}>
            <Field label="Message">
              <textarea className="ta" rows={11} value={corps} onChange={(e) => setCorps(e.target.value)} />
            </Field>
          </div>
          <div className="row" style={{ marginTop: 14, justifyContent: "space-between", flexWrap: "wrap" }}>
            <button className="btn ghost sm" onClick={saveCurrentAsTpl}>💾 Enregistrer comme modèle</button>
            <div className="row">
              <button className="btn ghost" onClick={copyEmails}>📋 Copier les adresses</button>
              <button className="btn gold" onClick={openMail}>✉️ Ouvrir dans ma boîte mail ({selected.length})</button>
            </div>
          </div>
          <div style={{ fontSize: 11.5, color: "#8593a8", marginTop: 10 }}>
            ℹ️ Le bouton ouvre votre boîte mail habituelle avec les destinataires en <b>Cci</b> (les clients ne voient pas les adresses des autres). Au-delà de ~50 destinataires, utilisez « Copier les adresses ».
          </div>
        </div>
      </div>
    </div>
  );
}

/* ================= IMPORT DE LISTES DE PROSPECTION (Excel / CSV) ================= */
function ImportListeModal({ users, onClose, onImport }) {
  const [rows, setRows] = useState(null);      // lignes brutes du fichier
  const [headers, setHeaders] = useState([]);  // en-têtes détectés
  const [mapping, setMapping] = useState({});  // colonne fichier → champ CRM
  const [owner, setOwner] = useState(users[0] ? users[0].id : "");
  const [fileName, setFileName] = useState("");

  const CHAMPS = [
    ["nom", "Nom *"], ["prenom", "Prénom"], ["profession", "Profession"],
    ["telephone", "Téléphone"], ["ville", "Ville"], ["email", "E-mail"], ["", "— Ignorer —"],
  ];

  /* Détection automatique des colonnes selon leur en-tête */
  const autoMap = (hdrs) => {
    const m = {};
    hdrs.forEach((h, i) => {
      const l = String(h || "").toLowerCase();
      if (/nom.*famille|^nom$|raison/.test(l) && !Object.values(m).includes("nom")) m[i] = "nom";
      else if (/pr[eé]nom/.test(l)) m[i] = "prenom";
      else if (/profession|m[eé]tier|activit|sp[eé]cialit/.test(l)) m[i] = "profession";
      else if (/t[eé]l|phone|portable|mobile/.test(l)) m[i] = "telephone";
      else if (/ville|commune|localit/.test(l)) m[i] = "ville";
      else if (/mail|courriel/.test(l)) m[i] = "email";
      else if (/^nom/.test(l) && !Object.values(m).includes("nom")) m[i] = "nom";
      else m[i] = "";
    });
    return m;
  };

  const readFile = (file) => {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        const nonEmpty = data.filter((r) => r.some((c) => String(c).trim()));
        if (nonEmpty.length < 2) { alert("Fichier vide ou illisible."); return; }
        const hdrs = nonEmpty[0].map((h) => String(h));
        setHeaders(hdrs);
        setRows(nonEmpty.slice(1));
        setMapping(autoMap(hdrs));
      } catch { alert("Impossible de lire ce fichier. Formats acceptés : .xlsx, .xls, .csv"); }
    };
    reader.readAsArrayBuffer(file);
  };

  const doImport = () => {
    const batchId = uid();
    const nomCol = Object.keys(mapping).find((k) => mapping[k] === "nom");
    if (nomCol === undefined) { alert("Indiquez quelle colonne contient le NOM (obligatoire)."); return; }
    const entries = rows
      .map((r) => {
        const e = { nom: "", prenom: "", profession: "", telephone: "", ville: "", email: "" };
        Object.keys(mapping).forEach((i) => { if (mapping[i]) e[mapping[i]] = String(r[i] ?? "").trim(); });
        return e;
      })
      .filter((e) => e.nom)
      .map((e) => ({
        id: uid(), ownerId: owner,
        nom: e.nom, prenom: e.prenom, profession: (e.profession || "").toUpperCase(), telephone: (e.telephone || "").replace(/\s/g, ""), ville: e.ville, email: e.email, dateNaissance: "",
        per: "", prevoyance: "",
        dateAppel: todayISO(), repondu: "Non", statut: "À rappeler",
        dateRdv: "", heureRdv: "", qualitePrise: "", noteRdv: "",
        commentaire: "", createdAt: todayISO(), importe: true,
        batchId: batchId, batchNom: fileName, batchDate: todayISO(),
      }));
    if (!entries.length) { alert("Aucune ligne exploitable (colonne NOM vide ?)."); return; }
    if (!confirm(`Importer ${entries.length} prospect(s) et les attribuer à ${((users.find((u) => u.id === owner)) || {}).prenom} ?`)) return;
    onImport(entries);
  };

  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 760, maxHeight: "88vh", overflowY: "auto" }}>
        <h2>📥 Importer une liste de prospection</h2>

        {!rows && (
          <div style={{ marginTop: 14 }}>
            <p style={{ fontSize: 13.5, color: "#5b6b82" }}>
              Choisissez un fichier <b>Excel (.xlsx, .xls)</b> ou <b>CSV</b> contenant vos prospects
              (une ligne par prospect, avec au minimum une colonne Nom).
            </p>
            <input
              type="file" accept=".xlsx,.xls,.csv"
              onChange={(e) => e.target.files[0] && readFile(e.target.files[0])}
              style={{ marginTop: 10, fontSize: 14 }}
            />
          </div>
        )}

        {rows && (
          <>
            <div style={{ fontSize: 13, color: "#5b6b82", margin: "10px 0" }}>
              📄 <b>{fileName}</b> — {rows.length} ligne(s) détectée(s). Vérifiez la correspondance des colonnes :
            </div>
            <div style={{ overflowX: "auto" }}>
              <table className="t" style={{ fontSize: 12.5 }}>
                <thead>
                  <tr>
                    {headers.map((h, i) => (
                      <th key={i}>
                        <div style={{ fontWeight: 400, color: "#8593a8", marginBottom: 4 }}>{h || `Colonne ${i + 1}`}</div>
                        <select className="sel" style={{ fontSize: 12, width: "100%" }} value={mapping[i] || ""} onChange={(e) => setMapping({ ...mapping, [i]: e.target.value })}>
                          {CHAMPS.map(([v, l]) => <option key={v + l} value={v}>{l}</option>)}
                        </select>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.slice(0, 4).map((r, ri) => (
                    <tr key={ri}>{headers.map((_, i) => <td key={i}>{String(r[i] ?? "")}</td>)}</tr>
                  ))}
                  {rows.length > 4 && <tr><td colSpan={headers.length} style={{ color: "#8593a8" }}>… et {rows.length - 4} autre(s) ligne(s)</td></tr>}
                </tbody>
              </table>
            </div>
            <div className="row" style={{ marginTop: 14, alignItems: "flex-end" }}>
              <Field label="Attribuer cette liste à">
                <select className="sel" value={owner} onChange={(e) => setOwner(e.target.value)}>
                  {users.map((u) => <option key={u.id} value={u.id}>{u.prenom} {u.nom}</option>)}
                </select>
              </Field>
              <div style={{ fontSize: 12, color: "#8593a8", paddingBottom: 8 }}>
                Les prospects arriveront avec le statut « À rappeler ».
              </div>
            </div>
          </>
        )}

        <div className="row" style={{ marginTop: 16, justifyContent: "space-between" }}>
          <button className="btn ghost" onClick={rows ? () => { setRows(null); setHeaders([]); setFileName(""); } : onClose}>{rows ? "← Autre fichier" : "Annuler"}</button>
          {rows && <button className="btn gold" onClick={doImport}>Importer {rows.length} prospect(s)</button>}
        </div>
      </div>
    </div>
  );
}

/* ================= DÉCOMMISSIONS ================= */
function DecomPage({ sales, users }) {
  const rows = [];
  Object.keys(sales).sort().reverse().forEach((m) => {
    users.forEach((u) => {
      (((sales[m] || {})[u.id] || {}).rows || []).forEach((r) => {
        if ((r.nom || "").trim() && r.statut === "Décommissionné") rows.push({ ...r, month: m, commercial: `${u.prenom} ${u.nom}` });
      });
    });
  });
  const totRem = rows.reduce((s, r) => s + parseNum(r.remuneration), 0);
  const totVol = rows.reduce((s, r) => s + parseNum(r.volume), 0);

  return (
    <div>
      <div className="ph">
        <div>
          <h1>📉 Décommissions</h1>
          <div className="sub">
            Dossiers décommissionnés — pour retirer un dossier des commissions, passez son statut sur
            « Décommissionné » dans le tableau des ventes : il s'affiche en rouge et arrive automatiquement ici.
          </div>
        </div>
        <div className="row">
          <div className="kpi"><div className="n" style={{ color: "#B3261E" }}>{rows.length}</div><div className="l">Dossiers</div></div>
          <div className="kpi"><div className="n" style={{ color: "#B3261E" }}>{fmtEUR(totRem)}</div><div className="l">Rémunération décommissionnée</div></div>
        </div>
      </div>

      <div className="card">
        <table className="t">
          <thead>
            <tr><th>Mois</th><th>Commercial</th><th>Date création</th><th>Client</th><th>Type</th><th>Compagnie</th><th>Réf.</th><th>Volume</th><th>Rémunération</th><th>Commentaire</th></tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id} className="decom">
                <td style={{ fontSize: 12 }}>{monthLabel(r.month)}</td>
                <td style={{ fontSize: 12 }}>{r.commercial}</td>
                <td style={{ fontSize: 12 }}>{fmtDate(r.dateCreation)}</td>
                <td><b>{r.nom}</b></td>
                <td style={{ fontSize: 12 }}>{r.type}</td>
                <td style={{ fontSize: 12 }}>{r.compagnie}</td>
                <td style={{ fontSize: 12 }}>{r.ref}</td>
                <td style={{ fontSize: 12 }}>{r.volume}</td>
                <td style={{ fontSize: 12 }}><b>{r.remuneration}</b></td>
                <td style={{ fontSize: 11.5 }}>{r.commentaire}</td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr><td colSpan={10} style={{ color: "#8593a8", padding: 18 }}>Aucune décommission ✓ — si un dossier est décommissionné, changez son statut dans le tableau des ventes.</td></tr>
            )}
          </tbody>
        </table>
        {totVol > 0 && <div style={{ fontSize: 12.5, color: "#5b6b82", marginTop: 10 }}>Volume total concerné : <b>{fmtEUR(totVol)}</b></div>}
      </div>
    </div>
  );
}

/* ================= ACCÈS REFUSÉ (téléprospecteurs) ================= */
function AccessDenied({ goBack }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: "60vh" }}>
      <div className="card" style={{ maxWidth: 460, textAlign: "center", padding: "40px 32px", border: `2px solid ${GOLD}` }}>
        <div style={{ fontSize: 52 }}>🔒</div>
        <h2 style={{ fontSize: 20, color: NAVY, margin: "14px 0 8px" }}>Vous n'avez pas accès à cet espace</h2>
        <p style={{ fontSize: 13.5, color: "#5b6b82", lineHeight: 1.6 }}>
          Votre profil <b>Téléprospecteur</b> vous donne accès à la prospection, à l'agenda
          et au tableau de bord — mais pas aux ventes ni aux fiches clients.
          <br />Contactez votre manager si vous pensez qu'il s'agit d'une erreur.
        </p>
        <button className="btn gold" style={{ marginTop: 16 }} onClick={goBack}>🎯 Retour à la prospection</button>
      </div>
    </div>
  );
}

/* ================= VUE ANNUELLE CONSOLIDÉE ================= */
function AnnualView({ sales, users, me }) {
  const monthsAll = Object.keys(sales).sort();
  const years = [...new Set(monthsAll.map((m) => m.slice(0, 4)))].sort().reverse();
  const [year, setYear] = useState(years[0]);

  /* Lignes réelles (hors annulé, hors recopies automatiques pour ne rien compter deux fois) */
  const rowsOf = (mk, uid2) => ((((sales[mk] || {})[uid2] || {}).rows) || [])
    .filter((r) => (r.nom || "").trim() && r.statut !== "Annulé" && !r.mirrorOf);

  const monthsY = monthsAll.filter((m) => m.startsWith(year));
  const monthsPrev = monthsAll.filter((m) => m.startsWith(String(Number(year) - 1)));

  /* Graphique : volume par mois, une barre par commercial */
  const chartData = monthsY.map((mk) => {
    const d = { mois: MONTH_NAMES[Number(mk.slice(5, 7)) - 1].slice(0, 4) };
    users.forEach((u) => { d[u.prenom] = rowsOf(mk, u.id).reduce((s, r) => s + parseNum(r.volume), 0); });
    return d;
  });
  const USER_COLORS = [NAVY, GOLD, "#1b6e5a", "#8e44ad", "#d35400", "#1d6fb8"];

  /* Cumuls par commercial */
  const cumuls = users.map((u) => {
    let contrats = 0, volume = 0, rem = 0;
    monthsY.forEach((mk) => rowsOf(mk, u.id).forEach((r) => { contrats++; volume += parseNum(r.volume); rem += parseNum(r.remuneration); }));
    return { u, contrats, volume, rem };
  }).sort((a, b) => b.volume - a.volume);

  /* Répartitions produit / compagnie */
  const parType = {}, parCompagnie = {};
  monthsY.forEach((mk) => users.forEach((u) => rowsOf(mk, u.id).forEach((r) => {
    if (r.type) { (parType[r.type] = parType[r.type] || { n: 0, v: 0 }); parType[r.type].n++; parType[r.type].v += parseNum(r.volume); }
    if (r.compagnie) { (parCompagnie[r.compagnie] = parCompagnie[r.compagnie] || { n: 0, v: 0 }); parCompagnie[r.compagnie].n++; parCompagnie[r.compagnie].v += parseNum(r.volume); }
  })));
  const triV = (o) => Object.entries(o).sort((a, b) => b[1].v - a[1].v);

  /* Totaux N vs N-1 */
  const totOf = (mks) => {
    let c = 0, v = 0;
    mks.forEach((mk) => users.forEach((u) => rowsOf(mk, u.id).forEach((r) => { c++; v += parseNum(r.volume); })));
    return { c, v };
  };
  const totN = totOf(monthsY), totP = totOf(monthsPrev);
  const evol = (a, b) => (b ? Math.round(((a - b) / b) * 100) : null);
  const evolC = evol(totN.c, totP.c), evolV = evol(totN.v, totP.v);
  const Evo = ({ e }) => e === null ? null : (
    <span style={{ fontSize: 12, fontWeight: 700, color: e >= 0 ? "#1b7a3d" : "#B3261E", marginLeft: 6 }}>
      {e >= 0 ? "▲" : "▼"} {Math.abs(e)} % vs {Number(year) - 1}
    </span>
  );

  const maxTypeV = Math.max(1, ...Object.values(parType).map((x) => x.v));
  const maxCieV = Math.max(1, ...Object.values(parCompagnie).map((x) => x.v));
  const BarLine = ({ label, n, v, max, color }) => (
    <div style={{ marginBottom: 8 }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
        <b>{label}</b><span style={{ color: "#5b6b82" }}>{n} contrat(s) · {fmtEUR(v)}</span>
      </div>
      <div style={{ height: 8, background: "#eef1f6", borderRadius: 4, marginTop: 3, overflow: "hidden" }}>
        <div style={{ height: "100%", width: `${Math.round((v / max) * 100)}%`, background: color, borderRadius: 4 }} />
      </div>
    </div>
  );

  return (
    <div>
      <div className="row" style={{ marginBottom: 16, justifyContent: "space-between" }}>
        <div className="row">
          <select className="sel" style={{ width: 120 }} value={year} onChange={(e) => setYear(e.target.value)}>
            {years.map((y) => <option key={y}>{y}</option>)}
          </select>
        </div>
        <div className="row">
          <div className="kpi"><div className="n">{totN.c}<Evo e={evolC} /></div><div className="l">Contrats {year}</div></div>
          <div className="kpi"><div className="n">{fmtEUR(totN.v)}<Evo e={evolV} /></div><div className="l">Volume {year}</div></div>
        </div>
      </div>

      <div className="card" style={{ marginBottom: 16 }}>
        <h2 style={{ fontSize: 16, marginBottom: 12 }}>📊 Volume par mois et par commercial — {year}</h2>
        <div style={{ width: "100%", height: 300 }}>
          <ResponsiveContainer>
            <BarChart data={chartData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e3e8f0" />
              <XAxis dataKey="mois" tick={{ fontSize: 11, fill: "#5b6b82" }} />
              <YAxis tick={{ fontSize: 11, fill: "#5b6b82" }} tickFormatter={(v) => v.toLocaleString("fr-FR")} />
              <Tooltip formatter={(v, name) => [fmtEUR(v), name]} />
              {users.map((u, i) => <Bar key={u.id} dataKey={u.prenom} stackId="a" fill={USER_COLORS[i % USER_COLORS.length]} radius={i === users.length - 1 ? [5, 5, 0, 0] : [0, 0, 0, 0]} />)}
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        <div className="card">
          <h2 style={{ fontSize: 15, marginBottom: 12 }}>🧑‍💼 Cumuls {year} par commercial</h2>
          <table className="t">
            <thead><tr><th>Commercial</th><th>Contrats</th><th>Volume</th><th>Rémunération</th></tr></thead>
            <tbody>
              {cumuls.map(({ u, contrats, volume, rem }) => (
                <tr key={u.id}>
                  <td><b>{u.prenom} {u.nom}</b></td>
                  <td>{contrats}</td>
                  <td>{fmtEUR(volume)}</td>
                  <td>{fmtEUR(rem)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="card">
          <h2 style={{ fontSize: 15, marginBottom: 12 }}>📦 Répartition par produit</h2>
          {triV(parType).map(([t, x]) => <BarLine key={t} label={t} n={x.n} v={x.v} max={maxTypeV} color={GOLD} />)}
          {!Object.keys(parType).length && <div style={{ color: "#8593a8", fontSize: 13 }}>Aucune donnée sur {year}.</div>}
        </div>
      </div>

      <div className="card">
        <h2 style={{ fontSize: 15, marginBottom: 12 }}>🏦 Répartition par compagnie <span style={{ fontSize: 12, color: "#8593a8", fontWeight: 400 }}>— utile pour vos négociations de barèmes</span></h2>
        {triV(parCompagnie).map(([c, x]) => <BarLine key={c} label={c} n={x.n} v={x.v} max={maxCieV} color={NAVY2} />)}
        {!Object.keys(parCompagnie).length && <div style={{ color: "#8593a8", fontSize: 13 }}>Aucune donnée sur {year}.</div>}
      </div>
    </div>
  );
}

/* ================= GÉNÉRATEUR DE COURRIERS PRÉ-REMPLIS ================= */
const COURRIER_TEMPLATES = [
  {
    id: "transfert", nom: "Demande de transfert de contrat",
    objet: "Demande de transfert — contrat n° {NUMERO}",
    corps: "Madame, Monsieur,\n\nJe soussigné(e) {PRENOM} {NOM}, titulaire du contrat {TYPE} n° {NUMERO} souscrit auprès de {COMPAGNIE}, vous demande par la présente de procéder au transfert de l'intégralité de mon contrat.\n\nJe vous remercie de bien vouloir me transmettre le relevé de situation ainsi que les documents nécessaires à cette opération dans les meilleurs délais.\n\nDans l'attente de votre retour, je vous prie d'agréer, Madame, Monsieur, l'expression de mes salutations distinguées.",
  },
  {
    id: "rib", nom: "Changement de coordonnées bancaires (RIB)",
    objet: "Changement de domiciliation bancaire — contrat n° {NUMERO}",
    corps: "Madame, Monsieur,\n\nJe soussigné(e) {PRENOM} {NOM}, titulaire du contrat {TYPE} n° {NUMERO}, vous informe du changement de mes coordonnées bancaires.\n\nVous trouverez ci-joint mon nouveau relevé d'identité bancaire. Je vous remercie de bien vouloir prendre en compte cette modification pour l'ensemble des prélèvements et versements à venir.\n\nJe vous prie d'agréer, Madame, Monsieur, l'expression de mes salutations distinguées.",
  },
  {
    id: "vp", nom: "Modification du versement programmé",
    objet: "Modification du versement programmé — contrat n° {NUMERO}",
    corps: "Madame, Monsieur,\n\nJe soussigné(e) {PRENOM} {NOM}, titulaire du contrat {TYPE} n° {NUMERO}, souhaite modifier le montant de mon versement programmé.\n\nJe vous demande de bien vouloir porter ce versement à ________ € par mois, à compter du prochain prélèvement.\n\nJe vous prie d'agréer, Madame, Monsieur, l'expression de mes salutations distinguées.",
  },
  {
    id: "rachat", nom: "Demande de rachat partiel",
    objet: "Demande de rachat partiel — contrat n° {NUMERO}",
    corps: "Madame, Monsieur,\n\nJe soussigné(e) {PRENOM} {NOM}, titulaire du contrat {TYPE} n° {NUMERO}, vous demande de procéder à un rachat partiel d'un montant de ________ € sur mon contrat.\n\nJe vous remercie de bien vouloir créditer cette somme sur le compte bancaire dont le RIB est joint à la présente.\n\nJe vous prie d'agréer, Madame, Monsieur, l'expression de mes salutations distinguées.",
  },
  {
    id: "beneficiaire", nom: "Modification de la clause bénéficiaire",
    objet: "Modification de la clause bénéficiaire — contrat n° {NUMERO}",
    corps: "Madame, Monsieur,\n\nJe soussigné(e) {PRENOM} {NOM}, titulaire du contrat {TYPE} n° {NUMERO}, souhaite modifier la clause bénéficiaire de mon contrat comme suit :\n\n________________________________________\n\nJe vous remercie de bien vouloir m'adresser un avenant confirmant cette modification.\n\nJe vous prie d'agréer, Madame, Monsieur, l'expression de mes salutations distinguées.",
  },
];

function CourrierModal({ client, onClose }) {
  const contrats = client.contrats || [];
  const [tplId, setTplId] = useState(COURRIER_TEMPLATES[0].id);
  const [contratId, setContratId] = useState(contrats[0] ? contrats[0].id : "");

  const fill = (txt) => {
    const k = contrats.find((x) => x.id === contratId) || {};
    return txt
      .replace(/{NOM}/g, (client.nom || "").toUpperCase())
      .replace(/{PRENOM}/g, client.prenom || "")
      .replace(/{TYPE}/g, k.type || "________")
      .replace(/{COMPAGNIE}/g, k.compagnie || "________")
      .replace(/{NUMERO}/g, k.numero || "________");
  };
  const tpl = COURRIER_TEMPLATES.find((t) => t.id === tplId);
  const [objet, setObjet] = useState(fill(tpl.objet));
  const [corps, setCorps] = useState(fill(tpl.corps));
  const pick = (id) => {
    setTplId(id);
    const t = COURRIER_TEMPLATES.find((x) => x.id === id);
    setObjet(fill(t.objet)); setCorps(fill(t.corps));
  };
  const rePick = (cid) => {
    setContratId(cid);
    setTimeout(() => { const t = COURRIER_TEMPLATES.find((x) => x.id === tplId); setObjet(fill(t.objet)); setCorps(fill(t.corps)); }, 0);
  };

  const imprimer = () => {
    const k = contrats.find((x) => x.id === contratId) || {};
    const dateFR = new Date().toLocaleDateString("fr-FR", { day: "numeric", month: "long", year: "numeric" });
    const w = window.open("", "_blank");
    w.document.write(`<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><title>${objet}</title>
      <style>
        body { font-family: Georgia, "Times New Roman", serif; color: #1a1a1a; max-width: 700px; margin: 40px auto; line-height: 1.7; font-size: 14.5px; }
        .exp { font-size: 13px; }
        .dest { text-align: right; margin: 30px 0; font-size: 13px; }
        .objet { font-weight: bold; margin: 30px 0 20px; }
        .corps { white-space: pre-wrap; text-align: justify; }
        .sign { margin-top: 50px; text-align: right; }
        @media print { body { margin: 20mm; } }
      </style></head><body>
      <div class="exp"><b>${(client.prenom || "")} ${(client.nom || "").toUpperCase()}</b><br>${client.ville || ""}<br>${client.telephone || ""}</div>
      <div class="dest">${k.compagnie || "________"}<br>Service Gestion<br><br>${client.ville ? client.ville + ", le " : "Le "}${dateFR}</div>
      <div class="objet">Objet : ${objet}</div>
      <div class="corps">${corps.replace(/&/g, "&amp;").replace(/</g, "&lt;")}</div>
      <div class="sign">${(client.prenom || "")} ${(client.nom || "").toUpperCase()}<br><br><i>Signature :</i></div>
      </body></html>`);
    w.document.close();
    setTimeout(() => w.print(), 400);
  };

  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 680, maxHeight: "88vh", overflowY: "auto" }}>
        <h2>📄 Générer un courrier — {client.nom.toUpperCase()} {client.prenom}</h2>
        <div className="row" style={{ marginTop: 12 }}>
          <Field label="Type de courrier">
            <select className="sel" value={tplId} onChange={(e) => pick(e.target.value)}>
              {COURRIER_TEMPLATES.map((t) => <option key={t.id} value={t.id}>{t.nom}</option>)}
            </select>
          </Field>
          <Field label="Contrat concerné">
            <select className="sel" value={contratId} onChange={(e) => rePick(e.target.value)}>
              {contrats.map((k) => <option key={k.id} value={k.id}>{k.type} {k.compagnie} — n° {k.numero || "?"}</option>)}
              {!contrats.length && <option value="">Aucun contrat sur la fiche</option>}
            </select>
          </Field>
        </div>
        <div style={{ marginTop: 10 }}>
          <Field label="Objet"><input className="in" value={objet} onChange={(e) => setObjet(e.target.value)} /></Field>
        </div>
        <div style={{ marginTop: 10 }}>
          <Field label="Corps du courrier (modifiable — remplissez les ________ si besoin)">
            <textarea className="ta" rows={13} value={corps} onChange={(e) => setCorps(e.target.value)} style={{ fontSize: 13.5, lineHeight: 1.6 }} />
          </Field>
        </div>
        <div className="row" style={{ marginTop: 16, justifyContent: "space-between" }}>
          <button className="btn ghost" onClick={onClose}>Fermer</button>
          <div className="row">
            <button className="btn ghost" onClick={async () => { try { await navigator.clipboard.writeText(corps); alert("Courrier copié ✓"); } catch { } }}>📋 Copier</button>
            <button className="btn gold" onClick={imprimer}>🖨️ Imprimer / PDF</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ================= AUDIT PATRIMONIAL ELYON ================= */
/* Synthèse du document papier : uniquement les informations décisives */
const AUDIT_OBJECTIFS = [
  "Protéger vos proches et vous-même", "Valoriser votre patrimoine", "Préparer l'avenir de vos enfants",
  "Compléter votre retraite", "Protéger son conjoint survivant", "Organiser votre transmission (DMTG)",
  "Financer vos projets personnels", "Optimiser votre fiscalité (IRPP, IFI)", "Développer votre entreprise",
  "Prévoyance / Dépendance",
];
const AUDIT_PROFILS = ["Stratégie prudente", "Stratégie équilibrée", "Stratégie dynamique"];
const AUDIT_VIDE = {
  regimeMat: "", parts: "", revenus: "", tmi: "", impot: "",
  residence: "Propriétaire", immobilier: "", credits: "",
  epargneDispo: "", assuranceVie: "", retraite: "", prevoyance: "",
  objectifs: [], horizon: "", epargneMens: "", capital: "", profil: "", notes: "",
};

function AuditModal({ client, onClose, onSave }) {
  const [f, setF] = useState({ ...AUDIT_VIDE, ...((client.audit || {}).f || {}) });
  const set = (k, v) => setF({ ...f, [k]: v });
  const toggleObj = (o) => set("objectifs", f.objectifs.includes(o) ? f.objectifs.filter((x) => x !== o) : [...f.objectifs, o]);

  const Section = ({ t }) => (
    <div style={{ background: NAVY, color: "#fff", fontSize: 12, fontWeight: 700, letterSpacing: 1.5, textTransform: "uppercase", padding: "7px 12px", borderRadius: 6, margin: "18px 0 12px", borderRight: `4px solid ${GOLD}` }}>{t}</div>
  );

  return (
    <div className="modal-bg" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 720, maxHeight: "90vh", overflowY: "auto" }}>
        <h2>🩺 Audit patrimonial — {client.nom.toUpperCase()} {client.prenom}</h2>

        <Section t="État civil et fiscalité" />
        <div className="fgrid">
          <Field label="Régime matrimonial"><input className="in" value={f.regimeMat} onChange={(e) => set("regimeMat", e.target.value)} placeholder="ex : marié, communauté réduite aux acquêts" /></Field>
          <Field label="Parts fiscales"><input className="in" value={f.parts} onChange={(e) => set("parts", e.target.value)} placeholder="ex : 3" /></Field>
          <Field label="Revenus annuels nets du foyer (€)"><input className="in" value={f.revenus} onChange={(e) => set("revenus", e.target.value)} placeholder="ex : 78 000" /></Field>
          <Field label="TMI (%)"><input className="in" value={f.tmi} onChange={(e) => set("tmi", e.target.value)} placeholder="ex : 30" /></Field>
          <Field label="Impôt net à payer (€)"><input className="in" value={f.impot} onChange={(e) => set("impot", e.target.value)} placeholder="ex : 6 400" /></Field>
        </div>

        <Section t="Patrimoine" />
        <div className="fgrid">
          <Field label="Résidence principale">
            <select className="sel" value={f.residence} onChange={(e) => set("residence", e.target.value)}>
              <option>Propriétaire</option><option>Locataire</option><option>Hébergé à titre gracieux</option>
            </select>
          </Field>
          <Field label="Patrimoine immobilier total (€)"><input className="in" value={f.immobilier} onChange={(e) => set("immobilier", e.target.value)} placeholder="ex : 420 000" /></Field>
          <Field label="Crédits en cours (CRD / mensualités)"><input className="in" value={f.credits} onChange={(e) => set("credits", e.target.value)} placeholder="ex : 180 k€ · 950 €/mois jusque 2038" /></Field>
          <Field label="Épargne disponible (livrets…) (€)"><input className="in" value={f.epargneDispo} onChange={(e) => set("epargneDispo", e.target.value)} placeholder="ex : 25 000" /></Field>
          <Field label="Assurance vie (encours) (€)"><input className="in" value={f.assuranceVie} onChange={(e) => set("assuranceVie", e.target.value)} placeholder="ex : 60 000" /></Field>
          <Field label="Retraite : PER / PERP / Madelin (€)"><input className="in" value={f.retraite} onChange={(e) => set("retraite", e.target.value)} placeholder="ex : PER 30 000" /></Field>
          <Field label="Prévoyance en place"><input className="in" value={f.prevoyance} onChange={(e) => set("prevoyance", e.target.value)} placeholder="ex : April, IJ + invalidité, 45 €/mois" /></Field>
        </div>

        <Section t="Objectifs patrimoniaux" />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
          {AUDIT_OBJECTIFS.map((o) => (
            <label key={o} style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 13, cursor: "pointer", padding: "5px 8px", borderRadius: 6, background: f.objectifs.includes(o) ? "#fdf9f0" : "transparent", border: f.objectifs.includes(o) ? `1px solid ${GOLD}` : "1px solid transparent" }}>
              <input type="checkbox" checked={f.objectifs.includes(o)} onChange={() => toggleObj(o)} />
              {o}
            </label>
          ))}
        </div>

        <Section t="Moyens et profil investisseur" />
        <div className="fgrid">
          <Field label="Horizon de temps"><input className="in" value={f.horizon} onChange={(e) => set("horizon", e.target.value)} placeholder="ex : 10 ans (retraite)" /></Field>
          <Field label="Épargne mensuelle possible (€)"><input className="in" value={f.epargneMens} onChange={(e) => set("epargneMens", e.target.value)} placeholder="ex : 400" /></Field>
          <Field label="Capital mobilisable (€)"><input className="in" value={f.capital} onChange={(e) => set("capital", e.target.value)} placeholder="ex : 15 000" /></Field>
          <Field label="Profil investisseur">
            <select className="sel" value={f.profil} onChange={(e) => set("profil", e.target.value)}>
              <option value="">—</option>
              {AUDIT_PROFILS.map((p) => <option key={p}>{p}</option>)}
            </select>
          </Field>
        </div>

        <Section t="Synthèse du rendez-vous" />
        <textarea className="ta" rows={3} value={f.notes} onChange={(e) => set("notes", e.target.value)} placeholder="Notes, critères de solution idéale, points d'attention…" />

        <div className="row" style={{ marginTop: 18, justifyContent: "flex-end" }}>
          <button className="btn ghost" onClick={onClose}>Annuler</button>
          <button className="btn gold" onClick={() => onSave(f)}>✓ Enregistrer l'audit</button>
        </div>
      </div>
    </div>
  );
}

/* Synthèse visible sur la fiche client */
function AuditSynthese({ a }) {
  const items = [
    ["🧾", "Fiscalité", [a.tmi && `TMI ${a.tmi} %`, a.impot && `Impôt ${a.impot} €`, a.parts && `${a.parts} parts`].filter(Boolean).join(" · ")],
    ["💶", "Revenus foyer", a.revenus && `${a.revenus} €/an`],
    ["🏠", "Immobilier", [a.residence, a.immobilier && `${a.immobilier} €`].filter(Boolean).join(" · ")],
    ["🏦", "Épargne", [a.epargneDispo && `Dispo ${a.epargneDispo} €`, a.assuranceVie && `AV ${a.assuranceVie} €`, a.retraite && `Retraite ${a.retraite}`].filter(Boolean).join(" · ")],
    ["🎯", "Objectifs", (a.objectifs || []).slice(0, 3).join(" · ") + ((a.objectifs || []).length > 3 ? "…" : "")],
    ["📊", "Profil", [a.profil, a.horizon && `horizon ${a.horizon}`].filter(Boolean).join(" · ")],
    ["💰", "Moyens", [a.epargneMens && `${a.epargneMens} €/mois`, a.capital && `capital ${a.capital} €`].filter(Boolean).join(" · ")],
  ].filter((x) => (x[2] || "").trim());
  if (!items.length) return <div style={{ color: "#8593a8", fontSize: 13.5 }}>Audit vide.</div>;
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8 }}>
      {items.map(([ic, l, v]) => (
        <div key={l} style={{ background: "#f8f9fc", borderRadius: 8, padding: "8px 10px", borderLeft: `3px solid ${GOLD}` }}>
          <div style={{ fontSize: 10.5, color: "#8593a8", textTransform: "uppercase", letterSpacing: 0.5 }}>{ic} {l}</div>
          <div style={{ fontSize: 12.5, color: NAVY, marginTop: 2 }}>{v}</div>
        </div>
      ))}
    </div>
  );
}

/* PDF « Audit patrimonial Elyon » — mise en page fidèle au document du cabinet */
function auditPDF(client) {
  const a = (client.audit || {}).f || {};
  const dateFR = new Date().toLocaleDateString("fr-FR");
  const L = (label, val) => `<div class="line"><span class="lab">${label}</span><span class="val">${val || "—"}</span></div>`;
  const w = window.open("", "_blank");
  w.document.write(`<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8">
  <title>Audit patrimonial Elyon</title>
  <style>
    body { font-family: "Helvetica Neue", Arial, sans-serif; color:#1a1a1a; max-width: 760px; margin: 24px auto; font-size: 13px; }
    .head { background:#111; color:#fff; padding: 16px 22px; display:flex; justify-content:space-between; align-items:center; }
    .head b { letter-spacing: 3px; font-size: 15px; }
    .head span { font-size: 10px; letter-spacing: 2px; color:#C9A24B; }
    h1 { font-family: Georgia, serif; font-weight: 500; font-size: 26px; margin: 18px 0 2px; }
    .gold { height:3px; width:120px; background:#C9A24B; margin-bottom: 14px; }
    .meta { font-size: 11.5px; color:#555; margin-bottom: 6px; }
    .sec { background:#111; color:#fff; font-size: 11px; letter-spacing: 2px; padding: 6px 12px; margin: 16px 0 8px; border-right: 6px solid #C9A24B; text-transform: uppercase; }
    .line { display:flex; border-bottom: 1px dotted #bbb; padding: 5px 2px; }
    .lab { width: 260px; color:#444; font-weight: 600; font-size: 11.5px; flex-shrink:0; }
    .val { flex:1; }
    .obj { display:inline-block; border:1px solid #C9A24B; border-radius: 4px; padding: 2px 8px; margin: 2px 4px 2px 0; font-size: 11.5px; }
    .notes { border:1px solid #ddd; padding: 10px; min-height: 60px; white-space: pre-wrap; }
    @media print { body { margin: 10mm; } }
  </style></head><body>
  <div class="head"><b>ELYON &nbsp;&amp;&nbsp; ASSOCIÉS</b><span>AUDIT PATRIMONIAL</span></div>
  <h1>Audit Patrimonial</h1><div class="gold"></div>
  <div class="meta">Client : <b>${(client.nom || "").toUpperCase()} ${client.prenom || ""}</b> · ${client.profession || ""} · Édité le ${dateFR}</div>
  <div class="sec">État civil et fiscalité</div>
  ${L("Situation / régime matrimonial", [client.situation, a.regimeMat].filter(Boolean).join(" — "))}
  ${L("Enfants", client.enfants)}
  ${L("Parts fiscales", a.parts)}
  ${L("Revenus annuels nets du foyer", a.revenus && a.revenus + " €")}
  ${L("TMI", a.tmi && a.tmi + " %")}
  ${L("Impôt net à payer", a.impot && a.impot + " €")}
  <div class="sec">Patrimoine</div>
  ${L("Résidence principale", a.residence)}
  ${L("Patrimoine immobilier total", a.immobilier && a.immobilier + " €")}
  ${L("Crédits en cours", a.credits)}
  ${L("Épargne disponible", a.epargneDispo && a.epargneDispo + " €")}
  ${L("Assurance vie", a.assuranceVie && a.assuranceVie + " €")}
  ${L("Retraite (PER, PERP, Madelin)", a.retraite)}
  ${L("Prévoyance", a.prevoyance)}
  <div class="sec">Objectifs patrimoniaux</div>
  <div>${(a.objectifs || []).map((o) => `<span class="obj">${o}</span>`).join("") || "—"}</div>
  <div class="sec">Moyens et profil investisseur</div>
  ${L("Horizon de temps", a.horizon)}
  ${L("Épargne mensuelle possible", a.epargneMens && a.epargneMens + " €/mois")}
  ${L("Capital mobilisable", a.capital && a.capital + " €")}
  ${L("Profil investisseur", a.profil)}
  <div class="sec">Synthèse du rendez-vous</div>
  <div class="notes">${(a.notes || "—").replace(/&/g, "&amp;").replace(/</g, "&lt;")}</div>
  </body></html>`);
  w.document.close();
  setTimeout(() => w.print(), 400);
}
