# CRM ELYON & Associés — Déploiement sur Railway

## Ce que contient ce dossier
- `src/CRM.jsx` — votre CRM (version 7, identique à l'artifact)
- `src/main.jsx` — l'adaptateur qui relie le CRM au serveur
- `server.js` — le serveur : héberge le CRM + sauvegarde les données + protection par code d'accès
- `package.json`, `index.html`, `favicon.png` — configuration et icône ELYON

## Déploiement (15 minutes, une seule fois)

### 1. Mettre le code sur GitHub
1. Créez un compte sur github.com si besoin
2. Cliquez sur **New repository** → nom : `elyon-crm` → cochez **Private** → Create
3. Cliquez sur **uploading an existing file** → glissez TOUS les fichiers de ce dossier → Commit

### 2. Déployer sur Railway
1. Sur railway.app → **New Project** → **Deploy from GitHub repo** → choisissez `elyon-crm`
2. Railway détecte Node.js et construit tout automatiquement

### 3. Rendre les données PERMANENTES (étape cruciale)
1. Cliquez sur votre service → onglet **Settings** → section **Volumes** → **Add Volume**
2. Mount path : `/data`
> Sans volume, les données seraient effacées à chaque redéploiement !

### 4. Configurer les variables (onglet Variables)
- `DATA_DIR` = `/data`
- `ACCESS_CODE` = un code de votre choix (ex : `Elyon2026!`)

### 5. Obtenir votre adresse
Settings → **Networking** → **Generate Domain** → vous obtenez une URL du type
`elyon-crm-production.up.railway.app`

### 6. Premier accès
Ouvrez l'URL : le navigateur demande un identifiant.
- Utilisateur : n'importe quoi (ex : `elyon`)
- Mot de passe : votre ACCESS_CODE

### 7. Rapatrier vos données
1. Dans l'ANCIEN CRM (artifact Claude) : connectez-vous en Quentin → 💾 Exporter les données
2. Dans le NOUVEAU CRM (Railway) : connectez-vous → 📥 Importer une sauvegarde → choisissez le fichier
3. Tout est transféré : clients, ventes, prospection, objectifs ✓

## Coût
Railway : environ 5 $/mois (plan Hobby) pour ce type d'application.

## Mises à jour futures
Remplacez le fichier `src/CRM.jsx` sur GitHub par la nouvelle version → Railway
redéploie automatiquement. Les données ne bougent pas (elles sont sur le volume).
