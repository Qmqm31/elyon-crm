#!/bin/bash
cd "$(dirname "$0")"
echo "📦 Envoi des modifications vers GitHub + Railway..."
git add -A
git commit -m "MAJ CRM $(date '+%d/%m/%Y %H:%M')" || { echo "ℹ️  Rien à publier."; exit 0; }
git push
echo "✅ Envoyé ! Railway redéploie automatiquement (2-3 min)."
