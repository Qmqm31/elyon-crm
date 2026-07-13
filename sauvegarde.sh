#!/bin/bash
URL="https://elyon-crm-production.up.railway.app"
CODE="VOTRE_CODE"
cd "$(dirname "$0")"
mkdir -p sauvegardes
FICHIER="sauvegardes/ELYON_donnees_$(date '+%Y-%m-%d_%Hh%M').json"
echo "💾 Sauvegarde des données du CRM..."
{
  echo '{'
  PREM=1
  for KEY in crm-users crm-clients crm-sales crm-docs crm-bordereaux crm-prospection crm-objectifs crm-trash crm-rdv-clients crm-mailtpl crm-settings crm-objectifs-prospection; do
    VAL=$(curl -s -u "elyon:$CODE" "$URL/api/storage/$KEY")
    if [ "$PREM" = "1" ]; then PREM=0; else echo ','; fi
    printf '"%s": %s' "$KEY" "${VAL:-null}"
  done
  echo ''
  echo '}'
} > "$FICHIER"
if grep -q "crm-clients" "$FICHIER"; then
  echo "✅ Sauvegarde créée : $FICHIER"
else
  echo "❌ Échec — vérifiez URL et CODE dans ce script."
  rm -f "$FICHIER"
fi
