#!/usr/bin/env bash
# Actualización semanal del portal de transparencia.
#   1. Copia el Excel más reciente desde Google Drive a data/
#   2. Regenera index.html
#   3. Publica en GitHub Pages
#
# Uso:  ./actualizar.sh
set -euo pipefail
cd "$(dirname "$0")"

DRIVE="$HOME/Library/CloudStorage/GoogleDrive-jbarbozameca@gmail.com/My Drive/1. TRAINING SYSTEMATIC REVIEW/INFORME TRANSPARENCIA"
XLSX="TRANSPARENCIA - INVESTIGACIONES  Tau 2024-2026.xlsx"

if [ -f "$DRIVE/$XLSX" ]; then
  cp "$DRIVE/$XLSX" "data/$XLSX"
  echo "✓ Excel actualizado desde Drive"
else
  echo "⚠ No encontré el Excel en Drive; uso la copia que ya está en data/"
fi

python3 -c "import openpyxl" 2>/dev/null || pip3 install --quiet openpyxl
python3 tools/build_portal.py "data/$XLSX"

git add -A
if git diff --cached --quiet; then
  echo "Sin cambios que publicar."
  exit 0
fi
git commit -m "Actualización $(date +%d.%m.%Y)"
git push
echo "✓ Publicado en https://jbarbozameca.github.io/publicacionestau/ (tarda 1–2 minutos en refrescarse)"
