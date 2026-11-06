#!/usr/bin/env bash
# Rebuilds every page from the workbooks in data/ into the published folders.
# Run from the repo root:  bash build/build_all.sh
set -e
cd "$(dirname "$0")/.."
mkdir -p akiro Kayao Matilda Ayayay
python3 -I build/build_weekly_report.py data/Exodo_weekly_sales_log.xlsx build/template_weekly_report.html index.html final
python3 -I build/build_akiro.py   data/TEST_Akiro_v2.xlsx   build/template_akiro.html   akiro/index.html
python3 -I build/build_kayao.py   data/TEST_Kayao_v2.xlsx   build/template_kayao.html   Kayao/index.html
python3 -I build/build_matilda.py data/TEST_Matilda_v2.xlsx build/template_matilda.html Matilda/index.html
if [ -f build/template_ayayay.html ]; then
  python3 -I build/build_ayayay.py data/TEST_Ayayay_v2.xlsx build/template_ayayay.html Ayayay/index.html
else
  echo "Ayayay: template missing - page not rebuilt"
fi
