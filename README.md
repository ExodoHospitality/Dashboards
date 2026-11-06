# Exodo Hospitality – Dashboards

Live site: the Weekly Report is `index.html`; the four invoice dashboards are `akiro/`, `Kayao/`, `Matilda/`, `Ayayay/`.

| Folder | What is in it | Edited by |
|---|---|---|
| `data/` | the five Excel workbooks (sales log + four TEST_ invoice workbooks) – the only source of every number | you, in Excel |
| `build/` | the builder scripts and the HTML templates | never by hand |
| root + four folders | the published pages | rebuilt by `bash build/build_all.sh` |

Routine: type the batch in Excel → send the workbook to Claude → Claude replaces it in `data/`, runs the build and commits. Every version stays in the history (file › History).
