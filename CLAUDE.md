@AGENTS.md

Дополнительно для Claude Code:

- Единый источник содержимого отчётов — `cases/*/content.py`; собранные
  docx руками не править.
- Локальный стандарт кейс-отчётов — `REPORT_STANDARDS.md`; базовый —
  `../big-data-labs/templates/REPORT_STANDARDS.md`.
- Цикл сборки: `python3 tools/build_reports.py` → `tools/convert_pdf.sh`
  → `python3 tools/build_reports.py --update-toc` → `tools/convert_pdf.sh`.
