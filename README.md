# ai-ethics-cases — «Этика в сфере искусственного интеллекта»

Решения 4 кейс-заданий дисциплины «Искусственный интеллект, большие
данные и этика» (Университет «Синергия», группа ДМЭ-101цэии).
Продукт — по одному отчёту (docx + pdf) на кейс в `itog/`.

## Состав сдачи

| Кейс | Файл сдачи |
|---|---|
| 1. Беспристрастный HR | `ДМЭ-101ЦЭИИ.Шатов.Кейс_1_Беспристрастный_HR.docx` (+ `.pdf`) |
| 2. Медицинский ассистент | `ДМЭ-101ЦЭИИ.Шатов.Кейс_2_Медицинский_ассистент.docx` (+ `.pdf`) |
| 3. Код без автора | `ДМЭ-101ЦЭИИ.Шатов.Кейс_3_Код_без_автора.docx` (+ `.pdf`) |
| 4. Deepfake-проверка | `ДМЭ-101ЦЭИИ.Шатов.Кейс_4_Deepfake_проверка.docx` (+ `.pdf`) |

Оформление — клон ВКР Синергии по ГОСТ 7.32-2017: Times New Roman 14,
поля 3/1,5/2/2, титульный лист «РЕШЕНИЕ КЕЙСА № N», СОДЕРЖАНИЕ,
«Вывод:» в каждом разделе, сквозная нумерация таблиц. Стандарт:
[REPORT_STANDARDS.md](REPORT_STANDARDS.md) (база —
[`../big-data-labs/templates/REPORT_STANDARDS.md`](../big-data-labs/templates/REPORT_STANDARDS.md)).

## Структура

```
assignments/keisy_etika_ii.md   тексты кейс-заданий (read-only)
cases/case-01..04/content.py    содержимое отчётов (тексты + таблицы)
cases/case-01..04/output/       собранные docx
templates/synergy_docx.py       движок стилей (копия big-data-labs + work_type)
tools/build_reports.py          сборщик docx
tools/convert_pdf.sh            docx -> pdf (soffice)
itog/                           файлы сдачи
```

## Сборка

Требуется Python 3 с `python-docx` (`pip install python-docx`).

```bash
python3 tools/build_reports.py            # docx всех 4 кейсов -> itog/
tools/convert_pdf.sh                      # docx -> pdf (нужен soffice)
python3 tools/build_reports.py --update-toc   # фактические № страниц в СОДЕРЖАНИИ по PDF
tools/convert_pdf.sh                      # финальный pdf
```

Собрать один кейс: `python3 tools/build_reports.py --case 2`.

**PDF без soffice на хосте** — через Docker (Liberation Serif метрически
совместим с Times New Roman, разбивка страниц не меняется):

```bash
docker build -t lo-pdf /tmp/kilo/lo-pdf   # или готовый образ с soffice
docker run --rm --pull never -u lo -v "$(pwd)/itog:/work" lo-pdf \
    soffice --headless --convert-to pdf --outdir /work "/work/ДМЭ-101ЦЭИИ.Шатов.Кейс_1_Беспристрастный_HR.docx"
```

## Проверка перед сдачей

1. `python3 ../synergy-agent/scripts/sanitize_ooxml.py itog/*.docx`
2. `python3 ../synergy-agent/scripts/validate_office.py itog/*.docx`
3. Визуально: титул (подчёркнутые значения), СОДЕРЖАНИЕ с номерами
   страниц, «(таблица N)» и «Вывод:» в каждом разделе.
4. Чек-лист `REPORT_STANDARDS.md` (раздел 5) по каждому файлу.

## Правила для агентов

См. [AGENTS.md](AGENTS.md): язык, единый источник содержимого,
запрет править `assignments/`, `itog/` и соседние репозитории вручную.
