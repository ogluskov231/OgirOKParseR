# ОгрОК_UA

## 1 клік (Windows)
1) Покладіть DOCX в `input/documents/`
2) Запустіть `RUN_GUI.bat` (або `RUN_WINDOWS.bat` якщо використовуєте старий сценарій)
3) Натисніть **Сканувати** → **Обробити**
4) Результати:
- Excel: `output/exports/excel/N_filled_*.xlsx`
- JSON: `output/exports/json/combined_*.json` (опційно)
- TXT: `output/exports/txt/summary_*.txt` (опційно)
- Невідоме: `output/diagnostics/unknowns_*.json`

## CLI (без GUI)
Запустіть `RUN_CLI.bat` або вручну:

- Windows:
  - `py run_parser_orders_cli.py --input "input\documents" --excel --json --txt`
- Linux/WSL:
  - `python3 run_parser_orders_cli.py --input ./input/documents --excel --json --txt`

## Важливо про Excel-форматування
Автозаповнення тепер:
- копіює **стилі та формули** з першого рядка даних (як "шаблон рядка")
- автоматично **розширює Data Validation** (випадаючі списки), якщо заповнення виходить за межі початкових діапазонів
- якщо у файлі є **Excel Table (ListObject)**, розширює `Table.ref` під нові рядки

## Самонавчання
Вкладка **🧠 Самонавчання** дозволяє додавати правила (підрозділи/причини/джерела) і ключові слова подій.

## Вкладка 👤 Люди (A–Я)
Показує **всю інформацію по вибраній людині** (усі записи/фрагменти), відсортовано **за алфавітом ПІБ**. Є експорт у TXT/JSON.


## Збірка .exe (Windows)

1. Встановіть Python 3.10+.
2. Запустіть `scripts\\build_exe_windows.bat`.
3. Готовий реліз буде у `dist\\OgrOK_UA\\` (папка з `OgrOK_UA.exe`).

> Іконка/лого: `assets\\icon.ico` та `assets\\logo.png`.

## JSON Schema парсера

Результат `parse_docx_to_structured()` містить:
- `structured_schema` — структурований документ за `config/schema/military_order.schema.json`;
- `schema_validation.valid` — результат перевірки;
- `schema_validation.errors` — перелік помилок схеми.

Під час парсингу зберігається також старий `doc_struct_v4`/`combined_v4`, щоб не ламати наявний GUI та Excel-автозаповнення.

Schema: `config/schema/military_order.schema.json`.
