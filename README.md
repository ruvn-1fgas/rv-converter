# RV Converter (v2026.02.15.2)

Универсальный конвертер данных между форматами JSON, JSONL, CSV и XLSX.

## Возможности

- Конвертация между JSON, JSONL, CSV и XLSX
- Обратная конвертация: XLSX/CSV → JSON/JSONL
- Автоматическое разбиение больших данных на листы (лимиты Excel)
- Сортировка и фильтрация данных
- Подсчёт заполненных ячеек в каждой колонке
- Автофильтры и закреплённые заголовки
- Поддержка метаданных Excel-файла
- Локализация интерфейса (русский/английский)
- Потоковая обработка больших JSON (с `ijson`)

## Установка

```bash
pip install rv-converter
```

### Установка из исходников

```bash
git clone https://github.com/ruvn-1fgas/rv-converter.git
cd rv-converter
pip install .
```

Или

```bash
pip install git+https://github.com/ruvn-1fgas/rv-converter.git
```

Для потоковой обработки больших JSON-файлов:

```bash
pip install ijson
```

## Использование

### Командная строка

```bash
rv_converter data.json                       # output: data.xlsx
rv_converter data.json -o output.xlsx        # указать имя выходного файла
rv_converter data.json -s name               # сортировка по ключу "name"
rv_converter data.json -s name -o out.xlsx   # сортировка + имя файла
rv_converter data.json -j "; "               # разделитель для списков
rv_converter data.json -t csv                # вывод в CSV
rv_converter data.xlsx -t json               # XLSX → JSON
rv_converter data.csv -t xlsx                # CSV → XLSX
rv_converter data.xlsx -t jsonl              # XLSX → JSONL
rv_converter data.json --sort-by name --output result.xlsx  # полные имена
```

**Опции:**

- `-o`, `--output` — имя выходного файла
- `-s`, `--sort-by` — ключ для сортировки
- `-j`, `--join` — разделитель для множественных значений
- `-t`, `--type` — формат выходного файла (xlsx, csv, json, jsonl; по умолчанию: xlsx)

### Python API

```python
from rv_converter import converter

data = [
    {"name": "Alice", "age": 30},
    {"name": "Bob", "age": 25},
]

# Базовое использование
converter.save_file(data, "output.xlsx")

# С сортировкой по ключу
converter.save_file(data, "output.xlsx", sort_by="name")

# С сортировкой по нескольким ключам
converter.save_file(data, "output.xlsx", sort_by=["age", "name"])

# С пользовательской сортировкой
converter.save_file(data, "output.xlsx", sort_by=lambda x: x["name"].lower())

# С фильтрацией
converter.save_file(
    data, 
    "output.xlsx", 
    filter_function=lambda x: x["age"] > 25
)

# С метаданными файла
converter.save_file(
    data,
    "output.xlsx",
    xlsx_metadata={"title": "Отчёт", "creator": "RV Converter"}
)

# Сохранение в CSV
converter.save_file(data, "output.csv", type="csv")

# Сохранение в JSON
converter.save_file(data, "output.json", type="json")

# Сохранение в JSONL
converter.save_file(data, "output.jsonl", type="jsonl")
```

### Конвертация между форматами

```python
from rv_converter import converter

# XLSX -> JSON
converter.convert_file("data.xlsx", "data.json")

# CSV -> XLSX
converter.convert_file("data.csv", "data.xlsx")

# JSON -> CSV
converter.convert_file("data.json", "data.csv")

# XLSX -> JSONL
converter.convert_file("data.xlsx", "data.jsonl")
```

### Загрузка данных из файла

```python
from rv_converter import converter
from rv_converter.utils import load_data

data = load_data("input.json")   # или input.jsonl / input.csv / input.xlsx
converter.save_file(data, "output.xlsx")
```

## API

### `converter.save_file()`

| Параметр           | Тип                            | Описание                                       |
|--------------------|--------------------------------|------------------------------------------------|
| `data`             | `List[Dict]`                   | Список словарей для записи                     |
| `filename`         | `str`                          | Путь к выходному файлу                         |
| `sort_by`          | `str \| List[str] \| Callable` | Ключ, список ключей или функция сортировки     |
| `filter_function`  | `Callable[[Dict], bool]`       | Функция фильтрации                             |
| `join_multivalued` | `str`                          | Разделитель для списков (по умолчанию: `", "`) |
| `type`             | `str`                          | Формат файла (`"xlsx"`, `"csv"`, `"json"`, `"jsonl"`) |
| `xlsx_metadata`    | `Dict[str, Any]`               | Метаданные Excel (title, creator, etc.)        |

### `converter.convert_file()`

| Параметр           | Тип                            | Описание                                       |
|--------------------|--------------------------------|------------------------------------------------|
| `input_file`       | `str`                          | Путь ко входному файлу (json/jsonl/csv/xlsx)    |
| `output_file`      | `str`                          | Путь к выходному файлу (json/jsonl/csv/xlsx)   |
| `sort_by`          | `str \| List[str] \| Callable` | Ключ, список ключей или функция сортировки     |
| `filter_function`  | `Callable[[Dict], bool]`       | Функция фильтрации                             |
| `join_multivalued` | `str`                          | Разделитель для списков (по умолчанию: `", "`) |
| `xlsx_metadata`    | `Dict[str, Any]`               | Метаданные Excel (title, creator, etc.)        |

## Основные зависимости

| Название       | Ссылка                                               |
|----------------|------------------------------------------------------|
| Python >= 3.11 | [Ссылка](https://www.python.org/)                    |
| openpyxl       | [Ссылка](https://openpyxl.readthedocs.io/en/stable/) |
| pandas         | [Ссылка](https://pandas.pydata.org/)                 |
| tqdm           | [Ссылка](https://tqdm.github.io/)                    |

## Лицензия

MIT

## Changelog

### v2026.02.15.2

- Добавлена поддержка CSV (чтение и запись)
- Добавлена обратная конвертация XLSX/CSV → JSON/JSONL
- Добавлен метод `convert_file()` для конвертации между любыми форматами
- Расширен CLI: поддержка входных форматов CSV/XLSX и выходных json/jsonl/csv/xlsx

### v2026.02.15.1

- Заменены статичные числа в строке итогов на формулы SUBTOTAL (подсчёт автоматически обновляется при фильтрации)
