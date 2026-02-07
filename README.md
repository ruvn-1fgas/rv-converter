# JSON -> XLSX Converter

Конвертер JSON/JSONL файлов в Excel (XLSX).

## Возможности

- Конвертация JSON и JSONL файлов в Excel
- Автоматическое разбиение больших данных на листы (лимиты Excel)
- Сортировка и фильтрация данных
- Подсчёт заполненных ячеек в каждой колонке
- Автофильтры и закреплённые заголовки
- Поддержка метаданных Excel-файла
- Локализация интерфейса (русский/английский)
- Потоковая обработка больших JSON (с `ijson`)

## Установка

```bash
pip install rv_converter
```

Для потоковой обработки больших JSON-файлов:

```bash
pip install rv_converter ijson
```

## Использование

### Командная строка

```bash
rv_converter data.json                       # output: data.xlsx
rv_converter data.json -o output.xlsx        # указать имя выходного файла
rv_converter data.json -s name               # сортировка по ключу "name"
rv_converter data.json -s name -o out.xlsx   # сортировка + имя файла
rv_converter data.json -j "; "               # разделитель для списков
rv_converter data.json --sort-by name --output result.xlsx  # полные имена
```

**Опции:**

- `-o`, `--output` — имя выходного файла
- `-s`, `--sort-by` — ключ для сортировки
- `-j`, `--join` — разделитель для множественных значений
- `-t`, `--type` — формат файла (по умолчанию: xlsx)

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
```

### Загрузка данных из файла

```python
from rv_converter import converter
from rv_converter.utils import load_data

data = load_data("input.json")  # или input.jsonl
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
| `type`             | `str`                          | Формат файла (`"xlsx"`)                        |
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
