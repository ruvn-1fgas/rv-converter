import locale

TRANSLATIONS = {
    "en": {
        "error_data_empty": "Error: File contains no data.",
        "error_data_format": "Error: Data should be a list of dictionaries.",
        "error_data_none": "Error: Failed to load data from file.",
        "file_invalid_json": "Error: File '{}' contains invalid JSON.",
        "file_invalid_jsonl": "Error: File '{}' contains invalid JSONL. Line error: {}",
        "file_not_exist": "Error: File '{}' does not exist.",
        "file_not_json": "Error: File '{}' is not a JSON or JSONL file.",
        "file_read_error": "Error: An error occurred while reading the file: {}",
        "filename_help": "Path to JSON/JSONL file",
        "index_out_of_range": "Index should be greater than 0",
        "join_multivalued_help": "String to join multiple values",
        "output_help": "Path to output file",
        "processing_sheets": "Processing sheets",
        "program_description": "VG JSON Converter -> XLSX",
        "sheet_name_with_chunk": "Sheet {}_{}",
        "sheet_name": "Sheet {}",
        "sort_by_help": "Key for sorting",
        "total_sheets_columns": "Total sheets: {}\nTotal columns: {}",
        "type_help": "Data save format (xlsx)",
        "unsupported_file_type": "Unsupported file type: {}. Supported types are: xlsx",
        "writing_rows_for": "Writing rows for {}",
    },
    "ru": {
        "error_data_empty": "Ошибка: Файл не содержит данных.",
        "error_data_format": "Ошибка: Данные должны быть списком словарей.",
        "error_data_none": "Ошибка: Не удалось загрузить данные из файла.",
        "file_invalid_json": "Ошибка: Файл '{}' содержит некорректный JSON.",
        "file_invalid_jsonl": "Ошибка: Файл '{}' содержит некорректный JSONL. Ошибка в строке: {}",
        "file_not_exist": "Ошибка: Файл '{}' не существует.",
        "file_not_json": "Ошибка: Файл '{}' не является JSON или JSONL файлом.",
        "file_read_error": "Ошибка: Произошла ошибка при чтении файла: {}",
        "filename_help": "Путь до JSON/JSONL файла",
        "index_out_of_range": "Индекс должен быть больше 0",
        "join_multivalued_help": "Строка для объединения множественных значений",
        "output_help": "Путь до выходного файла",
        "processing_sheets": "Обработка листов",
        "program_description": "VG Конвертер JSON -> XLSX",
        "sheet_name_with_chunk": "Лист {}_{}",
        "sheet_name": "Лист {}",
        "sort_by_help": "Ключ для сортировки",
        "total_sheets_columns": "Всего листов: {}\nВсего столбцов: {}",
        "type_help": "Формат сохранения данных (xlsx)",
        "unsupported_file_type": "Неподдерживаемый тип файла: {}. Поддерживаемые типы: xlsx",
        "writing_rows_for": "Запись строк для {}",
    },
}


def get_system_language() -> str:
    try:
        system_locale = locale.getdefaultlocale()[0]

        if system_locale:
            lang_code = system_locale.split("_")[0].lower()

            if lang_code in TRANSLATIONS:
                return lang_code
    except Exception:
        pass

    return "en"


class I18n:
    def __init__(self, lang: str = None):
        if lang is None:
            lang = get_system_language()

        self.lang = lang if lang in TRANSLATIONS else "en"
        self.translations = TRANSLATIONS[self.lang]

    def get(self, key: str, *args) -> str:
        text = self.translations.get(key, TRANSLATIONS["en"].get(key, key))

        if args:
            return text.format(*args)
        return text

    def __getitem__(self, key: str) -> str:
        return self.get(key)


_i18n = I18n()


def get_i18n() -> I18n:
    return _i18n


def t(key: str, *args) -> str:
    return _i18n.get(key, *args)
