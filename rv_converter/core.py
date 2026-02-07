from typing import Any, Callable, Dict, List

from openpyxl import Workbook
from openpyxl.cell.cell import KNOWN_TYPES
from pandas import DataFrame
from tqdm import tqdm

from rv_converter.constants import EXCEL_MAX_SHEETS

from .i18n import get_i18n
from .utils import (
    add_count,
    add_filters,
    make_bold_and_freeze,
    split_data,
    split_dataframe,
    split_dataframe_by_columns,
)

i18n = get_i18n()


class Converter:
    def save_file(
        self,
        data: List[Dict],
        filename: str,
        sort_by: Any = None,
        filter_function: Callable[[Dict], bool] = None,
        join_multivalued: str = ", ",
        type: str = "xlsx",
        xlsx_metadata: Dict[str, Any] = None,
    ) -> None:
        """Сохраняет данные в файл указанного формата.

        Args:
            data: Список словарей для записи.
            filename: Путь к выходному файлу.
            sort_by: Ключ, список ключей или функция для сортировки.
            filter_function: Функция фильтрации данных.
            join_multivalued: Разделитель для объединения множественных значений.
            type: Формат файла ('xlsx').
            xlsx_metadata: Метаданные для Excel-файла, см. openpyxl.packaging.core (DocumentProperties).
        """

        if type not in ["xlsx"]:
            raise ValueError(i18n.get("unsupported_file_type", type))

        if not data:
            return

        data = self._apply_filtering(data, filter_function)
        data = self._apply_sorting(data, sort_by)

        if type == "xlsx":
            df = DataFrame(data)
            chunked_dfs = split_dataframe(df)
            chunked_data = split_data(data)

            self._write_to_excel(
                chunked_dfs, chunked_data, filename, join_multivalued, xlsx_metadata
            )

    def _apply_filtering(
        self, data: List[Dict], filter_function: Callable[[Dict], bool]
    ) -> List[Dict]:
        """Применяет функцию фильтрации к данным.

        Args:
            data: Исходный список словарей.
            filter_function: Функция фильтрации.

        Returns:
            Отфильтрованный список словарей.
        """
        return list(filter(filter_function, data)) if filter_function else data

    def _apply_sorting(self, data: List[Dict], sort_by: Any) -> List[Dict]:
        """Сортирует данные по указанному критерию.

        Args:
            data: Список словарей для сортировки.
            sort_by: Ключ, список ключей или функция сортировки.

        Returns:
            Отсортированный список словарей.
        """
        if not sort_by:
            return data

        if isinstance(sort_by, str):
            return sorted(data, key=lambda x: x[sort_by])
        if isinstance(sort_by, list):
            return sorted(data, key=lambda x: tuple(x[sb] for sb in sort_by))
        if callable(sort_by):
            return sorted(data, key=sort_by)
        return data

    def _write_to_excel(
        self,
        dfs: List[DataFrame],
        data_chunks: List[List[Dict]],
        filename: str,
        join_multivalued: str,
        xlsx_metadata: Dict[str, Any] = None,
    ) -> None:
        """Записывает данные в Excel-файл с разбиением на листы.

        Args:
            dfs: Список DataFrame для записи.
            data_chunks: Список частей данных.
            filename: Путь к выходному файлу.
            join_multivalued: Разделитель для множественных значений.
            xlsx_metadata: Метаданные для файла.
        """
        wb = Workbook()
        wb.remove(wb.active)

        header_offset = 0
        filter_row, bold_row = 2 + header_offset, 3 + header_offset
        sheet_index = 1

        print(i18n.get("total_sheets_columns", len(data_chunks), len(dfs[0].columns)))
        total_sheets = 0

        for df, data_chunk in tqdm(
            zip(dfs, data_chunks), desc=i18n.get("processing_sheets")
        ):
            df_chunks = split_dataframe_by_columns(df)

            for chunk_index, df_chunk in enumerate(df_chunks, start=1):
                sheet_name = (
                    i18n.get("sheet_name_with_chunk", sheet_index, chunk_index)
                    if len(df_chunks) > 1
                    else i18n.get("sheet_name", sheet_index)
                )

                self._write_sheet(
                    wb,
                    sheet_name,
                    df_chunk,
                    data_chunk,
                    join_multivalued,
                    filter_row,
                    bold_row,
                )

                total_sheets += 1

                if total_sheets >= EXCEL_MAX_SHEETS:
                    wb.save(filename)
                    return

            sheet_index += 1

        if xlsx_metadata:
            props = wb.properties
            for key, value in xlsx_metadata.items():
                if hasattr(props, key):
                    setattr(props, key, value)

        wb.save(filename)

    def _write_sheet(
        self,
        wb: Workbook,
        sheet_name: str,
        df_chunk: DataFrame,
        data_chunk: List[Dict],
        join_multivalued: str,
        filter_row: int,
        bold_row: int,
    ) -> None:
        """Создаёт лист Excel и заполняет его данными.

        Args:
            wb: Рабочая книга Excel.
            sheet_name: Имя листа.
            df_chunk: Часть DataFrame для записи.
            data_chunk: Часть данных для записи.
            join_multivalued: Разделитель для множественных значений.
            filter_row: Номер строки для фильтров.
            bold_row: Номер строки для жирного шрифта.
        """
        ws = wb.create_sheet(sheet_name)

        add_count(
            ws,
            [
                df_chunk[col]
                .apply(lambda x: x if x not in [None, ""] else None)
                .count()
                for col in df_chunk.columns
            ],
        )

        ws.append(list(df_chunk.columns))

        add_filters(ws, len(df_chunk.columns), filter_row)
        make_bold_and_freeze(ws, bold_row)

        for entry in tqdm(data_chunk, desc=i18n.get("writing_rows_for", sheet_name)):
            ws.append(
                [
                    self.serialize_value(entry.get(col, ""), join_multivalued)
                    for col in df_chunk.columns
                ]
            )

    def serialize_value(self, value: Any, join_multivalued: str) -> Any:
        """Преобразует значение в формат, поддерживаемый Excel.

        Args:
            value: Значение для преобразования.
            join_multivalued: Разделитель для списков/кортежей.

        Returns:
            Преобразованное значение.
        """
        if isinstance(value, KNOWN_TYPES):
            return value

        if join_multivalued is not None and isinstance(value, (list, tuple)):
            try:
                return join_multivalued.join(value)
            except TypeError:
                pass

        return str(value)


converter = Converter()
