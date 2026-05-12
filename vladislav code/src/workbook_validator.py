from pathlib import Path
import warnings
import pandas as pd


class WorkbookValidator:
    """
    Проверяет структуру, заполненность и Excel-ошибки в модели кампуса.
    """

    REQUIRED_SHEETS = [
        "Предпосылки",
        "Площади",
        "График общий",
        "Графики основная деятельность",
        "Кап.затраты",
        "ФОТ и СВ",
        "Содержание",
        "Отчет о ДДС",
        "Сводная форма",
    ]

    MIN_NON_EMPTY_CELLS = {
        "Предпосылки": 20,
        "Площади": 20,
        "Кап.затраты": 20,
        "ФОТ и СВ": 20,
        "Содержание": 20,
        "Отчет о ДДС": 20,
        "Сводная форма": 5,
    }

    EXCEL_ERRORS = [
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#NULL!",
    ]

    NEGATIVE_CHECK_SHEETS = [
        "Отчет о ДДС",
        "Сводная форма",
        "Содержание",
        "Кап.затраты",
    ]

    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.errors = []
        self.warnings = []
        self.sheet_names = []

    def load_sheet_names(self):
        if not self.file_path.exists():
            self.errors.append(f"Файл не найден: {self.file_path}")
            return []

        xls = pd.ExcelFile(self.file_path)
        self.sheet_names = xls.sheet_names
        return self.sheet_names

    def validate_required_sheets(self):
        sheet_names = self.load_sheet_names()

        for sheet in self.REQUIRED_SHEETS:
            if sheet not in sheet_names:
                self.errors.append(f"Отсутствует обязательный лист: {sheet}")

        extra_sheets = [s for s in sheet_names if s not in self.REQUIRED_SHEETS]
        if extra_sheets:
            self.warnings.append(f"Дополнительные листы в модели: {extra_sheets}")

    def validate_sheet_content(self):
        for sheet, min_cells in self.MIN_NON_EMPTY_CELLS.items():
            if sheet not in self.sheet_names:
                continue

            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df = pd.read_excel(
                    self.file_path,
                    sheet_name=sheet,
                    header=None,
                )

            non_empty_cells = int(df.notna().sum().sum())

            if non_empty_cells < min_cells:
                self.errors.append(
                    f"Лист '{sheet}' выглядит незаполненным: "
                    f"{non_empty_cells} непустых ячеек"
                )
            else:
                self.warnings.append(
                    f"Лист '{sheet}' заполнен: "
                    f"{non_empty_cells} непустых ячеек"
                )

    def validate_excel_errors(self):
        for sheet in self.sheet_names:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df = pd.read_excel(
                    self.file_path,
                    sheet_name=sheet,
                    header=None,
                    dtype=str,
                )

            for error_value in self.EXCEL_ERRORS:
                count = int(df.eq(error_value).sum().sum())

                if count > 0:
                    self.errors.append(
                        f"На листе '{sheet}' найдено Excel-ошибок {error_value}: {count}"
                    )

    def validate_negative_values(self):
        for sheet in self.NEGATIVE_CHECK_SHEETS:
            if sheet not in self.sheet_names:
                continue

            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df = pd.read_excel(
                    self.file_path,
                    sheet_name=sheet,
                    header=None,
                )

            numeric_df = df.apply(pd.to_numeric, errors="coerce")
            negative_positions = []

            for row_idx, row in numeric_df.iterrows():
                for col_idx, value in row.items():
                    if pd.notna(value) and value < 0:
                        negative_positions.append(
                            {
                                "row": row_idx + 1,
                                "column": col_idx + 1,
                                "value": float(value),
                            }
                        )

            if negative_positions:
                examples = negative_positions[:10]

                self.warnings.append(
                    f"На листе '{sheet}' найдено отрицательных числовых значений: "
                    f"{len(negative_positions)}. "
                    f"Первые примеры: {examples}"
                )

    def inspect_all_sheets(self) -> dict:
        self.load_sheet_names()

        sheets_info = {}

        for sheet in self.sheet_names:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df = pd.read_excel(
                    self.file_path,
                    sheet_name=sheet,
                    header=None,
                )

            sheets_info[sheet] = {
                "rows": int(df.shape[0]),
                "columns": int(df.shape[1]),
                "non_empty_cells": int(df.notna().sum().sum()),
            }

        return sheets_info

    def validate(self) -> dict:
        self.validate_required_sheets()

        if not self.errors:
            self.validate_sheet_content()
            self.validate_excel_errors()
            self.validate_negative_values()

        return {
            "file": str(self.file_path),
            "is_valid": len(self.errors) == 0,
            "errors": self.errors,
            "warnings": self.warnings,
        }