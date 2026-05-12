from pathlib import Path
import warnings
import pandas as pd


class CheckedModelValidator:
    """
    Проверяет проверенную сводную модель:
    наличие листов, заполненность и Excel-ошибки.
    """

    REQUIRED_SHEETS = [
        "Паспорт",
        "ФинРез и инд-ры ВУЗ",
        "ФинРез Кампус",
        "Анализ чувствительности ",
        "Диаграмма эластичности",
    ]

    EXCEL_ERRORS = [
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#NULL!",
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

    def validate_sheet_content(self):
        for sheet in self.REQUIRED_SHEETS:
            if sheet not in self.sheet_names:
                continue

            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df = pd.read_excel(self.file_path, sheet_name=sheet, header=None)

            non_empty_cells = int(df.notna().sum().sum())

            if non_empty_cells == 0:
                self.errors.append(f"Лист '{sheet}' пустой")
            else:
                self.warnings.append(
                    f"Лист '{sheet}' заполнен: {non_empty_cells} непустых ячеек"
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

    def validate(self) -> dict:
        self.validate_required_sheets()

        if not self.errors:
            self.validate_sheet_content()
            self.validate_excel_errors()

        return {
            "file": str(self.file_path),
            "is_valid": len(self.errors) == 0,
            "errors": self.errors,
            "warnings": self.warnings,
        }