import pandas as pd


class DataValidator:
    """
    Проверяет, что данные подходят для дальнейшей загрузки в финансовую модель.
    """

    REQUIRED_COLUMNS = [
        "Symbol",
        "Name",
        "Last Sale",
        "Market Cap",
        "Sector",
        "Industry",
    ]

    def __init__(self, data: pd.DataFrame):
        self.data = data
        self.errors = []
        self.warnings = []

    def check_required_columns(self):
        for column in self.REQUIRED_COLUMNS:
            if column not in self.data.columns:
                self.errors.append(f"Отсутствует обязательная колонка: {column}")

    def check_empty_dataset(self):
        if self.data.empty:
            self.errors.append("Файл загружен, но таблица пустая.")

    def check_missing_values(self):
        for column in self.REQUIRED_COLUMNS:
            if column in self.data.columns:
                missing_count = self.data[column].isna().sum()
                if missing_count > 0:
                    self.warnings.append(
                        f"В колонке {column} есть пропуски: {missing_count}"
                    )

    def check_duplicates(self):
        if "Symbol" in self.data.columns:
            duplicated_count = self.data["Symbol"].duplicated().sum()
            if duplicated_count > 0:
                self.warnings.append(
                    f"Найдены повторяющиеся тикеры: {duplicated_count}"
                )

    def validate(self) -> dict:
        self.check_empty_dataset()
        self.check_required_columns()
        self.check_missing_values()
        self.check_duplicates()

        return {
            "is_valid": len(self.errors) == 0,
            "errors": self.errors,
            "warnings": self.warnings,
        }