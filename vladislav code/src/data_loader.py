from pathlib import Path
import pandas as pd


class DataLoader:
    """
    Модуль загрузки исходных данных для финансовой модели.
    Поддерживает CSV и Excel-файлы.
    """

    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.data = None

    def load(self) -> pd.DataFrame:
        """
        Загружает данные из файла и возвращает DataFrame.
        """

        if not self.file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {self.file_path}")

        if self.file_path.suffix.lower() == ".csv":
            self.data = pd.read_csv(self.file_path)

        elif self.file_path.suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
            self.data = pd.read_excel(self.file_path)

        else:
            raise ValueError(
                "Неподдерживаемый формат файла. Используйте CSV или Excel."
            )

        return self.data

    def get_basic_info(self) -> dict:
        """
        Возвращает базовую информацию о загруженных данных.
        """

        if self.data is None:
            raise ValueError("Сначала нужно загрузить данные через метод load().")

        return {
            "rows": len(self.data),
            "columns": len(self.data.columns),
            "column_names": list(self.data.columns),
            "missing_values": self.data.isna().sum().to_dict(),
        }


if __name__ == "__main__":
    loader = DataLoader("data/nasdaq_screener.csv")

    df = loader.load()
    info = loader.get_basic_info()

    print("Данные успешно загружены")
    print(df.head())
    print("\nИнформация о данных:")
    print(info)