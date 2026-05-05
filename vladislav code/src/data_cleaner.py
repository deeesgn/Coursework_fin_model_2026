import pandas as pd


class DataCleaner:
    """
    Очищает данные и приводит поля к формату, пригодному для финансовой модели.
    """

    def __init__(self, data: pd.DataFrame):
        self.data = data.copy()

    def select_model_columns(self):
        self.data = self.data[
            [
                "Symbol",
                "Name",
                "Last Sale",
                "Market Cap",
                "Country",
                "IPO Year",
                "Volume",
                "Sector",
                "Industry",
            ]
        ]
        return self

    def clean_text_fields(self):
        text_columns = ["Symbol", "Name", "Country", "Sector", "Industry"]

        for column in text_columns:
            self.data[column] = (
                self.data[column]
                .astype(str)
                .str.strip()
                .replace({"nan": None, "": None})
            )

        return self

    def clean_numeric_fields(self):
        self.data["Last Sale"] = (
            self.data["Last Sale"]
            .astype(str)
            .str.replace("$", "", regex=False)
            .str.replace(",", "", regex=False)
        )
        self.data["Last Sale"] = pd.to_numeric(self.data["Last Sale"], errors="coerce")

        self.data["Market Cap"] = pd.to_numeric(
            self.data["Market Cap"], errors="coerce"
        )

        self.data["IPO Year"] = pd.to_numeric(
            self.data["IPO Year"], errors="coerce"
        )

        self.data["Volume"] = pd.to_numeric(
            self.data["Volume"], errors="coerce"
        )

        return self

    def remove_invalid_rows(self):
        self.data = self.data.dropna(subset=["Symbol", "Name", "Last Sale"])
        return self

    def clean(self) -> pd.DataFrame:
        return (
            self.select_model_columns()
            .clean_text_fields()
            .clean_numeric_fields()
            .remove_invalid_rows()
            .data
        )