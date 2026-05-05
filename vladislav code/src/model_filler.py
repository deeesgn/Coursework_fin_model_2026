import pandas as pd
from pathlib import Path


class ModelFiller:
    """
    Заполняет Excel-файл подготовленными данными, агрегатами и финансовой моделью.
    """

    def __init__(self, output_path: str = "output/filled_financial_model.xlsx"):
        self.output_path = Path(output_path)
        self.output_path.parent.mkdir(parents=True, exist_ok=True)

    def fill(self, model_data: dict):
        with pd.ExcelWriter(self.output_path, engine="openpyxl") as writer:
            model_data["clean_data"].to_excel(writer, sheet_name="Clean Data", index=False)
            model_data["top_companies"].to_excel(writer, sheet_name="Top Companies", index=False)
            model_data["sector_summary"].to_excel(writer, sheet_name="Sector Summary", index=False)
            model_data["general_summary"].to_excel(writer, sheet_name="General Summary", index=False)

            if "inputs" in model_data:
                model_data["inputs"].to_excel(writer, sheet_name="Inputs", index=False)

            if "calculations" in model_data:
                model_data["calculations"].to_excel(writer, sheet_name="Calculations", index=False)

            if "dcf" in model_data:
                model_data["dcf"].to_excel(writer, sheet_name="DCF", index=False)

            if "financial_summary" in model_data:
                model_data["financial_summary"].to_excel(writer, sheet_name="Financial Summary", index=False)

        return self.output_path