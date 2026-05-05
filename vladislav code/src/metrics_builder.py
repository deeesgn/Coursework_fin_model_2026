import pandas as pd


class MetricsBuilder:
    """
    Формирует показатели и агрегаты, которые дальше можно загрузить в шаблон финмодели.
    """

    def __init__(self, data: pd.DataFrame):
        self.data = data.copy()

    def get_top_by_market_cap(self, n: int = 10) -> pd.DataFrame:
        return (
            self.data
            .sort_values(by="Market Cap", ascending=False)
            .head(n)
        )

    def get_sector_summary(self) -> pd.DataFrame:
        return (
            self.data
            .groupby("Sector", dropna=False)
            .agg(
                companies_count=("Symbol", "count"),
                total_market_cap=("Market Cap", "sum"),
                avg_market_cap=("Market Cap", "mean"),
                avg_last_sale=("Last Sale", "mean"),
                total_volume=("Volume", "sum"),
            )
            .reset_index()
            .sort_values(by="total_market_cap", ascending=False)
        )

    def get_general_summary(self) -> pd.DataFrame:
        summary = {
            "total_companies": len(self.data),
            "avg_last_sale": self.data["Last Sale"].mean(),
            "median_last_sale": self.data["Last Sale"].median(),
            "total_market_cap": self.data["Market Cap"].sum(),
            "avg_market_cap": self.data["Market Cap"].mean(),
            "total_volume": self.data["Volume"].sum(),
            "sectors_count": self.data["Sector"].nunique(),
            "countries_count": self.data["Country"].nunique(),
        }

        return pd.DataFrame(
            [{"Metric": key, "Value": value} for key, value in summary.items()]
        )

    def build_all(self) -> dict:
        return {
            "clean_data": self.data,
            "top_companies": self.get_top_by_market_cap(),
            "sector_summary": self.get_sector_summary(),
            "general_summary": self.get_general_summary(),
        }