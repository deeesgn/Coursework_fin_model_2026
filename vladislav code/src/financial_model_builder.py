import pandas as pd


class FinancialModelBuilder:
    """
    Строит упрощённую финансовую модель по компаниям:
    Inputs -> Calculations -> DCF -> Summary.
    """

    def __init__(
        self,
        data: pd.DataFrame,
        forecast_years: int = 5,
        revenue_to_market_cap: float = 0.35,
        revenue_growth: float = 0.05,
        ebitda_margin: float = 0.25,
        da_percent_of_revenue: float = 0.04,
        capex_percent_of_revenue: float = 0.06,
        tax_rate: float = 0.25,
        wacc: float = 0.12,
        terminal_growth: float = 0.03,
    ):
        self.data = data.copy()
        self.forecast_years = forecast_years
        self.revenue_to_market_cap = revenue_to_market_cap
        self.revenue_growth = revenue_growth
        self.ebitda_margin = ebitda_margin
        self.da_percent_of_revenue = da_percent_of_revenue
        self.capex_percent_of_revenue = capex_percent_of_revenue
        self.tax_rate = tax_rate
        self.wacc = wacc
        self.terminal_growth = terminal_growth

    def build_inputs(self) -> pd.DataFrame:
        return pd.DataFrame(
            {
                "Parameter": [
                    "Forecast years",
                    "Revenue / Market Cap",
                    "Revenue growth",
                    "EBITDA margin",
                    "D&A / Revenue",
                    "CAPEX / Revenue",
                    "Tax rate",
                    "WACC",
                    "Terminal growth",
                ],
                "Value": [
                    self.forecast_years,
                    self.revenue_to_market_cap,
                    self.revenue_growth,
                    self.ebitda_margin,
                    self.da_percent_of_revenue,
                    self.capex_percent_of_revenue,
                    self.tax_rate,
                    self.wacc,
                    self.terminal_growth,
                ],
            }
        )

    def build_calculations(self, top_n: int = 10) -> pd.DataFrame:
        companies = (
            self.data.sort_values(by="Market Cap", ascending=False)
            .head(top_n)
            .copy()
        )

        rows = []

        for _, company in companies.iterrows():
            symbol = company["Symbol"]
            name = company["Name"]
            market_cap = company["Market Cap"]

            base_revenue = market_cap * self.revenue_to_market_cap

            for year in range(1, self.forecast_years + 1):
                revenue = base_revenue * ((1 + self.revenue_growth) ** (year - 1))
                ebitda = revenue * self.ebitda_margin
                da = revenue * self.da_percent_of_revenue
                ebit = ebitda - da
                tax = ebit * self.tax_rate
                capex = revenue * self.capex_percent_of_revenue
                fcff = ebit * (1 - self.tax_rate) + da - capex

                rows.append(
                    {
                        "Symbol": symbol,
                        "Name": name,
                        "Year": year,
                        "Market Cap": market_cap,
                        "Revenue": revenue,
                        "EBITDA": ebitda,
                        "D&A": da,
                        "EBIT": ebit,
                        "Tax": tax,
                        "CAPEX": capex,
                        "FCFF": fcff,
                    }
                )

        return pd.DataFrame(rows)

    def build_dcf(self, calculations: pd.DataFrame) -> pd.DataFrame:
        rows = []

        for symbol, group in calculations.groupby("Symbol"):
            group = group.sort_values("Year")
            name = group["Name"].iloc[0]

            total_dcf = 0

            for _, row in group.iterrows():
                year = int(row["Year"])
                fcff = row["FCFF"]
                discount_factor = 1 / ((1 + self.wacc) ** year)
                discounted_fcff = fcff * discount_factor
                total_dcf += discounted_fcff

                rows.append(
                    {
                        "Symbol": symbol, 
                        "Name": name,
                        "Year": year,
                        "FCFF": fcff,
                        "Discount factor": discount_factor,
                        "Discounted FCFF": discounted_fcff,
                        "Terminal Value": None,
                        "Enterprise Value": None,
                    }
                )

            last_fcff = group["FCFF"].iloc[-1]
            terminal_value = (
                last_fcff * (1 + self.terminal_growth)
                / (self.wacc - self.terminal_growth)
            )
            discounted_terminal_value = terminal_value / (
                (1 + self.wacc) ** self.forecast_years
            )

            enterprise_value = total_dcf + discounted_terminal_value

            rows.append(
                {
                    "Symbol": symbol,
                    "Name": name,
                    "Year": "Terminal",
                    "FCFF": None,
                    "Discount factor": None,
                    "Discounted FCFF": discounted_terminal_value,
                    "Terminal Value": terminal_value,
                    "Enterprise Value": enterprise_value,
                }
            )

        return pd.DataFrame(rows)

    def build_summary(self, dcf: pd.DataFrame) -> pd.DataFrame:
        summary = dcf[dcf["Enterprise Value"].notna()][
            ["Symbol", "Name", "Enterprise Value"]
        ].copy()

        return summary.sort_values(by="Enterprise Value", ascending=False)

    def build_model(self, top_n: int = 10) -> dict:
        inputs = self.build_inputs()
        calculations = self.build_calculations(top_n=top_n)
        dcf = self.build_dcf(calculations)
        summary = self.build_summary(dcf)

        return {
            "inputs": inputs,
            "calculations": calculations,
            "dcf": dcf,
            "financial_summary": summary,
        }