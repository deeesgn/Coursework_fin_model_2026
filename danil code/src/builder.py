from openpyxl import Workbook

from src.sheets import params, revenue, expenses, pnl, cashflow, kpi, chart


class FinancialModelBuilder:
    def __init__(self, cfg: dict) -> None:
        self.cfg = cfg

    def build(self) -> Workbook:
        wb = Workbook()
        wb.remove(wb.active)

        params.build(wb, self.cfg)
        revenue.build(wb, self.cfg)
        expenses.build(wb, self.cfg)
        pnl.build(wb, self.cfg)
        cashflow.build(wb, self.cfg)
        kpi.build(wb, self.cfg)
        chart.build(wb, self.cfg)

        return wb

    def save(self, path: str) -> None:
        self.build().save(path)
