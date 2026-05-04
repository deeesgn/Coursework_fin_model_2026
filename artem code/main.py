from __future__ import annotations

import argparse
import html
import http.server
import json
import math
import re
import socketserver
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


DEFAULT_EXCEL_FILE = "ФМ_Юр лицо_урощен модель.xlsx"
DEFAULT_PROMPT_FILE = "prompt.txt"
DEFAULT_SHEET = "Финансовая модель"
DEFAULT_MODEL = "llama3.1"
DEFAULT_OLLAMA_URL = "http://localhost:11434/api/generate"
DEFAULT_WEB_HOST = "127.0.0.1"
DEFAULT_WEB_PORT = 8501

REVENUE_KEYWORDS = (
    "доход",
    "выруч",
    "revenue",
    "sales",
)

EXPENSE_KEYWORDS = (
    "расход",
    "себесто",
    "налог",
    "аренд",
    "cost",
    "expense",
    "cogs",
    "opex",
    "sg&a",
    "tax",
    "depreciation",
    "amortization",
    "interest",
)

KPI_KEYWORDS = (
    "чистая прибыль",
    "точка безубыточности",
    "прибыль",
    "profit",
    "income",
    "ebit",
    "ebitda",
    "ebt",
    "cash flow",
    "fcf",
    "npv",
    "irr",
    "wacc",
)


class FinancialModelError(RuntimeError):
    pass


@dataclass(frozen=True)
class YearLayout:
    header_row: int
    columns: dict[int, int]


@dataclass(frozen=True)
class SectionRow:
    group: str
    item: str
    unit: str
    values: dict[int, float]
    sources: dict[int, str]
    is_group_total: bool


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def is_year(value: Any) -> bool:
    if isinstance(value, str):
        value = value.strip()
        if not re.fullmatch(r"\d{4}", value):
            return False
    try:
        year = int(float(value))
    except (TypeError, ValueError):
        return False
    return 2000 <= year <= 2100


def to_number(value: Any) -> float | None:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, str):
        cleaned = value.replace("\xa0", "").replace(" ", "").replace(",", ".")
        try:
            return float(cleaned)
        except ValueError:
            return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def format_number(value: float) -> str:
    return f"{value:,.2f}".replace(",", " ")


def column_letter(col_idx: int) -> str:
    col_num = col_idx + 1
    letters = ""
    while col_num:
        col_num, remainder = divmod(col_num - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def cell_ref(sheet_name: str, row_idx: int, col_idx: int) -> str:
    safe_sheet = sheet_name.replace("'", "''")
    return f"'{safe_sheet}'!{column_letter(col_idx)}{row_idx + 1}"


def percent_change(start_value: float, end_value: float) -> float | None:
    if start_value == 0:
        return None
    return round((end_value - start_value) / abs(start_value) * 100, 2)


def extract_years_from_text(text: str, available_years: list[int]) -> list[int]:
    available = set(available_years)
    years = []
    for match in re.findall(r"\b(20\d{2}|21\d{2})\b", text):
        year = int(match)
        if year in available and year not in years:
            years.append(year)
    return years


class OllamaClient:
    def __init__(self, model: str = DEFAULT_MODEL, url: str = DEFAULT_OLLAMA_URL, timeout: int = 120):
        self.model = model
        self.url = url
        self.timeout = timeout

    def generate(self, prompt: str) -> str:
        payload = json.dumps(
            {"model": self.model, "prompt": prompt, "stream": False},
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            self.url,
            data=payload,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                data = json.loads(response.read().decode("utf-8"))
        except urllib.error.URLError as exc:
            raise FinancialModelError(f"Не удалось подключиться к Ollama по адресу {self.url}: {exc}") from exc
        return data.get("response", "").strip()


class FinancialModelAnalyzer:
    def __init__(self, excel_path: str | Path, sheet_name: str = DEFAULT_SHEET):
        self.excel_path = Path(excel_path).expanduser().resolve()
        self.sheet_name = sheet_name
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Excel-файл не найден: {self.excel_path}")

        self.excel_file = pd.ExcelFile(self.excel_path)
        if self.sheet_name not in self.excel_file.sheet_names:
            if self.sheet_name == DEFAULT_SHEET:
                self.sheet_name = self.choose_best_sheet()
            else:
                available = ", ".join(self.excel_file.sheet_names)
                raise FinancialModelError(f"Лист '{self.sheet_name}' не найден. Доступные листы: {available}")

        self.df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name, header=None)
        self.year_layout = self.detect_year_layout()

    def choose_best_sheet(self) -> str:
        preferred_names = (
            "финансовая модель",
            "income statement",
            "p&l",
            "profit and loss",
            "dcf",
            "cash flow",
        )
        candidates: list[tuple[int, str]] = []
        for sheet_name in self.excel_file.sheet_names:
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None)
            year_count = 0
            for row_idx in range(len(df)):
                year_count = max(year_count, sum(1 for value in df.iloc[row_idx].tolist() if is_year(value)))
            name = normalize_text(sheet_name)
            name_score = max((20 - idx for idx, word in enumerate(preferred_names) if word in name), default=0)
            candidates.append((year_count * 10 + name_score, sheet_name))

        candidates.sort(reverse=True)
        if not candidates or candidates[0][0] <= 0:
            available = ", ".join(self.excel_file.sheet_names)
            raise FinancialModelError(f"Не удалось автоматически выбрать лист. Доступные листы: {available}")
        return candidates[0][1]

    def detect_year_layout(self) -> YearLayout:
        best_row = -1
        best_columns: dict[int, int] = {}
        for row_idx in range(len(self.df)):
            columns: dict[int, int] = {}
            for col_idx, value in self.df.iloc[row_idx].items():
                if is_year(value):
                    columns[int(float(value))] = int(col_idx)
            if len(columns) > len(best_columns):
                best_row = row_idx
                best_columns = columns

        if not best_columns:
            raise FinancialModelError("Не удалось найти строку с годами в финансовой модели.")
        return YearLayout(header_row=best_row, columns=dict(sorted(best_columns.items())))

    def find_marker_row(self, marker: str, start: int = 0) -> int:
        needle = normalize_text(marker)
        for row_idx in range(start, len(self.df)):
            row_text = " ".join(normalize_text(value) for value in self.df.iloc[row_idx].dropna().tolist())
            if needle in row_text:
                return row_idx
        raise FinancialModelError(f"Не удалось найти маркер '{marker}' в листе '{self.sheet_name}'.")

    def section_rows(
        self,
        section_name: str,
        *,
        stop_before: str | None = None,
        stop_at_first_blank_after_data: bool = False,
    ) -> list[SectionRow]:
        start = self.find_marker_row(section_name)
        end = self.find_marker_row(stop_before, start + 1) if stop_before else len(self.df)
        rows: list[SectionRow] = []
        current_group = section_name

        for row_idx in range(start + 1, end):
            row = self.df.iloc[row_idx]
            col_a = row.iloc[0] if len(row) > 0 else None
            col_b = row.iloc[1] if len(row) > 1 else None
            unit = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""

            group_label = normalize_text(col_a)
            if group_label:
                current_group = str(col_a).strip()

            item = str(col_b).strip() if pd.notna(col_b) else ""
            is_group_total = False
            if not item and pd.notna(col_a):
                item = str(col_a).strip()
                is_group_total = True

            if not item:
                if stop_at_first_blank_after_data and rows:
                    break
                continue

            values: dict[int, float] = {}
            sources: dict[int, str] = {}
            for year, col_idx in self.year_layout.columns.items():
                if col_idx < len(row):
                    number = to_number(row.iloc[col_idx])
                    if number is not None:
                        values[year] = number
                        sources[year] = cell_ref(self.sheet_name, row_idx, col_idx)

            if values:
                rows.append(
                    SectionRow(
                        group=current_group,
                        item=item,
                        unit=unit,
                        values=values,
                        sources=sources,
                        is_group_total=is_group_total,
                    )
                )
            elif stop_at_first_blank_after_data and rows:
                break

        return rows

    def top_items(self, section_name: str, year: int, *, stop_before: str, limit: int) -> list[dict[str, Any]]:
        items = []
        for row in self.section_rows(section_name, stop_before=stop_before):
            value = row.values.get(year)
            if value is None:
                continue
            items.append(
                {
                    "group": row.group,
                    "item": row.item,
                    "unit": row.unit,
                    "value": round(value, 2),
                    "formatted_value": format_number(value),
                    "is_group_total": row.is_group_total,
                    "source": row.sources.get(year),
                }
            )

        total = sum(item["value"] for item in items if item["value"] > 0 and item["is_group_total"])
        if not total:
            total = sum(item["value"] for item in items if item["value"] > 0)
        for item in items:
            item["share_of_section_total_pct"] = round(item["value"] / total * 100, 2) if total else None
        return sorted(items, key=lambda item: abs(item["value"]), reverse=True)[:limit]

    def kpis(self, year: int) -> list[dict[str, Any]]:
        rows = [
            *self.section_rows("Финансовый результат", stop_before="Финановые показатели"),
            *self.section_rows("Финановые показатели", stop_at_first_blank_after_data=True),
        ]
        kpis = []
        for row in rows:
            value = row.values.get(year)
            if value is not None:
                kpis.append(
                    {
                        "name": row.item,
                        "unit": row.unit,
                        "value": round(value, 2),
                        "formatted_value": format_number(value),
                        "source": row.sources.get(year),
                    }
                )
        return kpis

    def build_context(
        self,
        user_question: str,
        year: int | None = None,
        compare_year: int | None = None,
    ) -> dict[str, Any]:
        available_years = list(self.year_layout.columns)
        mentioned_years = extract_years_from_text(user_question, available_years)
        selected_year = year or (mentioned_years[-1] if mentioned_years else min(self.year_layout.columns))
        if selected_year not in self.year_layout.columns:
            available = ", ".join(str(year) for year in self.year_layout.columns)
            raise FinancialModelError(f"Год {selected_year} не найден. Доступные годы: {available}")

        base_year = None
        target_year = None
        if compare_year is not None:
            base_year = selected_year
            target_year = compare_year
        elif len(mentioned_years) >= 2:
            base_year = mentioned_years[0]
            target_year = mentioned_years[1]
            selected_year = target_year

        if target_year is not None and target_year not in self.year_layout.columns:
            available = ", ".join(str(year) for year in self.year_layout.columns)
            raise FinancialModelError(f"Год {target_year} не найден. Доступные годы: {available}")

        try:
            top_expenses = self.top_items("Расходы", selected_year, stop_before="Финансовый результат", limit=12)
            top_revenue_items = self.top_items("Доходы", selected_year, stop_before="Расходы", limit=8)
            kpis = self.kpis(selected_year)
            extraction_mode = "structured_russian_financial_model"
        except FinancialModelError:
            generic = self.generic_context(selected_year)
            top_expenses = generic["top_expenses"]
            top_revenue_items = generic["top_revenue_items"]
            kpis = generic["kpis"]
            extraction_mode = "generic_year_table_fallback"

        return {
            "workbook": self.excel_path.name,
            "sheet": self.sheet_name,
            "extraction_mode": extraction_mode,
            "available_years": list(self.year_layout.columns),
            "selected_year": selected_year,
            "user_question": user_question,
            "comparison": self.compare_years(base_year, target_year) if base_year and target_year else None,
            "top_expenses": top_expenses,
            "top_revenue_items": top_revenue_items,
            "kpis": kpis,
        }

    def compare_years(self, base_year: int, target_year: int) -> dict[str, Any]:
        try:
            return {
                "base_year": base_year,
                "target_year": target_year,
                "mode": "structured_russian_financial_model",
                "kpis": self.compare_section_metrics(
                    self.kpis_rows(),
                    base_year,
                    target_year,
                    limit=12,
                ),
                "expense_changes": self.compare_section_metrics(
                    self.section_rows("Расходы", stop_before="Финансовый результат"),
                    base_year,
                    target_year,
                    limit=12,
                ),
                "revenue_changes": self.compare_section_metrics(
                    self.section_rows("Доходы", stop_before="Расходы"),
                    base_year,
                    target_year,
                    limit=8,
                ),
            }
        except FinancialModelError:
            rows = self.generic_rows_for_all_years()
            return {
                "base_year": base_year,
                "target_year": target_year,
                "mode": "generic_year_table_fallback",
                "kpis": self.compare_generic_metrics(
                    [row for row in rows if contains_any(row["name"], KPI_KEYWORDS)],
                    base_year,
                    target_year,
                    limit=12,
                ),
                "expense_changes": self.compare_generic_metrics(
                    [row for row in rows if contains_any(row["name"], EXPENSE_KEYWORDS)],
                    base_year,
                    target_year,
                    limit=12,
                ),
                "revenue_changes": self.compare_generic_metrics(
                    [row for row in rows if contains_any(row["name"], REVENUE_KEYWORDS)],
                    base_year,
                    target_year,
                    limit=8,
                ),
            }

    def kpis_rows(self) -> list[SectionRow]:
        return [
            *self.section_rows("Финансовый результат", stop_before="Финановые показатели"),
            *self.section_rows("Финановые показатели", stop_at_first_blank_after_data=True),
        ]

    def compare_section_metrics(
        self,
        rows: list[SectionRow],
        base_year: int,
        target_year: int,
        *,
        limit: int,
    ) -> list[dict[str, Any]]:
        changes = []
        for row in rows:
            start_value = row.values.get(base_year)
            end_value = row.values.get(target_year)
            if start_value is None or end_value is None:
                continue
            delta = end_value - start_value
            changes.append(
                {
                    "group": row.group,
                    "name": row.item,
                    "unit": row.unit,
                    "base_value": round(start_value, 2),
                    "base_formatted_value": format_number(start_value),
                    "base_source": row.sources.get(base_year),
                    "target_value": round(end_value, 2),
                    "target_formatted_value": format_number(end_value),
                    "target_source": row.sources.get(target_year),
                    "delta_abs": round(delta, 2),
                    "delta_abs_formatted": format_number(delta),
                    "delta_pct": percent_change(start_value, end_value),
                    "is_group_total": row.is_group_total,
                }
            )
        return sorted(changes, key=lambda item: abs(item["delta_abs"]), reverse=True)[:limit]

    def compare_generic_metrics(
        self,
        rows: list[dict[str, Any]],
        base_year: int,
        target_year: int,
        *,
        limit: int,
    ) -> list[dict[str, Any]]:
        changes = []
        for row in rows:
            values = row["values"]
            if base_year not in values or target_year not in values:
                continue
            start_value = values[base_year]
            end_value = values[target_year]
            delta = end_value - start_value
            changes.append(
                {
                    "name": row["name"],
                    "unit": "",
                    "base_value": round(start_value, 2),
                    "base_formatted_value": format_number(start_value),
                    "base_source": row["sources"].get(base_year),
                    "target_value": round(end_value, 2),
                    "target_formatted_value": format_number(end_value),
                    "target_source": row["sources"].get(target_year),
                    "delta_abs": round(delta, 2),
                    "delta_abs_formatted": format_number(delta),
                    "delta_pct": percent_change(start_value, end_value),
                }
            )
        return sorted(changes, key=lambda item: abs(item["delta_abs"]), reverse=True)[:limit]

    def generic_context(self, year: int) -> dict[str, list[dict[str, Any]]]:
        rows = self.generic_year_rows(year)
        revenue_rows = [row for row in rows if contains_any(row["name"], REVENUE_KEYWORDS)]
        expense_rows = [row for row in rows if contains_any(row["name"], EXPENSE_KEYWORDS)]
        kpi_rows = [row for row in rows if contains_any(row["name"], KPI_KEYWORDS)]

        return {
            "top_revenue_items": sorted(revenue_rows, key=lambda row: abs(row["value"]), reverse=True)[:8],
            "top_expenses": sorted(expense_rows, key=lambda row: abs(row["value"]), reverse=True)[:12],
            "kpis": sorted(kpi_rows, key=lambda row: row["name"])[:12],
        }

    def generic_year_rows(self, year: int) -> list[dict[str, Any]]:
        year_col = self.year_layout.columns[year]
        rows = []
        seen: set[tuple[str, float]] = set()
        for row_idx in range(self.year_layout.header_row + 1, len(self.df)):
            row = self.df.iloc[row_idx]
            value = to_number(row.iloc[year_col]) if year_col < len(row) else None
            if value is None:
                continue

            label = ""
            for col_idx in range(0, year_col):
                cell = row.iloc[col_idx]
                if pd.notna(cell) and not is_year(cell) and not to_number(cell):
                    label = str(cell).strip()
            if not label:
                continue

            normalized = normalize_text(label)
            if normalized in {"% growth", "% of gross sales", "% of beverage sales"}:
                continue
            if normalized.startswith("%"):
                continue

            key = (normalized, round(value, 6))
            if key in seen:
                continue
            seen.add(key)
            rows.append(
                {
                    "name": label,
                    "unit": "",
                    "value": round(value, 2),
                    "formatted_value": format_number(value),
                    "source": cell_ref(self.sheet_name, row_idx, year_col),
                }
            )
        return rows

    def generic_rows_for_all_years(self) -> list[dict[str, Any]]:
        rows = []
        seen: set[str] = set()
        for row_idx in range(self.year_layout.header_row + 1, len(self.df)):
            row = self.df.iloc[row_idx]
            label = ""
            for col_idx in range(0, min(self.year_layout.columns.values())):
                cell = row.iloc[col_idx]
                if pd.notna(cell) and not is_year(cell) and not to_number(cell):
                    label = str(cell).strip()
            if not label:
                continue
            normalized = normalize_text(label)
            if normalized in {"% growth", "% of gross sales", "% of beverage sales"} or normalized.startswith("%"):
                continue
            if normalized in seen:
                continue
            seen.add(normalized)

            values = {}
            sources = {}
            for year, col_idx in self.year_layout.columns.items():
                if col_idx < len(row):
                    value = to_number(row.iloc[col_idx])
                    if value is not None:
                        values[year] = value
                        sources[year] = cell_ref(self.sheet_name, row_idx, col_idx)
            if values:
                rows.append({"name": label, "values": values, "sources": sources})
        return rows


def contains_any(text: str, keywords: tuple[str, ...]) -> bool:
    normalized = normalize_text(text)
    return any(keyword in normalized for keyword in keywords)


def comparison_summary(comparison: dict[str, Any]) -> str:
    base_year = comparison["base_year"]
    target_year = comparison["target_year"]
    lines = [f"Сравнение {base_year} и {target_year} годов:"]

    if comparison.get("kpis"):
        lines.append("KPI:")
        for item in comparison["kpis"][:5]:
            unit = f" {item['unit']}" if item.get("unit") else ""
            pct = "н/д" if item["delta_pct"] is None else f"{item['delta_pct']}%"
            lines.append(
                "- "
                f"{item['name']}: {base_year} = {item['base_formatted_value']}{unit} "
                f"({item['base_source']}), {target_year} = {item['target_formatted_value']}{unit} "
                f"({item['target_source']}), изменение = {item['delta_abs_formatted']}{unit}, {pct}."
            )

    if comparison.get("expense_changes"):
        lines.append("Крупнейшие изменения расходов:")
        for item in comparison["expense_changes"][:5]:
            unit = f" {item['unit']}" if item.get("unit") else ""
            pct = "н/д" if item["delta_pct"] is None else f"{item['delta_pct']}%"
            lines.append(
                "- "
                f"{item['name']}: {base_year} = {item['base_formatted_value']}{unit} "
                f"({item['base_source']}), {target_year} = {item['target_formatted_value']}{unit} "
                f"({item['target_source']}), изменение = {item['delta_abs_formatted']}{unit}, {pct}."
            )

    return "\n".join(lines)


def read_user_prompt(path: str | Path) -> str:
    prompt_path = Path(path)
    if not prompt_path.exists():
        raise FileNotFoundError(f"Файл с prompt не найден: {prompt_path}")
    prompt = prompt_path.read_text(encoding="utf-8").strip()
    if not prompt:
        raise FinancialModelError(f"Файл с prompt пустой: {prompt_path}")
    return prompt


def build_llm_prompt(context: dict[str, Any]) -> str:
    if context.get("comparison"):
        prompt_context = {
            "workbook": context["workbook"],
            "sheet": context["sheet"],
            "extraction_mode": context["extraction_mode"],
            "user_question": context["user_question"],
            "computed_summary": comparison_summary(context["comparison"]),
        }
    else:
        prompt_context = context

    facts = json.dumps(prompt_context, ensure_ascii=False, indent=2)
    return f"""
Ты финансовый аналитик и отвечаешь только по данным из Excel-финмодели.

Правила:
- используй только факты из блока "Данные";
- если в данных есть поле "computed_summary", используй его как готовую рассчитанную сводку; не меняй числа, годы, названия строк и источники;
- не придумывай числа, причины, сравнения, активы, рынок или управленческие выводы;
- не переводи и не переименовывай названия строк из Excel; цитируй labels ровно как в данных;
- если используешь число, по возможности укажи его источник из поля "source", "base_source" или "target_source";
- если delta_pct равен null, напиши, что процентное изменение не рассчитано из-за нулевой базы;
- если для ответа не хватает данных, прямо напиши, чего именно не хватает;
- отвечай кратко, по-русски, в структурированном виде;
- все суммы уже указаны в единицах из поля "unit".

Данные:
{facts}

Вопрос пользователя:
{context["user_question"]}
""".strip()


def analyze_question(
    *,
    excel_path: str,
    question: str,
    sheet: str = DEFAULT_SHEET,
    year: int | None = None,
    compare_year: int | None = None,
    model: str = DEFAULT_MODEL,
    ollama_url: str = DEFAULT_OLLAMA_URL,
    no_llm: bool = False,
) -> dict[str, Any]:
    analyzer = FinancialModelAnalyzer(excel_path, sheet_name=sheet)
    context = analyzer.build_context(question, year=year, compare_year=compare_year)
    if no_llm:
        return {"answer": "", "context": context}

    llm_prompt = build_llm_prompt(context)
    answer = OllamaClient(model=model, url=ollama_url).generate(llm_prompt)
    return {"answer": answer, "context": context}


def collect_sources(value: Any) -> list[str]:
    sources = []
    if isinstance(value, dict):
        for key, item in value.items():
            if key in {"source", "base_source", "target_source"} and item:
                sources.append(str(item))
            else:
                sources.extend(collect_sources(item))
    elif isinstance(value, list):
        for item in value:
            sources.extend(collect_sources(item))
    return sorted(set(sources))


WEB_PAGE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Financial Model Copilot</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f7f9;
      --panel: #ffffff;
      --text: #1d2433;
      --muted: #667085;
      --line: #d9dee8;
      --accent: #2563eb;
      --accent-dark: #1d4ed8;
      --danger: #b42318;
      --soft: #eef4ff;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: var(--bg);
      color: var(--text);
    }
    main {
      max-width: 1180px;
      margin: 0 auto;
      padding: 28px;
    }
    header {
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      gap: 20px;
      margin-bottom: 20px;
    }
    h1 {
      margin: 0;
      font-size: 28px;
      line-height: 1.15;
    }
    .subtitle {
      margin-top: 8px;
      color: var(--muted);
      font-size: 14px;
    }
    .layout {
      display: grid;
      grid-template-columns: 380px minmax(0, 1fr);
      gap: 18px;
      align-items: start;
    }
    section {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
    }
    label {
      display: block;
      font-size: 13px;
      font-weight: 650;
      margin: 14px 0 6px;
    }
    label:first-child { margin-top: 0; }
    input, textarea, select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 10px 11px;
      font: inherit;
      color: var(--text);
      background: #fff;
    }
    textarea {
      min-height: 126px;
      resize: vertical;
    }
    .row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
    }
    button {
      margin-top: 16px;
      width: 100%;
      border: 0;
      border-radius: 6px;
      background: var(--accent);
      color: white;
      padding: 11px 14px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
    }
    button:hover { background: var(--accent-dark); }
    button:disabled { opacity: .65; cursor: wait; }
    .answer {
      min-height: 260px;
      white-space: pre-wrap;
      line-height: 1.55;
      font-size: 15px;
    }
    .empty {
      color: var(--muted);
    }
    .meta {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
      margin-bottom: 14px;
    }
    .pill {
      background: var(--soft);
      border: 1px solid #c7d7fe;
      border-radius: 6px;
      padding: 9px;
      font-size: 13px;
      overflow-wrap: anywhere;
    }
    .pill b {
      display: block;
      margin-bottom: 3px;
      color: #1e3a8a;
    }
    details {
      margin-top: 14px;
      border-top: 1px solid var(--line);
      padding-top: 12px;
    }
    summary {
      cursor: pointer;
      color: var(--accent-dark);
      font-weight: 700;
    }
    pre {
      overflow: auto;
      background: #111827;
      color: #e5e7eb;
      border-radius: 6px;
      padding: 12px;
      font-size: 12px;
      line-height: 1.45;
      max-height: 360px;
    }
    .sources {
      margin-top: 14px;
      color: var(--muted);
      font-size: 13px;
    }
    .sources code {
      display: inline-block;
      margin: 4px 4px 0 0;
      background: #f2f4f7;
      border: 1px solid var(--line);
      border-radius: 5px;
      padding: 4px 6px;
      color: #344054;
    }
    .error {
      color: var(--danger);
      font-weight: 650;
    }
    @media (max-width: 860px) {
      main { padding: 18px; }
      .layout { grid-template-columns: 1fr; }
      .meta { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <h1>Financial Model Copilot</h1>
        <div class="subtitle">Локальный помощник по Excel-финмодели: Python считает, Ollama объясняет.</div>
      </div>
    </header>

    <div class="layout">
      <section>
        <label for="excel">Путь до Excel-файла</label>
        <input id="excel" value="__DEFAULT_EXCEL__">

        <label for="sheet">Лист</label>
        <input id="sheet" value="__DEFAULT_SHEET__">

        <div class="row">
          <div>
            <label for="year">Год</label>
            <input id="year" type="number" value="2025">
          </div>
          <div>
            <label for="compareYear">Сравнить с</label>
            <input id="compareYear" type="number" placeholder="2035">
          </div>
        </div>

        <label for="model">Ollama model</label>
        <input id="model" value="__DEFAULT_MODEL__">

        <label for="question">Вопрос</label>
        <textarea id="question">Какие основные расходы и KPI в 2025 году? Дай короткий аналитический вывод.</textarea>

        <button id="ask">Спросить</button>
      </section>

      <section>
        <div class="meta" id="meta"></div>
        <div class="answer empty" id="answer">Ответ появится здесь.</div>
        <div class="sources" id="sources"></div>
        <details>
          <summary>Показать структурированный контекст</summary>
          <pre id="context">{}</pre>
        </details>
      </section>
    </div>
  </main>

  <script>
    const askButton = document.getElementById("ask");
    const answer = document.getElementById("answer");
    const contextBox = document.getElementById("context");
    const meta = document.getElementById("meta");
    const sources = document.getElementById("sources");

    function value(id) {
      return document.getElementById(id).value.trim();
    }

    function numberOrNull(id) {
      const raw = value(id);
      return raw ? Number(raw) : null;
    }

    function escapeHtml(text) {
      return String(text).replace(/[&<>"']/g, ch => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#039;"
      }[ch]));
    }

    async function ask() {
      askButton.disabled = true;
      answer.className = "answer empty";
      answer.textContent = "Думаю...";
      contextBox.textContent = "{}";
      meta.innerHTML = "";
      sources.innerHTML = "";

      const payload = {
        excel: value("excel"),
        sheet: value("sheet"),
        year: numberOrNull("year"),
        compare_year: numberOrNull("compareYear"),
        model: value("model"),
        question: value("question")
      };

      try {
        const response = await fetch("/api/analyze", {
          method: "POST",
          headers: {"Content-Type": "application/json"},
          body: JSON.stringify(payload)
        });
        const data = await response.json();
        if (!response.ok || data.error) {
          throw new Error(data.error || "Ошибка запроса");
        }

        answer.className = "answer";
        answer.textContent = data.answer || "Ответ пустой.";
        contextBox.textContent = JSON.stringify(data.context, null, 2);

        const c = data.context || {};
        meta.innerHTML = `
          <div class="pill"><b>Файл</b>${escapeHtml(c.workbook || "")}</div>
          <div class="pill"><b>Лист</b>${escapeHtml(c.sheet || "")}</div>
          <div class="pill"><b>Режим</b>${escapeHtml(c.extraction_mode || "")}</div>
        `;

        if (data.sources && data.sources.length) {
          sources.innerHTML = "<b>Источники:</b><br>" + data.sources
            .slice(0, 24)
            .map(src => `<code>${escapeHtml(src)}</code>`)
            .join("");
        }
      } catch (err) {
        answer.className = "answer error";
        answer.textContent = err.message;
      } finally {
        askButton.disabled = false;
      }
    }

    askButton.addEventListener("click", ask);
  </script>
</body>
</html>
"""


class CopilotRequestHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        if self.path not in {"/", "/index.html"}:
            self.send_json({"error": "Not found"}, status=404)
            return

        page = (
            WEB_PAGE.replace("__DEFAULT_EXCEL__", html.escape(DEFAULT_EXCEL_FILE))
            .replace("__DEFAULT_SHEET__", html.escape(DEFAULT_SHEET))
            .replace("__DEFAULT_MODEL__", html.escape(DEFAULT_MODEL))
        )
        body = page.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self) -> None:
        if self.path != "/api/analyze":
            self.send_json({"error": "Not found"}, status=404)
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
            question = str(payload.get("question", "")).strip()
            if not question:
                raise FinancialModelError("Вопрос пустой.")

            result = analyze_question(
                excel_path=str(payload.get("excel") or DEFAULT_EXCEL_FILE),
                question=question,
                sheet=str(payload.get("sheet") or DEFAULT_SHEET),
                year=payload.get("year"),
                compare_year=payload.get("compare_year"),
                model=str(payload.get("model") or DEFAULT_MODEL),
                ollama_url=str(payload.get("ollama_url") or DEFAULT_OLLAMA_URL),
            )
            result["sources"] = collect_sources(result["context"])
            self.send_json(result)
        except Exception as exc:
            self.send_json({"error": str(exc)}, status=400)

    def log_message(self, format: str, *args: Any) -> None:
        return

    def send_json(self, payload: dict[str, Any], status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def serve(host: str = DEFAULT_WEB_HOST, port: int = DEFAULT_WEB_PORT) -> None:
    with socketserver.ThreadingTCPServer((host, port), CopilotRequestHandler) as httpd:
        httpd.allow_reuse_address = True
        print(f"Financial Model Copilot: http://{host}:{port}")
        httpd.serve_forever()


def main() -> None:
    parser = argparse.ArgumentParser(description="Мини-backend для анализа Excel-финансовой модели через Ollama.")
    parser.add_argument("--excel", default=DEFAULT_EXCEL_FILE, help="Путь до Excel-финмодели")
    parser.add_argument("--prompt", default=DEFAULT_PROMPT_FILE, help="Путь до файла с вопросом пользователя")
    parser.add_argument("--question", default=None, help="Вопрос пользователя. Если указан, файл --prompt не читается")
    parser.add_argument("--year", type=int, default=None, help="Год для анализа")
    parser.add_argument("--compare-year", type=int, default=None, help="Второй год для сравнения с --year")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="Название листа с финансовой моделью")
    parser.add_argument("--model", default=DEFAULT_MODEL, help="Название модели в Ollama")
    parser.add_argument("--ollama-url", default=DEFAULT_OLLAMA_URL, help="URL Ollama generate API")
    parser.add_argument("--no-llm", action="store_true", help="Показать структурированные данные без запроса к Ollama")
    parser.add_argument("--serve", action="store_true", help="Запустить мини-интерфейс в браузере")
    parser.add_argument("--host", default=DEFAULT_WEB_HOST, help="Host для мини-интерфейса")
    parser.add_argument("--port", type=int, default=DEFAULT_WEB_PORT, help="Port для мини-интерфейса")
    args = parser.parse_args()

    if args.serve:
        serve(args.host, args.port)
        return

    user_question = args.question.strip() if args.question else read_user_prompt(args.prompt)
    if not user_question:
        raise FinancialModelError("Вопрос пользователя пустой.")
    result = analyze_question(
        excel_path=args.excel,
        question=user_question,
        sheet=args.sheet,
        year=args.year,
        compare_year=args.compare_year,
        model=args.model,
        ollama_url=args.ollama_url,
        no_llm=args.no_llm,
    )

    if args.no_llm:
        print(json.dumps(result["context"], ensure_ascii=False, indent=2))
        return

    print(result["answer"])


if __name__ == "__main__":
    main()
