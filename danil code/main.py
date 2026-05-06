from __future__ import annotations

import argparse
import sys
from datetime import datetime

try:
    import openpyxl  # noqa: F401
except ImportError:
    print("Установите зависимость: pip install openpyxl")
    sys.exit(1)

from src.builder import FinancialModelBuilder
from src.profiles import PROFILES, HORIZONS


def ask(prompt: str, default: str = "") -> str:
    hint = f" [{default}]" if default else ""
    val  = input(f"{prompt}{hint}: ").strip()
    return val or default


def ask_int(prompt: str, default: int) -> int:
    try:
        return int(ask(prompt, str(default)).replace(" ", "").replace(",", ""))
    except ValueError:
        return default


def ask_float(prompt: str, default: float) -> float:
    try:
        return float(ask(prompt, str(default)).replace(",", "."))
    except ValueError:
        return default


def build_cfg(args: argparse.Namespace) -> dict:
    if args.non_interactive:
        profile_key, profile_label = PROFILES.get(args.profile, PROFILES["1"])
        years        = HORIZONS.get(args.horizon, 5)
        budget_ratio = max(0.0, min(1.0, args.budget / 100))
        return {
            "name":          args.name,
            "profile":       profile_key,
            "profile_label": profile_label,
            "base_year":     datetime.now().year,
            "years":         years,
            "students":      args.students,
            "tuition":       args.tuition,
            "budget_ratio":  budget_ratio,
        }

    print("\n┌─────────────────────────────────────────┐")
    print("│  Шаблон финансовой модели университета  │")
    print("└─────────────────────────────────────────┘\n")

    name = ask("Название университета", "Государственный университет")

    print("\nПрофиль:  1 — Классический  2 — Исследовательский  3 — Коммерческий")
    profile_key, profile_label = PROFILES.get(ask("Выбор", "1"), PROFILES["1"])

    print("\nГоризонт: 1 — 3 года  2 — 5 лет  3 — 7 лет")
    years = HORIZONS.get(ask("Выбор", "2"), 5)

    students     = ask_int  ("Число студентов",               15_000)
    tuition      = ask_int  ("Стоимость обучения, руб./год", 180_000)
    budget_pct   = ask_float("Доля бюджетников, %",               45)
    budget_ratio = max(0.0, min(1.0, budget_pct / 100))

    return {
        "name":          name,
        "profile":       profile_key,
        "profile_label": profile_label,
        "base_year":     datetime.now().year,
        "years":         years,
        "students":      students,
        "tuition":       tuition,
        "budget_ratio":  budget_ratio,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Генератор шаблона финансовой модели университета"
    )
    parser.add_argument("--name",    default="Государственный университет")
    parser.add_argument("--profile", default="1", choices=["1", "2", "3"],
                        help="1=Классический 2=Исследовательский 3=Коммерческий")
    parser.add_argument("--horizon", default="2", choices=["1", "2", "3"],
                        help="1=3 года  2=5 лет  3=7 лет")
    parser.add_argument("--students", type=int,   default=15_000)
    parser.add_argument("--tuition",  type=int,   default=180_000,
                        help="Стоимость обучения, руб./год")
    parser.add_argument("--budget",   type=float, default=45.0,
                        help="Доля бюджетников, %%")
    parser.add_argument("--output",   default=None,
                        help="Имя выходного файла (по умолчанию — авто)")
    parser.add_argument("--non-interactive", action="store_true",
                        help="Использовать аргументы CLI без интерактивного ввода")
    args = parser.parse_args()

    cfg = build_cfg(args)

    print(f"\n  Генерация ({cfg['years'] + 1} лет, профиль «{cfg['profile_label']}»)...")

    filename = args.output or (
        f"FinModel_{cfg['name'].replace(' ', '_')[:20]}"
        f"_{datetime.now().strftime('%Y%m%d')}.xlsx"
    )
    FinancialModelBuilder(cfg).save(filename)

    print(f"  Готово → {filename}")
    print("  Листы: Параметры | Доходы | Расходы | P&L | Денежный поток | KPI | График\n")


if __name__ == "__main__":
    main()
