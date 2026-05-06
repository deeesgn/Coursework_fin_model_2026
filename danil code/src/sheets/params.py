from src.profiles import DA_RATIO, CAPEX_RATIO
from src.styles import BG, NUM, make_fill, make_font, make_border, make_align, set_col_widths


def build(wb, cfg: dict) -> None:
    ws = wb.create_sheet("Параметры")
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 30, "B": 28})

    ws.merge_cells("A1:B1")
    ws["A1"].value     = "ПАРАМЕТРЫ МОДЕЛИ"
    ws["A1"].fill      = make_fill(BG["navy"])
    ws["A1"].font      = make_font(bold=True, size=12, color="FFFFFF")
    ws["A1"].alignment = make_align("center")
    ws.row_dimensions[1].height = 28

    rows = [
        ("Университет",            cfg["name"]),
        ("Профиль",                cfg["profile_label"]),
        ("Базовый год",            cfg["base_year"]),
        ("Горизонт, лет",          cfg["years"]),
        ("Итоговый год",           cfg["base_year"] + cfg["years"]),
        ("Студентов всего",        cfg["students"]),
        ("Стоимость обучения, р.", cfg["tuition"]),
        ("Доля бюджетников",       cfg["budget_ratio"]),
        ("Инфляция (база)",        0.07),
        ("Амортизация / Доходы",   DA_RATIO[cfg["profile"]]),
        ("CapEx / Доходы",         CAPEX_RATIO[cfg["profile"]]),
    ]

    for i, (label, value) in enumerate(rows):
        r  = i + 2
        bg = "row" if i % 2 == 0 else "white"

        lbl = ws.cell(row=r, column=1, value=label)
        lbl.fill = make_fill(BG[bg]); lbl.font = make_font()
        lbl.border = make_border(); lbl.alignment = make_align()

        val = ws.cell(row=r, column=2, value=value)
        val.fill = make_fill(BG["white"]); val.font = make_font(bold=True)
        val.border = make_border(); val.alignment = make_align("right")

        if label == "Доля бюджетников":
            val.number_format = "0%"
        elif label in ("Инфляция (база)", "Амортизация / Доходы", "CapEx / Доходы"):
            val.number_format = "0.0%"
        elif label in ("Студентов всего", "Стоимость обучения, р."):
            val.number_format = NUM
