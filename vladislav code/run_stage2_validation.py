from pathlib import Path
import json

from src.workbook_validator import WorkbookValidator
from src.checked_model_validator import CheckedModelValidator


Path("output").mkdir(exist_ok=True)

validators = [
    (
        "Модель кампуса",
        WorkbookValidator("data/01_ФМ_Кампус Фрязино_до 2042 3003.xlsx"),
        "output/campus_model_validation_report.json",
    ),
    (
        "Проверенная сводная модель",
        CheckedModelValidator("data/Проверенная_ФМ_Кампус_Фрязино_Сводная_модель_3003.xlsx"),
        "output/checked_model_validation_report.json",
    ),
]

summary = []

for name, validator, output_path in validators:
    result = validator.validate()

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    summary.append(
        {
            "model": name,
            "file": result["file"],
            "is_valid": result["is_valid"],
            "errors_count": len(result["errors"]),
            "warnings_count": len(result["warnings"]),
        }
    )

with open("output/stage2_summary_report.json", "w", encoding="utf-8") as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)

with open("output/stage2_summary_report.txt", "w", encoding="utf-8") as f:
    f.write("СВОДНЫЙ ОТЧЁТ ПРОВЕРКИ МОДЕЛЕЙ\n")

    for item in summary:
        f.write(f"Модель: {item['model']}\n")
        f.write(f"Файл: {item['file']}\n")
        f.write(f"Статус: {'валидна' if item['is_valid'] else 'есть ошибки'}\n")
        f.write(f"Ошибок: {item['errors_count']}\n")
        f.write(f"Предупреждений: {item['warnings_count']}\n")
        f.write("-" * 60 + "\n")

print("Проверка второго этапа завершена")
print("Сводный отчёт:", "output/stage2_summary_report.txt")
