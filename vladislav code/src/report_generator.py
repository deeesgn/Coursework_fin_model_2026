from datetime import datetime
from pathlib import Path


class ReportGenerator:
    """
    Генерирует отчёт о загрузке и обработке данных.
    """

    def __init__(self, output_path="output/report.txt"):
        self.output_path = Path(output_path)
        self.output_path.parent.mkdir(parents=True, exist_ok=True)

    def generate(self, validation_result: dict, data_info: dict):
        lines = []

        lines.append("=== DATA LOADING REPORT ===")
        lines.append(f"Дата: {datetime.now()}")
        lines.append("")

        lines.append("=== DATA INFO ===")
        for key, value in data_info.items():
            lines.append(f"{key}: {value}")

        lines.append("")
        lines.append("=== VALIDATION ===")
        lines.append(f"Valid: {validation_result['is_valid']}")

        lines.append("")
        lines.append("Ошибки:")
        if validation_result["errors"]:
            for e in validation_result["errors"]:
                lines.append(f"- {e}")
        else:
            lines.append("Нет ошибок")

        lines.append("")
        lines.append("Предупреждения:")
        if validation_result["warnings"]:
            for w in validation_result["warnings"]:
                lines.append(f"- {w}")
        else:
            lines.append("Нет предупреждений")

        with open(self.output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

        return self.output_path