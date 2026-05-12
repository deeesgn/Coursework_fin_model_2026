from src.workbook_validator import WorkbookValidator


validator = WorkbookValidator(
    "data/01_ФМ_Кампус Фрязино_до 2042 3003.xlsx"
)

result = validator.validate()

print(result)