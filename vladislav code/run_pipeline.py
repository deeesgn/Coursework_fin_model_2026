from src.data_loader import DataLoader
from src.data_validator import DataValidator
from src.data_cleaner import DataCleaner
from src.metrics_builder import MetricsBuilder
from src.financial_model_builder import FinancialModelBuilder
from src.model_filler import ModelFiller
from src.report_generator import ReportGenerator


loader = DataLoader("data/nasdaq_screener.csv")
raw_df = loader.load()
info = loader.get_basic_info()

validator = DataValidator(raw_df)
validation_result = validator.validate()

clean_df = DataCleaner(raw_df).clean()

metrics = MetricsBuilder(clean_df)
model_data = metrics.build_all()

financial_model = FinancialModelBuilder(clean_df)
financial_data = financial_model.build_model(top_n=10)

model_data.update(financial_data)

filler = ModelFiller()
excel_path = filler.fill(model_data)

report = ReportGenerator()
report_path = report.generate(validation_result, info)

print("=== ГОТОВО ===")
print("Excel:", excel_path)
print("Report:", report_path)