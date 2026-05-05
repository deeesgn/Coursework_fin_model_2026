from src.data_loader import DataLoader
from src.data_cleaner import DataCleaner
from src.metrics_builder import MetricsBuilder
from src.model_filler import ModelFiller


loader = DataLoader("data/nasdaq_screener.csv")
raw_df = loader.load()

clean_df = DataCleaner(raw_df).clean()

metrics = MetricsBuilder(clean_df)
model_data = metrics.build_all()

filler = ModelFiller("output/filled_financial_model.xlsx")
output_file = filler.fill(model_data)

print("Финансовая модель заполнена и сохранена:")
print(output_file)