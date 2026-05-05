from src.data_loader import DataLoader
from src.data_cleaner import DataCleaner
from src.metrics_builder import MetricsBuilder


loader = DataLoader("data/nasdaq_screener.csv")
raw_df = loader.load()

clean_df = DataCleaner(raw_df).clean()

metrics = MetricsBuilder(clean_df)
result = metrics.build_all()

print("TOP COMPANIES:")
print(result["top_companies"].head())

print("\nSECTOR SUMMARY:")
print(result["sector_summary"].head())

print("\nGENERAL SUMMARY:")
print(result["general_summary"])