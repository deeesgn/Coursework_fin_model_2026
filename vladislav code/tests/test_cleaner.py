from src.data_loader import DataLoader
from src.data_cleaner import DataCleaner


loader = DataLoader("data/nasdaq_screener.csv")
df = loader.load()

cleaner = DataCleaner(df)
clean_df = cleaner.clean()

print(clean_df.head())
print(clean_df.dtypes)
print("Количество строк после очистки:", len(clean_df))
