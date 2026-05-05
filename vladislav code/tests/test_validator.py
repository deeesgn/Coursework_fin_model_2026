from src.data_loader import DataLoader
from src.data_validator import DataValidator


loader = DataLoader("data/nasdaq_screener.csv")
df = loader.load()

validator = DataValidator(df)
result = validator.validate()

print(result)