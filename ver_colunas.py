import pandas as pd

df = pd.read_excel("planilha_origem.xlsx")
print("Colunas encontradas na planilha:")
for col in df.columns:
    print(f"- {col}")
