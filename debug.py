import pandas as pd
from pathlib import Path

base_path = Path(r"C:\Users\Millene.Russo\workspace\sheets_logistica\nota_validada_20260423_162601.xlsx")

df_debug = pd.read_excel(
    base_path,
    sheet_name="Nota Validada",
    engine="openpyxl"
)

df_debug.columns = (
    df_debug.columns
    .astype(str)
    .str.replace("\u00A0", "", regex=False)
    .str.strip()
    .str.upper()
)


print("COLUNAS REAIS DO DATAFRAME:")
for c in df_debug.columns:
    print(repr(c))



print(
    df_debug["HSN"].head()
)
print(len(df_debug))
