import pandas as pd
from pathlib import Path

# =========================================================
# 1. Caminhos dos arquivos
# =========================================================
base_path = Path(r"C:\Users\Millene.Russo\OneDrive - Pandisc\base.xlsx")
nota_path = Path("nota_cliente.xlsx")

# =========================================================
# 2. Ler nota_cliente
#    ✅ CORREÇÃO: usar sheet_name=0 para evitar erro de nome
# =========================================================
df_nota = pd.read_excel(
    nota_path,
    sheet_name=0  # primeira aba
)

# Normalizar colunas
df_nota.columns = df_nota.columns.str.strip().str.upper()

# =========================================================
# 3. Extrair NOTAS FISCAIS distintas
# =========================================================
COL_NOTA_FISCAL = "NOTA FISCAL"

if COL_NOTA_FISCAL not in df_nota.columns:
    raise ValueError("Coluna NOTA FISCAL não encontrada em nota_cliente.xlsx")

notas = df_nota[COL_NOTA_FISCAL].dropna().unique()

# =========================================================
# 4. Ler base.xlsx SOMENTE com colunas necessárias
#    ✅ OTIMIZAÇÃO DE PERFORMANCE
# =========================================================
def normalizar_coluna(col):
    return (
        str(col)
        .replace("\u00A0", " ")
        .strip()
        .upper()
    )

colunas_base_desejadas = {
    "NF_ENTRADA",
    "TAG",
    "PARTNUMBER",
    "PPID_IN",
    "ADICIONAL 2",
    "ADICIONAL 4"
}

df_base = pd.read_excel(
    base_path,
    sheet_name="BASE_ATENDIMENTO",
    usecols=colunas_base_desejadas
)

df_base.columns = [normalizar_coluna(c) for c in df_base.columns]

print("Colunas carregadas:", df_base.columns.tolist())

# =========================================================
# 5. Filtrar base pelas notas fiscais
# =========================================================
df_base_filtrada = df_base[df_base["NF_ENTRADA"].isin(notas)]

# =========================================================
# 6. Mapeamento de colunas para comparação
# =========================================================
colunas_comparacao = {
    "HSN": "TAG",
    "DELL PN": "PARTNUMBER",
    "PPID": "PPID_IN",
    "DPS NUMBER": "ADICIONAL 2",
    "ORDER NUMBER": "ADICIONAL 4"
}

# =========================================================
# 7. Validação das colunas
# =========================================================
for col_nc in colunas_comparacao:
    if col_nc not in df_nota.columns:
        raise ValueError(f"Coluna {col_nc} não encontrada em nota_cliente.xlsx")

for col_base in colunas_comparacao.values():
    if col_base not in df_base_filtrada.columns:
        raise ValueError(f"Coluna {col_base} não encontrada em base.xlsx")

# =========================================================
# 8. Comparação usando nota_cliente como base
# =========================================================
resultados = []

for nota in notas:
    nota_rows = df_nota[df_nota[COL_NOTA_FISCAL] == nota]
    base_rows = df_base_filtrada[df_base_filtrada["NF_ENTRADA"] == nota]

    if base_rows.empty:
        # Nota não encontrada na base
        for col_nc in colunas_comparacao:
            resultados.append({
                "NOTA FISCAL": nota,
                "CAMPO": col_nc,
                "VALOR_NOTA_CLIENTE": nota_rows.iloc[0][col_nc],
                "VALOR_BASE": None,
                "RESULTADO": "NF NÃO ENCONTRADA NA BASE"
            })
        continue

    base_row = base_rows.iloc[0]

    for col_nc, col_base in colunas_comparacao.items():
        valor_nota = nota_rows.iloc[0][col_nc]
        valor_base = base_row[col_base]

        resultado = "MATCH" if valor_nota == valor_base else "DIFERENTE"

        resultados.append({
            "NOTA FISCAL": nota,
            "CAMPO": col_nc,
            "VALOR_NOTA_CLIENTE": valor_nota,
            "VALOR_BASE": valor_base,
            "RESULTADO": resultado
        })

# =========================================================
# 9. DataFrame final
# =========================================================
df_resultado = pd.DataFrame(resultados)

print(df_resultado)