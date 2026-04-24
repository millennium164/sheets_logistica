"""
Utilitário de debug: abre uma planilha gerada pela validação
(nota_validada_*.xlsx) e imprime as colunas e um preview da chave.

Uso:
    python debug.py                 # abre diálogo de arquivo
    python debug.py caminho.xlsx    # usa o arquivo passado
"""
import sys
from pathlib import Path

import pandas as pd


def escolher_arquivo() -> Path:
    if len(sys.argv) > 1:
        return Path(sys.argv[1])

    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        raise SystemExit(
            "Tkinter indisponível. Passe o caminho do arquivo como argumento."
        )

    root = tk.Tk()
    root.withdraw()
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha a inspecionar",
        filetypes=[("Excel", "*.xlsx")],
    )
    if not caminho:
        raise SystemExit("Nenhum arquivo selecionado.")
    return Path(caminho)


def main() -> None:
    arquivo = escolher_arquivo()
    if not arquivo.exists():
        raise SystemExit(f"Arquivo não encontrado: {arquivo}")

    print(f"Arquivo: {arquivo}\n")
    xls = pd.ExcelFile(arquivo, engine="openpyxl")
    print(f"Abas disponíveis: {xls.sheet_names}\n")

    for aba in xls.sheet_names:
        df = pd.read_excel(arquivo, sheet_name=aba, engine="openpyxl")
        df.columns = (
            df.columns.astype(str)
            .str.replace("\u00A0", "", regex=False)
            .str.strip()
            .str.upper()
        )
        print(f"--- Aba: {aba} ({len(df)} linhas) ---")
        for c in df.columns:
            print(f"  {c!r}")

        for coluna_interesse in ("HSN", "STATUS LINHA"):
            if coluna_interesse in df.columns:
                print(f"\n  Preview {coluna_interesse}:")
                print(df[coluna_interesse].head().to_string(index=False))
        print()


if __name__ == "__main__":
    main()
