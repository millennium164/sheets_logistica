import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
import re

# =========================================================
# Funções utilitárias
# =========================================================
def normalizar_valor(v):
    if pd.isna(v):
        return ""
    v = str(v)
    v = v.replace("\u00A0", " ")
    v = re.sub(r"\s+", " ", v)
    return v.strip().upper()

def limpar_colunas(df):
    df.columns = (
        df.columns
        .astype(str)
        .str.replace("\u00A0", "", regex=False)
        .str.strip()
        .str.upper()
    )
    return df

def detectar_header(path, sheet_name, colunas_esperadas=None):
    """
    Detecta automaticamente a linha de cabeçalho.
    Se colunas_esperadas for None, apenas retorna o primeiro header válido.
    """
    for i in range(10):
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, header=i)
            df.columns = df.columns.astype(str).str.upper()

            if colunas_esperadas is None:
                # acha qualquer tabela minimamente válida
                if df.shape[1] >= 5:
                    return i
            else:
                if any(
                    any(col_esp in c for c in df.columns)
                    for col_esp in colunas_esperadas
                ):
                    return i
        except Exception:
            pass

    raise ValueError(
        f"Não foi possível detectar automaticamente o cabeçalho da aba '{sheet_name}'."
    )

# =========================================================
# Função principal de validação (merge por HSN ↔ TAG)
# =========================================================
def validar(df_nota, df_base, col_key_nota, col_key_base, pares):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = f"nota_validada_{timestamp}.xlsx"

    # --- normalizar chave ---
    df_nota["_KEY"] = df_nota[col_key_nota].apply(normalizar_valor)
    df_base["_KEY"] = df_base[col_key_base].apply(normalizar_valor)

    # --- garantir BASE 1 → 1 ---
    df_base_unica = (
        df_base
        .drop_duplicates(subset="_KEY", keep="first")
    )

    # --- merge correto ---
    df_merge = df_nota.merge(
        df_base_unica,
        how="left",
        on="_KEY"
    )

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_nota.drop(columns=["_KEY"]).to_excel(
            writer, index=False, sheet_name="NOTA VALIDADA"
        )

        workbook = writer.book
        worksheet = writer.sheets["NOTA VALIDADA"]

        fmt_green = workbook.add_format({
            "bg_color": "#C6EFCE",
            "font_color": "#006100"
        })

        fmt_red = workbook.add_format({
            "bg_color": "#FFC7CE",
            "font_color": "#9C0006"
        })

        # --- comparação célula a célula ---
        for col_nota, col_base in pares:
            col_idx = df_nota.columns.get_loc(col_nota)

            for row_idx in range(len(df_merge)):
                v_nota = normalizar_valor(df_merge.loc[row_idx, col_nota])
                v_base = normalizar_valor(df_merge.loc[row_idx, col_base])

                fmt = fmt_green if v_nota == v_base else fmt_red

                valor_original = df_nota.iloc[row_idx, col_idx]
                valor_excel = "" if pd.isna(valor_original) else str(valor_original)

                worksheet.write(
                    row_idx + 1,
                    col_idx,
                    valor_excel,
                    fmt
                )

    messagebox.showinfo(
        "Concluído",
        f"Arquivo gerado com sucesso:\n{output}"
    )

# =========================================================
# INTERFACE GRÁFICA
# =========================================================
root = tk.Tk()
root.withdraw()

# ---------------- selecionar arquivos ----------------
nota_file = filedialog.askopenfilename(
    title="Selecione a planilha da NOTA FISCAL",
    filetypes=[("Excel", "*.xlsx")]
)
if not nota_file:
    raise SystemExit

base_file = filedialog.askopenfilename(
    title="Selecione a planilha da BASE",
    filetypes=[("Excel", "*.xlsx")]
)
if not base_file:
    raise SystemExit

nota_excel = pd.ExcelFile(nota_file)
base_excel = pd.ExcelFile(base_file)

# ---------------- selecionar abas ----------------
aba_win = tk.Toplevel()
aba_win.title("Selecionar Abas")

ttk.Label(aba_win, text="Aba da NOTA").grid(row=0, column=0)
cb_aba_nota = ttk.Combobox(
    aba_win,
    values=nota_excel.sheet_names,
    state="readonly",
    width=40
)
cb_aba_nota.grid(row=1, column=0)

ttk.Label(aba_win, text="Aba da BASE").grid(row=2, column=0)
cb_aba_base = ttk.Combobox(
    aba_win,
    values=base_excel.sheet_names,
    state="readonly",
    width=40
)
cb_aba_base.grid(row=3, column=0)

selecoes = {}

def confirmar_abas():
    if not cb_aba_nota.get() or not cb_aba_base.get():
        messagebox.showerror("Erro", "Selecione as duas abas.")
        return
    selecoes["nota"] = cb_aba_nota.get()
    selecoes["base"] = cb_aba_base.get()
    aba_win.destroy()

ttk.Button(aba_win, text="Confirmar", command=confirmar_abas)\
    .grid(row=4, column=0, pady=10)

aba_win.wait_window()

# ---------------- carregar dataframes com header correto ----------------
# NOTA: sabemos que tem HSN
header_nota = detectar_header(
    nota_file,
    selecoes["nota"],
    colunas_esperadas=["HSN"]
)

# BASE: qualquer tabela grande serve (tag/ppid/part number etc.)
header_base = detectar_header(
    base_file,
    selecoes["base"]
)


df_nota = pd.read_excel(
    nota_file,
    sheet_name=selecoes["nota"],
    header=header_nota
)
df_base = pd.read_excel(
    base_file,
    sheet_name=selecoes["base"],
    header=header_base
)

df_nota = limpar_colunas(df_nota)
df_base = limpar_colunas(df_base)

# ---------------- selecionar colunas chave ----------------
key_win = tk.Toplevel()
key_win.title("Colunas de Ligação")

ttk.Label(key_win, text="Coluna da NOTA (ex: HSN)").grid(row=0, column=0)
cb_key_nota = ttk.Combobox(
    key_win,
    values=df_nota.columns.tolist(),
    state="readonly",
    width=35
)
cb_key_nota.grid(row=1, column=0)

ttk.Label(key_win, text="Coluna da BASE (ex: TAG)").grid(row=2, column=0)
cb_key_base = ttk.Combobox(
    key_win,
    values=df_base.columns.tolist(),
    state="readonly",
    width=35
)
cb_key_base.grid(row=3, column=0)

keys = {}

def confirmar_keys():
    if not cb_key_nota.get() or not cb_key_base.get():
        messagebox.showerror("Erro", "Selecione as colunas de ligação.")
        return
    keys["nota"] = cb_key_nota.get()
    keys["base"] = cb_key_base.get()
    key_win.destroy()

ttk.Button(key_win, text="Confirmar", command=confirmar_keys)\
    .grid(row=4, column=0, pady=10)

key_win.wait_window()

# ---------------- mapeamento de colunas ----------------
pares_padrao = [
    ("DELL PN", "PART NUMBER"),
    ("PPID", "PPID IN"),
    ("DPS NUMBER", "ADICIONAL 2"),
    ("ORDER NUMBER", "ADICIONAL 4"),
]

map_win = tk.Toplevel()
map_win.title("Mapeamento de Colunas")

frame = ttk.Frame(map_win, padding=10)
frame.pack()

pares_widgets = []

def adicionar_par(cn=None, cb=None):
    r = len(pares_widgets)

    c_n = ttk.Combobox(
        frame,
        values=df_nota.columns.tolist(),
        state="readonly",
        width=30
    )
    c_n.grid(row=r, column=0, padx=5, pady=3)

    ttk.Label(frame, text="⇄").grid(row=r, column=1)

    c_b = ttk.Combobox(
        frame,
        values=df_base.columns.tolist(),
        state="readonly",
        width=30
    )
    c_b.grid(row=r, column=2, padx=5, pady=3)

    if cn in df_nota.columns:
        c_n.set(cn)
    if cb in df_base.columns:
        c_b.set(cb)

    pares_widgets.append((c_n, c_b))

def prosseguir():
    pares = []
    for cn, cb in pares_widgets:
        if not cn.get() or not cb.get():
            messagebox.showerror("Erro", "Todos os pares devem estar preenchidos.")
            return
        pares.append((cn.get(), cb.get()))

    map_win.destroy()
    validar(
        df_nota,
        df_base,
        keys["nota"],
        keys["base"],
        pares
    )

for cn, cb in pares_padrao:
    adicionar_par(cn, cb)

ttk.Button(frame, text="Adicionar par", command=lambda: adicionar_par())\
    .grid(row=100, column=0, pady=10)

ttk.Button(frame, text="Prosseguir", command=prosseguir)\
    .grid(row=100, column=2, pady=10)

map_win.wait_window()