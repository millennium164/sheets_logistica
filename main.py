import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
import re

# =========================================================
# Funções utilitárias
# =========================================================
def normalizar_valor(v):
    """
    Normaliza valores para comparação:
    - NaN / None viram string vazia
    - floats que são inteiros perdem o .0 (12345.0 -> '12345')
    - remove espaços invisíveis (\u00A0, \u200B, \ufeff) e colapsa whitespace
    - remove apóstrofos de texto (artefato do Excel 'text-format)
    - uppercase + strip
    """
    try:
        if v is None:
            return ""
        if isinstance(v, float) and pd.isna(v):
            return ""
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(v, float) and v.is_integer():
        v = int(v)

    s = str(v)
    # caracteres invisíveis comuns em export Excel/SAP
    for ch in ("\u00A0", "\u200B", "\u200C", "\u200D", "\ufeff"):
        s = s.replace(ch, " " if ch == "\u00A0" else "")
    s = re.sub(r"\s+", " ", s).strip().upper()

    # remove apóstrofo / aspas simples que às vezes prefixam números texto
    s = s.strip("'\"`")

    # Remove ".0" residual quando vier como "123.0" ou "123.00"
    if re.fullmatch(r"-?\d+\.0+", s):
        s = s.split(".")[0]

    return s


def normalizar_chave_estrita(v):
    """
    Normalização mais agressiva usada só para comparar chaves:
    - passa por normalizar_valor (trata NaN, float-int, espaços invisíveis)
    - remove pontuação/símbolos, mantendo apenas letras e dígitos
      (resolve '123-456' vs '123456', '12.345.678' vs '12345678')
    - se a chave for puramente numérica, remove zeros à esquerda
      (resolve '001443388' vs '1443388' — zero-padding do Protheus/SAP)
    """
    s = normalizar_valor(v)
    if not s:
        return ""
    s = re.sub(r"[^0-9A-Z]", "", s)
    if not s:
        return ""
    if s.isdigit():
        s = s.lstrip("0") or "0"
    return s


# Valores que, mesmo preenchidos, devem ser tratados como "sem chave"
# para disparar o fallback (p.ex. PPID ausente → usa HSN).
VALORES_VAZIOS_CHAVE = {
    "", "NA", "N/A", "#N/A", "NAN", "NULL", "NONE", "-", "--", "—",
}


def construir_chave_linha(row, colunas_ordem):
    """
    Recebe uma linha de DataFrame e uma lista ordenada de colunas.
    Retorna a primeira chave estrita não-vazia (a primeira coluna cujo
    valor não seja vazio nem '(N/A)'). Vazio se nenhuma tiver valor útil.
    """
    for col in colunas_ordem:
        if col not in row.index:
            continue
        chave = normalizar_chave_estrita(row[col])
        if chave and chave not in VALORES_VAZIOS_CHAVE:
            return chave
    return ""


def coluna_de_origem_linha(row, colunas_ordem):
    """Retorna o nome da coluna que gerou a chave (útil p/ debug)."""
    for col in colunas_ordem:
        if col not in row.index:
            continue
        chave = normalizar_chave_estrita(row[col])
        if chave and chave not in VALORES_VAZIOS_CHAVE:
            return col
    return ""


def valores_equivalentes(a, b):
    """
    Comparação tolerante para campos de negócio:
    - igualdade textual normalizada
    - igualdade em forma de chave estrita (remove máscara/separador/zeros)
    """
    va = normalizar_valor(a)
    vb = normalizar_valor(b)
    if va == vb:
        return True
    ka = normalizar_chave_estrita(a)
    kb = normalizar_chave_estrita(b)
    if ka and kb and ka == kb:
        return True
    return False


def sugerir_pares_colunas(df_nota, df_base, limite_sugestoes=8):
    """
    Sugere pares de colunas por score de correspondência de valores.
    Retorna lista: [(col_nota, col_base, score_float), ...]
    """
    resultados = []
    colunas_nota = [c for c in df_nota.columns if not str(c).startswith("_")]
    colunas_base = [c for c in df_base.columns if not str(c).startswith("_")]
    if df_nota.empty or df_base.empty:
        return resultados
    n = df_nota.head(4000).copy()
    b = df_base.head(4000).copy()

    for cn in colunas_nota:
        serie_n = n[cn].map(normalizar_valor)
        serie_n_key = n[cn].map(normalizar_chave_estrita)
        valid_n = serie_n[serie_n != ""]
        valid_n_key = serie_n_key[serie_n_key != ""]
        if valid_n.empty:
            continue
        set_n = set(valid_n.head(2500))
        set_n_key = set(valid_n_key.head(2500))
        freq_n = valid_n.value_counts().head(120)
        freq_n_key = valid_n_key.value_counts().head(120)
        cobertura_n = len(valid_n) / len(n)

        for cb in colunas_base:
            serie_b = b[cb].map(normalizar_valor)
            serie_b_key = b[cb].map(normalizar_chave_estrita)
            valid_b = serie_b[serie_b != ""]
            valid_b_key = serie_b_key[serie_b_key != ""]
            if valid_b.empty:
                continue
            set_b = set(valid_b.head(2500))
            set_b_key = set(valid_b_key.head(2500))
            freq_b = valid_b.value_counts().head(120)
            freq_b_key = valid_b_key.value_counts().head(120)
            cobertura_b = len(valid_b) / len(b)

            # Score 1: sobreposição de conjuntos (independente de posição)
            inter_txt = len(set_n & set_b)
            inter_key = len(set_n_key & set_b_key)
            jacc_txt = inter_txt / (len(set_n | set_b) or 1)
            jacc_key = inter_key / (len(set_n_key | set_b_key) or 1)
            cont_txt = inter_txt / (min(len(set_n), len(set_b)) or 1)
            cont_key = inter_key / (min(len(set_n_key), len(set_b_key)) or 1)
            score_set = max(jacc_txt, jacc_key, cont_txt, cont_key)

            # Score 2: hit dos valores mais frequentes (também sem posição)
            hits_txt = sum(1 for v in freq_n.index if v in set_b)
            hits_key = sum(1 for v in freq_n_key.index if v in set_b_key)
            hits_txt_rev = sum(1 for v in freq_b.index if v in set_n)
            hits_key_rev = sum(1 for v in freq_b_key.index if v in set_n_key)
            den = max(len(freq_n.index), len(freq_b.index), 1)
            score_freq = max(hits_txt, hits_key, hits_txt_rev, hits_key_rev) / den

            # Score 3: cobertura semelhante (penaliza coluna quase vazia)
            score_cob = 1.0 - abs(cobertura_n - cobertura_b)

            # Score 4: similaridade de nome (apenas guia fraco)
            tok_n = set(re.findall(r"[A-Z0-9]+", cn.upper()))
            tok_b = set(re.findall(r"[A-Z0-9]+", cb.upper()))
            if tok_n and tok_b:
                score_nome = len(tok_n & tok_b) / len(tok_n | tok_b)
            else:
                score_nome = 0.0

            # Nome NÃO é determinante: peso pequeno.
            score = (
                (0.55 * score_set)
                + (0.30 * score_freq)
                + (0.10 * score_cob)
                + (0.05 * score_nome)
            )
            resultados.append((cn, cb, float(score)))

    resultados.sort(key=lambda x: x[2], reverse=True)

    # matching guloso 1:1 para evitar coluna duplicada em várias sugestões
    usados_n = set()
    usados_b = set()
    sugestoes = []
    for cn, cb, score in resultados:
        if score < 0.08:
            continue
        if cn in usados_n or cb in usados_b:
            continue
        usados_n.add(cn)
        usados_b.add(cb)
        sugestoes.append((cn, cb, score))
        if len(sugestoes) >= limite_sugestoes:
            break

    return sugestoes

def limpar_colunas(df):
    df.columns = (
        df.columns
        .astype(str)
        .str.replace("\u00A0", "", regex=False)
        .str.strip()
        .str.upper()
    )
    return df


def remover_linhas_em_branco(df):
    """
    Remove linhas totalmente vazias (considerando espaços, NaN, etc.).
    Retorna (df_filtrado, qtd_removida).
    """
    if df.empty:
        return df.copy(), 0

    def _linha_vazia(row):
        for v in row:
            if normalizar_valor(v) != "":
                return False
        return True

    mask_vazia = df.apply(_linha_vazia, axis=1)
    qtd = int(mask_vazia.sum())
    return df.loc[~mask_vazia].copy(), qtd

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
# Função principal de validação
# Referência: nota_cliente (df_nota)
# Validar preenchimento no sistema: report_protheus (df_base)
#
# chaves_nota / chaves_base: LISTAS ordenadas de colunas.
#   A 1ª é a chave principal; as seguintes são fallback
#   usados quando a principal estiver vazia ou '(N/A)'.
#   Exemplo: chaves_nota=['PPID', 'HSN'] → usa PPID;
#   se PPID da linha for N/A, cai para HSN.
# =========================================================
def validar(df_nota, df_base, chaves_nota, chaves_base, pares):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = f"nota_validada_{timestamp}.xlsx"

    if not chaves_nota or not chaves_base:
        raise ValueError("Informe pelo menos uma coluna-chave de cada lado.")
    if len(chaves_nota) != len(chaves_base):
        raise ValueError(
            "Chaves-nota e chaves-base precisam ter a mesma quantidade "
            "(principal + fallbacks emparelhados)."
        )

    for col in chaves_nota:
        if col not in df_nota.columns:
            raise ValueError(f"Coluna chave '{col}' não existe na nota_cliente.")
    for col in chaves_base:
        if col not in df_base.columns:
            raise ValueError(f"Coluna chave '{col}' não existe no report_protheus.")
    for cn, cb in pares:
        if cn not in df_nota.columns:
            raise ValueError(f"Coluna '{cn}' não existe na nota_cliente.")
        if cb not in df_base.columns:
            raise ValueError(f"Coluna '{cb}' não existe no report_protheus.")

    df_nota = df_nota.copy().reset_index(drop=True)
    df_base = df_base.copy().reset_index(drop=True)
    df_nota, linhas_branco_nota = remover_linhas_em_branco(df_nota)
    df_base, linhas_branco_base = remover_linhas_em_branco(df_base)
    df_nota = df_nota.reset_index(drop=True)
    df_base = df_base.reset_index(drop=True)
    label_chaves_nota = " → ".join(chaves_nota)
    label_chaves_base = " → ".join(chaves_base)

    nomes_colunas_norm = {normalizar_chave_estrita(c) for c in df_base.columns}
    nomes_colunas_norm.discard("")

    # Nota: uma chave efetiva por linha (cascata principal->fallback)
    df_nota["_KEY"] = df_nota.apply(lambda r: construir_chave_linha(r, chaves_nota), axis=1)
    df_nota["_KEY_ORIGEM"] = df_nota.apply(lambda r: coluna_de_origem_linha(r, chaves_nota), axis=1)
    df_nota["_NOTA_IDX"] = df_nota.index

    # Base: uma linha pode gerar múltiplas entradas de chave
    base_keys_registros = []
    linhas_fantasma_idx = set()
    for base_idx in df_base.index:
        row = df_base.loc[base_idx]
        for prioridade, col in enumerate(chaves_base):
            k = normalizar_chave_estrita(row[col])
            if not k or k in VALORES_VAZIOS_CHAVE:
                continue
            if k in nomes_colunas_norm:
                linhas_fantasma_idx.add(base_idx)
                continue
            base_keys_registros.append(
                {
                    "_BASE_IDX": base_idx,
                    "_KEY": k,
                    "_KEY_ORIGEM_BASE": col,
                    "_PRIORIDADE": prioridade,
                }
            )

    linhas_fantasma = len(linhas_fantasma_idx)
    df_base_keys = pd.DataFrame(base_keys_registros) if base_keys_registros else pd.DataFrame(
        columns=["_BASE_IDX", "_KEY", "_KEY_ORIGEM_BASE", "_PRIORIDADE"]
    )
    linhas_duplicadas_base = max(0, len(df_base_keys) - len(df_base_keys.drop_duplicates(subset=["_KEY", "_BASE_IDX"])))
    if not df_base_keys.empty:
        df_base_keys = df_base_keys.sort_values(["_KEY", "_PRIORIDADE", "_BASE_IDX"]).reset_index(drop=True)
        # rank dentro da chave para parear duplicados de forma determinística
        df_base_keys["_RANK_KEY"] = df_base_keys.groupby("_KEY").cumcount() + 1
    else:
        df_base_keys["_RANK_KEY"] = pd.Series(dtype=int)

    df_base["_BASE_IDX"] = df_base.index
    df_base_enriq = df_base_keys.merge(df_base, on="_BASE_IDX", how="left")

    # rank da nota dentro da chave para parear duplicados
    df_nota = df_nota.sort_values(["_KEY", "_NOTA_IDX"]).reset_index(drop=True)
    df_nota["_RANK_KEY"] = df_nota.groupby("_KEY").cumcount() + 1

    # 1) match nota -> base por (_KEY, rank)
    df_match = df_nota.merge(
        df_base_enriq,
        how="left",
        on=["_KEY", "_RANK_KEY"],
        indicator="_MATCH",
        suffixes=("", "__BASE_RAW"),
    ).reset_index(drop=True)

    # 2) sobras da base (não presentes na nota)
    chaves_usadas_nota = set(zip(df_match["_KEY"].fillna(""), df_match["_RANK_KEY"].fillna(0)))
    sobra_mask = ~df_base_enriq.apply(
        lambda r: (r["_KEY"], r["_RANK_KEY"]) in chaves_usadas_nota, axis=1
    )
    df_base_sobras = df_base_enriq[sobra_mask].copy().reset_index(drop=True)

    # Estatísticas por par
    stats_pares = {(cn, cb): {"ok": 0, "divergente": 0, "sem_base": 0, "nao_presente_nota": 0} for cn, cb in pares}

    # Montar linhas de saída unificadas
    registros_saida = []
    for idx in range(len(df_match)):
        row = df_match.iloc[idx]
        registro = {c: row[c] if c in df_nota.columns else "" for c in df_nota.columns if c not in {"_KEY", "_KEY_ORIGEM", "_NOTA_IDX", "_RANK_KEY"}}
        registro["_KEY"] = row["_KEY"]
        registro["_KEY_ORIGEM"] = row.get("_KEY_ORIGEM", "")
        registro["_MATCH"] = row["_MATCH"]
        registro["_KEY_ORIGEM_BASE"] = row.get("_KEY_ORIGEM_BASE", "")
        registro["_BASE_IDX"] = row.get("_BASE_IDX", None)
        linha_divergente = False

        if row["_MATCH"] == "left_only":
            status = "SEM CORRESPONDÊNCIA NA BASE"
            for cn, cb in pares:
                stats_pares[(cn, cb)]["sem_base"] += 1
                registro[f"__BASE__{cb}"] = ""
        else:
            for cn, cb in pares:
                cb_no_match = f"{cb}__BASE_RAW" if f"{cb}__BASE_RAW" in row.index else cb
                valor_base = row.get(cb_no_match, "")
                registro[f"__BASE__{cb}"] = valor_base
                if valores_equivalentes(row[cn], valor_base):
                    stats_pares[(cn, cb)]["ok"] += 1
                else:
                    stats_pares[(cn, cb)]["divergente"] += 1
                    linha_divergente = True
            status = "DIVERGENTE" if linha_divergente else "OK"

        registro["STATUS LINHA"] = status
        registros_saida.append(registro)

    sobras_vazias_ignoradas = 0
    colunas_contexto_base = list(dict.fromkeys(chaves_base + [cb for _, cb in pares]))

    for _, row in df_base_sobras.iterrows():
        # Ignora sobras que ficam totalmente vazias no contexto da validação
        # (chaves + colunas comparadas). Evita linhas "NÃO PRESENTE NA NOTA"
        # sem informação útil no output.
        if all(normalizar_valor(row.get(c, "")) == "" for c in colunas_contexto_base):
            sobras_vazias_ignoradas += 1
            continue

        registro = {c: "" for c in df_nota.columns if c not in {"_KEY", "_KEY_ORIGEM", "_NOTA_IDX", "_RANK_KEY"}}
        registro["_KEY"] = row["_KEY"]
        registro["_KEY_ORIGEM"] = ""
        registro["_MATCH"] = "right_only"
        registro["_KEY_ORIGEM_BASE"] = row.get("_KEY_ORIGEM_BASE", "")
        registro["_BASE_IDX"] = row.get("_BASE_IDX", None)
        for cn, cb in pares:
            registro[f"__BASE__{cb}"] = row.get(cb, "")
            stats_pares[(cn, cb)]["nao_presente_nota"] += 1
        registro["STATUS LINHA"] = "NÃO PRESENTE NA NOTA"
        registros_saida.append(registro)

    df_out = pd.DataFrame(registros_saida).reset_index(drop=True)
    cols_nota_saida = [c for c in df_nota.columns if c not in {"_KEY", "_KEY_ORIGEM", "_NOTA_IDX", "_RANK_KEY"}]
    df_excel = df_out[cols_nota_saida + ["STATUS LINHA"]].copy()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_excel.to_excel(writer, index=False, sheet_name="NOTA VALIDADA")

        workbook = writer.book
        worksheet = writer.sheets["NOTA VALIDADA"]

        fmt_green = workbook.add_format({
            "bg_color": "#C6EFCE",
            "font_color": "#006100",
        })
        fmt_red = workbook.add_format({
            "bg_color": "#FFC7CE",
            "font_color": "#9C0006",
        })
        fmt_yellow = workbook.add_format({
            "bg_color": "#FFEB9C",
            "font_color": "#9C5700",
        })

        # Índices das colunas-chave
        key_col_indices = {
            col: df_excel.columns.get_loc(col)
            for col in chaves_nota
            if col in df_excel.columns
        }
        status_col_idx = df_excel.columns.get_loc("STATUS LINHA")

        # --- pintar célula a célula ---
        for row_idx in range(len(df_out)):
            match = df_out.at[row_idx, "_MATCH"]

            # Amarelo na coluna-chave efetivamente usada (ou na principal)
            # quando a linha não encontrou par na base
            if match == "left_only":
                col_origem = df_out.at[row_idx, "_KEY_ORIGEM"] or chaves_nota[0]
                if col_origem in key_col_indices:
                    k_idx = key_col_indices[col_origem]
                    v_key = df_excel.iloc[row_idx, k_idx]
                    v_key_excel = "" if pd.isna(v_key) else str(v_key)
                    worksheet.write(row_idx + 1, k_idx, v_key_excel, fmt_yellow)

            # Pares de comparação
            for cn, cb in pares:
                col_idx = df_excel.columns.get_loc(cn)
                valor_original = df_excel.iloc[row_idx, col_idx]
                valor_excel = "" if pd.isna(valor_original) else str(valor_original)

                if match == "left_only":
                    # Sem par: a base não trouxe informação → divergência real
                    fmt = fmt_red
                elif match == "right_only":
                    fmt = fmt_red
                else:
                    fmt = (
                        fmt_green
                        if valores_equivalentes(
                            df_out.at[row_idx, cn],
                            df_out.at[row_idx, f"__BASE__{cb}"],
                        )
                        else fmt_red
                    )

                worksheet.write(row_idx + 1, col_idx, valor_excel, fmt)

            # Pintar a própria coluna STATUS LINHA
            status_val = df_out.at[row_idx, "STATUS LINHA"]
            if status_val == "OK":
                fmt_status = fmt_green
            elif status_val in {"DIVERGENTE", "NÃO PRESENTE NA NOTA"}:
                fmt_status = fmt_red
            else:
                fmt_status = fmt_yellow
            worksheet.write(row_idx + 1, status_col_idx, status_val, fmt_status)

        # --- aba RESUMO ---
        total = len(df_out)
        matched = int((df_out["_MATCH"] == "both").sum())
        sem_match = int((df_out["_MATCH"] == "left_only").sum())
        nao_presente_nota = int((df_out["_MATCH"] == "right_only").sum())
        linhas_ok = int((df_out["STATUS LINHA"] == "OK").sum())
        linhas_div = sum(1 for s in df_out["STATUS LINHA"] if s == "DIVERGENTE")

        # Diagnóstico por origem (principal vs fallback)
        diag_origem = []
        for col_n in chaves_nota:
            mask = df_out["_KEY_ORIGEM"] == col_n
            total_col = int(mask.sum())
            casadas_col = int((mask & (df_out["_MATCH"] == "both")).sum())
            diag_origem.append((col_n, total_col, casadas_col))

        diag_str = " | ".join(
            f"{col}: {cas}/{tot}" for col, tot, cas in diag_origem
        )

        resumo_cab = pd.DataFrame(
            [
                ("Arquivo gerado em", datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
                ("Total de linhas na nota (referência)", total),
                ("Linhas casadas com a base", matched),
                ("Linhas sem correspondência na base", sem_match),
                ("Linhas NÃO PRESENTES NA NOTA (sobras Protheus)", nao_presente_nota),
                ("Linhas OK (todos os pares conferem)", linhas_ok),
                ("Linhas DIVERGENTES", linhas_div),
                ("Linhas em branco removidas da nota", linhas_branco_nota),
                ("Linhas em branco removidas da base", linhas_branco_base),
                ("Linhas fantasma removidas da base (chave = nome de coluna)",
                    linhas_fantasma),
                ("Linhas da base descartadas por chave duplicada",
                    linhas_duplicadas_base),
                ("Sobras vazias da base ignoradas", sobras_vazias_ignoradas),
                ("Chaves nota (principal → fallback)", label_chaves_nota),
                ("Chaves base (principal → fallback)", label_chaves_base),
                ("Casadas por coluna-origem (nota)", diag_str),
            ],
            columns=["Métrica", "Valor"],
        )
        resumo_cab.to_excel(writer, index=False, sheet_name="RESUMO", startrow=0)

        resumo_pares = pd.DataFrame(
            [
                {
                    "Coluna nota_cliente": cn,
                    "Coluna report_protheus": cb,
                    "OK": stats_pares[(cn, cb)]["ok"],
                    "Divergente": stats_pares[(cn, cb)]["divergente"],
                    "Sem base": stats_pares[(cn, cb)]["sem_base"],
                    "Nao presente nota": stats_pares[(cn, cb)]["nao_presente_nota"],
                }
                for cn, cb in pares
            ]
        )
        start_pares = len(resumo_cab) + 3
        ws_resumo = writer.sheets["RESUMO"]
        ws_resumo.write(start_pares - 1, 0, "Detalhe por par de colunas")
        resumo_pares.to_excel(
            writer, index=False, sheet_name="RESUMO", startrow=start_pares
        )

        ws_resumo.set_column(0, 0, 38)
        ws_resumo.set_column(1, 4, 22)
        worksheet.set_column(0, len(df_excel.columns) - 1, 20)

        # --- aba DEBUG CHAVES (sempre gerada, essencial p/ troubleshoot) ---
        amostra = 30

        def _montar_tabela_chaves_nota(df, colunas):
            cols_existentes = [c for c in colunas if c in df.columns]
            out = df[cols_existentes].copy()
            out.columns = [f"{c} (bruto)" for c in cols_existentes]
            for c in cols_existentes:
                out[f"{c} (estrita)"] = df[c].apply(normalizar_chave_estrita)
            out["_KEY_EFETIVA"] = df["_KEY"]
            out["_COLUNA_USADA"] = df["_KEY_ORIGEM"]
            return out

        def _montar_tabela_chaves_base(df, colunas):
            """Para a base, mostra todas as chaves candidatas de cada linha."""
            cols_existentes = [c for c in colunas if c in df.columns]
            out = df[cols_existentes].copy()
            out.columns = [f"{c} (bruto)" for c in cols_existentes]
            for c in cols_existentes:
                out[f"{c} (estrita)"] = df[c].apply(normalizar_chave_estrita)
            return out

        keys_nota = _montar_tabela_chaves_nota(df_nota, chaves_nota)
        keys_base = _montar_tabela_chaves_base(df_base, chaves_base)

        set_nota = set(df_nota["_KEY"]) - {""}
        set_base = set(df_base_enriq["_KEY"]) - {""}
        intersec = sorted(set_nota & set_base)
        so_nota = sorted(set_nota - set_base)[:amostra]
        so_base = sorted(set_base - set_nota)[:amostra]

        origem_counts_nota = df_out[df_out["_MATCH"] != "right_only"]["_KEY_ORIGEM"].value_counts().to_dict()
        origem_counts_base = (
            df_base_enriq["_KEY_ORIGEM_BASE"].value_counts().to_dict()
            if not df_base_enriq.empty
            else {}
        )
        origem_str_nota = ", ".join(
            f"{k or '(vazio)'}={v}" for k, v in origem_counts_nota.items()
        ) or "(nenhuma linha com chave)"
        origem_str_base = ", ".join(
            f"{k or '(vazio)'}={v}" for k, v in origem_counts_base.items()
        ) or "(nenhuma linha com chave)"

        resumo_chaves = pd.DataFrame(
            [
                ("Chaves nota (principal → fallback)", label_chaves_nota),
                ("Chaves base (principal → fallback)", label_chaves_base),
                ("Linhas na nota", len(df_nota)),
                ("Linhas na base (original)", len(df_base)),
                ("Linhas em branco removidas na nota", linhas_branco_nota),
                ("Linhas em branco removidas na base", linhas_branco_base),
                ("Linhas fantasma removidas da base", linhas_fantasma),
                ("Linhas duplicadas descartadas na base", linhas_duplicadas_base),
                ("Sobras vazias ignoradas na base", sobras_vazias_ignoradas),
                ("Sobras da base adicionadas na saída", nao_presente_nota),
                ("Chaves únicas não-vazias na nota", len(set_nota)),
                ("Chaves únicas não-vazias na base", len(set_base)),
                ("Chaves em comum (após normalização estrita)", len(intersec)),
                ("Origem da chave usada (nota)", origem_str_nota),
                ("Origem da chave usada (base)", origem_str_base),
            ],
            columns=["Métrica", "Valor"],
        )
        resumo_chaves.to_excel(
            writer, index=False, sheet_name="DEBUG CHAVES", startrow=0
        )
        ws_dbg = writer.sheets["DEBUG CHAVES"]
        ws_dbg.set_column(0, 0, 45)
        ws_dbg.set_column(1, 10, 30)

        row_cursor = len(resumo_chaves) + 3
        ws_dbg.write(row_cursor - 1, 0, f"Amostra ({amostra}) chaves nota_cliente:")
        keys_nota.head(amostra).to_excel(
            writer, index=False, sheet_name="DEBUG CHAVES", startrow=row_cursor
        )
        row_cursor += amostra + 3

        ws_dbg.write(row_cursor - 1, 0, f"Amostra ({amostra}) chaves report_protheus:")
        keys_base.head(amostra).to_excel(
            writer, index=False, sheet_name="DEBUG CHAVES", startrow=row_cursor
        )
        row_cursor += amostra + 3

        ws_dbg.write(
            row_cursor - 1,
            0,
            f"Chaves presentes SÓ na nota (até {amostra}):",
        )
        pd.DataFrame({"chave estrita": so_nota}).to_excel(
            writer, index=False, sheet_name="DEBUG CHAVES", startrow=row_cursor
        )
        row_cursor += max(len(so_nota), 1) + 3

        ws_dbg.write(
            row_cursor - 1,
            0,
            f"Chaves presentes SÓ na base (até {amostra}):",
        )
        pd.DataFrame({"chave estrita": so_base}).to_excel(
            writer, index=False, sheet_name="DEBUG CHAVES", startrow=row_cursor
        )
        row_cursor += max(len(so_base), 1) + 3

        ws_dbg.write(
            row_cursor - 1,
            0,
            f"Exemplos de chaves em COMUM (até {amostra}):",
        )
        pd.DataFrame({"chave estrita": intersec[:amostra]}).to_excel(
            writer,
            index=False,
            sheet_name="DEBUG CHAVES",
            startrow=row_cursor,
        )

        # log no console
        print("\n=== DEBUG CHAVES ===")
        print(f"chaves nota: {label_chaves_nota}")
        print(f"chaves base: {label_chaves_base}")
        print(f"#nota={len(df_nota)}  #base={len(df_base)}")
        print(
            f"linhas fantasma removidas da base = {linhas_fantasma} | "
            f"duplicadas descartadas = {linhas_duplicadas_base}"
        )
        print(f"origem nota: {origem_str_nota}")
        print(f"origem base: {origem_str_base}")
        print(f"#chaves únicas nota={len(set_nota)}  base={len(set_base)}")
        print(f"#interseção = {len(intersec)}")
        print("===================\n")

    # --- feedback final ---
    extras = (
        f"\nLinhas em branco removidas da nota: {linhas_branco_nota}"
        f"\nLinhas em branco removidas da base: {linhas_branco_base}"
        f"\nLinhas fantasma removidas da base: {linhas_fantasma}"
        f"\nLinhas duplicadas descartadas: {linhas_duplicadas_base}"
        f"\nSobras vazias ignoradas da base: {sobras_vazias_ignoradas}"
        f"\nLinhas extras da base adicionadas: {nao_presente_nota}"
    )
    if matched == 0:
        amostra_nota = ", ".join(list(set_nota)[:5]) or "(vazio)"
        amostra_base = ", ".join(list(set_base)[:5]) or "(vazio)"
        messagebox.showwarning(
            "Nenhuma chave casou!",
            (
                f"Arquivo gerado:\n{output}\n\n"
                f"Nenhuma chave da nota ({label_chaves_nota}) bateu com "
                f"a base ({label_chaves_base}) após normalização.\n\n"
                f"Abra a aba DEBUG CHAVES para comparar os valores lado a lado.\n\n"
                f"Amostra nota: {amostra_nota}\n"
                f"Amostra base: {amostra_base}\n"
                f"{extras}"
            ),
        )
    else:
        messagebox.showinfo(
            "Concluído",
            (
                f"Arquivo gerado:\n{output}\n\n"
                f"Total de linhas: {total}\n"
                f"Casadas com a base: {matched}\n"
                f"Sem correspondência: {sem_match}\n"
                f"NÃO PRESENTE NA NOTA: {nao_presente_nota}\n"
                f"OK: {linhas_ok}\n"
                f"Divergentes: {linhas_div}"
                f"{extras}\n\n"
                f"Se algo parecer divergente indevidamente, confira a aba DEBUG CHAVES."
            ),
        )

# =========================================================
# INTERFACE GRÁFICA
# =========================================================
root = tk.Tk()
root.withdraw()

# ---------------- selecionar arquivos ----------------
nota_file = filedialog.askopenfilename(
    title="Selecione a planilha da NOTA DO CLIENTE (referência)",
    filetypes=[("Excel", "*.xlsx")]
)
if not nota_file:
    raise SystemExit

base_file = filedialog.askopenfilename(
    title="Selecione a planilha do REPORT PROTHEUS (a validar)",
    filetypes=[("Excel", "*.xlsx")]
)
if not base_file:
    raise SystemExit

nota_excel = pd.ExcelFile(nota_file)
base_excel = pd.ExcelFile(base_file)

# ---------------- selecionar abas ----------------
aba_win = tk.Toplevel()
aba_win.title("Selecionar Abas")

ttk.Label(aba_win, text="Aba da NOTA DO CLIENTE (referência)").grid(row=0, column=0)
cb_aba_nota = ttk.Combobox(
    aba_win,
    values=nota_excel.sheet_names,
    state="readonly",
    width=40
)
cb_aba_nota.grid(row=1, column=0)

ttk.Label(aba_win, text="Aba do REPORT PROTHEUS (a validar)").grid(row=2, column=0)
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
df_nota, _linhas_branco_nota_inicial = remover_linhas_em_branco(df_nota)
df_base, _linhas_branco_base_inicial = remover_linhas_em_branco(df_base)

# ---------------- selecionar colunas chave (principal + fallbacks) ----------------
key_win = tk.Toplevel()
key_win.title("Colunas de Ligação (principal + fallbacks)")

key_frame = ttk.Frame(key_win, padding=10)
key_frame.pack()

ttk.Label(
    key_frame,
    text=(
        "Defina a chave principal e, opcionalmente, colunas de fallback.\n"
        "Se a chave principal de uma linha estiver vazia ou '(N/A)', a\n"
        "próxima coluna na lista será usada. Ex.: PPID ↔ PPID IN "
        "(principal) + HSN ↔ TAG (fallback)."
    ),
    foreground="#555",
).grid(row=0, column=0, columnspan=4, pady=(0, 8), sticky="w")

ttk.Label(
    key_frame,
    text="Coluna NOTA DO CLIENTE",
    font=("TkDefaultFont", 9, "bold"),
).grid(row=1, column=0, padx=5)
ttk.Label(key_frame, text="").grid(row=1, column=1)
ttk.Label(
    key_frame,
    text="Coluna REPORT PROTHEUS",
    font=("TkDefaultFont", 9, "bold"),
).grid(row=1, column=2, padx=5)
ttk.Label(key_frame, text="").grid(row=1, column=3)

# Sugestões padrão: PPID principal, HSN fallback (se existirem)
chaves_padrao = [
    ("PPID", "PPID IN"),
    ("HSN", "TAG"),
]

chave_widgets = []


def _rotulo_par(idx):
    return "Principal" if idx == 0 else f"Fallback {idx}"


def adicionar_chave(cn=None, cb=None):
    r = len(chave_widgets) + 2  # +2 porque linhas 0/1 são texto e cabeçalho

    cn_cb = ttk.Combobox(
        key_frame,
        values=df_nota.columns.tolist(),
        state="readonly",
        width=32,
    )
    cn_cb.grid(row=r, column=0, padx=5, pady=2)

    ttk.Label(key_frame, text="⇄").grid(row=r, column=1)

    cb_cb = ttk.Combobox(
        key_frame,
        values=df_base.columns.tolist(),
        state="readonly",
        width=32,
    )
    cb_cb.grid(row=r, column=2, padx=5, pady=2)

    tipo_lbl = ttk.Label(key_frame, text=_rotulo_par(len(chave_widgets)))
    tipo_lbl.grid(row=r, column=3, padx=5)

    if cn and cn in df_nota.columns:
        cn_cb.set(cn)
    if cb and cb in df_base.columns:
        cb_cb.set(cb)

    chave_widgets.append((cn_cb, cb_cb, tipo_lbl))


keys = {}


def confirmar_keys():
    chaves_nota = []
    chaves_base = []
    for cn_cb, cb_cb, _ in chave_widgets:
        cn = cn_cb.get()
        cb = cb_cb.get()
        if not cn and not cb:
            continue  # linha em branco, ignora
        if not cn or not cb:
            messagebox.showerror(
                "Erro",
                "Cada linha de chave precisa ter coluna na NOTA e na BASE.",
            )
            return
        chaves_nota.append(cn)
        chaves_base.append(cb)

    if not chaves_nota:
        messagebox.showerror("Erro", "Defina pelo menos a chave principal.")
        return

    keys["nota"] = chaves_nota
    keys["base"] = chaves_base
    key_win.destroy()


for cn, cb in chaves_padrao:
    adicionar_chave(cn, cb)

btn_frame = ttk.Frame(key_frame)
btn_frame.grid(row=100, column=0, columnspan=4, pady=(10, 0))

ttk.Button(
    btn_frame,
    text="+ Adicionar fallback",
    command=lambda: adicionar_chave(),
).pack(side="left", padx=5)

ttk.Button(btn_frame, text="Confirmar", command=confirmar_keys)\
    .pack(side="left", padx=5)

key_win.wait_window()

# ---------------- mapeamento de colunas ----------------
# PPID e HSN não estão aqui porque foram promovidos a coluna-chave
# (principal + fallback). O usuário pode adicionar manualmente se quiser
# comparar mesmo assim.
pares_padrao = [
    ("NOTA FISCAL", "NOTA DE ENTRADA"),
    ("DELL PN", "PART NUMBER"),
    ("DPS NUMBER", "ADICIONAL 2"),
    ("ORDER NUMBER", "ADICIONAL 4"),
]

map_win = tk.Toplevel()
map_win.title("Mapeamento de Colunas — nota_cliente ⇄ report_protheus")

frame = ttk.Frame(map_win, padding=10)
frame.pack()

ttk.Label(
    frame,
    text="Coluna da NOTA DO CLIENTE (referência)",
    font=("TkDefaultFont", 9, "bold"),
).grid(row=0, column=0, padx=5, pady=(0, 5))
ttk.Label(frame, text="").grid(row=0, column=1)
ttk.Label(
    frame,
    text="Coluna do REPORT PROTHEUS (a validar)",
    font=("TkDefaultFont", 9, "bold"),
).grid(row=0, column=2, padx=5, pady=(0, 5))

pares_widgets = []

def adicionar_par(cn=None, cb=None):
    # +1 por causa do cabeçalho na linha 0
    r = len(pares_widgets) + 1

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


def aplicar_sugestoes_automaticas():
    sugestoes = sugerir_pares_colunas(df_nota, df_base, limite_sugestoes=10)
    if not sugestoes:
        messagebox.showwarning(
            "Sem sugestões",
            "Não foi possível inferir pares com score suficiente.",
        )
        return

    # Mapa das seleções atuais para não sobrescrever o que usuário já definiu.
    selecionados_nota = set()
    selecionados_base = set()
    for w_n, w_b in pares_widgets:
        if w_n.get():
            selecionados_nota.add(w_n.get())
        if w_b.get():
            selecionados_base.add(w_b.get())

    # Regra forte: preservar/forçar PPID ↔ PPID IN quando ambos existirem.
    if "PPID" in df_nota.columns and "PPID IN" in df_base.columns:
        ja_tem_ppid = any(
            (w_n.get() == "PPID" and w_b.get() == "PPID IN")
            for w_n, w_b in pares_widgets
        )
        if not ja_tem_ppid:
            alvo = None
            for i, (w_n, w_b) in enumerate(pares_widgets):
                if not w_n.get() and not w_b.get():
                    alvo = i
                    break
            if alvo is None:
                adicionar_par()
                alvo = len(pares_widgets) - 1
            w_n, w_b = pares_widgets[alvo]
            w_n.set("PPID")
            w_b.set("PPID IN")
            selecionados_nota.add("PPID")
            selecionados_base.add("PPID IN")

    # Preenche sugestões SOMENTE em linhas vazias (não sobrescreve manuais).
    for cn, cb, score in sugestoes:
        if cn in selecionados_nota or cb in selecionados_base:
            continue
        alvo = None
        for i, (w_n, w_b) in enumerate(pares_widgets):
            if not w_n.get() and not w_b.get():
                alvo = i
                break
        if alvo is None:
            adicionar_par()
            alvo = len(pares_widgets) - 1
        w_n, w_b = pares_widgets[alvo]
        if cn in df_nota.columns:
            w_n.set(cn)
        if cb in df_base.columns:
            w_b.set(cb)
        selecionados_nota.add(cn)
        selecionados_base.add(cb)

    top3 = "\n".join(
        f"- {cn} ↔ {cb} (score={score:.2f})"
        for cn, cb, score in sugestoes[:3]
    )
    messagebox.showinfo(
        "Sugestões aplicadas",
        (
            "Pares sugeridos por score foram preenchidos.\n\n"
            f"Top sugestões:\n{top3}"
        ),
    )

def prosseguir():
    pares = []
    for cn, cb in pares_widgets:
        if not cn.get() or not cb.get():
            messagebox.showerror("Erro", "Todos os pares devem estar preenchidos.")
            return
        pares.append((cn.get(), cb.get()))

    map_win.destroy()
    try:
        validar(
            df_nota,
            df_base,
            keys["nota"],
            keys["base"],
            pares
        )
    except ValueError as e:
        messagebox.showerror("Erro de validação", str(e))
    except Exception as e:
        messagebox.showerror(
            "Erro inesperado",
            f"{type(e).__name__}: {e}"
        )

for cn, cb in pares_padrao:
    adicionar_par(cn, cb)

ttk.Button(frame, text="Adicionar par", command=lambda: adicionar_par())\
    .grid(row=100, column=0, pady=10)

ttk.Button(frame, text="Sugerir pares (score)", command=aplicar_sugestoes_automaticas)\
    .grid(row=100, column=1, pady=10)

ttk.Button(frame, text="Prosseguir", command=prosseguir)\
    .grid(row=100, column=2, pady=10)

map_win.wait_window()