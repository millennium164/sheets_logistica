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
    # se restou apenas número, remove zeros à esquerda, caso haja
    if s.isdigit():
        s = s.lstrip("0") or "0"
    return s


# Valores que, mesmo preenchidos, devem ser tratados como "sem chave"
# para disparar o fallback (p.ex. PPID ausente → usa HSN).
VALORES_VAZIOS_CHAVE = {
    "", "NA", "N/A", "#N/A", "NAN", "NULL", "NONE", "-", "--", "—",
}

def eh_vazio_semantico(v):
    """
    Considera vazio:
    - NaN/None/""
    - e tokens como N/A, NA, NULL, '-', etc (VALORES_VAZIOS_CHAVE)
    """
    nv = normalizar_valor(v)  # já vira upper/strip e trata NaN
    return (nv == "") or (nv in VALORES_VAZIOS_CHAVE)

def filtrar_base_por_nota(df_nota, df_base, col_filtro_nota, col_filtro_base):
    """
    Filtra df_base usando os valores únicos de df_nota[col_filtro_nota],
    comparando com df_base[col_filtro_base].

    Usa normalização estrita para evitar problemas de máscara/pontuação/zeros.
    Ignora valores vazios/N/A.
    Retorna: (df_base_filtrado, qtd_valores_nota, qtd_linhas_base_antes, qtd_linhas_base_depois)
    """
    if col_filtro_nota not in df_nota.columns:
        raise ValueError(f"Coluna de filtro da NOTA '{col_filtro_nota}' não existe.")
    if col_filtro_base not in df_base.columns:
        raise ValueError(f"Coluna de filtro da BASE '{col_filtro_base}' não existe.")

    # Valores únicos (normalizados) da nota
    serie_n = df_nota[col_filtro_nota].map(normalizar_chave_estrita)
    valores_nota = {v for v in serie_n.dropna().unique() if v and v not in VALORES_VAZIOS_CHAVE}

    antes = len(df_base)

    if not valores_nota:
        # Nada para filtrar
        return df_base.copy(), 0, antes, antes

    # Normaliza a coluna da base e filtra por membership
    serie_b = df_base[col_filtro_base].map(normalizar_chave_estrita)
    mask = serie_b.isin(valores_nota)
    df_filtrado = df_base.loc[mask].copy()

    depois = len(df_filtrado)
    return df_filtrado, len(valores_nota), antes, depois


# Define qual coluna será utilizada como chave de merge, para determinada linha. Retorna o valor da linha nessa coluna
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

# Retorna o nome da coluna a ser utilizada naquela linha
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


def sugerir_pares_colunas(df_nota, df_base, limite_sugestoes=None, min_score=0.70, top_debug=40):
    """
    Sugere pares de colunas por score de correspondência de valores.

    Retorna TODOS os pares com score >= min_score.
    - min_score padrão: 0.70 (70%)
    - limite_sugestoes=None -> sem limite (retorna tudo acima do min_score)
    - limite_sugestoes=int  -> retorna apenas top-N pares acima do min_score
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

            score = (
                (0.55 * score_set)
                + (0.30 * score_freq)
                + (0.10 * score_cob)
                + (0.05 * score_nome)
            )

            resultados.append((cn, cb, float(score)))

    resultados.sort(key=lambda x: x[2], reverse=True)

    # DEBUG: mostrar top pares por score (antes do filtro)
    if top_debug and top_debug > 0:
        print("\n=== DEBUG SCORE: TOP PARES (antes do filtro min_score) ===")
        for cn, cb, sc in resultados[:top_debug]:
            print(f"{cn:30s}  <->  {cb:30s}   score={sc:.4f}")
        print("=========================================================\n")

    # ✅ filtro 70%+
    sugestoes = [(cn, cb, sc) for (cn, cb, sc) in resultados if sc >= min_score]

    # limite opcional (se quiser top-N)
    if isinstance(limite_sugestoes, int) and limite_sugestoes > 0:
        sugestoes = sugestoes[:limite_sugestoes]

    return sugestoes

def metricas_unicidade(df, col, amostra=8000):
    """
    Métricas para chave:
    - cobertura: % de linhas com valor útil (não vazio/N/A) dentro da amostra
    - unicidade: % de valores distintos entre os preenchidos (nunique / n_valid)
      (1.0 = todos diferentes -> ótimo candidato a identificador)
    Retorna: (unicidade, cobertura, n_valid, nunique)
    """
    if col not in df.columns:
        return 0.0, 0.0, 0, 0

    d = df.head(amostra) if amostra else df
    total = len(d)
    if total == 0:
        return 0.0, 0.0, 0, 0

    s = d[col].map(normalizar_chave_estrita)
    # remove vazios semânticos
    s = s[~s.map(eh_vazio_semantico)]

    n_valid = int(len(s))
    if n_valid == 0:
        return 0.0, 0.0, 0, 0

    nunique = int(s.nunique(dropna=True))
    cobertura = n_valid / total
    unicidade = nunique / n_valid
    return float(unicidade), float(cobertura), n_valid, nunique


def sugerir_chaves_por_unicidade(
    df_nota,
    df_base,
    limite=2,
    min_score_match=0.25,     # apenas para garantir que as colunas "correspondem"
    min_cobertura=0.20,       # evita escolher coluna quase vazia
    min_unicidade=0.60,       # exige que seja bem identificador
    amostra=8000,
    top_debug=10,
):
    """
    Sugere pares de colunas para CHAVE (principal + fallback) priorizando UNICIDADE.

    Estratégia:
    1) Gera pares candidatos por compatibilidade (sugerir_pares_colunas) com threshold leve.
    2) Ranqueia candidatos por unicidade (principal fator), com penalização leve de cobertura.
       key_score = min(unic_nota, unic_base) * min(cob_nota, cob_base)

    Retorna lista ordenada:
      [(col_nota, col_base, key_score, unic_nota, cob_nota, unic_base, cob_base, match_score), ...]
    """

    # 1) candidatos por compatibilidade (não é ranking final, só filtro)
    candidatos = sugerir_pares_colunas(
        df_nota,
        df_base,
        limite_sugestoes=None,
        min_score=min_score_match,
        top_debug=0
    )

    # cache de métricas por coluna
    cache_nota = {}
    cache_base = {}

    def mnota(c):
        if c not in cache_nota:
            cache_nota[c] = metricas_unicidade(df_nota, c, amostra=amostra)
        return cache_nota[c]

    def mbase(c):
        if c not in cache_base:
            cache_base[c] = metricas_unicidade(df_base, c, amostra=amostra)
        return cache_base[c]

    ranqueados = []
    for cn, cb, match_score in candidatos:
        unic_n, cob_n, _, _ = mnota(cn)
        unic_b, cob_b, _, _ = mbase(cb)

        # filtros mínimos (chave precisa "existir" em quantidade e ser identificadora)
        if cob_n < min_cobertura or cob_b < min_cobertura:
            continue
        if unic_n < min_unicidade or unic_b < min_unicidade:
            continue

        # 2) ranking por unicidade (principal fator) + cobertura como gate/penalização
        key_score = min(unic_n, unic_b) * min(cob_n, cob_b)

        ranqueados.append((cn, cb, float(key_score), unic_n, cob_n, unic_b, cob_b, float(match_score)))

    ranqueados.sort(key=lambda x: x[2], reverse=True)

    # 3) escolhe principal + fallback sem repetir colunas
    usados_n = set()
    usados_b = set()
    out = []
    for cn, cb, ksc, unic_n, cob_n, unic_b, cob_b, msc in ranqueados:
        if cn in usados_n or cb in usados_b:
            continue
        usados_n.add(cn)
        usados_b.add(cb)
        out.append((cn, cb, ksc, unic_n, cob_n, unic_b, cob_b, msc))
        if len(out) >= limite:
            break

    if top_debug:
        print("\n=== DEBUG CHAVES (UNICIDADE) ===")
        for item in out[:top_debug]:
            cn, cb, ksc, unic_n, cob_n, unic_b, cob_b, msc = item
            print(
                f"{cn:28s} <-> {cb:28s} | key_score={ksc:.3f} | "
                f"unic_n={unic_n:.3f} cob_n={cob_n:.2f} | unic_b={unic_b:.3f} cob_b={cob_b:.2f} | match={msc:.3f}"
            )
        print("================================\n")

    return out

def _serie_key_normalizada(df, col, amostra=5000):
    """Retorna uma Series normalizada estrita (para chave), limitada a amostra."""
    s = df[col]
    if amostra is not None:
        s = s.head(amostra)
    s = s.map(normalizar_chave_estrita)
    # remove vazios semânticos
    s = s[~s.map(eh_vazio_semantico)]
    return s


def score_identificador_unico(df, col, amostra=5000):
    """
    Mede 'qualidade de chave' para uma coluna:
    - cobertura: % de linhas com valor útil (não vazio/N/A)
    - unicidade: % de distintos entre os preenchidos
    Retorna um score [0..1] e métricas auxiliares.
    """
    total = min(len(df), amostra) if amostra is not None else len(df)
    if total <= 0:
        return 0.0, {"cobertura": 0.0, "unicidade": 0.0, "n_valid": 0, "n_total": 0}

    s = _serie_key_normalizada(df, col, amostra=amostra)
    n_valid = int(len(s))
    if n_valid == 0:
        return 0.0, {"cobertura": 0.0, "unicidade": 0.0, "n_valid": 0, "n_total": total}

    cobertura = n_valid / total
    nunique = int(s.nunique(dropna=True))
    unicidade = nunique / n_valid  # 1.0 = todos diferentes

    # penaliza colunas com baixa cobertura (ex.: quase sempre vazio)
    # e também penaliza colunas extremamente pouco únicas (ex.: commodity)
    score = (0.55 * unicidade) + (0.45 * cobertura)
    return float(score), {
        "cobertura": float(cobertura),
        "unicidade": float(unicidade),
        "n_valid": int(n_valid),
        "n_total": int(total),
    }


def sugerir_colunas_chave(df_nota, df_base, limite=2, min_match=0.55, min_key_quality=0.55, amostra=5000, top_debug=15):
    """
    Sugere pares (col_nota, col_base) para serem usados como CHAVE (principal + fallbacks).
    Critérios:
      - match_score alto (compatibilidade entre colunas)
      - score de identificador único alto em ambos os lados (qualidade de chave)
    Retorna lista ordenada: [(col_nota, col_base, key_score, match_score, qual_nota, qual_base), ...]
    """
    # pega muitos pares candidatos por match (sem limitar e com threshold baixo)
    candidatos = sugerir_pares_colunas(
        df_nota, df_base,
        limite_sugestoes=None,
        min_score=0.0,         # não filtra aqui; filtra abaixo com min_match
        top_debug=0            # evita duplicar print
    )

    # calcula qualidade de chave por coluna (cache)
    qual_nota_cache = {}
    qual_base_cache = {}

    def qual_nota(c):
        if c not in qual_nota_cache:
            qual_nota_cache[c] = score_identificador_unico(df_nota, c, amostra=amostra)
        return qual_nota_cache[c]

    def qual_base(c):
        if c not in qual_base_cache:
            qual_base_cache[c] = score_identificador_unico(df_base, c, amostra=amostra)
        return qual_base_cache[c]

    ranqueados = []
    for cn, cb, match_score in candidatos:
        if match_score < min_match:
            continue

        qn, meta_n = qual_nota(cn)
        qb, meta_b = qual_base(cb)

        # precisa ser "boa chave" nos dois lados
        if qn < min_key_quality or qb < min_key_quality:
            continue

        # composição do score final de chave
        # (dá mais peso ao match entre colunas, mas exige qualidade)
        key_score = (0.65 * match_score) + (0.35 * min(qn, qb))

        ranqueados.append((cn, cb, float(key_score), float(match_score), float(qn), float(qb)))

    ranqueados.sort(key=lambda x: x[2], reverse=True)

    # evita recomendar o mesmo par repetido e tenta manter diversidade
    usados_n = set()
    usados_b = set()
    out = []
    for cn, cb, key_score, match_score, qn, qb in ranqueados:
        if cn in usados_n or cb in usados_b:
            continue
        usados_n.add(cn)
        usados_b.add(cb)
        out.append((cn, cb, key_score, match_score, qn, qb))
        if len(out) >= limite:
            break

    if top_debug and out:
        print("\n=== DEBUG CHAVES (auto) ===")
        for cn, cb, ksc, msc, qn, qb in out[:top_debug]:
            print(f"{cn:30s} <-> {cb:30s} | key_score={ksc:.3f} | match={msc:.3f} | qual_nota={qn:.3f} | qual_base={qb:.3f}")
        print("===========================\n")

    return out


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
    Remove linhas totalmente vazias (considerando espaços, NaN, N/A, '-', etc.).
    Retorna (df_filtrado, qtd_removida).
    """
    if df.empty:
        return df.copy(), 0

    def _linha_vazia(row):
        for v in row:
            if not eh_vazio_semantico(v):
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
    # --- filtro extra: remove linhas totalmente vazias (muito comum no fim do Excel) ---
    # Considera "vazio" também N/A, NA, NULL, '-', espaços invisíveis etc.
    cols_nota_check = [c for c in df_nota.columns if not str(c).startswith("_")]

    mask_linha_util_nota = ~df_nota[cols_nota_check].apply(
        lambda r: all(eh_vazio_semantico(v) for v in r),
        axis=1
    )

    removidas_extra = int((~mask_linha_util_nota).sum())
    if removidas_extra:
        df_nota = df_nota.loc[mask_linha_util_nota].copy().reset_index(drop=True)
        linhas_branco_nota += removidas_extra
    
    # --- opcional: truncar df_nota até a última linha com algum valor útil ---
    has_any = df_nota[cols_nota_check].apply(
        lambda r: any(not eh_vazio_semantico(v) for v in r),
        axis=1
    )
    if has_any.any():
        last_valid = int(has_any[has_any].index.max())
        df_nota = df_nota.loc[:last_valid].copy().reset_index(drop=True)
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

        # Dados "brutos" da nota (sem colunas internas)
        cols_nota_saida = [c for c in df_nota.columns if c not in {"_KEY", "_KEY_ORIGEM", "_NOTA_IDX", "_RANK_KEY"}]
        registro = {c: row[c] if c in row.index else "" for c in cols_nota_saida}

        registro["_KEY"] = row.get("_KEY", "")
        registro["_KEY_ORIGEM"] = row.get("_KEY_ORIGEM", "")
        registro["_MATCH"] = row["_MATCH"]
        registro["_KEY_ORIGEM_BASE"] = row.get("_KEY_ORIGEM_BASE", "")
        registro["_BASE_IDX"] = row.get("_BASE_IDX", None)

        # ✅ Se a linha inteira da NOTA está vazia/N/A -> IGNORA (não entra no output)
        if all(eh_vazio_semantico(registro.get(c, "")) for c in cols_nota_saida):
            continue

        linha_divergente = False
        tem_celula_vazia = False

        # Se não tem correspondência na base
        if row["_MATCH"] == "left_only":
            # Se a chave está vazia/N/A, isso não é divergência -> é dado ausente
            if eh_vazio_semantico(registro["_KEY"]):
                status = "CÉLULA VAZIA"
                tem_celula_vazia = True
            else:
                status = "SEM CORRESPONDÊNCIA NA BASE"

            # Base vazia para todos os pares
            for cn, cb in pares:
                stats_pares[(cn, cb)]["sem_base"] += 1
                registro[f"__BASE__{cb}"] = ""

            registro["STATUS LINHA"] = status
            registros_saida.append(registro)
            continue

        # Se veio da base mas não existe na nota (sobras) -> mantém comportamento
        if row["_MATCH"] == "right_only":
            # (em geral este caso nem passa por df_match, mas mantemos por segurança)
            for cn, cb in pares:
                registro[f"__BASE__{cb}"] = row.get(cb, "")
                stats_pares[(cn, cb)]["nao_presente_nota"] += 1
            registro["STATUS LINHA"] = "NÃO PRESENTE NA NOTA"
            registros_saida.append(registro)
            continue

        # ✅ Caso normal: match "both"
        for cn, cb in pares:
            cb_no_match = f"{cb}__BASE_RAW" if f"{cb}__BASE_RAW" in row.index else cb
            valor_base = row.get(cb_no_match, "")
            registro[f"__BASE__{cb}"] = valor_base

            valor_nota = row.get(cn, "")

            # ✅ Se a célula da NOTA está vazia/N/A -> não compara (vira amarelo depois)
            if eh_vazio_semantico(valor_nota):
                tem_celula_vazia = True
                # ✅ conta como "não presente na nota" (campo vazio na nota)
                stats_pares[(cn, cb)]["nao_presente_nota"] += 1
                # não compara e não entra como ok/divergente
                continue
            # se a nota tem valor, compara normal
            if valores_equivalentes(valor_nota, valor_base):
                stats_pares[(cn, cb)]["ok"] += 1
            else:
                stats_pares[(cn, cb)]["divergente"] += 1
                linha_divergente = True

        # ✅ Prioridade de status:
        # - Se divergente em alguma célula preenchida -> DIVERGENTE
        # - Senão se teve célula vazia -> CÉLULA VAZIA
        # - Senão -> OK
        if linha_divergente:
            status = "DIVERGENTE"
        elif tem_celula_vazia:
            status = "CÉLULA VAZIA"
        else:
            status = "OK"

        registro["STATUS LINHA"] = status
        registros_saida.append(registro)

    sobras_vazias_ignoradas = 0
    # Inclui também _KEY para não aceitar sobras com chave vazia/semântica
    # Contexto real: chaves base + colunas comparadas (não inclui _KEY)
    colunas_contexto_base = list(dict.fromkeys(chaves_base + [cb for _, cb in pares]))

    for _, row in df_base_sobras.iterrows():
        # Ignora sobras totalmente vazias no contexto da validação
        # (chaves + colunas comparadas). Se não há nada útil, não entra no output.
        if all(eh_vazio_semantico(row.get(c, "")) for c in colunas_contexto_base):
            sobras_vazias_ignoradas += 1
            continue

        registro = {c: "" for c in df_nota.columns if c not in {"_KEY", "_KEY_ORIGEM", "_NOTA_IDX", "_RANK_KEY"}}
        registro["_KEY"] = row.get("_KEY", "")
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

    # ======================================================
    # ✅ FILTRO FINAL: remove linhas do output onde a "nota" está totalmente vazia/N/A
    # (isso elimina o rabicho do Excel que ainda aparece amarelo no fim)
    # ======================================================
    cols_nota_saida = [c for c in df_nota.columns if c not in {"_KEY", "_KEY_ORIGEM", "_NOTA_IDX", "_RANK_KEY"}]
    cols_presentes = [c for c in cols_nota_saida if c in df_out.columns]

    if cols_presentes:
        mask_out_util = ~df_out[cols_presentes].apply(
            lambda r: all(eh_vazio_semantico(v) for v in r),
            axis=1
        )
        removidas_no_out = int((~mask_out_util).sum())
        if removidas_no_out:
            print(f"[DEBUG] Removendo {removidas_no_out} linhas totalmente vazias/N/A do output (df_out).")
        df_out = df_out.loc[mask_out_util].copy().reset_index(drop=True)

    # Agora monta o df_excel já filtrado
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

                # ✅ Se a célula da NOTA é vazia/N/A -> amarelo e NÃO compara
                if eh_vazio_semantico(df_out.at[row_idx, cn]):
                    fmt = fmt_yellow
                    worksheet.write(row_idx + 1, col_idx, valor_excel, fmt)
                    continue

                if match == "left_only":
                    # Sem par na base (mas nota tem valor): isso é divergência real
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
            elif status_val in {"CÉLULA VAZIA", "SEM CORRESPONDÊNCIA NA BASE"}:
                fmt_status = fmt_yellow
            else:
                fmt_status = fmt_yellow
        # --- aba RESUMO ---
        total = len(df_out)
        matched = int((df_out["_MATCH"] == "both").sum())
        sem_match = int((df_out["_MATCH"] == "left_only").sum())
        nao_presente_nota = int((df_out["_MATCH"] == "right_only").sum())
        linhas_ok = int((df_out["STATUS LINHA"] == "OK").sum())
        linhas_div = sum(1 for s in df_out["STATUS LINHA"] if s == "DIVERGENTE")
        linhas_celula_vazia = int((df_out["STATUS LINHA"] == "CÉLULA VAZIA").sum())

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
                ("Linhas com CÉLULA VAZIA (não comparadas)", linhas_celula_vazia),
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

    # --- feedback final ---
    set_nota = set(df_nota["_KEY"]) - {""}
    set_base = set(df_base_enriq["_KEY"]) - {""}

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

# ---------------- selecionar colunas de filtro (NOTA -> filtra BASE) ----------------
# Ideal: filtrar a base usando os números de nota existentes na planilha de nota.
# Ex.: NOTA FISCAL (nota) -> NOTA DE ENTRADA (base)

filtro_win = tk.Toplevel()
filtro_win.title("Filtro da Base (opcional) — usar valores da Nota")

f_frame = ttk.Frame(filtro_win, padding=10)
f_frame.pack()

ttk.Label(
    f_frame,
    text=(
        "Escolha uma coluna na NOTA e uma coluna na BASE para filtrar a BASE.\n"
        "O programa vai pegar os valores únicos da NOTA e manter na BASE apenas as linhas\n"
        "cujo valor da coluna escolhida exista nesse conjunto (comparação com normalização estrita).\n\n"
        "Exemplo comum: NOTA FISCAL (NOTA) ↔ NOTA DE ENTRADA (BASE)."
    ),
    foreground="#555",
    justify="left",
).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

ttk.Label(f_frame, text="Coluna de filtro na NOTA").grid(row=1, column=0, sticky="w")
cb_filtro_nota = ttk.Combobox(
    f_frame,
    values=df_nota.columns.tolist(),
    state="readonly",
    width=40
)
cb_filtro_nota.grid(row=2, column=0, padx=(0, 10), pady=(0, 8), sticky="w")

ttk.Label(f_frame, text="Coluna de filtro na BASE").grid(row=1, column=1, sticky="w")
cb_filtro_base = ttk.Combobox(
    f_frame,
    values=df_base.columns.tolist(),
    state="readonly",
    width=40
)
cb_filtro_base.grid(row=2, column=1, pady=(0, 8), sticky="w")

usar_filtro_var = tk.BooleanVar(value=True)
chk = ttk.Checkbutton(
    f_frame,
    text="Aplicar filtro na BASE (recomendado quando a BASE é maior)",
    variable=usar_filtro_var
)
chk.grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 8))

# Tentativa de autopreenchimento (se existirem)
# Ajuste os nomes se quiser, mas deixei só “tentativas”.
for candidato in ["NOTA FISCAL", "NF", "NF_ENTRADA", "NUMERO NOTA"]:
    if candidato in df_nota.columns:
        cb_filtro_nota.set(candidato)
        break

for candidato in ["NOTA DE ENTRADA", "NF ENTRADA", "NF_ENTRADA", "NOTA FISCAL"]:
    if candidato in df_base.columns:
        cb_filtro_base.set(candidato)
        break

filtro_selecoes = {}

def confirmar_filtro():
    if not usar_filtro_var.get():
        filtro_selecoes["aplicar"] = False
        filtro_win.destroy()
        return

    coln = cb_filtro_nota.get()
    colb = cb_filtro_base.get()
    if not coln or not colb:
        messagebox.showerror("Erro", "Selecione as duas colunas de filtro (NOTA e BASE) ou desmarque a opção de aplicar filtro.")
        return

    filtro_selecoes["aplicar"] = True
    filtro_selecoes["nota"] = coln
    filtro_selecoes["base"] = colb
    filtro_win.destroy()

ttk.Button(f_frame, text="Continuar", command=confirmar_filtro).grid(row=4, column=0, columnspan=2, pady=10)

filtro_win.wait_window()

# Aplica o filtro (se selecionado)
if filtro_selecoes.get("aplicar"):
    try:
        df_base_filtrado, qtd_vals, antes, depois = filtrar_base_por_nota(
            df_nota, df_base,
            filtro_selecoes["nota"],
            filtro_selecoes["base"]
        )

        # Se o filtro zerar a base, oferece continuar sem filtro
        if depois == 0 and antes > 0:
            resp = messagebox.askyesno(
                "Filtro resultou em 0 linhas",
                (
                    f"O filtro removeu todas as linhas da BASE.\n\n"
                    f"Coluna NOTA: {filtro_selecoes['nota']}\n"
                    f"Coluna BASE: {filtro_selecoes['base']}\n"
                    f"Valores únicos na NOTA (não vazios): {qtd_vals}\n"
                    f"Linhas BASE antes: {antes}\n"
                    f"Linhas BASE depois: {depois}\n\n"
                    f"Deseja continuar SEM aplicar filtro?"
                )
            )
            if resp:
                # mantém df_base original
                pass
            else:
                raise SystemExit
        else:
            df_base = df_base_filtrado
            messagebox.showinfo(
                "Filtro aplicado",
                (
                    f"Filtro aplicado com sucesso.\n\n"
                    f"Coluna NOTA: {filtro_selecoes['nota']}\n"
                    f"Coluna BASE: {filtro_selecoes['base']}\n"
                    f"Valores únicos na NOTA (não vazios): {qtd_vals}\n"
                    f"Linhas BASE antes: {antes}\n"
                    f"Linhas BASE depois: {depois}"
                )
            )
    except Exception as e:
        messagebox.showerror("Erro ao aplicar filtro", f"{type(e).__name__}: {e}")
        raise

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


# Sugestão automática de chaves por UNICIDADE (principal + fallback)
sug_chaves = sugerir_chaves_por_unicidade(
    df_nota,
    df_base,
    limite=2,               # principal + 1 fallback
    min_score_match=0.25,   # apenas garante que as colunas se "parecem"
    min_cobertura=0.20,     # evita coluna quase vazia
    min_unicidade=0.60,     # identificador bem único
    amostra=8000,
    top_debug=10
)

if sug_chaves:
    for cn, cb, *_ in sug_chaves:
        adicionar_chave(cn, cb)
else:
    # se não inferir nada confiável, deixa 1 linha vazia pro usuário escolher
    adicionar_chave()


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

# ---------------- mapeamento de colunas (APENAS AUTOMÁTICO) ----------------
# O usuário não recebe mais pares fixos. Apenas o botão "Sugerir pares (score)"
# preenche automaticamente os pares acima de 70%.

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
    # Apenas sugestões acima de 70% (min_score=0.70)
    sugestoes = sugerir_pares_colunas(
        df_nota,
        df_base,
        limite_sugestoes=None,   # não limita
        min_score=0.70,          # 70%+
        top_debug=40             # printa top scores no console (opcional)
    )

    if not sugestoes:
        messagebox.showwarning(
            "Sem sugestões",
            "Não foi possível inferir pares com score ≥ 0.70.",
        )
        return

    # Pares já existentes na UI (para evitar duplicar o MESMO par)
    pares_existentes = set()
    for w_n, w_b in pares_widgets:
        cn_atual = (w_n.get() or "").strip()
        cb_atual = (w_b.get() or "").strip()
        if cn_atual and cb_atual:
            pares_existentes.add((cn_atual, cb_atual))

    # Queremos "mostrar tudo" — inclusive os que já estavam configurados.
    # Então vamos:
    # - garantir que exista uma linha para cada sugestão NOVA (não duplicada)
    sugestoes_novas = [(cn, cb, score) for (cn, cb, score) in sugestoes if (cn, cb) not in pares_existentes]

    adicionadas = 0
    for cn, cb, score in sugestoes_novas:
        # tenta usar uma linha vazia primeiro
        alvo = None
        for i, (w_n, w_b) in enumerate(pares_widgets):
            if not w_n.get() and not w_b.get():
                alvo = i
                break

        if alvo is None:
            adicionar_par()
            alvo = len(pares_widgets) - 1

        w_n, w_b = pares_widgets[alvo]
        w_n.set(cn)
        w_b.set(cb)
        adicionadas += 1

    # Pop-up mostrando TODAS (ou limite para não ficar gigante)
    MAX_MOSTRAR = 50  # ajuste; coloque None para mostrar tudo
    lista_popup = sugestoes[:MAX_MOSTRAR] if MAX_MOSTRAR is not None else sugestoes
    extra = (len(sugestoes) - len(lista_popup)) if MAX_MOSTRAR is not None else 0

    texto = "\n".join(
        f"- {cn} ↔ {cb} (score={score:.2f})"
        for cn, cb, score in lista_popup
    )
    if extra > 0:
        texto += f"\n... (+{extra} pares)"

    messagebox.showinfo(
        "Sugestões aplicadas",
        (
            f"Foram encontradas {len(sugestoes)} sugestões com score ≥ 0.70.\n"
            f"Foram adicionadas {adicionadas} novas linhas na tela.\n\n"
            f"Sugestões:\n{texto}"
        ),
    )

def prosseguir():
    pares = []
    for cn, cb in pares_widgets:
        if not cn.get() or not cb.get():
            # Agora, como é automático, vamos permitir ignorar linhas em branco
            # (p.ex. se houver sobras vazias na UI).
            continue
        pares.append((cn.get(), cb.get()))

    if not pares:
        messagebox.showerror("Erro", "Nenhum par de colunas foi definido. Clique em 'Sugerir pares (score)'.")
        return

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

# Opcional: começar com 1 linha vazia para o usuário ver o layout
# (se você quiser zero linhas iniciais, apague esta linha)
adicionar_par()

# Botões: agora não tem mais "Adicionar par" por padrão (somente automático).
ttk.Button(frame, text="Sugerir pares (score ≥ 0.70)", command=aplicar_sugestoes_automaticas)\
    .grid(row=100, column=0, pady=10, padx=5)

ttk.Button(frame, text="Prosseguir", command=prosseguir)\
    .grid(row=100, column=2, pady=10, padx=5)

map_win.wait_window()