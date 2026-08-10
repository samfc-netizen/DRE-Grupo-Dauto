# GERAL.py
import re
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px
import os
import glob
import csv
from pathlib import Path
from io import StringIO

# =========================
# Normalização de texto (para filtros robustos)
# =========================
_DED_EXCL_DRE = "02.07.008-ICMS- SUBSTITUIÇÃO TRIBUTARIA"

def _norm_txt(s: object) -> str:
    """Normaliza texto: remove acentos, padroniza hífens/espaços e coloca em minúsculas."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    t = str(s)
    # NBSP e hífens diferentes
    t = t.replace("\u00a0", " ").replace("–", "-").replace("—", "-")
    # remove acentos
    t = unicodedata.normalize("NFKD", t)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    # normaliza espaços
    t = re.sub(r"\s+", " ", t).strip().lower()
    return t


MESES_PT = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
MES_NUM_TO_PT = {1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
                 7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"}
MES_PT_TO_NUM = {v: k for k, v in MES_NUM_TO_PT.items()}


# =========================
# Helpers
# =========================
def to_num(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip()
    if s == "":
        return 0.0
    s = s.replace("\u00a0", " ").replace("R$", "").strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def format_brl(x) -> str:
    try:
        return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"


def fmt_pct(x) -> str:
    try:
        return f"{float(x):,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00%"


def fmt_brl_display(x) -> str:
    return f"R$ {format_brl(x)}"


def inject_sticky_table_css():
    st.markdown(
        """
        <style>
        .sticky-table-wrap {
            overflow-x: auto;
            border: 1px solid rgba(49, 51, 63, 0.2);
            border-radius: 8px;
            background: white;
            margin-bottom: 0.5rem;
        }
        .sticky-table {
            border-collapse: separate;
            border-spacing: 0;
            min-width: 100%;
            font-size: 0.92rem;
        }
        .sticky-table th, .sticky-table td {
            padding: 8px 10px;
            border-bottom: 1px solid rgba(49, 51, 63, 0.12);
            white-space: nowrap;
            text-align: right;
        }
        .sticky-table th {
            position: sticky;
            top: 0;
            z-index: 3;
            background: #f6f8fb;
            font-weight: 700;
        }
        .sticky-table th:first-child, .sticky-table td:first-child {
            position: sticky;
            left: 0;
            z-index: 2;
            text-align: left;
            background: white;
            min-width: 260px;
            max-width: 260px;
            white-space: normal;
        }
        .sticky-table th:first-child {
            z-index: 4;
            background: #f6f8fb;
        }
        .sticky-table tr:hover td {
            background: #fafafa;
        }
        .sticky-table tr:hover td:first-child {
            background: #f0f3f9;
        }
        .sticky-table .row-strong td:first-child {
            font-weight: 800;
        }
        .sticky-table .pos-strong {
            color: #1f4e79;
            font-weight: 800;
        }
        .sticky-table .neg-strong {
            color: #c00000;
            font-weight: 800;
        }
        .sticky-table .text-left { text-align: left; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_sticky_table(df: pd.DataFrame, value_cols=None, pct_cols=None, highlight_row_label=None):
    value_cols = set(value_cols or [])
    pct_cols = set(pct_cols or [])
    inject_sticky_table_css()

    cols = list(df.columns)
    html = ['<div class="sticky-table-wrap"><table class="sticky-table"><thead><tr>']
    for c in cols:
        cls = 'text-left' if c == cols[0] else ''
        html.append(f'<th class="{cls}">{c}</th>')
    html.append('</tr></thead><tbody>')

    for _, row in df.iterrows():
        is_highlight = str(row.iloc[0]) == str(highlight_row_label) if highlight_row_label is not None else False
        tr_cls = 'row-strong' if is_highlight else ''
        html.append(f'<tr class="{tr_cls}">')
        for j, c in enumerate(cols):
            val = row[c]
            classes = []
            if j == 0:
                classes.append('text-left')
            if c in value_cols:
                num = to_num(val)
                display = fmt_brl_display(num)
                if is_highlight:
                    classes.append('neg-strong' if num < 0 else 'pos-strong')
            elif c in pct_cols:
                num = to_num(val)
                display = fmt_pct(num)
                if is_highlight:
                    classes.append('neg-strong' if num < 0 else 'pos-strong')
            else:
                display = '' if pd.isna(val) else str(val)
            html.append(f'<td class="{" ".join(classes)}">{display}</td>')
        html.append('</tr>')
    html.append('</tbody></table></div>')
    st.markdown(''.join(html), unsafe_allow_html=True)


def parse_mes(v):
    """Aceita 1..12, '01', 'JAN', 'Janeiro' e devolve mes_num."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip().upper()
    if s.isdigit():
        m = int(s)
        return m if 1 <= m <= 12 else None
    mapa = {
        "JANEIRO": 1, "JAN": 1,
        "FEVEREIRO": 2, "FEV": 2,
        "MARCO": 3, "MARÇO": 3, "MAR": 3,
        "ABRIL": 4, "ABR": 4,
        "MAIO": 5, "MAI": 5,
        "JUNHO": 6, "JUN": 6,
        "JULHO": 7, "JUL": 7,
        "AGOSTO": 8, "AGO": 8,
        "SETEMBRO": 9, "SET": 9,
        "OUTUBRO": 10, "OUT": 10,
        "NOVEMBRO": 11, "NOV": 11,
        "DEZEMBRO": 12, "DEZ": 12,
    }
    return mapa.get(s)


@st.cache_data(show_spinner=False)
def get_sheet_names(excel_path: str, sig):
    try:
        return pd.ExcelFile(excel_path).sheet_names
    except Exception:
        return []

@st.cache_data(show_spinner=False)
def read_sheet(excel_path: str, sheet_name: str, sig):
    """Lê uma aba do Excel com cache (melhora muito a navegação no Streamlit)."""
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except Exception:
        return None
    df.columns = [str(c).strip() for c in df.columns]
    return df

@st.cache_data(show_spinner=False)
def prep_geral_year(excel_path: str, ano_ref: int, sig):
    """Carrega e prepara a aba DRE E DFC GERAL 1x por ano (parse de datas/valores)."""
    df = read_sheet(excel_path, "DRE E DFC GERAL", sig)
    if df is None:
        return None
    g = df.copy()
    g["_dt"] = pd.to_datetime(g.get("DTA.PAG"), errors="coerce", dayfirst=True)
    g["_ano"] = g["_dt"].dt.year
    g["_mes"] = g["_dt"].dt.month
    g["_v"] = g.get("VAL.PAG").apply(to_num) if "VAL.PAG" in g.columns else 0.0
    g = g[g["_ano"] == int(ano_ref)]
    return g
@st.cache_data(show_spinner=False)
def prep_dre_sheet_year(excel_path: str, ano_ref: int, sig):
    """Carrega e prepara a aba DRE (por ano).
    Tenta usar DTA.PAG (se existir). Caso contrário, tenta MES/MÊS e ANO.
    """
    df = read_sheet(excel_path, "DRE", sig)
    if df is None:
        return None
    d = df.copy()

    # Datas / Mês / Ano
    if "DTA.PAG" in d.columns:
        dt = pd.to_datetime(d.get("DTA.PAG"), errors="coerce", dayfirst=True)
        d["_ano"] = dt.dt.year
        d["_mes"] = dt.dt.month
    else:
        # Aceita variações comuns de colunas
        col_ano = "ANO" if "ANO" in d.columns else ("Ano" if "Ano" in d.columns else None)
        col_mes = None
        for c in ["MÊS", "MES", "MÊS.", "MES.", "MÊS ", "MES "]:
            if c in d.columns:
                col_mes = c
                break
        if col_ano is None or col_mes is None:
            # sem base para ano/mês
            d["_ano"] = pd.NA
            d["_mes"] = pd.NA
        else:
            d["_ano"] = pd.to_numeric(d[col_ano], errors="coerce").astype("Int64")
            d["_mes"] = pd.to_numeric(d[col_mes], errors="coerce").astype("Int64")

    # Valor
    if "VAL.PAG" in d.columns:
        d["_v"] = d["VAL.PAG"].apply(to_num)
    elif "VALOR" in d.columns:
        d["_v"] = d["VALOR"].apply(to_num)
    elif "VAL" in d.columns:
        d["_v"] = d["VAL"].apply(to_num)
    else:
        d["_v"] = 0.0

    # Filtra ano
    try:
        d = d[d["_ano"].astype("Int64") == int(ano_ref)]
    except Exception:
        d = d.iloc[0:0]
    return d


@st.cache_data(show_spinner=False)
def prep_impostos_folha_dre(excel_path: str, ano_ref: int, sig):
    """IMPOSTOS E FOLHA para DRE: considera shift +1 mês e filtra pelo ano de referência."""
    df = read_sheet(excel_path, "IMPOSTOS E FOLHA", sig)
    if df is None:
        return None
    i = df.copy()
    d = pd.to_datetime(i.get("DTA.PAG"), errors="coerce", dayfirst=True)
    d_ref = d + pd.offsets.MonthBegin(1)
    i["_ano_ref"] = d_ref.dt.year
    i["_mes_ref"] = d_ref.dt.month
    i["_v"] = i.get("VAL.PAG").apply(to_num) if "VAL.PAG" in i.columns else 0.0
    i = i[i["_ano_ref"] == int(ano_ref)]
    return i

@st.cache_data(show_spinner=False)
def prep_impostos_folha_dfc(excel_path: str, ano_ref: int, sig):
    """IMPOSTOS E FOLHA para DFC: usa o mês/ano do pagamento (sem shift)."""
    df = read_sheet(excel_path, "IMPOSTOS E FOLHA", sig)
    if df is None:
        return None
    i = df.copy()
    d = pd.to_datetime(i.get("DTA.PAG"), errors="coerce", dayfirst=True)
    i["_ano"] = d.dt.year
    i["_mes"] = d.dt.month
    i["_v"] = i.get("VAL.PAG").apply(to_num) if "VAL.PAG" in i.columns else 0.0
    i = i[i["_ano"] == int(ano_ref)]
    return i

# (Compat) não usar mais diretamente: mantido só para não quebrar imports antigos
def read_sheet_xls(xls: pd.ExcelFile, sheet_name: str):
    return None


def agg_by_month_from_ano_mes(df, col_value, col_ano="ANO", col_mes="MÊS", ano_ref=None):
    """
    Agrega por mês usando colunas ANO e MÊS.
    Se ano_ref for informado, filtra ANO == ano_ref.
    """
    if col_value not in df.columns or col_mes not in df.columns:
        return None

    tmp = df.copy()

    if col_ano in tmp.columns:
        tmp["_ano"] = pd.to_numeric(tmp[col_ano], errors="coerce")
        if ano_ref is not None:
            tmp = tmp[tmp["_ano"] == int(ano_ref)]
    else:
        tmp["_ano"] = None

    tmp["_mes"] = tmp[col_mes].apply(parse_mes)
    tmp = tmp[tmp["_mes"].notna()].copy()
    tmp["_v"] = tmp[col_value].apply(to_num)

    grp = tmp.groupby("_mes")["_v"].sum()
    return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}


def sintetizar_despesa(nome: str) -> str:
    """
    Ex.: '02.02.007-INSS + IRRF (3 - DESPESAS)' -> '02.02.007-INSS + IRRF'
    Remove sufixos do tipo '(n - DESPESAS)' e parênteses finais.
    """
    if nome is None or (isinstance(nome, float) and pd.isna(nome)):
        return "—"
    s = str(nome).strip()
    s = re.sub(r"\s*\(\s*\d+\s*-\s*DESPESAS\s*\)\s*$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*\([^)]*\)\s*$", "", s).strip()
    s = re.sub(r"\s{2,}", " ", s)
    return s if s else "—"


def safe_topn_slider(label: str, n_items: int, default: int = 15, cap: int = 50) -> int:
    """Evita erro quando min == max no slider."""
    if n_items <= 1:
        return n_items
    max_v = min(cap, n_items)
    if max_v <= 5:
        return st.slider(label, 1, max_v, min(default, max_v))
    return st.slider(label, 5, max_v, min(default, max_v))


def pick_hist_key(df: pd.DataFrame) -> str | None:
    """Escolhe a melhor coluna para sintetizar histórico."""
    for c in ["HISTÓRICO", "FAVORECIDO", "DESPESA", "DUPLICATA"]:
        if c in df.columns:
            return c
    return None


def sum_by_prefix_month(df_base: pd.DataFrame, prefix: str, ano_ref: int):
    """
    Soma por mês com base em DTA.PAG e CONTA DE RESULTADO prefixo.
    df_base precisa ter colunas: CONTA DE RESULTADO, DTA.PAG, VAL.PAG.
    """
    tmp = df_base.copy()
    tmp["_dt"] = pd.to_datetime(tmp["DTA.PAG"], errors="coerce", dayfirst=True)
    tmp["_ano"] = tmp["_dt"].dt.year
    tmp["_mes"] = tmp["_dt"].dt.month
    tmp["_v"] = tmp["VAL.PAG"].apply(to_num)
    tmp = tmp[tmp["_ano"] == int(ano_ref)]
    mask = tmp["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)
    grp = tmp[mask].groupby("_mes")["_v"].sum()
    return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}


def sum_by_prefix_prepped(g: pd.DataFrame, prefix: str):
    """Soma por mês usando dataframe já preparado (com _mes e _v)."""
    mask = g["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)
    grp = g[mask].groupby("_mes")["_v"].sum()
    return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}


def _mask_outras_receitas(df: pd.DataFrame) -> pd.Series:
    """Identifica Outras Receitas na aba DRE E DFC GERAL / CONTA DE RESULTADO."""
    conta = df.get("CONTA DE RESULTADO", pd.Series([""] * len(df), index=df.index)).astype(str)
    conta_norm = conta.apply(_norm_txt)
    mask = conta.str.strip().str.startswith("00003 -", na=False)
    mask = mask | conta_norm.str.contains("outras receitas", na=False)
    return mask


def sum_outras_receitas_prepped(g: pd.DataFrame):
    """Soma Outras Receitas por mês usando dataframe já preparado (com _mes e _v)."""
    if g is None or g.empty:
        return {m: 0.0 for m in range(1, 13)}
    mask = _mask_outras_receitas(g)
    grp = g[mask].groupby("_mes")["_v"].sum()
    return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}


def dfc_prefix_map():
    """
    Plano de contas do DFC (conforme você informou):
    - FORNECEDORES = 00012
    """
    return {
        "FORNECEDORES": "00012 -",                   # ✅ AJUSTADO
        "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)": "00004 -",
        "DESPESAS COM PESSOAL": "00006 -",
        "DESPESAS ADMINISTRATIVAS": "00007 -",
        "DESPESAS COMERCIAIS": "00009 -",
        "DESPESAS FINANCEIRAS": "00011 -",
        "RETIRADAS SÓCIOS": "00016 -",
        "INVESTIMENTOS": "00015 -",
        "DESPESAS OPERACIONAIS": "00017 -",
    }


# =========================
# Página 1: DRE Geral
# =========================
def pagina_dre_geral(excel_path, ano_ref, meses_pt_sel=None):
    st.title("DRE Geral — (DRE e DFC GERAL)")

    # Meses selecionados no filtro lateral
    meses_pt = (meses_pt_sel or [])
    meses_pt = meses_pt if len(meses_pt) > 0 else MESES_PT
    meses_nums = [MES_PT_TO_NUM[m] for m in meses_pt]

    df_receita = read_sheet(excel_path, "RECEITA", sig)
    df_nfs = read_sheet(excel_path, "NOTAS EMITIDAS", sig)
    df_geral = read_sheet(excel_path, "DRE E DFC GERAL", sig)
    df_if = read_sheet(excel_path, "IMPOSTOS E FOLHA", sig)

    missing = [n for n, df in [("RECEITA", df_receita), ("NOTAS EMITIDAS", df_nfs),
                               ("DRE E DFC GERAL", df_geral)] if df is None]
    if missing:
        st.error(f"Faltam abas no Excel: {', '.join(missing)}")
        return

    if "RECEITA GRUPO" not in df_receita.columns or "MÊS" not in df_receita.columns:
        st.error("Na aba RECEITA preciso das colunas: 'RECEITA GRUPO' e 'MÊS'.")
        return
    receita_by_month = agg_by_month_from_ano_mes(df_receita, "RECEITA GRUPO", "ANO", "MÊS", ano_ref)

    if "NFS EMITIDAS" not in df_nfs.columns or "MÊS" not in df_nfs.columns:
        st.error("Na aba NOTAS EMITIDAS preciso das colunas: 'NFS EMITIDAS' e 'MÊS'.")
        return
    compras_by_month = agg_by_month_from_ano_mes(df_nfs, "NFS EMITIDAS", "ANO", "MÊS", ano_ref)

    # IMPOSTOS E FOLHA é OPCIONAL nesta página (DRE Geral).
    # Você pediu para DEDUÇÕES e PESSOAL puxarem da aba "DRE E DFC GERAL" (mês +1),
    # então NÃO exigimos esta aba aqui.
    i = None
    if df_if is None:
        pass
    else:
        req_if = {"CONTA DE RESULTADO", "DTA.PAG", "VAL.PAG"}
        if not req_if.issubset(set(df_if.columns)):
            st.warning("Aba IMPOSTOS E FOLHA encontrada, mas faltam colunas (CONTA DE RESULTADO, DTA.PAG, VAL.PAG). Ela será ignorada nesta página.")
        else:
            i = prep_impostos_folha_dre(excel_path, ano_ref, sig)
            if i is None:
                st.warning("Aba IMPOSTOS E FOLHA encontrada, mas não foi possível processá-la. Ela será ignorada nesta página.")

    # ===== DRE: Deduções e Pessoal (mês +1) =====
    # Agora puxamos essas duas linhas da aba "DRE E DFC GERAL", coluna "CONTA DE RESULTADO",
    # sempre com o mês "à frente": exibe mês m usando dados do mês (m+1).
    g_cur = prep_geral_year(excel_path, ano_ref, sig)
    g_next = prep_geral_year(excel_path, int(ano_ref) + 1, sig)
    g = g_cur
    if g is None:
        st.error("Não encontrei a aba DRE E DFC GERAL.")
        return

    def _sum_by_prefix_shift(prefix: str, exclude_icmsst: bool = False) -> dict:
        d = g[g["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)].copy()
        if exclude_icmsst and (not d.empty):
            target_norm = _norm_txt(_DED_EXCL_DRE)
            # tenta excluir pelo texto completo ou pelo código 02.07.008 (apenas no DRE)
            for c in ["DESPESA", "CONTA DE RESULTADO", "HISTÓRICO", "HISTORICO"]:
                if c in d.columns:
                    s_norm = d[c].astype(str).apply(_norm_txt)
                    d = d[~s_norm.str.contains(target_norm, na=False)]
                    d = d[~s_norm.str.contains("02.07.008", na=False)]
                    break
        src = d.groupby("_mes")["_v"].sum()
        # Shift: exibe m usando dados de m+1 (se não existir, zera)
        # Shift: exibe mês m usando dados de (m+1). Em dezembro, busca janeiro do próximo ano.
        src_next = None
        if "g_next" in locals() and g_next is not None:
            dn = g_next[g_next["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)].copy()
            if exclude_icmsst and (not dn.empty):
                target_norm = _norm_txt(_DED_EXCL_DRE)
                for c2 in ["DESPESA", "CONTA DE RESULTADO", "HISTÓRICO", "HISTORICO"]:
                    if c2 in dn.columns:
                        s2 = dn[c2].astype(str).apply(_norm_txt)
                        dn = dn[~s2.str.contains(target_norm, na=False)]
                        dn = dn[~s2.str.contains("02.07.008", na=False)]
                        break
            src_next = dn.groupby("_mes")["_v"].sum()
        out = {}
        for m in range(1, 13):
            if m < 12:
                out[m] = float(src.get(m + 1, 0.0))
            else:
                out[m] = float((src_next.get(1, 0.0) if src_next is not None else 0.0))
        return out

    deducoes_by_month = _sum_by_prefix_shift("00004 -", exclude_icmsst=True)
    pessoal_by_month  = _sum_by_prefix_shift("00006 -", exclude_icmsst=False)

    # Geral por prefixos

    def sum_by_prefix(prefix: str):
        mask = g["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)
        grp = g[mask].groupby("_mes")["_v"].sum()
        return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}

    adm_by_month = sum_by_prefix("00007 -")
    com_by_month = sum_by_prefix("00009 -")
    fin_by_month = sum_by_prefix("00011 -")
    inv_by_month = sum_by_prefix("00015 -")
    op_by_month = sum_by_prefix("00017 -")
    ret_by_month = sum_by_prefix("00016 -")
    outras_receitas_by_month = sum_outras_receitas_prepped(g)
    receita_total_by_month = {m: float(receita_by_month.get(m, 0.0)) + float(outras_receitas_by_month.get(m, 0.0)) for m in range(1, 13)}

    resultado_by_month = {}
    for m in range(1, 13):
        outros = (compras_by_month[m] + deducoes_by_month[m] + pessoal_by_month[m] +
                  adm_by_month[m] + com_by_month[m] + fin_by_month[m] + inv_by_month[m] + op_by_month[m] + ret_by_month[m])
        resultado_by_month[m] = receita_total_by_month[m] - outros


    # Resultado antes das retiradas e despesas financeiras (volta essas duas linhas no resultado)
    resultado_antes_by_month = {m: float(resultado_by_month.get(m, 0.0)) + float(fin_by_month.get(m, 0.0)) + float(ret_by_month.get(m, 0.0)) for m in range(1, 13)}

    linhas = [
        ("+ RECEITA", receita_by_month),
        ("+ OUTRAS RECEITAS", outras_receitas_by_month),
        ("- COMPRAS EMISSÃO", compras_by_month),
        ("- DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", deducoes_by_month),
        ("- DESPESAS COM PESSOAL", pessoal_by_month),
        ("- DESPESAS ADMINISTRATIVAS", adm_by_month),
        ("- DESPESAS COMERCIAIS", com_by_month),
        ("- DESPESAS FINANCEIRAS", fin_by_month),
        ("- RETIRADAS SÓCIOS", ret_by_month),
        ("- INVESTIMENTOS", inv_by_month),
        ("- DESPESAS OPERACIONAIS", op_by_month),
        ("RESULTADO ANTES DAS RETIRADAS E DESP. FINANCEIRAS", resultado_antes_by_month),
        ("RESULTADO OPERACIONAL", resultado_by_month),
    ]

    rows = []
    for nome, by_month in linhas:
        row = {"Linha": nome}
        for m in meses_nums:
            v = float(by_month.get(m, 0.0))
            rec = float(receita_total_by_month.get(m, 0.0))
            pct = (v / rec * 100.0) if rec != 0 else 0.0
            mes_pt = MES_NUM_TO_PT[m]
            row[mes_pt] = v
            row[f"%{mes_pt}"] = pct
        rows.append(row)
    dre = pd.DataFrame(rows)
    # Coluna de acumulado (soma no período selecionado)
    if len(meses_pt) > 0:
        dre["ACUMULADO"] = dre[meses_pt].sum(axis=1, skipna=True)
    else:
        dre["ACUMULADO"] = 0.0

    # % Acumulado sobre Receita (no período selecionado)
    receita_acum = float(sum(receita_total_by_month.get(m, 0.0) for m in meses_nums))
    dre["%ACUMULADO"] = (dre["ACUMULADO"] / receita_acum * 100.0) if receita_acum != 0 else 0.0


    st.subheader("DRE (JAN–DEZ) — Valores em R$ e % sobre Receita")

    def style_resultado(row):
        styles = [""] * len(row)
        if str(row.get("Linha", "")) == "RESULTADO OPERACIONAL":
            for j, col in enumerate(row.index):
                if (col in meses_pt) or (col == "ACUMULADO") or (col == "%ACUMULADO"):
                    val = row[col]
                    if pd.notna(val):
                        if float(val) < 0:
                            styles[j] = "color: #c00000; font-weight: 800;"
                        else:
                            styles[j] = "color: #1f4e79; font-weight: 800;"
                if col == "Linha":
                    styles[j] = "font-weight: 900;"
        return styles

    fmt_map = {}
    for m in meses_pt:
        fmt_map[m] = lambda x: f"R$ {format_brl(x)}"
        fmt_map[f"%{m}"] = lambda x: fmt_pct(x)

    fmt_map["ACUMULADO"] = lambda x: f"R$ {format_brl(x)}"
    fmt_map["%ACUMULADO"] = lambda x: fmt_pct(x)
    fmt_map["%ACUMULADO"] = lambda x: fmt_pct(x)

    value_cols_dre = list(meses_pt) + ["ACUMULADO"]
    pct_cols_dre = [f"%{m}" for m in meses_pt] + ["%ACUMULADO"]
    render_sticky_table(dre, value_cols=value_cols_dre, pct_cols=pct_cols_dre, highlight_row_label="RESULTADO OPERACIONAL")

    # Indicadores por Linha (Soma e Média) — respeita Ano/Meses do filtro lateral
    st.markdown("### Indicadores por linha (Soma e Média)")
    _linhas_kpi = list(dre["Linha"].dropna().unique()) if "Linha" in dre.columns else []
    if _linhas_kpi:
        _linha_sel = st.selectbox("Linha (DRE)", options=_linhas_kpi, key="kpi_linha_dre")
        _row = dre.loc[dre["Linha"] == _linha_sel].iloc[0]
        _vals = pd.Series({m: _row.get(m, 0.0) for m in meses_pt}, dtype="float64").fillna(0.0)
        _soma = float(_vals.sum())
        _media = float(_soma / max(len(meses_pt), 1))
        _c1, _c2 = st.columns(2)
        _c1.metric("Soma no período (R$)", "R$ " + format_brl(_soma))
        _c2.metric("Média mensal (R$)", "R$ " + format_brl(_media))
    else:
        st.info("Não foi possível montar o indicador por linha (coluna 'Linha' não encontrada).")

    # Drill DRE (mantém)
    st.divider()
    st.subheader("Drill (DRE): Contas → Despesas (sintetizadas) + Histórico")

    grupos = [
        "OUTRAS RECEITAS",
        "COMPRAS EMISSÃO",
        "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)",
        "DESPESAS COM PESSOAL",
        "DESPESAS ADMINISTRATIVAS",
        "DESPESAS COMERCIAIS",
        "DESPESAS FINANCEIRAS",
        "RETIRADAS SÓCIOS",
        "INVESTIMENTOS",
        "DESPESAS OPERACIONAIS",
    ]

    c1, c2 = st.columns([2, 1])
    with c1:
        grupo_sel = st.selectbox("Conta (grupo)", grupos, key="dre_grupo")
    with c2:
        mes_opt = ["TODOS"] + list(meses_pt)
        mes_sel = st.selectbox("Mês", options=mes_opt, index=0, key="dre_mes")

    meses_nums_drill = meses_nums if mes_sel == 'TODOS' else [MES_PT_TO_NUM[mes_sel]]
    receita_mes = float(sum(float(receita_total_by_month.get(m, 0.0)) for m in meses_nums_drill))

    def _sum_months(by_month):
        return float(sum(float(by_month.get(m, 0.0)) for m in meses_nums_drill))

    contas_mes = {
        "Outras Receitas": _sum_months(outras_receitas_by_month),
        "Compras": _sum_months(compras_by_month),
        "Deduções": _sum_months(deducoes_by_month),
        "Pessoal": _sum_months(pessoal_by_month),
        "Administrativas": _sum_months(adm_by_month),
        "Comerciais": _sum_months(com_by_month),
        "Financeiras": _sum_months(fin_by_month),
        "Retiradas Sócios": _sum_months(ret_by_month),
        "Investimentos": _sum_months(inv_by_month),
        "Operacionais": _sum_months(op_by_month),
    }
    pie_df = pd.DataFrame({"Conta": list(contas_mes.keys()), "Valor": list(contas_mes.values())})
    pie_df = pie_df[pie_df["Valor"] != 0].copy()
    pie_df["% Receita"] = (pie_df["Valor"] / receita_mes * 100.0) if receita_mes != 0 else 0.0

    pc1, pc2 = st.columns([1.2, 1])
    with pc1:
        if not pie_df.empty:
            fig = px.pie(pie_df, names="Conta", values="Valor",
                         title=f"Contas sobre Receita — {mes_sel}",
                         hover_data={"% Receita": True, "Valor": True})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sem valores no mês selecionado para o gráfico.")
    with pc2:
        val_grupo_mes_map = {
            "OUTRAS RECEITAS": _sum_months(outras_receitas_by_month),
            "COMPRAS EMISSÃO": _sum_months(compras_by_month),
            "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)": _sum_months(deducoes_by_month),
            "DESPESAS COM PESSOAL": _sum_months(pessoal_by_month),
            "DESPESAS ADMINISTRATIVAS": _sum_months(adm_by_month),
            "DESPESAS COMERCIAIS": _sum_months(com_by_month),
            "DESPESAS FINANCEIRAS": _sum_months(fin_by_month),
            "RETIRADAS SÓCIOS": _sum_months(ret_by_month),
            "INVESTIMENTOS": _sum_months(inv_by_month),
            "DESPESAS OPERACIONAIS": _sum_months(op_by_month),
        }
        val_grupo_mes = val_grupo_mes_map.get(grupo_sel, 0.0)
        pct_grupo = (val_grupo_mes / receita_mes * 100.0) if receita_mes != 0 else 0.0
        st.metric(f"{grupo_sel} ({mes_sel})", f"R$ {format_brl(val_grupo_mes)}", fmt_pct(pct_grupo))

    if grupo_sel == "COMPRAS EMISSÃO":
        st.info("Compras vêm da aba NOTAS EMITIDAS (NFS EMITIDAS). Drill de despesas/histórico de compras depende de detalhamento por fornecedor/nota.")
        return

    if grupo_sel == "OUTRAS RECEITAS":
        base_raw = g.copy()
        base_raw = base_raw[base_raw["_mes"].isin(meses_nums_drill)].copy()
        base_raw = base_raw[_mask_outras_receitas(base_raw)]
    elif grupo_sel in {"DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", "DESPESAS COM PESSOAL"}:
        # Drill dessas duas contas vem da aba DRE E DFC GERAL com mês à frente (m+1).
        meses_src = [m + 1 for m in meses_nums_drill if m is not None and int(m) < 12]
        base_raw = g[g["_mes"].isin(meses_src)].copy()
        if grupo_sel == "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)":
            base_raw = base_raw[base_raw["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith("00004 -")]
            # Excluir ICMS-ST na composição do DRE (apenas aqui)
            if not base_raw.empty:
                target_norm = _norm_txt(_DED_EXCL_DRE)
                for c in ["DESPESA", "CONTA DE RESULTADO", "HISTÓRICO", "HISTORICO"]:
                    if c in base_raw.columns:
                        s_norm = base_raw[c].astype(str).apply(_norm_txt)
                        base_raw = base_raw[~s_norm.str.contains(target_norm, na=False)]
                        base_raw = base_raw[~s_norm.str.contains("02.07.008", na=False)]
                        break
        else:
            base_raw = base_raw[base_raw["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith("00006 -")]
    else:
        base_raw = g.copy()
        base_raw = base_raw[base_raw["_mes"].isin(meses_nums_drill)].copy()
        prefix_map = {
            "DESPESAS ADMINISTRATIVAS": "00007 -",
            "DESPESAS COMERCIAIS": "00009 -",
            "DESPESAS FINANCEIRAS": "00011 -",
            "RETIRADAS SÓCIOS": "00016 -",
            "INVESTIMENTOS": "00015 -",
            "DESPESAS OPERACIONAIS": "00017 -",
        }
        prefix = prefix_map.get(grupo_sel)
        if prefix:
            base_raw = base_raw[base_raw["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)]

    if base_raw.empty:
        st.info("Sem lançamentos para esse grupo/mês.")
        return

    if "DESPESA" not in base_raw.columns:
        base_raw["DESPESA"] = "—"
    if "HISTÓRICO" not in base_raw.columns:
        base_raw["HISTÓRICO"] = "—"
    if "_v" not in base_raw.columns:
        base_raw["_v"] = base_raw["VAL.PAG"].apply(to_num)

    base_raw["DESPESA_SINT"] = base_raw["DESPESA"].apply(sintetizar_despesa)

    det_agg = (base_raw.groupby("DESPESA_SINT", dropna=False)["_v"]
               .sum().reset_index().rename(columns={"_v": "Valor"}))
    det_agg["% Receita"] = (det_agg["Valor"] / receita_mes * 100.0) if receita_mes != 0 else 0.0
    det_agg = det_agg.sort_values("Valor", ascending=False)

    top_n = safe_topn_slider("Top N despesas no gráfico", n_items=len(det_agg), default=15, cap=50)
    det_top = det_agg.head(top_n).copy()

    fig_bar = px.bar(det_top, x="Valor", y="DESPESA_SINT", orientation="h",
                     title=f"{grupo_sel} — Top {top_n} despesas ({mes_sel})",
                     hover_data={"% Receita": True})
    st.plotly_chart(fig_bar, use_container_width=True)

    st.dataframe(det_agg.style.format({"Valor": lambda x: f"R$ {format_brl(x)}", "% Receita": lambda x: fmt_pct(x)}).hide(axis="index"),
                 use_container_width=True)

    st.markdown("### Histórico — sintetizado e detalhado")
    desp_sel = st.selectbox("Selecione a despesa (sintetizada)", options=det_agg["DESPESA_SINT"].tolist(), key="dre_desp_sel")
    raw_sel = base_raw[base_raw["DESPESA_SINT"] == desp_sel].copy()

    raw_sel["_dt_sort"] = pd.to_datetime(raw_sel["DTA.PAG"], errors="coerce", dayfirst=True)
    raw_sel = raw_sel.sort_values(["_dt_sort"], ascending=False).drop(columns=["_dt_sort"])

    soma_sel = float(raw_sel["_v"].sum())
    pct_sel = (soma_sel / receita_mes * 100.0) if receita_mes != 0 else 0.0
    st.metric("Total da despesa selecionada", f"R$ {format_brl(soma_sel)}", fmt_pct(pct_sel))

    tab_sint, tab_fav, tab_det = st.tabs(["Histórico sintetizado", "Histórico sintetizado por Favorecido", "Histórico detalhado"])
    with tab_sint:
        key = pick_hist_key(raw_sel)
        if key is None:
            st.info("Não encontrei coluna para sintetizar (HISTÓRICO/FAVORECIDO/DESPESA).")
        else:
            tmp = raw_sel.copy()
            tmp[key] = tmp[key].astype(str).str.strip().replace({"": "—"})
            tmp["_valor"] = tmp.get("VAL.PAG", tmp["_v"]).apply(to_num)
            hist_sint = (tmp.groupby(key, dropna=False)["_valor"].sum().reset_index().rename(columns={"_valor": "Valor"}))
            hist_sint["% Receita"] = (hist_sint["Valor"] / receita_mes * 100.0) if receita_mes != 0 else 0.0
            hist_sint = hist_sint.sort_values("Valor", ascending=False)
            st.caption(f"Sintetizado por: **{key}**")
            st.dataframe(hist_sint.style.format({"Valor": lambda x: f"R$ {format_brl(x)}", "% Receita": lambda x: fmt_pct(x)}).hide(axis="index"),
                         use_container_width=True)
    with tab_fav:
        if "FAVORECIDO" not in raw_sel.columns:
            st.info("Não existe coluna 'FAVORECIDO' para sintetizar por favorecido.")
        else:
            tmp = raw_sel.copy()
            tmp["FAVORECIDO"] = tmp["FAVORECIDO"].astype(str).str.strip().replace({"": "—"})
            tmp["_valor"] = tmp.get("VAL.PAG", tmp["_v"]).apply(to_num)

            denom = receita_mes if "receita_mes" in locals() else receb_mes
            pct_label = "% Receita" if "receita_mes" in locals() else "% Recebimentos"

            fav_sint = (tmp.groupby("FAVORECIDO", dropna=False)["_valor"].sum()
                        .reset_index().rename(columns={"_valor": "Valor"}))
            fav_sint[pct_label] = (fav_sint["Valor"] / denom * 100.0) if denom != 0 else 0.0
            fav_sint = fav_sint.sort_values("Valor", ascending=False)

            topn_fav = safe_topn_slider("Top N (Favorecido)", len(fav_sint), default=15, cap=80)
            st.dataframe(
                fav_sint.head(topn_fav).style.format(
                    {"Valor": lambda x: f"R$ {format_brl(x)}", pct_label: lambda x: fmt_pct(x)}
                ).hide(axis="index"),
                use_container_width=True,
            )

    with tab_det:
        cols = [c for c in ["DTA.PAG", "CONTA DE RESULTADO", "DESPESA", "FAVORECIDO", "DUPLICATA", "HISTÓRICO", "VAL.PAG"] if c in raw_sel.columns]
        view = raw_sel[cols].copy() if cols else raw_sel.copy()
        st.dataframe(view.style.format({"VAL.PAG": lambda x: f"R$ {format_brl(to_num(x))}"}).hide(axis="index"),
                     use_container_width=True)


# =========================
# Página 2: DFC (FORNECEDORES = 00012)
# =========================
def pagina_dfc_geral(excel_path, ano_ref, meses_pt_sel=None):
    st.title("DFC Geral — (DRE e DFC GERAL)")

    # Meses selecionados no filtro lateral
    meses_pt = (meses_pt_sel or [])
    meses_pt = meses_pt if len(meses_pt) > 0 else MESES_PT
    meses_nums = [MES_PT_TO_NUM[m] for m in meses_pt]

    df_rec = read_sheet(excel_path, "RECEBIMENTO", sig)
    df_geral = read_sheet(excel_path, "DRE E DFC GERAL", sig)

    missing = [n for n, df in [("RECEBIMENTO", df_rec), ("DRE E DFC GERAL", df_geral)] if df is None]
    if missing:
        st.error(f"Faltam abas no Excel: {', '.join(missing)}")
        return

    req_r = {"MÊS", "ANO", "RECEBIMENTO"}
    if not req_r.issubset(set(df_rec.columns)):
        st.error("Na aba RECEBIMENTO preciso das colunas: 'MÊS', 'ANO', 'RECEBIMENTO'.")
        return
    receb_by_month = agg_by_month_from_ano_mes(df_rec, "RECEBIMENTO", "ANO", "MÊS", ano_ref)

    req_g = {"CONTA DE RESULTADO", "DTA.PAG", "VAL.PAG"}
    if not req_g.issubset(set(df_geral.columns)):
        st.error("Na aba DRE E DFC GERAL preciso das colunas: 'CONTA DE RESULTADO', 'DTA.PAG', 'VAL.PAG'.")
        return

    g_cur = prep_geral_year(excel_path, ano_ref, sig)
    g_next = prep_geral_year(excel_path, int(ano_ref) + 1, sig)
    g = g_cur
    if g is None:
        st.error("Não encontrei a aba DRE E DFC GERAL.")
        return

    pmap = dfc_prefix_map()
    fornec_by_month = sum_by_prefix_prepped(g, pmap["FORNECEDORES"])
    ded_by_month = sum_by_prefix_prepped(g, pmap["DEDUÇÕES (IMPOSTOS SOBRE VENDAS)"])
    pessoal_by_month = sum_by_prefix_prepped(g, pmap["DESPESAS COM PESSOAL"])
    adm_by_month = sum_by_prefix_prepped(g, pmap["DESPESAS ADMINISTRATIVAS"])
    com_by_month = sum_by_prefix_prepped(g, pmap["DESPESAS COMERCIAIS"])
    fin_by_month = sum_by_prefix_prepped(g, pmap["DESPESAS FINANCEIRAS"])
    ret_by_month = sum_by_prefix_prepped(g, '00016 -')
    inv_by_month = sum_by_prefix_prepped(g, pmap["INVESTIMENTOS"])
    op_by_month = sum_by_prefix_prepped(g, pmap["DESPESAS OPERACIONAIS"])
    outras_receitas_by_month = sum_outras_receitas_prepped(g)
    receb_total_by_month = {m: float(receb_by_month.get(m, 0.0)) + float(outras_receitas_by_month.get(m, 0.0)) for m in range(1, 13)}

    saldo_by_month = {}
    for m in range(1, 13):
        saidas = (fornec_by_month[m] + ded_by_month[m] + pessoal_by_month[m] + adm_by_month[m] +
                  com_by_month[m] + fin_by_month[m] + inv_by_month[m] + op_by_month[m] + ret_by_month[m])
        saldo_by_month[m] = receb_total_by_month[m] - saidas

    # Resultado antes das retiradas e despesas financeiras (volta essas duas linhas no resultado)
    resultado_antes_by_month = {m: float(saldo_by_month.get(m, 0.0)) + float(fin_by_month.get(m, 0.0)) + float(ret_by_month.get(m, 0.0)) for m in range(1, 13)}

    linhas = [
        ("+ RECEBIMENTOS", receb_by_month),
        ("+ OUTRAS RECEITAS", outras_receitas_by_month),
        ("- FORNECEDORES", fornec_by_month),
        ("- DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", ded_by_month),
        ("- DESPESAS COM PESSOAL", pessoal_by_month),
        ("- DESPESAS ADMINISTRATIVAS", adm_by_month),
        ("- DESPESAS COMERCIAIS", com_by_month),
        ("- DESPESAS FINANCEIRAS", fin_by_month),
        ("- RETIRADAS SÓCIOS", ret_by_month),
        ("- INVESTIMENTOS", inv_by_month),
        ("- DESPESAS OPERACIONAIS", op_by_month),
        ("RESULTADO ANTES DAS RETIRADAS E DESP. FINANCEIRAS", resultado_antes_by_month),
        ("SALDO OPERACIONAL", saldo_by_month),
    ]

    rows = []
    for nome, by_month in linhas:
        row = {"Linha": nome}
        for m in meses_nums:
            v = float(by_month.get(m, 0.0))
            rec = float(receb_total_by_month.get(m, 0.0))
            pct = (v / rec * 100.0) if rec != 0 else 0.0
            mes_pt = MES_NUM_TO_PT[m]
            row[mes_pt] = v
            row[f"%{mes_pt}"] = pct
        rows.append(row)

    dfc = pd.DataFrame(rows)
    # Coluna de acumulado (soma no período selecionado)
    if len(meses_pt) > 0:
        dfc["ACUMULADO"] = dfc[meses_pt].sum(axis=1, skipna=True)
    else:
        dfc["ACUMULADO"] = 0.0

    # % Acumulado sobre Recebimentos (no período selecionado)
    receb_acum = float(sum(receb_total_by_month.get(m, 0.0) for m in meses_nums))
    dfc["%ACUMULADO"] = (dfc["ACUMULADO"] / receb_acum * 100.0) if receb_acum != 0 else 0.0

    st.subheader("DFC (JAN–DEZ) — Valores em R$ e % sobre Recebimentos")

    def style_saldo(row):
        styles = [""] * len(row)
        if str(row.get("Linha", "")) == "SALDO OPERACIONAL":
            for j, col in enumerate(row.index):
                if (col in meses_pt) or (col == "ACUMULADO") or (col == "%ACUMULADO"):
                    val = row[col]
                    if pd.notna(val):
                        if float(val) < 0:
                            styles[j] = "color: #c00000; font-weight: 800;"
                        else:
                            styles[j] = "color: #1f4e79; font-weight: 800;"
                if col == "Linha":
                    styles[j] = "font-weight: 900;"
        return styles

    fmt_map = {}
    for m in meses_pt:
        fmt_map[m] = lambda x: f"R$ {format_brl(x)}"
        fmt_map[f"%{m}"] = lambda x: fmt_pct(x)

    fmt_map["ACUMULADO"] = lambda x: f"R$ {format_brl(x)}"
    fmt_map["%ACUMULADO"] = lambda x: fmt_pct(x)

    value_cols_dfc = list(meses_pt) + ["ACUMULADO"]
    pct_cols_dfc = [f"%{m}" for m in meses_pt] + ["%ACUMULADO"]
    render_sticky_table(dfc, value_cols=value_cols_dfc, pct_cols=pct_cols_dfc, highlight_row_label="SALDO OPERACIONAL")

    # Indicadores por Linha (Soma e Média) — respeita Ano/Meses do filtro lateral
    st.markdown("### Indicadores por linha (Soma e Média)")
    _linhas_kpi = list(dfc["Linha"].dropna().unique()) if "Linha" in dfc.columns else []
    if _linhas_kpi:
        _linha_sel = st.selectbox("Linha (DFC)", options=_linhas_kpi, key="kpi_linha_dfc")
        _row = dfc.loc[dfc["Linha"] == _linha_sel].iloc[0]
        _vals = pd.Series({m: _row.get(m, 0.0) for m in meses_pt}, dtype="float64").fillna(0.0)
        _soma = float(_vals.sum())
        _media = float(_soma / max(len(meses_pt), 1))
        _c1, _c2 = st.columns(2)
        _c1.metric("Soma no período (R$)", "R$ " + format_brl(_soma))
        _c2.metric("Média mensal (R$)", "R$ " + format_brl(_media))
    else:
        st.info("Não foi possível montar o indicador por linha (coluna 'Linha' não encontrada).")

    # Drill DFC (mesma experiência do DRE)
    st.divider()
    st.subheader("Drill (DFC): Contas → Despesas (sintetizadas) + Histórico")

    grupos = [
        "OUTRAS RECEITAS",
        "FORNECEDORES",
        "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)",
        "DESPESAS COM PESSOAL",
        "DESPESAS ADMINISTRATIVAS",
        "DESPESAS COMERCIAIS",
        "DESPESAS FINANCEIRAS",
        "RETIRADAS SÓCIOS",
        "INVESTIMENTOS",
        "DESPESAS OPERACIONAIS",
    ]

    c1, c2 = st.columns([2, 1])
    with c1:
        grupo_sel = st.selectbox("Conta (grupo)", grupos, key="dfc_grupo")
    with c2:
        mes_opt = ["TODOS"] + list(meses_pt)
        mes_sel = st.selectbox("Mês", options=mes_opt, index=0, key="dfc_mes")

    meses_nums_drill = meses_nums if mes_sel == 'TODOS' else [MES_PT_TO_NUM[mes_sel]]
    receb_mes = float(sum(float(receb_total_by_month.get(m, 0.0)) for m in meses_nums_drill))

    def _sum_months(by_month):
        return float(sum(float(by_month.get(m, 0.0)) for m in meses_nums_drill))

    contas_mes = {
        "Outras Receitas": _sum_months(outras_receitas_by_month),
        "Fornecedores": _sum_months(fornec_by_month),
        "Deduções": _sum_months(ded_by_month),
        "Pessoal": _sum_months(pessoal_by_month),
        "Administrativas": _sum_months(adm_by_month),
        "Comerciais": _sum_months(com_by_month),
        "Financeiras": _sum_months(fin_by_month),
        "Retiradas Sócios": _sum_months(ret_by_month),
        "Investimentos": _sum_months(inv_by_month),
        "Operacionais": _sum_months(op_by_month),
    }
    pie_df = pd.DataFrame({"Conta": list(contas_mes.keys()), "Valor": list(contas_mes.values())})
    pie_df = pie_df[pie_df["Valor"] != 0].copy()
    pie_df["% Recebimentos"] = (pie_df["Valor"] / receb_mes * 100.0) if receb_mes != 0 else 0.0

    pc1, pc2 = st.columns([1.2, 1])
    with pc1:
        if not pie_df.empty:
            fig = px.pie(pie_df, names="Conta", values="Valor",
                         title=f"Contas sobre Recebimentos — {mes_sel}",
                         hover_data={"% Recebimentos": True, "Valor": True})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sem valores no mês selecionado para o gráfico.")
    with pc2:
        val_map = {
            "OUTRAS RECEITAS": _sum_months(outras_receitas_by_month),
            "FORNECEDORES": _sum_months(fornec_by_month),
            "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)": _sum_months(ded_by_month),
            "DESPESAS COM PESSOAL": _sum_months(pessoal_by_month),
            "DESPESAS ADMINISTRATIVAS": _sum_months(adm_by_month),
            "DESPESAS COMERCIAIS": _sum_months(com_by_month),
            "DESPESAS FINANCEIRAS": _sum_months(fin_by_month),
            "RETIRADAS SÓCIOS": _sum_months(ret_by_month),
            "INVESTIMENTOS": _sum_months(inv_by_month),
            "DESPESAS OPERACIONAIS": _sum_months(op_by_month),
        }
        val_grp = val_map.get(grupo_sel, 0.0)
        pct_grp = (val_grp / receb_mes * 100.0) if receb_mes != 0 else 0.0
        st.metric(f"{grupo_sel} ({mes_sel})", f"R$ {format_brl(val_grp)}", fmt_pct(pct_grp))

    prefix = dfc_prefix_map().get(grupo_sel)
    base_raw = g.copy()
    base_raw = base_raw[base_raw["_mes"].isin(meses_nums_drill)].copy()
    if grupo_sel == "OUTRAS RECEITAS":
        base_raw = base_raw[_mask_outras_receitas(base_raw)]
    elif prefix:
        base_raw = base_raw[base_raw["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)]

    if base_raw.empty:
        st.info("Sem lançamentos para esse grupo/mês.")
        return

    if "DESPESA" not in base_raw.columns:
        base_raw["DESPESA"] = "—"
    if "HISTÓRICO" not in base_raw.columns:
        base_raw["HISTÓRICO"] = "—"
    if "_v" not in base_raw.columns:
        base_raw["_v"] = base_raw["VAL.PAG"].apply(to_num)

    base_raw["DESPESA_SINT"] = base_raw["DESPESA"].apply(sintetizar_despesa)

    det_agg = (base_raw.groupby("DESPESA_SINT", dropna=False)["_v"]
               .sum().reset_index().rename(columns={"_v": "Valor"}))
    det_agg["% Recebimentos"] = (det_agg["Valor"] / receb_mes * 100.0) if receb_mes != 0 else 0.0
    det_agg = det_agg.sort_values("Valor", ascending=False)

    top_n = safe_topn_slider("Top N despesas no gráfico", n_items=len(det_agg), default=15, cap=50)
    det_top = det_agg.head(top_n).copy()

    fig_bar = px.bar(det_top, x="Valor", y="DESPESA_SINT", orientation="h",
                     title=f"{grupo_sel} — Top {top_n} despesas ({mes_sel})",
                     hover_data={"% Recebimentos": True})
    st.plotly_chart(fig_bar, use_container_width=True)

    st.dataframe(det_agg.style.format({"Valor": lambda x: f"R$ {format_brl(x)}", "% Recebimentos": lambda x: fmt_pct(x)}).hide(axis="index"),
                 use_container_width=True)

    st.markdown("### Histórico — sintetizado e detalhado")
    desp_sel = st.selectbox("Selecione a despesa (sintetizada)", options=det_agg["DESPESA_SINT"].tolist(), key="dfc_desp_sel")
    raw_sel = base_raw[base_raw["DESPESA_SINT"] == desp_sel].copy()

    raw_sel["_dt_sort"] = pd.to_datetime(raw_sel["DTA.PAG"], errors="coerce", dayfirst=True)
    raw_sel = raw_sel.sort_values(["_dt_sort"], ascending=False).drop(columns=["_dt_sort"])

    soma_sel = float(raw_sel["_v"].sum())
    pct_sel = (soma_sel / receb_mes * 100.0) if receb_mes != 0 else 0.0
    st.metric("Total da despesa selecionada", f"R$ {format_brl(soma_sel)}", fmt_pct(pct_sel))

    tab_sint, tab_fav, tab_det = st.tabs(["Histórico sintetizado", "Histórico sintetizado por Favorecido", "Histórico detalhado"])
    with tab_sint:
        key = pick_hist_key(raw_sel)
        if key is None:
            st.info("Não encontrei coluna para sintetizar (HISTÓRICO/FAVORECIDO/DESPESA).")
        else:
            tmp = raw_sel.copy()
            tmp[key] = tmp[key].astype(str).str.strip().replace({"": "—"})
            tmp["_valor"] = tmp.get("VAL.PAG", tmp["_v"]).apply(to_num)
            hist_sint = (tmp.groupby(key, dropna=False)["_valor"].sum().reset_index().rename(columns={"_valor": "Valor"}))
            hist_sint["% Recebimentos"] = (hist_sint["Valor"] / receb_mes * 100.0) if receb_mes != 0 else 0.0
            hist_sint = hist_sint.sort_values("Valor", ascending=False)
            st.caption(f"Sintetizado por: **{key}**")
            st.dataframe(hist_sint.style.format({"Valor": lambda x: f"R$ {format_brl(x)}", "% Recebimentos": lambda x: fmt_pct(x)}).hide(axis="index"),
                         use_container_width=True)
    with tab_fav:
        if "FAVORECIDO" not in raw_sel.columns:
            st.info("Não existe coluna 'FAVORECIDO' para sintetizar por favorecido.")
        else:
            tmp = raw_sel.copy()
            tmp["FAVORECIDO"] = tmp["FAVORECIDO"].astype(str).str.strip().replace({"": "—"})
            tmp["_valor"] = tmp.get("VAL.PAG", tmp["_v"]).apply(to_num)

            denom = receita_mes if "receita_mes" in locals() else receb_mes
            pct_label = "% Receita" if "receita_mes" in locals() else "% Recebimentos"

            fav_sint = (tmp.groupby("FAVORECIDO", dropna=False)["_valor"].sum()
                        .reset_index().rename(columns={"_valor": "Valor"}))
            fav_sint[pct_label] = (fav_sint["Valor"] / denom * 100.0) if denom != 0 else 0.0
            fav_sint = fav_sint.sort_values("Valor", ascending=False)

            topn_fav = safe_topn_slider("Top N (Favorecido)", len(fav_sint), default=15, cap=80)
            st.dataframe(
                fav_sint.head(topn_fav).style.format(
                    {"Valor": lambda x: f"R$ {format_brl(x)}", pct_label: lambda x: fmt_pct(x)}
                ).hide(axis="index"),
                use_container_width=True,
            )

    with tab_det:
        cols = [c for c in ["DTA.PAG", "CONTA DE RESULTADO", "DESPESA", "FAVORECIDO", "DUPLICATA", "HISTÓRICO", "VAL.PAG"] if c in raw_sel.columns]
        view = raw_sel[cols].copy() if cols else raw_sel.copy()
        st.dataframe(view.style.format({"VAL.PAG": lambda x: f"R$ {format_brl(to_num(x))}"}).hide(axis="index"),
                     use_container_width=True)


# =========================
# Página 3: Faturamento
# =========================
def pagina_faturamento(excel_path, ano_ref, meses_pt_sel=None):
    st.title("Faturamento por canal")

    meses_pt = (meses_pt_sel or [])
    meses_pt = meses_pt if len(meses_pt) > 0 else MESES_PT
    meses_nums = [MES_PT_TO_NUM[m] for m in meses_pt]

    df_receita = read_sheet(excel_path, "RECEITA", sig)
    if df_receita is None:
        st.error("Não encontrei a aba RECEITA.")
        return

    canais_base = [
        "OFICINAS DAUTO final",
        "ZEMA",
        "BOX RÁPIDO",
        "SAGA",
        "LOJAS SOCIEDADE",
        "CANAL DIRETO",
        "OUTRAS *negociações",
    ]
    col_abastecimento = "ABASTECIMENTO LOJAS DAUTO TINTAS"
    col_fat_unica = "FATURAMENTO ÚNICA"
    col_fat_dauto_serv = "FATURAMENTO LOJAS DAUTO + SERVIÇO"
    col_receita_grupo = "RECEITA GRUPO"
    col_fat_logistico = "FATURAMENTO LOGÍSTICO"

    req = [
        "MÊS", "ANO", *canais_base,
        col_abastecimento, col_fat_unica, col_fat_dauto_serv,
        col_receita_grupo, col_fat_logistico,
    ]
    missing = [c for c in req if c not in df_receita.columns]
    if missing:
        st.error("Na aba RECEITA faltam as colunas: " + ", ".join(missing))
        return

    base = df_receita.copy()
    base["_ano"] = pd.to_numeric(base["ANO"], errors="coerce").astype("Int64")
    base["_mes"] = base["MÊS"].apply(parse_mes)
    base = base[(base["_ano"] == int(ano_ref)) & (base["_mes"].isin(meses_nums))].copy()

    cols_numericas = [*canais_base, col_abastecimento, col_fat_unica, col_fat_dauto_serv, col_receita_grupo, col_fat_logistico]
    for c in cols_numericas:
        base[c] = base[c].apply(to_num)

    base = base.sort_values("_mes")
    if base.empty:
        st.info("Não há dados de faturamento para os filtros selecionados.")
        return

    # Regras novas da página:
    # - FATURAMENTO ÚNICA = soma dos canais-base
    # - RECEITA GRUPO = FATURAMENTO ÚNICA + FATURAMENTO LOJAS DAUTO + SERVIÇO
    # - FATURAMENTO LOGÍSTICO = FATURAMENTO ÚNICA + ABASTECIMENTO LOJAS DAUTO TINTAS
    base[col_fat_unica] = base[canais_base].sum(axis=1)
    base[col_receita_grupo] = base[col_fat_unica] + base[col_fat_dauto_serv]
    base[col_fat_logistico] = base[col_fat_unica] + base[col_abastecimento]

    canais_tabela_principal = [
        *canais_base,
        col_fat_unica,
        col_fat_dauto_serv,
        col_receita_grupo,
    ]

    tabela_principal = base[["MÊS", *canais_tabela_principal]].copy().rename(columns={"MÊS": "Mês"})
    st.subheader("Faturamento mensal por canal")
    render_sticky_table(tabela_principal, value_cols=canais_tabela_principal)

    totais_principal = tabela_principal[canais_tabela_principal].sum(axis=0).reset_index()
    totais_principal.columns = ["Canal", "Acumulado"]
    total_receita_grupo = float(totais_principal.loc[totais_principal["Canal"] == col_receita_grupo, "Acumulado"].sum())
    totais_principal["% Receita Grupo"] = totais_principal["Acumulado"].apply(
        lambda x: (x / total_receita_grupo * 100.0) if total_receita_grupo != 0 else 0.0
    )
    totais_principal = totais_principal.sort_values("Acumulado", ascending=False).reset_index(drop=True)

    st.markdown("### Acumulado por canal no período selecionado")
    render_sticky_table(
        totais_principal,
        value_cols=["Acumulado"],
        pct_cols=["% Receita Grupo"],
        highlight_row_label=col_receita_grupo,
    )

    st.markdown("### Base logística")
    tabela_logistica = base[["MÊS", col_abastecimento, col_fat_logistico]].copy().rename(columns={"MÊS": "Mês"})
    render_sticky_table(tabela_logistica, value_cols=[col_abastecimento, col_fat_logistico])

    st.markdown("### Drill do faturamento logístico")
    total_logistico = float(base[col_fat_logistico].sum())
    linhas_drill = []
    for canal in canais_base:
        valor = float(base[canal].sum())
        linhas_drill.append({
            "Canal": canal,
            "Acumulado": valor,
            "% Faturamento Logístico": (valor / total_logistico * 100.0) if total_logistico != 0 else 0.0,
        })

    valor_fat_unica = float(base[col_fat_unica].sum())
    linhas_drill.append({
        "Canal": col_fat_unica,
        "Acumulado": valor_fat_unica,
        "% Faturamento Logístico": (valor_fat_unica / total_logistico * 100.0) if total_logistico != 0 else 0.0,
    })

    valor_abastecimento = float(base[col_abastecimento].sum())
    linhas_drill.append({
        "Canal": col_abastecimento,
        "Acumulado": valor_abastecimento,
        "% Faturamento Logístico": (valor_abastecimento / total_logistico * 100.0) if total_logistico != 0 else 0.0,
    })

    linhas_drill.append({
        "Canal": col_fat_logistico,
        "Acumulado": total_logistico,
        "% Faturamento Logístico": 100.0 if total_logistico != 0 else 0.0,
    })

    drill_logistico = pd.DataFrame(linhas_drill)
    render_sticky_table(
        drill_logistico,
        value_cols=["Acumulado"],
        pct_cols=["% Faturamento Logístico"],
        highlight_row_label=col_fat_logistico,
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Receita Grupo acumulada", fmt_brl_display(total_receita_grupo))
    c2.metric("Faturamento Logístico acumulado", fmt_brl_display(total_logistico))
    c3.metric("Média mensal Receita Grupo", fmt_brl_display(total_receita_grupo / max(len(base), 1)))

    st.markdown("### Evolução mensal")
    canal_sel = st.selectbox(
        "Canal",
        options=[*canais_tabela_principal, col_abastecimento, col_fat_logistico],
        index=[*canais_tabela_principal, col_abastecimento, col_fat_logistico].index(col_receita_grupo),
        key="fat_canal",
    )
    evo = base[["MÊS", "_mes", canal_sel]].copy().sort_values("_mes")
    fig = px.bar(evo, x="MÊS", y=canal_sel, title=f"Evolução mensal — {canal_sel}")
    fig.update_layout(xaxis_title="Mês", yaxis_title="Valor (R$)")
    st.plotly_chart(fig, use_container_width=True)


# =========================
# Página 4: Controles fiscais
# =========================
MAPA_EMPRESAS_FISCAL = {
    1: "GUARÁ",
    4: "ADE",
    6: "GAMA",
    8: "LUZIÂNIA",
    9: "ÚNICA",
    12: "SOFNORTE",
    13: "CEILÂNDIA",
    14: "S IA",
    15: "UNAÍ",
    16: "AG LINDAS",
    20: "DAUTO SERVIÇO",
    22: "GUARÁ",
    24: "LUZIÂNIA",
}

ORDEM_LOJAS_FISCAL = [
    "ADE", "AG LINDAS", "CEILÂNDIA", "DAUTO SERVIÇO", "GAMA",
    "GUARÁ", "LUZIÂNIA", "S IA", "SOFNORTE", "UNAÍ", "ÚNICA",
]

MESES_NOME_FISCAL = {
    1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
    7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ",
}


def _fiscal_localizar_arquivo(nome: str) -> Path:
    pasta_app = Path(__file__).resolve().parent
    candidatos = [
        pasta_app / nome,
        Path.cwd() / nome,
        pasta_app / "dados" / nome,
        Path.cwd() / "dados" / nome,
    ]
    for caminho in candidatos:
        if caminho.exists():
            return caminho
    raise FileNotFoundError(
        f"Não encontrei '{nome}'. Coloque o arquivo na mesma pasta do app "
        "ou dentro da pasta 'dados' no repositório."
    )


def _fiscal_ler_texto(caminho: Path) -> str:
    for enc in ("utf-8-sig", "cp1252", "latin1"):
        try:
            return caminho.read_text(encoding=enc)
        except UnicodeDecodeError:
            pass
    return caminho.read_text(encoding="latin1", errors="ignore")


def _fiscal_valor(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    s = str(v).strip().replace('"', '').replace("R$", "").replace(" ", "")
    if not s:
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def _fiscal_codigo(v):
    try:
        return int(float(str(v).strip().strip('"')))
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def _fiscal_carregar_cartoes(caminho_str: str, assinatura):
    caminho = Path(caminho_str)
    linhas = _fiscal_ler_texto(caminho).splitlines()
    inicio = next((i for i, l in enumerate(linhas) if l.startswith("DUPLICATA;EMP;DTA.CAD;")), None)
    if inicio is None:
        raise ValueError("Cabeçalho do relatório de cartões não localizado.")

    df = pd.read_csv(
        StringIO("\n".join(linhas[inicio:])),
        sep=";", dtype=str, engine="python", on_bad_lines="skip"
    )
    df["COD_EMPRESA"] = df["EMP"].apply(_fiscal_codigo)
    df["DATA"] = pd.to_datetime(df["DTA.CAD"], dayfirst=True, errors="coerce")
    df = df[df["COD_EMPRESA"].isin(MAPA_EMPRESAS_FISCAL) & df["DATA"].notna()].copy()

    for c in ["VLR.BRU", "VLR.LÍQ", "VLR.TAX"]:
        df[c] = df[c].apply(_fiscal_valor) if c in df.columns else 0.0

    df["LOJA"] = df["COD_EMPRESA"].map(MAPA_EMPRESAS_FISCAL)
    df["ANO"] = df["DATA"].dt.year.astype("Int64")
    df["MES_NUM"] = df["DATA"].dt.month.astype("Int64")
    return df


@st.cache_data(show_spinner=False)
def _fiscal_carregar_saidas(caminho_str: str, assinatura):
    caminho = Path(caminho_str)
    linhas = _fiscal_ler_texto(caminho).splitlines()
    inicio = next((i for i, l in enumerate(linhas) if l.startswith("EMPRESA;DATA;VR. CON;")), None)
    if inicio is None:
        raise ValueError("Bloco 'REGISTRO DE SAÍDAS - RESUMO POR DATA' não localizado.")

    cab = next(csv.reader([linhas[inicio]], delimiter=";", quotechar='"'))
    regs = []
    for linha in linhas[inicio + 1:]:
        if not linha.strip():
            if regs:
                break
            continue
        try:
            campos = next(csv.reader([linha], delimiter=";", quotechar='"'))
        except Exception:
            break
        if len(campos) != len(cab):
            break
        if not campos[0].strip() or "-" not in campos[0]:
            break
        regs.append(campos)

    if not regs:
        raise ValueError("O bloco de saídas foi localizado, mas não contém registros.")

    df = pd.DataFrame(regs, columns=cab)
    df["COD_EMPRESA"] = (
        df["EMPRESA"].astype(str).str.extract(r"^(\d+)", expand=False).apply(_fiscal_codigo)
    )
    df["DATA"] = pd.to_datetime(df["DATA"], dayfirst=True, errors="coerce")
    df["VLR_CONTABIL"] = df["VR. CON"].apply(_fiscal_valor)
    df = df[df["COD_EMPRESA"].isin(MAPA_EMPRESAS_FISCAL) & df["DATA"].notna()].copy()
    df["LOJA"] = df["COD_EMPRESA"].map(MAPA_EMPRESAS_FISCAL)
    df["ANO"] = df["DATA"].dt.year.astype("Int64")
    df["MES_NUM"] = df["DATA"].dt.month.astype("Int64")
    return df


def _fiscal_conciliacao(cartoes, saidas):
    c = (
        cartoes.groupby(["LOJA", "ANO", "MES_NUM"], as_index=False)
        .agg(
            CARTAO_BRUTO=("VLR.BRU", "sum"),
            CARTAO_LIQUIDO=("VLR.LÍQ", "sum"),
            TAXAS=("VLR.TAX", "sum"),
            LANCAMENTOS=("DUPLICATA", "count"),
        )
    )
    s = (
        saidas.groupby(["LOJA", "ANO", "MES_NUM"], as_index=False)
        .agg(VLR_CONTABIL=("VLR_CONTABIL", "sum"))
    )
    d = s.merge(c, on=["LOJA", "ANO", "MES_NUM"], how="outer")
    for col in ["VLR_CONTABIL", "CARTAO_BRUTO", "CARTAO_LIQUIDO", "TAXAS", "LANCAMENTOS"]:
        d[col] = pd.to_numeric(d.get(col, 0), errors="coerce").fillna(0)
    d["DIFERENCA"] = d["VLR_CONTABIL"] - d["CARTAO_BRUTO"]
    d["PERC_CARTAO"] = (
        d["CARTAO_BRUTO"].div(d["VLR_CONTABIL"].replace(0, pd.NA)).mul(100).fillna(0)
    )
    d["MES"] = d["MES_NUM"].map(MESES_NOME_FISCAL)
    return d


def _fiscal_style_tabela(df):
    moeda_cols = ["VLR Contábil", "Cartão Bruto", "Diferença", "Cartão Líquido", "Taxas"]
    pct_cols = ["% Cartão"]
    fmt = {c: lambda x: fmt_brl_display(x) for c in moeda_cols if c in df.columns}
    fmt.update({c: lambda x: fmt_pct(x) for c in pct_cols if c in df.columns})
    if "Lançamentos" in df.columns:
        fmt["Lançamentos"] = lambda x: f"{int(x):,}".replace(",", ".")

    sty = df.style.format(fmt)

    if "Cartão Bruto" in df.columns and "VLR Contábil" in df.columns:
        def _cor(row):
            try:
                maior = float(row["Cartão Bruto"]) > float(row["VLR Contábil"])
                bg = "background-color: rgba(220,53,69,.14); color:#a61b29;" if maior \
                     else "background-color: rgba(31,78,121,.10); color:#1f4e79;"
                return [bg] * len(row)
            except Exception:
                return [""] * len(row)
        sty = sty.apply(_cor, axis=1)
    return sty


def pagina_controles_fiscais():
    st.markdown(
        """
        <style>
        .fiscal-hero {
            padding: 1.1rem 1.3rem;
            border: 1px solid rgba(49,51,63,.14);
            border-radius: 16px;
            background: linear-gradient(135deg, rgba(31,78,121,.08), rgba(255,255,255,.96));
            margin-bottom: 1rem;
        }
        .fiscal-hero h1 { margin: 0; font-size: 2rem; }
        .fiscal-hero p { margin: .3rem 0 0 0; opacity: .75; }
        </style>
        <div class="fiscal-hero">
          <h1>Controles fiscais</h1>
          <p>Conciliação mensal entre o valor contábil das notas emitidas e os cartões passados.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    try:
        arq_cartoes = _fiscal_localizar_arquivo("cartões passados.csv")
        arq_saidas = _fiscal_localizar_arquivo("registro saídas.csv")
        sig_c = (arq_cartoes.stat().st_mtime_ns, arq_cartoes.stat().st_size)
        sig_s = (arq_saidas.stat().st_mtime_ns, arq_saidas.stat().st_size)
        cartoes = _fiscal_carregar_cartoes(str(arq_cartoes), sig_c)
        saidas = _fiscal_carregar_saidas(str(arq_saidas), sig_s)
        base = _fiscal_conciliacao(cartoes, saidas)
    except Exception as e:
        st.error(f"Não foi possível carregar os arquivos fiscais: {e}")
        st.info(
            "No repositório, deixe `cartões passados.csv` e `registro saídas.csv` "
            "na mesma pasta do PY ou dentro da pasta `dados/`."
        )
        return

    anos = sorted(pd.to_numeric(base["ANO"], errors="coerce").dropna().astype(int).unique().tolist())
    if not anos:
        st.warning("Não encontrei anos válidos nos arquivos fiscais.")
        return

    f1, f2 = st.columns([1, 3])
    with f1:
        ano = st.selectbox("Ano", anos, index=len(anos)-1, key="fiscal_ano")
    with f2:
        st.caption(
            f"Arquivos: {arq_cartoes.name} • {arq_saidas.name}"
        )

    ano_df = base[pd.to_numeric(base["ANO"], errors="coerce") == int(ano)].copy()
    if ano_df.empty:
        st.info("Sem dados para o ano selecionado.")
        return

    # Acumulado por loja
    acum = (
        ano_df.groupby("LOJA", as_index=False)
        .agg(
            VLR_CONTABIL=("VLR_CONTABIL", "sum"),
            CARTAO_BRUTO=("CARTAO_BRUTO", "sum"),
            CARTAO_LIQUIDO=("CARTAO_LIQUIDO", "sum"),
            TAXAS=("TAXAS", "sum"),
            LANCAMENTOS=("LANCAMENTOS", "sum"),
        )
    )
    acum["DIFERENCA"] = acum["VLR_CONTABIL"] - acum["CARTAO_BRUTO"]
    acum["PERC_CARTAO"] = (
        acum["CARTAO_BRUTO"].div(acum["VLR_CONTABIL"].replace(0, pd.NA)).mul(100).fillna(0)
    )
    acum["_ord"] = acum["LOJA"].apply(
        lambda x: ORDEM_LOJAS_FISCAL.index(x) if x in ORDEM_LOJAS_FISCAL else 999
    )
    acum = acum.sort_values("_ord").drop(columns="_ord")

    total_cont = float(acum["VLR_CONTABIL"].sum())
    total_cart = float(acum["CARTAO_BRUTO"].sum())
    total_dif = total_cont - total_cart
    pct_total = (total_cart / total_cont * 100) if total_cont else 0.0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("VLR contábil acumulado", fmt_brl_display(total_cont))
    k2.metric("Cartão bruto acumulado", fmt_brl_display(total_cart))
    k3.metric("Diferença acumulada", fmt_brl_display(total_dif))
    k4.metric("% vendas em cartão", fmt_pct(pct_total))

    st.divider()
    st.markdown("## Drill por loja")
    lojas = [x for x in ORDEM_LOJAS_FISCAL if x in ano_df["LOJA"].dropna().unique().tolist()]
    if not lojas:
        st.info("Nenhuma loja encontrada.")
        return

    loja_sel = st.selectbox(
        "Selecione a loja",
        options=lojas,
        index=0,
        key="fiscal_drill_loja",
        help="Ao trocar a loja, a tabela mensal abaixo é atualizada automaticamente.",
    )

    loja = ano_df[ano_df["LOJA"] == loja_sel].sort_values("MES_NUM").copy()
    loja_view = loja[
        ["MES", "VLR_CONTABIL", "CARTAO_BRUTO", "DIFERENCA",
         "PERC_CARTAO", "CARTAO_LIQUIDO", "TAXAS", "LANCAMENTOS"]
    ].rename(columns={
        "MES": "Mês",
        "VLR_CONTABIL": "VLR Contábil",
        "CARTAO_BRUTO": "Cartão Bruto",
        "DIFERENCA": "Diferença",
        "PERC_CARTAO": "% Cartão",
        "CARTAO_LIQUIDO": "Cartão Líquido",
        "TAXAS": "Taxas",
        "LANCAMENTOS": "Lançamentos",
    })

    lc = float(loja["VLR_CONTABIL"].sum())
    lcb = float(loja["CARTAO_BRUTO"].sum())
    ld = lc - lcb
    lp = (lcb / lc * 100) if lc else 0.0
    a1, a2, a3, a4 = st.columns(4)
    a1.metric(f"{loja_sel} • Contábil", fmt_brl_display(lc))
    a2.metric(f"{loja_sel} • Cartões", fmt_brl_display(lcb))
    a3.metric(f"{loja_sel} • Diferença", fmt_brl_display(ld))
    a4.metric(f"{loja_sel} • % Cartão", fmt_pct(lp))

    st.dataframe(
        _fiscal_style_tabela(loja_view).hide(axis="index"),
        use_container_width=True,
        height=min(520, 86 + len(loja_view) * 36),
    )
    st.caption("Vermelho: cartão maior que VLR Contábil. Azul: cartão menor ou igual ao VLR Contábil.")

    if not loja.empty:
        chart = loja[["MES_NUM", "MES", "VLR_CONTABIL", "CARTAO_BRUTO"]].copy()
        chart = chart.melt(
            id_vars=["MES_NUM", "MES"],
            value_vars=["VLR_CONTABIL", "CARTAO_BRUTO"],
            var_name="Indicador", value_name="Valor"
        )
        chart["Indicador"] = chart["Indicador"].replace(
            {"VLR_CONTABIL": "VLR Contábil", "CARTAO_BRUTO": "Cartão Bruto"}
        )
        fig = px.bar(
            chart.sort_values("MES_NUM"),
            x="MES", y="Valor", color="Indicador", barmode="group",
            title=f"Evolução mensal — {loja_sel}",
        )
        fig.update_layout(xaxis_title="", yaxis_title="Valor (R$)", legend_title="")
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.markdown("## Acumulado por lojas")
    acum_view = acum[
        ["LOJA", "VLR_CONTABIL", "CARTAO_BRUTO", "DIFERENCA",
         "PERC_CARTAO", "CARTAO_LIQUIDO", "TAXAS", "LANCAMENTOS"]
    ].rename(columns={
        "LOJA": "Loja",
        "VLR_CONTABIL": "VLR Contábil",
        "CARTAO_BRUTO": "Cartão Bruto",
        "DIFERENCA": "Diferença",
        "PERC_CARTAO": "% Cartão",
        "CARTAO_LIQUIDO": "Cartão Líquido",
        "TAXAS": "Taxas",
        "LANCAMENTOS": "Lançamentos",
    })
    st.dataframe(
        _fiscal_style_tabela(acum_view).hide(axis="index"),
        use_container_width=True,
        height=min(640, 86 + len(acum_view) * 36),
    )

    st.markdown("### Comparativo acumulado por loja")
    graf_acum = acum[["LOJA", "VLR_CONTABIL", "CARTAO_BRUTO"]].melt(
        id_vars="LOJA",
        value_vars=["VLR_CONTABIL", "CARTAO_BRUTO"],
        var_name="Indicador", value_name="Valor"
    )
    graf_acum["Indicador"] = graf_acum["Indicador"].replace(
        {"VLR_CONTABIL": "VLR Contábil", "CARTAO_BRUTO": "Cartão Bruto"}
    )
    fig2 = px.bar(
        graf_acum, x="LOJA", y="Valor", color="Indicador",
        barmode="group", title=f"Acumulado por loja — {ano}"
    )
    fig2.update_layout(xaxis_title="", yaxis_title="Valor (R$)", legend_title="")
    st.plotly_chart(fig2, use_container_width=True)


# =========================
# Main
# =========================
st.set_page_config(page_title="GERAL", layout="wide")
st.sidebar.title("Menu")

pagina = st.sidebar.radio(
    "Selecione:",
    ["DRE Geral", "DFC Geral", "Faturamento", "Controles fiscais"],
)

if pagina == "Controles fiscais":
    if st.sidebar.button("Atualizar dados fiscais", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    pagina_controles_fiscais()
    st.stop()


# As páginas abaixo dependem do Excel.
def _auto_find_excel() -> str | None:
    preferred = ["DRE E DFC GERAL.xlsx", "DRE_E_DFC_GERAL.xlsx", "BASE.xlsx", "BASE .xlsx"]
    for fn in preferred:
        if os.path.exists(fn):
            return fn
    files = []
    for pat in ["*.xlsx", "*.xlsm", "*.xls"]:
        files.extend(glob.glob(pat))
    files = [f for f in files if os.path.isfile(f)]
    if not files:
        return None
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return files[0]


excel_path = _auto_find_excel()
if not excel_path:
    st.sidebar.error("Não encontrei nenhum Excel (.xlsx/.xlsm/.xls) na mesma pasta do app.")
    st.stop()


EXCEL_PATH = excel_path

def excel_signature(path: str):
    stt = os.stat(path)
    return (stt.st_mtime_ns, stt.st_size)


sig = excel_signature(EXCEL_PATH)
EXCEL_SIG = sig

st.sidebar.caption(f"Excel: **{excel_path}**")
sheet_names = get_sheet_names(excel_path, sig)
if not sheet_names:
    st.sidebar.error(f"Não consegui abrir '{excel_path}'.")
    st.stop()
st.sidebar.success("Excel carregado")

meses_pt_sel = st.sidebar.multiselect("Meses", options=MESES_PT, default=MESES_PT)

anos = set()
for sheet in ["RECEITA", "NOTAS EMITIDAS", "RECEBIMENTO"]:
    df_tmp = read_sheet(excel_path, sheet, sig)
    if df_tmp is not None and "ANO" in df_tmp.columns:
        anos |= set(pd.to_numeric(df_tmp["ANO"], errors="coerce").dropna().astype(int).unique().tolist())

df_tmp = read_sheet(excel_path, "DRE E DFC GERAL", sig)
if df_tmp is not None and "DTA.PAG" in df_tmp.columns:
    d = pd.to_datetime(df_tmp["DTA.PAG"], errors="coerce", dayfirst=True)
    anos |= set(d.dt.year.dropna().astype(int).unique().tolist())

anos = sorted(list(anos))
if not anos:
    st.error("Não encontrei nenhum ANO válido no Excel.")
    st.stop()

ano_ref = st.sidebar.selectbox("Ano de referência", options=anos, index=len(anos) - 1)

if pagina == "DRE Geral":
    pagina_dre_geral(excel_path, ano_ref, meses_pt_sel)
elif pagina == "DFC Geral":
    pagina_dfc_geral(excel_path, ano_ref, meses_pt_sel)
else:
    pagina_faturamento(excel_path, ano_ref, meses_pt_sel)
