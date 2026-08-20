# GERAL.py
import re
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px
import os
import glob
import csv
from io import StringIO, BytesIO
import base64
from PIL import Image
from pathlib import Path


# Logos incorporadas ao arquivo (Única + Dauto)
LOGO_UNICA_B64 = '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'
LOGO_DAUTO_B64 = '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'

def _img_from_b64(data):
    return Image.open(BytesIO(base64.b64decode(data)))

def _inject_modern_ui():
    st.markdown("""
    <style>
    .stApp { background:#f6f8fc; }
    [data-testid="stMainBlockContainer"] {padding-top:1.5rem; max-width:1800px;}
    [data-testid="stSidebar"] {background:#eef2f7; border-right:1px solid #e1e7ef;}
    [data-testid="stSidebar"] .stButton button {border-radius:10px;}
    [data-testid="stSidebar"] div[data-testid="stCheckbox"] {
        background:#ffffff;
        border:1px solid #dfe5ee;
        border-radius:8px;
        padding:5px 8px;
        margin-bottom:5px;
        min-height:38px;
        display:flex;
        align-items:center;
    }
    [data-testid="stSidebar"] div[data-testid="stCheckbox"]:hover {
        border-color:#b8c4d6;
        box-shadow:0 1px 4px rgba(15,23,42,.05);
    }
    div[data-testid="stMetric"] {background:white;border:1px solid #e3e8f0;border-radius:14px;padding:16px 18px;box-shadow:0 2px 10px rgba(15,23,42,.04);}
    div[data-testid="stMetric"] label {font-weight:700;color:#64748b;}
    div[data-testid="stMetricValue"] {font-weight:800;color:#0f172a;}
    .eyebrow {font-size:.78rem;font-weight:800;letter-spacing:.08em;color:#64748b;text-transform:uppercase;margin-bottom:4px;}
    .hero-title {font-size:2rem;font-weight:850;color:#0f172a;line-height:1.1;margin:0;}
    .hero-sub {color:#64748b;margin-top:6px;margin-bottom:16px;}
    h1,h2,h3 {color:#0f172a;}
    hr {border-color:#e7ebf1;}
    [data-testid="stDataFrame"], .sticky-table-wrap {box-shadow:0 2px 10px rgba(15,23,42,.035);}
    .block-container {padding-bottom:3rem;}
    </style>
    """, unsafe_allow_html=True)

def _sidebar_brand():
    st.markdown(
        '<div style="display:flex;gap:10px;align-items:center;justify-content:center;margin:2px 0 18px;">'
        f'<img src="data:image/png;base64,{LOGO_UNICA_B64}" style="width:66px;height:66px;border-radius:50%;object-fit:cover;">'
        f'<img src="data:image/png;base64,{LOGO_DAUTO_B64}" style="width:66px;height:66px;border-radius:50%;object-fit:cover;">'
        '</div>', unsafe_allow_html=True)

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
    st.markdown('<div class="eyebrow">Financeiro • Visão Gerencial</div><div class="hero-title">DRE Gerencial</div><div class="hero-sub">Visão consolidada do resultado da empresa</div>', unsafe_allow_html=True)

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


    # Total de todas as despesas exibidas acima
    total_despesas_by_month = {
        m: (float(compras_by_month.get(m, 0.0)) + float(deducoes_by_month.get(m, 0.0)) +
            float(pessoal_by_month.get(m, 0.0)) + float(adm_by_month.get(m, 0.0)) +
            float(com_by_month.get(m, 0.0)) + float(fin_by_month.get(m, 0.0)) +
            float(ret_by_month.get(m, 0.0)) + float(inv_by_month.get(m, 0.0)) +
            float(op_by_month.get(m, 0.0)))
        for m in range(1, 13)
    }

    # Resultado antes das retiradas, despesas financeiras e investimentos
    # (volta essas três linhas no resultado operacional)
    resultado_antes_by_month = {
        m: (float(resultado_by_month.get(m, 0.0)) + float(fin_by_month.get(m, 0.0)) +
            float(ret_by_month.get(m, 0.0)) + float(inv_by_month.get(m, 0.0)))
        for m in range(1, 13)
    }

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
        ("TOTAL DESPESAS", total_despesas_by_month),
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


    # ===== Cockpit executivo (mesmos dados do DRE; apenas nova apresentação) =====
    receita_acum_c = float(sum(receita_total_by_month.get(m, 0.0) for m in meses_nums))
    compras_acum_c = float(sum(compras_by_month.get(m, 0.0) for m in meses_nums))
    despesas_acum_c = float(sum(deducoes_by_month.get(m,0.0)+pessoal_by_month.get(m,0.0)+adm_by_month.get(m,0.0)+com_by_month.get(m,0.0)+fin_by_month.get(m,0.0)+inv_by_month.get(m,0.0)+op_by_month.get(m,0.0)+ret_by_month.get(m,0.0) for m in meses_nums))
    resultado_acum_c = float(sum(resultado_by_month.get(m, 0.0) for m in meses_nums))
    lucro_bruto_c = receita_acum_c - compras_acum_c - float(sum(deducoes_by_month.get(m,0.0) for m in meses_nums))
    margem_c = (resultado_acum_c / receita_acum_c * 100.0) if receita_acum_c else 0.0

    st.markdown('<div class="eyebrow">Cockpit Executivo</div>', unsafe_allow_html=True)
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("Receita", fmt_brl_display(receita_acum_c))
    k2.metric("Lucro bruto", fmt_brl_display(lucro_bruto_c))
    k3.metric("Despesas totais", fmt_brl_display(despesas_acum_c))
    k4.metric("Resultado operacional", fmt_brl_display(resultado_acum_c))
    k5.metric("Margem operacional", fmt_pct(margem_c))

    evo = pd.DataFrame({
        "Mês": [MES_NUM_TO_PT[m] for m in meses_nums],
        "Receita": [receita_total_by_month.get(m,0.0) for m in meses_nums],
        "Lucro Bruto": [receita_total_by_month.get(m,0.0)-compras_by_month.get(m,0.0)-deducoes_by_month.get(m,0.0) for m in meses_nums],
        "Resultado": [resultado_by_month.get(m,0.0) for m in meses_nums],
    }).melt(id_vars="Mês", var_name="Indicador", value_name="Valor")
    despesas_rank = pd.DataFrame({
        "Conta":["Compras","Pessoal","Administrativas","Comerciais","Financeiras","Operacionais","Investimentos","Retiradas Sócios"],
        "Valor":[sum(compras_by_month.get(m,0.0) for m in meses_nums),sum(pessoal_by_month.get(m,0.0) for m in meses_nums),sum(adm_by_month.get(m,0.0) for m in meses_nums),sum(com_by_month.get(m,0.0) for m in meses_nums),sum(fin_by_month.get(m,0.0) for m in meses_nums),sum(op_by_month.get(m,0.0) for m in meses_nums),sum(inv_by_month.get(m,0.0) for m in meses_nums),sum(ret_by_month.get(m,0.0) for m in meses_nums)]
    }).sort_values("Valor", ascending=True)
    g1,g2=st.columns([1.35,1])
    with g1:
        fig_evo=px.line(evo,x="Mês",y="Valor",color="Indicador",markers=True,title="Evolução mensal")
        fig_evo.update_layout(height=340,margin=dict(l=10,r=10,t=55,b=10),legend_title="",paper_bgcolor="white",plot_bgcolor="white")
        st.plotly_chart(fig_evo,use_container_width=True)
    with g2:
        fig_rank=px.bar(despesas_rank,x="Valor",y="Conta",orientation="h",title="Maiores despesas • acumulado")
        fig_rank.update_layout(height=340,margin=dict(l=10,r=10,t=55,b=10),showlegend=False,paper_bgcolor="white",plot_bgcolor="white")
        st.plotly_chart(fig_rank,use_container_width=True)

    st.markdown('<div class="eyebrow">DRE Gerencial</div>', unsafe_allow_html=True)
    st.subheader("Demonstrativo de Resultado — Valores em R$ e % sobre Receita")

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
            _rank = pie_df.sort_values("Valor", ascending=True)
            fig = px.bar(_rank, x="Valor", y="Conta", orientation="h",
                         title=f"Composição das contas — {mes_sel}",
                         hover_data={"% Receita": True, "Valor": True})
            fig.update_layout(showlegend=False, height=360, margin=dict(l=10,r=10,t=55,b=10), paper_bgcolor="white", plot_bgcolor="white")
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

    # Total de todas as despesas/saídas exibidas acima
    total_despesas_by_month = {
        m: (float(fornec_by_month.get(m, 0.0)) + float(ded_by_month.get(m, 0.0)) +
            float(pessoal_by_month.get(m, 0.0)) + float(adm_by_month.get(m, 0.0)) +
            float(com_by_month.get(m, 0.0)) + float(fin_by_month.get(m, 0.0)) +
            float(ret_by_month.get(m, 0.0)) + float(inv_by_month.get(m, 0.0)) +
            float(op_by_month.get(m, 0.0)))
        for m in range(1, 13)
    }

    # Resultado antes das retiradas, despesas financeiras e investimentos
    # (volta essas três linhas no saldo operacional)
    resultado_antes_by_month = {
        m: (float(saldo_by_month.get(m, 0.0)) + float(fin_by_month.get(m, 0.0)) +
            float(ret_by_month.get(m, 0.0)) + float(inv_by_month.get(m, 0.0)))
        for m in range(1, 13)
    }

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
        ("TOTAL DESPESAS", total_despesas_by_month),
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
    inicio = next((
        i for i, linha in enumerate(linhas)
        if [ _norm_txt(c) for c in linha.split(";")[:3] ]
        == ["duplicata", "emp", "dta.cad"]
    ), None)
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
    inicio = next((
        i for i, linha in enumerate(linhas)
        if [ _norm_txt(c) for c in linha.split(";")[:3] ]
        == ["empresa", "data", "vr. con"]
    ), None)
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


def _fiscal_resumo_impostos(excel_path, assinatura, saidas, ano_pagto, mes_pagto):
    """Impostos pagos no mês versus notas/receita do mês de competência anterior."""
    geral = read_sheet(excel_path, "DRE E DFC GERAL", assinatura)
    receita = read_sheet(excel_path, "RECEITA", assinatura)
    if geral is None or receita is None:
        raise ValueError("Não encontrei as abas 'DRE E DFC GERAL' e 'RECEITA'.")

    req_geral = {"CONTA DE RESULTADO", "DESPESA", "DTA.PAG", "VAL.PAG"}
    if not req_geral.issubset(geral.columns):
        faltam = sorted(req_geral.difference(geral.columns))
        raise ValueError("Faltam colunas no Excel: " + ", ".join(faltam))
    if not {"ANO", "MÊS", "RECEITA GRUPO"}.issubset(receita.columns):
        raise ValueError("Na aba RECEITA preciso de ANO, MÊS e RECEITA GRUPO.")

    referencia = pd.Timestamp(int(ano_pagto), int(mes_pagto), 1) - pd.DateOffset(months=1)
    ano_comp, mes_comp = int(referencia.year), int(referencia.month)

    g = geral.copy()
    g["_DATA"] = pd.to_datetime(g["DTA.PAG"], dayfirst=True, errors="coerce")
    g["_VALOR"] = pd.to_numeric(g["VAL.PAG"], errors="coerce").fillna(0.0)
    g["_CONTA"] = g["CONTA DE RESULTADO"].map(_norm_txt)
    g["_DESPESA"] = g["DESPESA"].map(_norm_txt)
    g = g[
        g["_CONTA"].str.startswith("00004 -")
        & (g["_DATA"].dt.year == int(ano_pagto))
        & (g["_DATA"].dt.month == int(mes_pagto))
        & ~g["_DESPESA"].str.contains(r"substituicao tributaria|icms\s*-?\s*st", regex=True)
    ].copy()

    regras = [
        ("Simples", r"simples"),
        ("PIS", r"\bpis\b"),
        ("COFINS", r"cofins"),
        ("ICMS", r"\bicms\b"),
        ("IRPJ", r"irpj"),
        ("CSLL", r"csll|contribuicao social sobre o lucro"),
    ]
    valores = []
    for imposto, padrao in regras:
        valor = float(g.loc[g["_DESPESA"].str.contains(padrao, regex=True), "_VALOR"].sum())
        valores.append({"Imposto": imposto, "Valor pago": valor})

    notas = saidas[
        (pd.to_numeric(saidas["ANO"], errors="coerce") == ano_comp)
        & (pd.to_numeric(saidas["MES_NUM"], errors="coerce") == mes_comp)
    ]
    total_notas = float(pd.to_numeric(notas["VLR_CONTABIL"], errors="coerce").fillna(0).sum())

    r = receita.copy()
    r["_ANO"] = pd.to_numeric(r["ANO"], errors="coerce")
    r["_MES"] = r["MÊS"].apply(parse_mes)
    r["_RECEITA"] = pd.to_numeric(r["RECEITA GRUPO"], errors="coerce").fillna(0.0)
    total_receita = float(r.loc[(r["_ANO"] == ano_comp) & (r["_MES"] == mes_comp), "_RECEITA"].sum())

    resumo = pd.DataFrame(valores)
    resumo["% sobre notas"] = resumo["Valor pago"].div(total_notas).mul(100) if total_notas else 0.0
    resumo["% sobre receita"] = resumo["Valor pago"].div(total_receita).mul(100) if total_receita else 0.0
    return resumo, total_notas, total_receita, referencia


def pagina_controles_fiscais(excel_path=None, assinatura_excel=None):
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

    st.divider()
    st.markdown("## Impostos pagos × faturamento do mês anterior")
    st.caption(
        "Consolidado do grupo. ICMS-ST não entra. O mês escolhido é o do pagamento; "
        "notas emitidas e Receita Grupo são buscadas no mês imediatamente anterior."
    )
    if not excel_path:
        st.warning("Não encontrei 'DRE E DFC GERAL.xlsx' para montar o comparativo tributário.")
        return

    base_periodos = read_sheet(excel_path, "DRE E DFC GERAL", assinatura_excel).copy()
    conta_periodo = base_periodos.get("CONTA DE RESULTADO", pd.Series(index=base_periodos.index, dtype=str)).map(_norm_txt)
    despesa_periodo = base_periodos.get("DESPESA", pd.Series(index=base_periodos.index, dtype=str)).map(_norm_txt)
    mask_periodos = (
        conta_periodo.str.startswith("00004 -")
        & despesa_periodo.str.contains(r"simples|\bpis\b|cofins|\bicms\b|irpj|csll", regex=True)
        & ~despesa_periodo.str.contains(r"substituicao tributaria|icms\s*-?\s*st", regex=True)
    )
    datas_impostos = pd.to_datetime(
        base_periodos.loc[mask_periodos, "DTA.PAG"], dayfirst=True, errors="coerce"
    )
    periodos = sorted({(int(d.year), int(d.month)) for d in datas_impostos.dropna()})
    if not periodos:
        st.info("Não encontrei meses de pagamento no plano de contas.")
        return
    periodo_sel = st.selectbox(
        "Mês de pagamento dos impostos",
        options=periodos,
        index=len(periodos) - 1,
        format_func=lambda p: f"{MESES_NOME_FISCAL[p[1]]}/{p[0]}",
        key="fiscal_periodo_impostos",
    )
    try:
        resumo_imp, total_notas, total_receita, competencia = _fiscal_resumo_impostos(
            excel_path, assinatura_excel, saidas, periodo_sel[0], periodo_sel[1]
        )
    except Exception as e:
        st.error(f"Não foi possível calcular os impostos: {e}")
        return

    total_impostos = float(resumo_imp["Valor pago"].sum())
    pct_notas = total_impostos / total_notas * 100 if total_notas else 0.0
    pct_receita = total_impostos / total_receita * 100 if total_receita else 0.0
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Impostos pagos", fmt_brl_display(total_impostos))
    c2.metric(f"Notas {MESES_NOME_FISCAL[competencia.month]}/{competencia.year}", fmt_brl_display(total_notas))
    c3.metric("Receita Grupo", fmt_brl_display(total_receita))
    c4.metric("Impostos / notas", fmt_pct(pct_notas))
    c5.metric("Impostos / receita", fmt_pct(pct_receita))

    tabela_imp = pd.concat([
        resumo_imp,
        pd.DataFrame([{
            "Imposto": "TOTAL DE IMPOSTOS",
            "Valor pago": total_impostos,
            "% sobre notas": pct_notas,
            "% sobre receita": pct_receita,
        }]),
    ], ignore_index=True).rename(columns={
        "Valor pago": "Valor pago (R$)",
        "% sobre notas": "% sobre notas emitidas",
        "% sobre receita": "% sobre Receita Grupo",
    })
    st.dataframe(
        tabela_imp.style.format({
            "Valor pago (R$)": fmt_brl_display,
            "% sobre notas emitidas": fmt_pct,
            "% sobre Receita Grupo": fmt_pct,
        }).hide(axis="index"),
        use_container_width=True,
        height=300,
    )





# =========================
# Navegação multipage (arquivo único)
# =========================
st.set_page_config(
    page_title="Painel Geral | Dauto & Única",
    page_icon=_img_from_b64(LOGO_DAUTO_B64),
    layout="wide",
)
_inject_modern_ui()

def _localizar_excel_principal():
    pasta_app = Path(__file__).resolve().parent
    candidatos = [
        pasta_app / "DRE E DFC GERAL.xlsx",
        Path.cwd() / "DRE E DFC GERAL.xlsx",
    ]
    for p in candidatos:
        if p.exists():
            return str(p)
    return None

excel_path = _localizar_excel_principal()
sig = None
if excel_path:
    try:
        sig = os.path.getmtime(excel_path)
    except Exception:
        sig = None

anos_disponiveis = []
if excel_path:
    try:
        _tmp = read_sheet(excel_path, "RECEITA", sig)
        if _tmp is not None and "ANO" in _tmp.columns:
            anos_disponiveis = sorted(
                pd.to_numeric(_tmp["ANO"], errors="coerce").dropna().astype(int).unique().tolist()
            )
    except Exception:
        anos_disponiveis = []

if not anos_disponiveis:
    anos_disponiveis = [pd.Timestamp.today().year]

with st.sidebar:
    _sidebar_brand()
    st.markdown("## Filtros gerais")
    ano_ref = st.selectbox("Ano", anos_disponiveis, index=len(anos_disponiveis)-1)

    # Meses em checklist
    st.markdown("**Meses**")
    meses_pt_sel = []
    _mes_cols = st.columns(2, gap="small")
    for _idx, _mes in enumerate(MESES_PT):
        with _mes_cols[_idx % 2]:
            if st.checkbox(_mes, value=True, key=f"filtro_mes_{_mes}"):
                meses_pt_sel.append(_mes)

    st.divider()
    if st.button("Atualizar dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

def _page_dre():
    if not excel_path:
        st.error("Não encontrei 'DRE E DFC GERAL.xlsx' no repositório.")
        return
    pagina_dre_geral(excel_path, ano_ref, meses_pt_sel)

def _page_dfc():
    if not excel_path:
        st.error("Não encontrei 'DRE E DFC GERAL.xlsx' no repositório.")
        return
    pagina_dfc_geral(excel_path, ano_ref, meses_pt_sel)

def _page_faturamento():
    if not excel_path:
        st.error("Não encontrei 'DRE E DFC GERAL.xlsx' no repositório.")
        return
    pagina_faturamento(excel_path, ano_ref, meses_pt_sel)

def _page_controles_fiscais():
    pagina_controles_fiscais(excel_path, sig)

paginas = {
    "Financeiro": [
        st.Page(_page_dre, title="DRE Geral", icon="📈"),
        st.Page(_page_dfc, title="DFC Geral", icon="💵"),
        st.Page(_page_faturamento, title="Faturamento", icon="🧾"),
    ],
    "Controles": [
        st.Page(_page_controles_fiscais, title="Controles Fiscais", icon="🧮"),
    ],
}

pg = st.navigation(paginas)
pg.run()
