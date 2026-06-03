# GERAL.py
import re
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px
import os
import glob

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


    render_agente_bi(excel_path, ano_ref, meses_pt_sel, contexto_padrao="DRE")

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


    render_agente_bi(excel_path, ano_ref, meses_pt_sel, contexto_padrao="DFC")

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
# Agente de BI — DRE / DFC
# =========================
def _periodo_meses_from_text(pergunta: str, meses_padrao=None):
    """Extrai meses citados na pergunta. Se encontrar 'de X a Y', devolve intervalo."""
    meses_padrao = meses_padrao or MESES_PT
    txt = _norm_txt(pergunta)
    aliases = {
        "jan": 1, "janeiro": 1,
        "fev": 2, "fevereiro": 2,
        "mar": 3, "marco": 3, "março": 3,
        "abr": 4, "abril": 4,
        "mai": 5, "maio": 5,
        "jun": 6, "junho": 6,
        "jul": 7, "julho": 7,
        "ago": 8, "agosto": 8,
        "set": 9, "setembro": 9,
        "out": 10, "outubro": 10,
        "nov": 11, "novembro": 11,
        "dez": 12, "dezembro": 12,
    }
    encontrados = []
    for nome, num in aliases.items():
        if re.search(rf"\b{re.escape(nome)}\b", txt):
            encontrados.append(num)
    # números soltos 1..12 quando acompanhados de mês/mes
    nums = re.findall(r"\b(?:mes|mês)\s*(\d{1,2})\b|\b(\d{1,2})\s*(?:mes|mês)\b", txt)
    for a, b in nums:
        n = int(a or b)
        if 1 <= n <= 12:
            encontrados.append(n)
    encontrados = sorted(set(encontrados))
    if len(encontrados) >= 2:
        ini, fim = encontrados[0], encontrados[-1]
        if ini <= fim:
            return list(range(ini, fim + 1))
    if len(encontrados) == 1:
        return encontrados
    return [MES_PT_TO_NUM[m] for m in (meses_padrao if meses_padrao else MESES_PT)]


def _detectar_visao_bi(pergunta: str):
    txt = _norm_txt(pergunta)
    tem_dre = "dre" in txt or "competencia" in txt or "resultado operacional" in txt
    tem_dfc = "dfc" in txt or "caixa" in txt or "recebimento" in txt or "saldo operacional" in txt
    if tem_dre and not tem_dfc:
        return "DRE"
    if tem_dfc and not tem_dre:
        return "DFC"
    return None


def _inferir_intencao_bi(pergunta: str):
    txt = _norm_txt(pergunta)
    if any(w in txt for w in ["lucro", "resultado", "saldo operacional", "resultado operacional"]):
        return "resultado"
    if any(w in txt for w in ["comparativo", "comparar", "evolucao", "evolução", "mes a mes", "mês a mês"]):
        return "comparativo"
    if any(w in txt for w in ["%", "percentual", "por cento", "representa", "participacao", "participação"]):
        return "percentual"
    if any(w in txt for w in ["maior despesa", "top despesa", "principais despesas", "ranking"]):
        return "maior_despesa"
    if any(w in txt for w in ["receita", "faturamento"]):
        return "receita"
    if "recebimento" in txt or "recebimentos" in txt:
        return "recebimento"
    if "outras receitas" in txt or "outra receita" in txt:
        return "outras_receitas"
    return "diagnostico"


def _bi_prefixos_por_visao(visao: str):
    base = {
        "OUTRAS RECEITAS": "00003 -",
        "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)": "00004 -",
        "DESPESAS COM PESSOAL": "00006 -",
        "DESPESAS ADMINISTRATIVAS": "00007 -",
        "DESPESAS COMERCIAIS": "00009 -",
        "DESPESAS FINANCEIRAS": "00011 -",
        "RETIRADAS SÓCIOS": "00016 -",
        "INVESTIMENTOS": "00015 -",
        "DESPESAS OPERACIONAIS": "00017 -",
    }
    if visao == "DFC":
        base = {"FORNECEDORES": "00012 -", **base}
    else:
        base = {"COMPRAS EMISSÃO": None, **base}
    return base


def _match_opcao_texto(pergunta: str, opcoes: list[str]) -> str | None:
    txt = _norm_txt(pergunta)
    melhor = None
    melhor_score = 0
    for op in opcoes:
        opn = _norm_txt(op)
        palavras = [p for p in re.split(r"\W+", opn) if len(p) > 3 and p not in {"despesas", "receitas", "sobre", "vendas"}]
        score = sum(1 for p in palavras if p in txt)
        if opn in txt:
            score += 5
        if score > melhor_score:
            melhor, melhor_score = op, score
    return melhor if melhor_score > 0 else None


def _montar_bases_bi(excel_path, ano_ref: int, meses_pt_sel=None):
    """Monta bases mensais de DRE e DFC para o Agente de BI."""
    meses_pt = (meses_pt_sel or []) or MESES_PT
    meses_nums = [MES_PT_TO_NUM[m] for m in meses_pt]
    g = prep_geral_year(excel_path, ano_ref, sig)
    g_next = prep_geral_year(excel_path, int(ano_ref) + 1, sig)
    df_receita = read_sheet(excel_path, "RECEITA", sig)
    df_nfs = read_sheet(excel_path, "NOTAS EMITIDAS", sig)
    df_rec = read_sheet(excel_path, "RECEBIMENTO", sig)
    if g is None:
        return None

    receita_by_month = agg_by_month_from_ano_mes(df_receita, "RECEITA GRUPO", "ANO", "MÊS", ano_ref) if df_receita is not None else {m: 0.0 for m in range(1,13)}
    compras_by_month = agg_by_month_from_ano_mes(df_nfs, "NFS EMITIDAS", "ANO", "MÊS", ano_ref) if df_nfs is not None else {m: 0.0 for m in range(1,13)}
    receb_by_month = agg_by_month_from_ano_mes(df_rec, "RECEBIMENTO", "ANO", "MÊS", ano_ref) if df_rec is not None else {m: 0.0 for m in range(1,13)}

    def sum_prefix(prefix: str, base=None):
        base = g if base is None else base
        if base is None or base.empty:
            return {m: 0.0 for m in range(1, 13)}
        mask = base["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)
        grp = base[mask].groupby("_mes")["_v"].sum()
        return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}

    def sum_shift_dre(prefix: str, exclude_icmsst: bool = False):
        d = g[g["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)].copy()
        dn = g_next[g_next["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)].copy() if g_next is not None else pd.DataFrame()
        if exclude_icmsst:
            for base in [d, dn]:
                if base is not None and not base.empty:
                    target_norm = _norm_txt(_DED_EXCL_DRE)
                    for c in ["DESPESA", "CONTA DE RESULTADO", "HISTÓRICO", "HISTORICO"]:
                        if c in base.columns:
                            s_norm = base[c].astype(str).apply(_norm_txt)
                            base.drop(base[s_norm.str.contains(target_norm, na=False) | s_norm.str.contains("02.07.008", na=False)].index, inplace=True)
                            break
        src = d.groupby("_mes")["_v"].sum() if not d.empty else pd.Series(dtype="float64")
        src_next = dn.groupby("_mes")["_v"].sum() if dn is not None and not dn.empty else pd.Series(dtype="float64")
        out = {}
        for m in range(1, 13):
            out[m] = float(src.get(m + 1, 0.0)) if m < 12 else float(src_next.get(1, 0.0))
        return out

    outras_receitas = sum_outras_receitas_prepped(g)
    receita_total = {m: float(receita_by_month.get(m, 0)) + float(outras_receitas.get(m, 0)) for m in range(1,13)}
    receb_total = {m: float(receb_by_month.get(m, 0)) + float(outras_receitas.get(m, 0)) for m in range(1,13)}

    dre = {
        "+ RECEITA": receita_by_month,
        "+ OUTRAS RECEITAS": outras_receitas,
        "- COMPRAS EMISSÃO": compras_by_month,
        "- DEDUÇÕES (IMPOSTOS SOBRE VENDAS)": sum_shift_dre("00004 -", True),
        "- DESPESAS COM PESSOAL": sum_shift_dre("00006 -", False),
        "- DESPESAS ADMINISTRATIVAS": sum_prefix("00007 -"),
        "- DESPESAS COMERCIAIS": sum_prefix("00009 -"),
        "- DESPESAS FINANCEIRAS": sum_prefix("00011 -"),
        "- RETIRADAS SÓCIOS": sum_prefix("00016 -"),
        "- INVESTIMENTOS": sum_prefix("00015 -"),
        "- DESPESAS OPERACIONAIS": sum_prefix("00017 -"),
    }
    resultado_operacional = {}
    for m in range(1,13):
        saidas = sum(float(dre[k].get(m,0)) for k in dre if k.startswith("-"))
        resultado_operacional[m] = float(receita_total.get(m,0)) - saidas
    dre["RESULTADO ANTES DAS RETIRADAS E DESP. FINANCEIRAS"] = {m: resultado_operacional[m] + dre["- DESPESAS FINANCEIRAS"][m] + dre["- RETIRADAS SÓCIOS"][m] for m in range(1,13)}
    dre["RESULTADO OPERACIONAL"] = resultado_operacional

    pmap = dfc_prefix_map()
    dfc = {
        "+ RECEBIMENTOS": receb_by_month,
        "+ OUTRAS RECEITAS": outras_receitas,
        "- FORNECEDORES": sum_prefix(pmap["FORNECEDORES"]),
        "- DEDUÇÕES (IMPOSTOS SOBRE VENDAS)": sum_prefix(pmap["DEDUÇÕES (IMPOSTOS SOBRE VENDAS)"]),
        "- DESPESAS COM PESSOAL": sum_prefix(pmap["DESPESAS COM PESSOAL"]),
        "- DESPESAS ADMINISTRATIVAS": sum_prefix(pmap["DESPESAS ADMINISTRATIVAS"]),
        "- DESPESAS COMERCIAIS": sum_prefix(pmap["DESPESAS COMERCIAIS"]),
        "- DESPESAS FINANCEIRAS": sum_prefix(pmap["DESPESAS FINANCEIRAS"]),
        "- RETIRADAS SÓCIOS": sum_prefix("00016 -"),
        "- INVESTIMENTOS": sum_prefix(pmap["INVESTIMENTOS"]),
        "- DESPESAS OPERACIONAIS": sum_prefix(pmap["DESPESAS OPERACIONAIS"]),
    }
    saldo_operacional = {}
    for m in range(1,13):
        saidas = sum(float(dfc[k].get(m,0)) for k in dfc if k.startswith("-"))
        saldo_operacional[m] = float(receb_total.get(m,0)) - saidas
    dfc["RESULTADO ANTES DAS RETIRADAS E DESP. FINANCEIRAS"] = {m: saldo_operacional[m] + dfc["- DESPESAS FINANCEIRAS"][m] + dfc["- RETIRADAS SÓCIOS"][m] for m in range(1,13)}
    dfc["SALDO OPERACIONAL"] = saldo_operacional

    return {"DRE": dre, "DFC": dfc, "g": g, "receita_total": receita_total, "receb_total": receb_total, "meses_pt": meses_pt, "meses_nums": meses_nums}


def _bi_tabela_mensal(label: str, by_month: dict, denom: dict, meses_nums: list[int], pct_label: str):
    rows = []
    for m in meses_nums:
        valor = float(by_month.get(m, 0.0))
        base = float(denom.get(m, 0.0))
        rows.append({"Mês": MES_NUM_TO_PT[m], "Valor": valor, pct_label: (valor / base * 100.0) if base else 0.0})
    df = pd.DataFrame(rows)
    total = float(df["Valor"].sum()) if not df.empty else 0.0
    base_total = float(sum(denom.get(m, 0.0) for m in meses_nums))
    df.loc[len(df)] = {"Mês": "ACUMULADO", "Valor": total, pct_label: (total / base_total * 100.0) if base_total else 0.0}
    st.markdown(f"**{label}**")
    render_sticky_table(df, value_cols=["Valor"], pct_cols=[pct_label], highlight_row_label="ACUMULADO")
    return total, (total / base_total * 100.0) if base_total else 0.0


def _bi_detalhar_despesas(excel_path, ano_ref, visao, grupo, meses_nums, denom_total):
    g = prep_geral_year(excel_path, ano_ref, sig)
    if g is None or g.empty:
        return pd.DataFrame()
    base = g[g["_mes"].isin(meses_nums)].copy()
    if grupo == "OUTRAS RECEITAS":
        base = base[_mask_outras_receitas(base)]
    else:
        prefix = _bi_prefixos_por_visao(visao).get(grupo)
        if prefix:
            base = base[base["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith(prefix)]
    if base.empty:
        return pd.DataFrame()
    if "DESPESA" not in base.columns:
        base["DESPESA"] = "—"
    base["DESPESA_SINT"] = base["DESPESA"].apply(sintetizar_despesa)
    det = base.groupby("DESPESA_SINT", dropna=False)["_v"].sum().reset_index().rename(columns={"_v": "Valor", "DESPESA_SINT": "Despesa"})
    det["% Base"] = det["Valor"].apply(lambda x: (x / denom_total * 100.0) if denom_total else 0.0)
    return det.sort_values("Valor", ascending=False).reset_index(drop=True)


def render_agente_bi(excel_path, ano_ref, meses_pt_sel=None, contexto_padrao="DRE"):
    st.divider()
    st.subheader("Agente de BI — Perguntas gerenciais")
    st.caption("Pergunte sobre lucro, resultado, receita, recebimentos, outras receitas, comparativos por conta, maior despesa e participação percentual.")

    bases = _montar_bases_bi(excel_path, ano_ref, meses_pt_sel)
    if bases is None:
        st.info("Não consegui montar a base do Agente de BI com as abas disponíveis.")
        return

    exemplos = [
        "Qual foi o lucro no DRE?",
        "Qual foi o lucro no DFC?",
        "Faça um comparativo de despesas administrativas de janeiro a maio",
        "Quanto representa despesas comerciais no DRE?",
        "Qual foi a maior despesa da conta despesas administrativas?",
        "Quanto foi outras receitas?",
        "Qual foi a receita?",
        "Qual foi o recebimento?",
    ]
    with st.expander("Exemplos de perguntas", expanded=False):
        st.write("\n".join([f"- {e}" for e in exemplos]))

    pergunta = st.text_input("Digite sua pergunta para o Agente de BI", key=f"agente_bi_pergunta_{contexto_padrao}", placeholder="Ex.: Faça um comparativo de despesas administrativas de janeiro a maio")
    if not pergunta:
        return

    intencao = _inferir_intencao_bi(pergunta)
    visao = _detectar_visao_bi(pergunta) or contexto_padrao
    if _detectar_visao_bi(pergunta) is None and intencao == "resultado":
        visao = st.radio("Esse lucro/resultado é no DRE ou no DFC?", ["DRE", "DFC"], horizontal=True, key=f"agente_bi_visao_{contexto_padrao}")

    meses_nums = _periodo_meses_from_text(pergunta, meses_pt_sel)
    meses_nums = [m for m in meses_nums if 1 <= int(m) <= 12]
    if not meses_nums:
        meses_nums = bases["meses_nums"]
    pct_label = "% Receita" if visao == "DRE" else "% Recebimentos"
    denom = bases["receita_total"] if visao == "DRE" else bases["receb_total"]
    base_total = float(sum(denom.get(m, 0.0) for m in meses_nums))

    st.markdown(f"**Intenção identificada:** {intencao.replace('_', ' ').title()} | **Visão:** {visao} | **Período:** {', '.join(MES_NUM_TO_PT[m] for m in meses_nums)} / {ano_ref}")

    linhas = bases[visao]
    opcoes_contas = list(linhas.keys())
    opcoes_grupos = list(_bi_prefixos_por_visao(visao).keys())

    if intencao == "resultado":
        if visao == "DFC":
            total1, pct1 = _bi_tabela_mensal("RESULTADO ANTES DAS RETIRADAS E DESP. FINANCEIRAS", linhas["RESULTADO ANTES DAS RETIRADAS E DESP. FINANCEIRAS"], denom, meses_nums, pct_label)
            total2, pct2 = _bi_tabela_mensal("SALDO OPERACIONAL", linhas["SALDO OPERACIONAL"], denom, meses_nums, pct_label)
            st.success(f"Leitura gerencial: antes das retiradas e despesas financeiras, o caixa gerou {fmt_brl_display(total1)} no período. Após todas as saídas operacionais, o saldo operacional ficou em {fmt_brl_display(total2)} ({fmt_pct(pct2)} da base de recebimentos).")
        else:
            total1, pct1 = _bi_tabela_mensal("RESULTADO ANTES DAS RETIRADAS E DESP. FINANCEIRAS", linhas["RESULTADO ANTES DAS RETIRADAS E DESP. FINANCEIRAS"], denom, meses_nums, pct_label)
            total2, pct2 = _bi_tabela_mensal("RESULTADO OPERACIONAL", linhas["RESULTADO OPERACIONAL"], denom, meses_nums, pct_label)
            st.success(f"Leitura gerencial: antes das retiradas e despesas financeiras, o negócio gerou {fmt_brl_display(total1)}. Considerando todas as linhas do DRE, o resultado operacional ficou em {fmt_brl_display(total2)} ({fmt_pct(pct2)} da receita).")
        return

    if intencao in {"receita", "recebimento", "outras_receitas"}:
        if intencao == "receita":
            visao_local, linha = "DRE", "+ RECEITA"
            denom_local, pct_local = bases["receita_total"], "% Receita"
        elif intencao == "recebimento":
            visao_local, linha = "DFC", "+ RECEBIMENTOS"
            denom_local, pct_local = bases["receb_total"], "% Recebimentos"
        else:
            visao_local, linha = visao, "+ OUTRAS RECEITAS"
            denom_local, pct_local = denom, pct_label
        _bi_tabela_mensal(linha, bases[visao_local][linha], denom_local, meses_nums, pct_local)
        return

    conta_detectada = _match_opcao_texto(pergunta, opcoes_contas)
    grupo_detectado = _match_opcao_texto(pergunta, opcoes_grupos)

    if intencao in {"comparativo", "percentual"}:
        conta = conta_detectada
        if conta is None:
            grupo = grupo_detectado
            if grupo:
                conta = next((c for c in opcoes_contas if _norm_txt(grupo) in _norm_txt(c)), None)
        if conta is None:
            conta = st.selectbox("Não identifiquei a conta. Selecione a conta para análise:", opcoes_contas, key=f"agente_bi_conta_{contexto_padrao}")
        total, pct = _bi_tabela_mensal(conta, linhas[conta], denom, meses_nums, pct_label)
        if intencao == "percentual":
            st.info(f"No acumulado selecionado, **{conta}** representa **{fmt_pct(pct)}** da base de {'receita' if visao == 'DRE' else 'recebimentos'}, com valor de **{fmt_brl_display(total)}**.")
        else:
            vals = [float(linhas[conta].get(m, 0.0)) for m in meses_nums]
            maior_m = meses_nums[vals.index(max(vals))] if vals else None
            menor_m = meses_nums[vals.index(min(vals))] if vals else None
            if maior_m and menor_m:
                st.info(f"Leitura gerencial: maior valor em **{MES_NUM_TO_PT[maior_m]}** e menor valor em **{MES_NUM_TO_PT[menor_m]}**. Acumulado do período: **{fmt_brl_display(total)}**.")
        return

    if intencao == "maior_despesa":
        grupo = grupo_detectado or st.selectbox("Qual conta/grupo deseja abrir?", opcoes_grupos, key=f"agente_bi_grupo_{contexto_padrao}")
        det = _bi_detalhar_despesas(excel_path, ano_ref, visao, grupo, meses_nums, base_total)
        if det.empty:
            st.info("Não encontrei lançamentos detalhados para essa conta no período selecionado.")
            return
        top = det.iloc[0]
        st.success(f"Maior despesa em **{grupo}**: **{top['Despesa']}**, com **{fmt_brl_display(top['Valor'])}**, representando **{fmt_pct(top['% Base'])}** da base.")
        st.dataframe(det.head(20).style.format({"Valor": lambda x: fmt_brl_display(x), "% Base": lambda x: fmt_pct(x)}).hide(axis="index"), use_container_width=True)
        return

    st.warning("Não consegui fixar a intenção com segurança. Trata-se de Conta de resultado ou despesa? Informe também o período, o mês e o ano.")
    c1, c2, c3 = st.columns(3)
    with c1:
        visao_sel = st.selectbox("Visão", ["DRE", "DFC"], index=0 if contexto_padrao == "DRE" else 1, key=f"agente_bi_fallback_visao_{contexto_padrao}")
    with c2:
        tipo_sel = st.selectbox("Tipo", ["Conta de resultado", "Despesa detalhada"], key=f"agente_bi_fallback_tipo_{contexto_padrao}")
    with c3:
        meses_sel = st.multiselect("Período", MESES_PT, default=[MES_NUM_TO_PT[m] for m in meses_nums], key=f"agente_bi_fallback_meses_{contexto_padrao}")
    meses_fb = [MES_PT_TO_NUM[m] for m in meses_sel] or meses_nums
    denom_fb = bases["receita_total"] if visao_sel == "DRE" else bases["receb_total"]
    pct_fb = "% Receita" if visao_sel == "DRE" else "% Recebimentos"
    if tipo_sel == "Conta de resultado":
        conta_fb = st.selectbox("Conta", list(bases[visao_sel].keys()), key=f"agente_bi_fallback_conta_{contexto_padrao}")
        _bi_tabela_mensal(conta_fb, bases[visao_sel][conta_fb], denom_fb, meses_fb, pct_fb)
    else:
        grupo_fb = st.selectbox("Despesa/Grupo", list(_bi_prefixos_por_visao(visao_sel).keys()), key=f"agente_bi_fallback_grupo_{contexto_padrao}")
        det = _bi_detalhar_despesas(excel_path, ano_ref, visao_sel, grupo_fb, meses_fb, float(sum(denom_fb.get(m, 0.0) for m in meses_fb)))
        if det.empty:
            st.info("Sem lançamentos para essa seleção.")
        else:
            st.dataframe(det.head(30).style.format({"Valor": lambda x: fmt_brl_display(x), "% Base": lambda x: fmt_pct(x)}).hide(axis="index"), use_container_width=True)

# =========================
# Main: lê Excel 1x e usa nas páginas
# =========================
st.set_page_config(page_title="GERAL", layout="wide")
st.sidebar.title("Menu")

# Leitura automática do Excel (mesma pasta do app)
def _auto_find_excel() -> str | None:
    # Prioriza nomes comuns
    preferred = ["DRE E DFC GERAL.xlsx", "DRE_E_DFC_GERAL.xlsx", "BASE.xlsx", "BASE .xlsx"]
    for fn in preferred:
        if os.path.exists(fn):
            return fn
    # Qualquer xlsx/xlsm na pasta (pega o mais recente)
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



EXCEL_PATH = excel_path  # alias padrão
def excel_signature(path: str):
    """
    Assinatura do arquivo para invalidar caches quando o Excel for atualizado (mesmo mantendo o mesmo nome).
    Retorna (mtime_ns, size).
    """
    stt = os.stat(path)
    return (stt.st_mtime_ns, stt.st_size)


# Assinatura atual do arquivo (usada para invalidar st.cache_data quando o Excel muda)
sig = excel_signature(EXCEL_PATH)
EXCEL_SIG = sig

st.sidebar.caption(f"Excel: **{excel_path}**")
sheet_names = get_sheet_names(excel_path, sig)
if not sheet_names:
    st.sidebar.error(f"Não consegui abrir '{excel_path}'.")
    st.stop()
st.sidebar.success("Excel carregado")

# Filtros gerais
meses_pt_sel = st.sidebar.multiselect("Meses", options=MESES_PT, default=MESES_PT)

# Descobre anos disponíveis

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

pagina = st.sidebar.radio("Selecione:", ["DRE Geral", "DFC Geral", "Faturamento"])

if pagina == "DRE Geral":
    pagina_dre_geral(excel_path, ano_ref, meses_pt_sel)
elif pagina == "DFC Geral":
    pagina_dfc_geral(excel_path, ano_ref, meses_pt_sel)
else:
    pagina_faturamento(excel_path, ano_ref, meses_pt_sel)