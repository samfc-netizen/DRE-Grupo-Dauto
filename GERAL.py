# GERAL.py
import re
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px
import os
import glob

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
                               ("IMPOSTOS E FOLHA", df_if), ("DRE E DFC GERAL", df_geral)] if df is None]
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

    # IMPOSTOS E FOLHA (para DRE mantém shift +1 mês)
    req_if = {"CONTA DE RESULTADO", "DTA.PAG", "VAL.PAG"}
    if not req_if.issubset(set(df_if.columns)):
        st.error("Na aba IMPOSTOS E FOLHA preciso das colunas: 'CONTA DE RESULTADO', 'DTA.PAG', 'VAL.PAG'.")
        return
    i = prep_impostos_folha_dre(excel_path, ano_ref, sig)
    if i is None:
        st.error("Não encontrei a aba IMPOSTOS E FOLHA.")
        return

    # Exceção solicitada: na DRE, dentro de DEDUÇÕES (IMPOSTOS SOBRE VENDAS),
    # desconsiderar a despesa '02.07.008-ICMS- SUBSTITUIÇÃO TRIBUTARIA' (no DFC permanece igual).
    _DED_EXCL_DRE = "02.07.008-ICMS- SUBSTITUIÇÃO TRIBUTARIA"
    def _norm_txt(x):
        s = "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x)
        # Normaliza para comparação robusta (remove acentos/diacríticos e padroniza hífens/espaços)
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        s = s.replace("–", "-").replace("—", "-")
        s = re.sub(r"\s*[-]+\s*", "-", s)     # remove espaços ao redor de hífens
        s = re.sub(r"\s+", " ", s).strip().upper()
        return s


    ded_mask = i["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith("00004 -")
    pes_mask = i["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith("00006 -")
    i_ded = i[ded_mask].copy()
    if "DESPESA" in i_ded.columns:
        i_ded = i_ded[i_ded["DESPESA"].apply(_norm_txt) != _norm_txt(_DED_EXCL_DRE)]
    deducoes_by_month = {m: float(i_ded.groupby("_mes_ref")["_v"].sum().get(m, 0.0)) for m in range(1, 13)}
    pessoal_by_month = {m: float(i[pes_mask].groupby("_mes_ref")["_v"].sum().get(m, 0.0)) for m in range(1, 13)}

    # Geral por prefixos
    g = prep_geral_year(excel_path, ano_ref, sig)
    if g is None:
        st.error("Não encontrei a aba DRE E DFC GERAL.")
        return

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

    resultado_by_month = {}
    for m in range(1, 13):
        outros = (compras_by_month[m] + deducoes_by_month[m] + pessoal_by_month[m] +
                  adm_by_month[m] + com_by_month[m] + fin_by_month[m] + inv_by_month[m] + op_by_month[m] + ret_by_month[m])
        resultado_by_month[m] = receita_by_month[m] - outros


    # Resultado antes das retiradas e despesas financeiras (volta essas duas linhas no resultado)
    resultado_antes_by_month = {m: float(resultado_by_month.get(m, 0.0)) + float(fin_by_month.get(m, 0.0)) + float(ret_by_month.get(m, 0.0)) for m in range(1, 13)}

    linhas = [
        ("+ RECEITA", receita_by_month),
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
            rec = float(receita_by_month.get(m, 0.0))
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
    receita_acum = float(sum(receita_by_month.get(m, 0.0) for m in meses_nums))
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

    st.dataframe(dre.style.format(fmt_map).apply(style_resultado, axis=1).hide(axis="index"),
                 use_container_width=True)

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
    receita_mes = float(sum(float(receita_by_month.get(m, 0.0)) for m in meses_nums_drill))

    def _sum_months(by_month):
        return float(sum(float(by_month.get(m, 0.0)) for m in meses_nums_drill))

    contas_mes = {
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

    if grupo_sel in {"DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", "DESPESAS COM PESSOAL"}:
        base_raw = i.copy()
        base_raw = base_raw[base_raw["_mes_ref"].isin(meses_nums_drill)].copy()
        if grupo_sel == "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)":
            base_raw = base_raw[base_raw["CONTA DE RESULTADO"].astype(str).str.strip().str.startswith("00004 -")]
            if "DESPESA" in base_raw.columns:
                base_raw = base_raw[base_raw["DESPESA"].apply(_norm_txt) != _norm_txt(_DED_EXCL_DRE)]
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

    g = prep_geral_year(excel_path, ano_ref, sig)
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

    saldo_by_month = {}
    for m in range(1, 13):
        saidas = (fornec_by_month[m] + ded_by_month[m] + pessoal_by_month[m] + adm_by_month[m] +
                  com_by_month[m] + fin_by_month[m] + inv_by_month[m] + op_by_month[m] + ret_by_month[m])
        saldo_by_month[m] = receb_by_month[m] - saidas

    # Resultado antes das retiradas e despesas financeiras (volta essas duas linhas no resultado)
    resultado_antes_by_month = {m: float(saldo_by_month.get(m, 0.0)) + float(fin_by_month.get(m, 0.0)) + float(ret_by_month.get(m, 0.0)) for m in range(1, 13)}

    linhas = [
        ("+ RECEBIMENTOS", receb_by_month),
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
            rec = float(receb_by_month.get(m, 0.0))
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
    receb_acum = float(sum(receb_by_month.get(m, 0.0) for m in meses_nums))
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

    st.dataframe(dfc.style.format(fmt_map).apply(style_saldo, axis=1).hide(axis="index"),
                 use_container_width=True)

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
    receb_mes = float(sum(float(receb_by_month.get(m, 0.0)) for m in meses_nums_drill))

    def _sum_months(by_month):
        return float(sum(float(by_month.get(m, 0.0)) for m in meses_nums_drill))

    contas_mes = {
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

pagina = st.sidebar.radio("Selecione:", ["DRE Geral", "DFC Geral"])

if pagina == "DRE Geral":
    pagina_dre_geral(excel_path, ano_ref, meses_pt_sel)
else:
    pagina_dfc_geral(excel_path, ano_ref, meses_pt_sel)
