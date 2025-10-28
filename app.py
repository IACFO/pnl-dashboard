# app.py
# -----------------------------------------------------------
# P&L – Projeção, Realizados, Comparativos e Highlights (com abas)
# -----------------------------------------------------------

import io
import re
import unicodedata
import numpy as np
import pandas as pd
import streamlit as st
import altair as alt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ==================== VISUAL / CSS ====================
CB = {
    "blue":  "#0033A0",
    "red":   "#E1002A",
    "yellow":"#FFCC00",
    "ink":   "#0E1A2A",
    "bg":    "#FFFFFF",
    "bg2":   "#F5F7FB",
    "muted": "#8A94A6",
    "green": "#0A8454",
    "gray":  "#8A94A6",
}

def inject_css():
    st.markdown(f"""
    <style>
    .block-container {{
        padding-top: 4.5rem !important;
        padding-bottom: 2rem !important;
        overflow: visible !important;
    }}
    header[data-testid="stHeader"] {{ background: {CB["bg"]}; }}

    .brand-section {{
        width: 100%; display: block; box-sizing: border-box;
        background: linear-gradient(90deg, {CB["blue"]} 0%, {CB["red"]} 100%);
        color: #fff; padding: 22px 24px; border-radius: 14px;
        font-weight: 800; letter-spacing: .2px; margin: 8px 0 14px 0;
        text-align: center; line-height: 1.25; font-size: clamp(18px, 2.0vw, 26px);
        box-shadow: 0 2px 10px rgba(0,0,0,.06);
    }}
    .table-wrap {{
        max-height: 70vh; overflow-y: auto; overflow-x: auto;
        border: 1px solid #e9edf4; border-radius: 10px;
        white-space: nowrap; /* evita wrap de texto */
    }}
    table.pnltbl {{ border-collapse: collapse; width: 100%; }}

    /* Cabeçalho sticky (1a linha) */
    table.pnltbl thead tr:nth-child(1) th {{
        position: sticky; top: 0; z-index: 3;
        background: {CB["blue"]}; color: #fff; font-weight: 700; padding: 8px;
        border-bottom: 1px solid #d0d7de;
    }}
    /* Cabeçalho sticky (2a linha - usado na visão diretoria) */
    table.pnltbl thead tr:nth-child(2) th {{
        position: sticky; top: 38px; z-index: 3;
        background: {CB["blue"]}; color: #fff; font-weight: 700; padding: 8px;
        border-bottom: 1px solid #d0d7de;
    }}
    /* Fallback quando só há 1 linha no thead */
    table.pnltbl th {{
        background: {CB["blue"]}; color: #fff; font-weight: 700; padding: 8px;
        border-bottom: 1px solid #d0d7de;
    }}
    table.pnltbl td {{ padding: 6px 8px; border-bottom: 1px solid #eee; }}
    table.pnltbl tr.parent {{ background: #f7f7f7; font-weight: 700; }}

    .delta-up   {{ color:{CB["green"]}; font-weight:700; }}
    .delta-down {{ color:{CB["red"]};   font-weight:700; }}
    .delta-zero {{ color:{CB["gray"]};  font-weight:700; }}

    .hl-card {{
        border: 1px solid #e9edf4; border-radius: 12px; padding: 12px 14px;
        background: #fff; margin: 8px 0; box-shadow: 0 1px 6px rgba(0,0,0,.04);
    }}
    .hl-sub {{ color: #5b667a; font-size: 0.96rem; }}
    .hl-bad {{ color: #E1002A; font-weight: 800; }}
    </style>
    """, unsafe_allow_html=True)

def section(title: str):
    inject_css()
    st.markdown(f'<div class="brand-section">{title}</div>', unsafe_allow_html=True)

# ==================== SETUP ====================
st.set_page_config(page_title="P&L – Projeção e Comparativos", layout="wide")
section("📊 P&L – Projeção do mês e comparativos")

# ==================== HELPERS ====================
def _to_upper(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    alias = {
        "NV MERGE":"NV_MERGE",
        "KPI COMPACT":"KPI_COMPACT",
        "KPI_COMPACTO":"KPI_COMPACT",
        "KPI COMPACTO":"KPI_COMPACT",
        "%":"PCT",
        "ORDEM SETOR":"ORDEM SETOR",
    }
    for k,v in alias.items():
        if k in df.columns and v not in df.columns:
            df.rename(columns={k:v}, inplace=True)
    return df

def _norm_metric(s: pd.Series) -> pd.Series:
    return (s.astype(str).str.strip().str.lower()
            .replace({"orçado":"forecast","orcado":"forecast","fcst":"forecast",
                      "real":"realizado","realizado ":"realizado",
                      "projeção":"projecao","projeçao":"projecao","proj":"projecao"}))

def _norm_period(s: pd.Series) -> pd.Series:
    def norm(x):
        if pd.isna(x): return np.nan
        sx = str(x).strip().replace("/", "-")
        dt = pd.to_datetime(sx, errors="coerce", dayfirst=False)
        if pd.isna(dt): dt = pd.to_datetime(sx, errors="coerce", dayfirst=True)
        return np.nan if pd.isna(dt) else dt.strftime("%Y-%m")
    return s.apply(norm)

def _period_minus(p: str, m: int) -> str:
    return (pd.Period(p, freq="M") - m).strftime("%Y-%m")

def fmt_brl(v):
    if pd.isna(v): return ""
    v = float(v)
    s = f"{abs(v):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {'-' if v<0 else ''}{s}"

def fmt_pct_symbol(v, dec=2):
    if pd.isna(v): return ""
    return f"{float(v)*100:.{dec}f}".replace(".", ",") + "%"

def fmt_pp_value(v, dec=1):
    if pd.isna(v): return ""
    return f"{float(v)*100:.{dec}f}".replace(".", ",")

def decorate_delta_money(v):
    if pd.isna(v): return ""
    sign = 0 if v==0 else (1 if v>0 else -1)
    cls  = "delta-zero" if sign==0 else ("delta-up" if sign>0 else "delta-down")
    arrow= "▲" if sign>0 else ("▼" if sign<0 else "→")
    val  = fmt_brl(abs(v)) if sign!=0 else fmt_brl(v)
    return f"<span class='{cls}'>{arrow} {val}</span>"

def decorate_delta_pp_plain(v, dec=1):
    if pd.isna(v): return ""
    sign = 0 if v==0 else (1 if v>0 else -1)
    cls  = "delta-zero" if sign==0 else ("delta-up" if sign>0 else "delta-down")
    arrow= "▲" if sign>0 else ("▼" if sign<0 else "→")
    val  = f"{abs(float(v))*100:.{dec}f}".replace(".", ",")
    return f"<span class='{cls}'>{arrow} {val}</span>"

def remove_accents(s: str) -> str:
    return ''.join(ch for ch in unicodedata.normalize('NFKD', s) if not unicodedata.combining(ch))

def normalize_key(x: str) -> str:
    x = str(x)
    x = re.sub(r"\s+", " ", x).strip().upper()
    x = remove_accents(x)
    return x

def kpi_list_ordered(df_src: pd.DataFrame) -> list[str]:
    """Retorna lista de KPI_COMPACT ordenada pelo menor valor de ORDEM encontrado na base."""
    if "KPI_COMPACT" not in df_src.columns:
        return []
    tmp = df_src[["KPI_COMPACT","ORDEM"]].dropna(subset=["KPI_COMPACT"]).copy()
    tmp["ORDEM"] = pd.to_numeric(tmp["ORDEM"], errors="coerce")
    rank = (tmp.groupby("KPI_COMPACT", dropna=True)["ORDEM"]
              .min()
              .reset_index()
              .sort_values(["ORDEM","KPI_COMPACT"], na_position="last"))
    return rank["KPI_COMPACT"].tolist()

# ==================== CARGA / NORMALIZAÇÃO ====================
@st.cache_data(show_spinner=False)
def load_normalize(file_bytes: bytes, filename: str) -> pd.DataFrame:
    base = pd.read_csv(io.BytesIO(file_bytes)) if filename.lower().endswith(".csv") else pd.read_excel(io.BytesIO(file_bytes))
    base = _to_upper(base)

    # reforço para KPI_COMPACT
    if "KPI_COMPACT" not in base.columns:
        for alt in ["KPI_COMPACTO", "KPI COMPACTO", "KPI COMPACT", "KPI_COMPACTO "]:
            if alt in base.columns:
                base.rename(columns={alt:"KPI_COMPACT"}, inplace=True)
                break

    # numéricos
    for c in ["$", "PCT"]:
        if c in base.columns:
            base[c] = pd.to_numeric(base[c], errors="coerce")

    # normalizações
    if "METRICA" in base.columns: base["METRICA"] = _norm_metric(base["METRICA"])
    if "PERIODO" in base.columns: base["PERIODO"] = _norm_period(base["PERIODO"])

    # PCT pode vir “cheio”
    if "PCT" in base.columns:
        base["PCT"] = np.where(base["PCT"].abs() > 1, base["PCT"]/100.0, base["PCT"])

    # defaults
    for col in ["KPI","KPI_COMPACT","AGREG","TIPO","PRINCIPAL","BU","CATEGORIA","DIRETORIA","SINAL","FAMILIA"]:
        if col not in base.columns: base[col] = ""
        base[col] = base[col].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

    if "ORDEM" not in base.columns:
        base["ORDEM"] = 9999
    else:
        base["ORDEM"] = pd.to_numeric(base["ORDEM"], errors="coerce").fillna(9999)

    if "ORDEM SETOR" in base.columns:
        base["ORDEM SETOR"] = pd.to_numeric(base["ORDEM SETOR"], errors="coerce").fillna(9999)

    base["PRINCIPAL"] = base.get("PRINCIPAL","NAO").str.upper().replace({"NÃO":"NAO","NO":"NAO","TRUE":"SIM","FALSE":"NAO"})
    base["SINAL"] = pd.to_numeric(base.get("SINAL","1"), errors="coerce").fillna(1).astype(int)

    # chaves normalizadas para diretoria
    base["DIRETORIA_KEY"] = base["DIRETORIA"].apply(normalize_key)
    base["DIRETORIA_KEY"] = base["DIRETORIA_KEY"].replace({
        "CAUDA LONGA":"CAUDA",
        "INFO E PERIFÉRICOS":"INFO E PERIFERICOS"
    })

    return base

# --- substitua tudo que vai desde "uploaded = st.file_uploader(...)" até o primeiro uso de 'base' ---

import os

# caminho padrão da base no repositório (mesma pasta do app.py)
DEFAULT_DATA_PATH = os.path.join(os.path.dirname(__file__), "BASE_PNL.xlsx")

# opção no sidebar (se quiser permitir override por upload)
with st.sidebar:
    st.markdown("### Fonte de dados")
    use_repo_file = st.checkbox("Usar BASE_PNL.xlsx do repositório", value=True)
    uploaded = None
    if not use_repo_file:
        uploaded = st.file_uploader("Carregue uma base (XLSX/CSV)", type=["xlsx","xls","csv"])

# leitura dos bytes + nome do arquivo
if use_repo_file:
    if not os.path.exists(DEFAULT_DATA_PATH):
        st.error("Arquivo **BASE_PNL.xlsx** não encontrado no repositório. Coloque-o na mesma pasta do `app.py`.")
        st.stop()
    with open(DEFAULT_DATA_PATH, "rb") as f:
        file_bytes = f.read()
    filename = os.path.basename(DEFAULT_DATA_PATH)
else:
    if uploaded is None:
        st.info("Carregue um arquivo para continuar ou marque 'Usar BASE_PNL.xlsx do repositório'.")
        st.stop()
    file_bytes = uploaded.getvalue()
    filename = uploaded.name

# carrega + normaliza
base = load_normalize(file_bytes, filename)

# =============== MENU LATERAL (Abas) ===============
with st.sidebar:
    st.markdown("## Navegação")
    tab_choice = st.radio("Selecione a aba", ["Visão Geral", "Visão Diretoria", "Gráficos"], index=0, label_visibility="collapsed")

# =============== FILTROS ====================
c1, c2, c3, c4 = st.columns([1.3,1.2,1.3,1.0])

with c1:
    st.markdown("**Diretoria**")
    dir_keys_options = sorted(base["DIRETORIA_KEY"].dropna().unique().tolist())
    preferred_cons_keys = [k for k in ["", "CONSOLIDADO", "TOTAL"] if k in dir_keys_options]
    default_dir_key = preferred_cons_keys[0] if preferred_cons_keys else (dir_keys_options[0] if dir_keys_options else "")
    def _fmt_dir(k): return (k or "Consolidado").title()
    diretoria_sel_keys = st.multiselect("", dir_keys_options, default=[default_dir_key] if default_dir_key in dir_keys_options else [], format_func=_fmt_dir, label_visibility="collapsed")

with c2:
    st.markdown("**BU**")
    bu_vals = sorted([x for x in base["BU"].dropna().unique() if x])
    bu_sel = st.multiselect("", bu_vals, default=[], label_visibility="collapsed")

with c3:
    st.markdown("**Setor**")
    if "ORDEM SETOR" in base.columns:
        setores = base[["CATEGORIA","ORDEM SETOR"]].drop_duplicates().sort_values("ORDEM SETOR")
        setor_list = setores["CATEGORIA"].tolist()
    else:
        setor_list = sorted([x for x in base["CATEGORIA"].dropna().unique() if x])
    setor_sel = st.multiselect("", setor_list, default=[], label_visibility="collapsed")

with c4:
    st.markdown("**Mês vigente (P0)**")
    periods = sorted(base["PERIODO"].dropna().unique().tolist())
    p0 = st.selectbox("", periods, index=len(periods)-1 if periods else 0, label_visibility="collapsed")

st.write("")
t1, t2, t3, t4 = st.columns(4)
with t1:
    show_totais = st.checkbox("Mostrar apenas totais", value=False)
with t2:
    show_money = st.checkbox("Exibir colunas $", value=True)
with t3:
    show_percent = st.checkbox("Exibir colunas %RL", value=True)
with t4:
    show_principais = st.checkbox("Exibir apenas KPIs principais", value=False)

# Filtro "somente margens"
only_margins = st.checkbox("Exibir apenas linhas de Margem (MC#1 a MC#4)", value=False)

# --- aplica filtros base comuns ---
def apply_common_filters(df0: pd.DataFrame):
    d = df0.copy()
    if bu_sel:
        d = d[d["BU"].isin(bu_sel)]
    if setor_sel:
        d = d[d["CATEGORIA"].isin(setor_sel)]
    if show_principais:
        d = d[d["PRINCIPAL"]=="SIM"]
    if show_totais:
        d = d[d["AGREG"].str.lower()=="pai"]
    return d

# 1) DF principal (respeita filtro de diretoria selecionado)
df = apply_common_filters(base)
if diretoria_sel_keys:
    df = df[df["DIRETORIA_KEY"].isin(diretoria_sel_keys)]

# 2) DF para visão diretoria e gráficos comparativos (NÃO filtra diretoria)
df_all_dirs = apply_common_filters(base)

if df.empty:
    st.info("Sem dados para os filtros selecionados."); st.stop()

# --- P0 efetivo baseado no DF filtrado ---
periods_avail = sorted(df["PERIODO"].dropna().unique().tolist())
p0_eff = p0 if p0 in periods_avail else (periods_avail[-1] if periods_avail else p0)
if p0_eff != p0:
    st.warning(f"Sem dados para {p0} na seleção atual. Usando {p0_eff}.")
p0 = p0_eff

# períodos relativos
p_m1  = _period_minus(p0,1)
p_m2  = _period_minus(p0,2)
p_m3  = _period_minus(p0,3)
p_m12 = _period_minus(p0,12)
st.markdown(f"**Períodos:** P0=`{p0}` | M-1=`{p_m1}` | M-2=`{p_m2}` | M-3=`{p_m3}` | M-12=`{p_m12}`")

# ==================== PIVOT (EXATO) ====================
@st.cache_data(show_spinner=False)
def pivotize(df_in: pd.DataFrame, p0, p_m1, p_m2, p_m3, p_m12):
    index_cols = ["AGREG","KPI_COMPACT","KPI","SINAL","FAMILIA","ORDEM","CATEGORIA","TIPO","DIRETORIA","DIRETORIA_KEY"]

    df_key_sorted = df_in.sort_values(["AGREG","KPI_COMPACT","KPI","PERIODO","METRICA"])
    df_dedup = df_key_sorted.drop_duplicates(
        subset=index_cols + ["PERIODO","METRICA"],
        keep="first"
    )

    pv_money = pd.pivot_table(df_dedup, index=index_cols, columns=["PERIODO","METRICA"], values="$",  aggfunc="first")
    pv_pct   = pd.pivot_table(df_dedup, index=index_cols, columns=["PERIODO","METRICA"], values="PCT", aggfunc="first")

    def col_get(pv, period, metric):
        try:    return pv[(period, metric)]
        except: return pd.Series(index=pv.index, dtype=float)

    m = pd.DataFrame(index=pv_money.index).assign(
        real_m3 = col_get(pv_money, p_m3, "realizado"),
        real_m2 = col_get(pv_money, p_m2, "realizado"),
        real_m1 = col_get(pv_money, p_m1, "realizado"),
        real_m12= col_get(pv_money, p_m12,"realizado"),
        proj    = col_get(pv_money, p0,   "projecao"),
        fcst    = col_get(pv_money, p0,   "forecast"),
        p_proj  = col_get(pv_pct,   p0,   "projecao"),
        p_m3v   = col_get(pv_pct,   p_m3, "realizado"),
        p_m2v   = col_get(pv_pct,   p_m2, "realizado"),
        p_m1v   = col_get(pv_pct,   p_m1, "realizado"),
        p_m12v  = col_get(pv_pct,   p_m12,"realizado"),
        p_fcst  = col_get(pv_pct,   p0,   "forecast"),
    ).reset_index()

    m["d_m1"]   = m["proj"] - m["real_m1"]
    m["d_m12"]  = m["proj"] - m["real_m12"]
    m["d_fc"]   = m["proj"] - m["fcst"]
    m["pd_m1"]  = m["p_proj"] - m["p_m1v"]
    m["pd_m12"] = m["p_proj"] - m["p_m12v"]
    m["pd_fc"]  = m["p_proj"] - m["p_fcst"]

    return m

m  = pivotize(df,          p0, p_m1, p_m2, p_m3, p_m12)  # respeita diretoria
mD = pivotize(df_all_dirs, p0, p_m1, p_m2, p_m3, p_m12)  # todas diretorias

def dedup_kpi(df_in: pd.DataFrame) -> pd.DataFrame:
    pais   = df_in[df_in["AGREG"].str.lower()=="pai"].sort_values("ORDEM")
    filhos = df_in[df_in["AGREG"].str.lower()=="filho"].sort_values("ORDEM")
    pais   = pais.drop_duplicates(subset=["KPI_COMPACT"], keep="first")
    filhos = filhos.drop_duplicates(subset=["KPI"], keep="first")
    out = pd.concat([pais, filhos], ignore_index=True).sort_values("ORDEM")
    return out

# filtro de margens (se marcado)
if only_margins:
    margin_names = [
        "Margem Contribuição #1 (CashMerc + Bonif. + Demais Rec)",
        "Margem Contribuição #2",
        "Margem Contribuição #3",
        "Margem Contribuição #4",
    ]
    m  = m[m["KPI_COMPACT"].isin(margin_names)]
    mD = mD[mD["KPI_COMPACT"].isin(margin_names)]

# ======= Fallback PP (M-1) quando p_m1v não veio do pivô =======
def _fallback_pct_real_m1(df_raw: pd.DataFrame, r_agreg, r_kpi_compact, r_kpi, p_m1):
    d = df_raw[
        (df_raw["AGREG"]==r_agreg) &
        (df_raw["KPI_COMPACT"]==r_kpi_compact) &
        (df_raw["KPI"]==r_kpi) &
        (df_raw["PERIODO"]==p_m1) &
        (df_raw["METRICA"].str.lower()=="realizado")
    ]
    colpct = "PCT" if "PCT" in d.columns else ("%")
    if colpct in d.columns and not d.empty:
        val = pd.to_numeric(d[colpct], errors="coerce")
        if val.notna().any():
            v = float(val.iloc[0])
            return v/100.0 if abs(v)>1 else v
    return np.nan

# ==================== FUNÇÃO ROBUSTA – CONTRIBUIÇÃO POR SETOR ====================
def sector_contribution_delta_m1(df_raw: pd.DataFrame, kpi_compact: str, p0: str, p_m1: str) -> pd.Series:
    """
    Calcula delta por setor (CATEGORIA) em Consolidado (se existir) ou agregado geral,
    considerando SINAL. Robustez:
      - normaliza chaves e categorias
      - ignora TOTAL/blank
      - 3 tentativas (filhos, depois pais, depois todos)
      - fallback para maiores absolutos quando não houver negativos
    Retorna uma Series ordenada (negativos primeiro quando existirem).
    """
    if df_raw.empty:
        return pd.Series(dtype=float)

    d0 = df_raw.copy()

    # normalizações leves
    d0["DIR_KEY_N"] = d0["DIRETORIA_KEY"].astype(str).str.strip().str.upper().fillna("")
    d0["AGREG_N"]   = d0["AGREG"].astype(str).str.strip().str.lower()
    d0["CAT_N"]     = d0["CATEGORIA"].astype(str).str.strip()
    d0["CAT_UP"]    = d0["CAT_N"].str.upper()

    # Consolidado se existir; senão usa tudo
    cons_keys = {"", "CONSOLIDADO", "TOTAL"}
    have_cons = d0["DIR_KEY_N"].isin(cons_keys).any()
    d = d0[d0["DIR_KEY_N"].isin(cons_keys)] if have_cons else d0

    # somente o KPI alvo
    d = d[d["KPI_COMPACT"] == kpi_compact]
    if d.empty:
        return pd.Series(dtype=float)

    def _sum_rs(g):
        vals = pd.to_numeric(g["$"], errors="coerce").fillna(0)
        sinal = pd.to_numeric(g["SINAL"], errors="coerce").fillna(1)
        return (vals * sinal).sum(min_count=1)

    def _make_series(dsub: pd.DataFrame) -> pd.Series:
        proj = dsub[(dsub["PERIODO"] == p0) & (dsub["METRICA"] == "projecao")].groupby("CAT_UP").apply(_sum_rs)
        m1   = dsub[(dsub["PERIODO"] == p_m1) & (dsub["METRICA"] == "realizado")].groupby("CAT_UP").apply(_sum_rs)
        s = (proj - m1)
        s = s.dropna()
        s = s[~s.index.isin(["", "TOTAL"])]
        s = s[abs(s) > 0]  # remove zeros exatos
        return s

    # Tentativa 1: só filhos
    s = _make_series(d[d["AGREG_N"] == "filho"])
    # Tentativa 2: só pais
    if s.empty:
        s = _make_series(d[d["AGREG_N"] == "pai"])
    # Tentativa 3: todos
    if s.empty:
        s = _make_series(d)

    if s.empty:
        return s  # vazio; o chamador trata

    # Ordenação: negativos primeiro (puxando para baixo). Se não houver, usa maiores absolutos.
    if (s < 0).any():
        return s.sort_values(ascending=True)
    return s.reindex(s.abs().sort_values(ascending=False).index)

# ==================== RENDERIZAÇÃO TABELAS ====================
def render_table_general(m_df: pd.DataFrame, df_raw: pd.DataFrame, table_id="pnltbl_general") -> str:
    m_consol = m_df.copy()
    tbl = dedup_kpi(m_consol)
    tbl["_PAI"] = (tbl["AGREG"].str.lower()=="pai").astype(int)
    tbl["DRE"] = np.where(tbl["_PAI"]==1, "**"+tbl["KPI_COMPACT"]+"**", tbl["KPI"])

    # lookup TIPO
    tipo_lookup = (
        df_raw[["AGREG","KPI_COMPACT","KPI","TIPO"]]
        .drop_duplicates()
        .assign(AGREG=lambda d: d["AGREG"].astype(str).str.lower())
        .rename(columns={"TIPO":"TIPO_SRC"})
    )
    dfv = tbl.copy()
    dfv["AGREG"] = dfv["AGREG"].astype(str).str.lower()
    dfv = dfv.merge(tipo_lookup, on=["AGREG","KPI_COMPACT","KPI"], how="left")
    dfv["TIPO"] = dfv["TIPO_SRC"].fillna("VALOR")

    cols = ["DRE"]
    if show_money:
        cols += ["real_m3","real_m2","real_m1","proj","d_m1","d_m12","d_fc"]
    if show_percent:
        cols += ["p_proj","pd_m1","pd_m12","pd_fc"]

    rename = {
        "real_m3":"Real M-3", "real_m2":"Real M-2", "real_m1":"Real M-1", "proj":"Projeção",
        "d_m1":"Δ vs M-1", "d_m12":"Δ vs M-12", "d_fc":"Δ vs Forecast",
        "p_proj":"Projeção %RL", "pd_m1":"Δ vs M-1 %RL", "pd_m12":"Δ vs M-12 %RL", "pd_fc":"Δ vs Forecast %RL"
    }

    df_show = dfv[["_PAI","DRE","TIPO","AGREG","KPI_COMPACT","KPI"] + [c for c in cols if c!="DRE"]].copy()

    if show_money:
        if "real_m3" in df_show.columns:
            df_show["real_m3"] = ["" if str(t).upper()=="PP" else fmt_brl(v) for v,t in zip(df_show["real_m3"], df_show["TIPO"])]
        if "real_m2" in df_show.columns:
            df_show["real_m2"] = ["" if str(t).upper()=="PP" else fmt_brl(v) for v,t in zip(df_show["real_m2"], df_show["TIPO"])]

        p_m1v = m_consol.set_index(["AGREG","KPI_COMPACT","KPI"]).get("p_m1v", pd.Series(dtype=float))
        p_proj= m_consol.set_index(["AGREG","KPI_COMPACT","KPI"]).get("p_proj", pd.Series(dtype=float))
        idx_list = list(zip(dfv["AGREG"], dfv["KPI_COMPACT"], dfv["KPI"]))
        pm1_vals = [p_m1v.get(i, np.nan) for i in idx_list]
        ppr_vals = [p_proj.get(i, np.nan) for i in idx_list]

        if "real_m1" in df_show.columns:
            def _fmt_real_m1_pp(v, p, t, a, kc, k):
                if str(t).upper()!="PP":
                    return fmt_brl(v)
                if pd.isna(p):
                    p = _fallback_pct_real_m1(df_raw, a, kc, k, p_m1)
                return "" if pd.isna(p) else fmt_pp_value(p)
            df_show["real_m1"] = [
                _fmt_real_m1_pp(v, p, t, a, kc, k)
                for v,p,t,a,kc,k in zip(
                    df_show["real_m1"], pm1_vals, df_show["TIPO"], dfv["AGREG"], dfv["KPI_COMPACT"], dfv["KPI"]
                )
            ]
        if "proj" in df_show.columns:
            df_show["proj"] = [fmt_pp_value(p) if str(t).upper()=="PP" else fmt_brl(v)
                               for v,p,t in zip(df_show["proj"], ppr_vals, df_show["TIPO"])]

        if "d_m1" in df_show.columns:
            df_show["d_m1"] = [decorate_delta_pp_plain(v) if str(t).upper()=="PP" else decorate_delta_money(v)
                               for v,t in zip(df_show["d_m1"], df_show["TIPO"])]
        if "d_m12" in df_show.columns:
            df_show["d_m12"] = [decorate_delta_pp_plain(v) if str(t).upper()=="PP" else decorate_delta_money(v)
                                for v,t in zip(df_show["d_m12"], df_show["TIPO"])]
        if "d_fc" in df_show.columns:
            df_show["d_fc"] = [decorate_delta_pp_plain(v) if str(t).upper()=="PP" else decorate_delta_money(v)
                               for v,t in zip(df_show["d_fc"], df_show["TIPO"])]

    if show_percent:
        for c in ["p_proj","pd_m1","pd_m12","pd_fc"]:
            if c in df_show.columns:
                df_show[c] = df_show[c].apply(lambda x: "" if pd.isna(x) else fmt_pct_symbol(x))

    ordered = ["DRE"]
    if show_money:   ordered += ["real_m3","real_m2","real_m1","proj","d_m1","d_m12","d_fc"]
    if show_percent: ordered += ["p_proj","pd_m1","pd_m12","pd_fc"]

    headers = "".join(f"<th>{rename.get(h,h)}</th>" for h in ordered)
    rows_html = []
    for _, r in df_show.iterrows():
        klass = "parent" if int(r["_PAI"])==1 else ""
        tds = "".join(f"<td>{r.get(h,'')}</td>" for h in ordered)
        rows_html.append(f"<tr class='{klass}'>{tds}</tr>")

    html = f"<div class='table-wrap'><table id='{table_id}' class='pnltbl'><thead><tr>{headers}</tr></thead><tbody>{''.join(rows_html)}</tbody></table></div>"
    return html

def render_table_diretoria(m_df: pd.DataFrame, table_id="pnltbl_dir") -> str:
    all_keys = m_df["DIRETORIA_KEY"].fillna("").astype(str).unique().tolist()
    priors = ["", "CONSOLIDADO", "TOTAL", "CAUDA", "INFO", "INFO E PERIFERICOS"]
    dir_order = []
    for p in priors:
        if p in all_keys and p not in dir_order:
            dir_order.append(p)
    for d in sorted(all_keys):
        if d not in dir_order:
            dir_order.append(d)

    key_to_label = (
        m_df[["DIRETORIA_KEY","DIRETORIA"]]
        .drop_duplicates()
        .set_index("DIRETORIA_KEY")["DIRETORIA"]
        .to_dict()
    )
    def disp_label(k):
        if not k: return "Consolidado"
        return (key_to_label.get(k) or k).title()

    base_rows = dedup_kpi(m_df)
    base_rows["_PAI"] = (base_rows["AGREG"].str.lower()=="pai").astype(int)
    base_rows["DRE"] = np.where(base_rows["_PAI"]==1, "**"+base_rows["KPI_COMPACT"]+"**", base_rows["KPI"])

    headers = ["<th>KPI</th>"] + [f"<th colspan='2'>{disp_label(k)}</th>" for k in dir_order]
    subhdr  = ["<th></th>"] + sum([[f"<th>Projeção</th>", f"<th>Δ vs M-1</th>"] for _ in dir_order], [])

    rows_html = []
    for _, r in base_rows.iterrows():
        row = [f"<td>{r['DRE']}</td>"]
        for k in dir_order:
            sub = m_df[
                (m_df["AGREG"]==r["AGREG"]) &
                (m_df["KPI_COMPACT"]==r["KPI_COMPACT"]) &
                (m_df["KPI"]==r["KPI"]) &
                (m_df["DIRETORIA_KEY"]==k)
            ]
            if sub.empty:
                row += ["<td></td>", "<td></td>"]
            else:
                t = str(sub["TIPO"].iloc[0]).upper()
                v_proj = sub["proj"].iloc[0]
                v_dm1  = sub["d_m1"].iloc[0]
                proj_txt = fmt_pp_value(sub["p_proj"].iloc[0]) if t=="PP" else fmt_brl(v_proj)
                dm1_txt  = decorate_delta_pp_plain(v_dm1) if t=="PP" else decorate_delta_money(v_dm1)
                row += [f"<td>{proj_txt}</td>", f"<td>{dm1_txt}</td>"]
        rows_html.append(f"<tr>{''.join(row)}</tr>")

    html = f"""
    <div class='table-wrap'>
      <table id="{table_id}" class='pnltbl'>
        <thead>
          <tr>{''.join(headers)}</tr>
          <tr>{''.join(subhdr)}</tr>
        </thead>
        <tbody>
          {''.join(rows_html)}
        </tbody>
      </table>
    </div>
    """
    return html

# ==================== GRÁFICOS ====================
def draw_kpi_evolution(m_df: pd.DataFrame, keys_source_df: pd.DataFrame, kpi_name: str, diretoria_sel_keys: list[str]):
    """Plota evolução do KPI (VALOR ou %RL) respeitando as diretorias de interesse."""
    desired_order = ["M-12","M-3","M-2","M-1","Projeção"]

    # Decide diretorias a mostrar
    keys_want = list(diretoria_sel_keys or [])
    if not keys_want:
        for k in ["", "CONSOLIDADO", "TOTAL", "CAUDA", "INFO", "INFO E PERIFERICOS"]:
            if k in set(keys_source_df["DIRETORIA_KEY"]):
                keys_want.append(k)
        if not keys_want:
            keys_want = sorted(keys_source_df["DIRETORIA_KEY"].dropna().unique().tolist())[:3]

    sub_all = m_df[(m_df["KPI_COMPACT"]==kpi_name)]
    if sub_all.empty:
        st.info(f"KPI **{kpi_name}** sem dados para os filtros."); return

    t = str(sub_all["TIPO"].dropna().iloc[0]).upper() if sub_all["TIPO"].notna().any() else "VALOR"
    y_label = "% da Receita Líquida" if t=="PP" else "R$"

    rows = []
    for k in keys_want:
        sdir = sub_all[sub_all["DIRETORIA_KEY"]==k]
        if sdir.empty: 
            continue
        sdir = dedup_kpi(sdir)
        r = sdir.iloc[0]
        series_vals = [
            ("M-12", r.get("p_m12v" if t=="PP" else "real_m12", np.nan)),
            ("M-3",  r.get("p_m3v"  if t=="PP" else "real_m3",  np.nan)),
            ("M-2",  r.get("p_m2v"  if t=="PP" else "real_m2",  np.nan)),
            ("M-1",  r.get("p_m1v"  if t=="PP" else "real_m1",  np.nan)),
            ("Projeção", r.get("p_proj" if t=="PP" else "proj", np.nan)),
        ]
        for lab, val in series_vals:
            if pd.notna(val):
                y = float(val)*100.0 if t=="PP" else float(val)
                label = f"{y:.1f}%" if t=="PP" else f"{fmt_brl(y)}".replace("R$ ","")
                rows.append({"Diretoria": (k or "Consolidado").title(), "Período": lab, "Valor": y, "YLabel": y_label, "Label": label})
    if not rows:
        st.info(f"KPI **{kpi_name}** sem pontos válidos."); return

    chart_df = pd.DataFrame(rows)
    chart_df["Período"] = pd.Categorical(chart_df["Período"], categories=desired_order, ordered=True)

    base = alt.Chart(chart_df).mark_line(
        point=alt.OverlayMarkDef(size=110),
        interpolate='monotone',
        strokeWidth=3
    ).encode(
        x=alt.X('Período', sort=desired_order, title=''),
        y=alt.Y('Valor:Q', title=y_label),
        color=alt.Color('Diretoria:N',
                        legend=alt.Legend(
                            title='Diretoria',
                            labelFontWeight='bold', titleFontWeight='bold',
                            labelFontSize=13, titleFontSize=14,
                            symbolSize=300, symbolStrokeWidth=2
                        ))
    )
    labels = alt.Chart(chart_df).mark_text(
        align='left', dx=8, dy=-8,
        fontWeight='bold', fontSize=13
    ).encode(
        x='Período', y='Valor:Q', text=alt.Text('Label'), color='Diretoria:N'
    )
    st.markdown(f"**{kpi_name} – Evolução**")
    st.altair_chart(base + labels, use_container_width=True)

def draw_margin_block(m_df: pd.DataFrame, keys_source_df: pd.DataFrame, diretoria_sel_keys: list[str]):
    """Mantém o bloco padrão de Margens quando nenhum KPI foi selecionado."""
    desired_order = ["M-12","M-3","M-2","M-1","Projeção"]
    margin_names = [
        "Margem Contribuição #1 (CashMerc + Bonif. + Demais Rec)",
        "Margem Contribuição #2",
        "Margem Contribuição #3",
        "Margem Contribuição #4",
    ]
    keys_want = list(diretoria_sel_keys or [])
    if not keys_want:
        for k in ["", "CONSOLIDADO", "TOTAL", "CAUDA", "INFO", "INFO E PERIFERICOS"]:
            if k in set(keys_source_df["DIRETORIA_KEY"]): keys_want.append(k)
        if not keys_want:
            keys_want = sorted(keys_source_df["DIRETORIA_KEY"].dropna().unique().tolist())[:3]

    for name in margin_names:
        sub_all = m_df[m_df["KPI_COMPACT"]==name]
        if sub_all.empty:
            continue
        rows = []
        for k in keys_want:
            sdir = sub_all[sub_all["DIRETORIA_KEY"]==k]
            if sdir.empty: continue
            sdir = dedup_kpi(sdir); r = sdir.iloc[0]
            series_vals = [("M-12", r.get("p_m12v", np.nan)), ("M-3", r.get("p_m3v", np.nan)),
                           ("M-2", r.get("p_m2v", np.nan)), ("M-1", r.get("p_m1v", np.nan)),
                           ("Projeção", r.get("p_proj", np.nan))]
            for lab, val in series_vals:
                if pd.notna(val):
                    y = float(val)*100.0
                    rows.append({"Diretoria": (k or "Consolidado").title(), "Período": lab,
                                 "Valor": y, "YLabel": "% da Receita Líquida", "Label": f"{y:.1f}%"})
        if not rows: 
            continue
        chart_df = pd.DataFrame(rows)
        chart_df["Período"] = pd.Categorical(chart_df["Período"], categories=desired_order, ordered=True)
        base = alt.Chart(chart_df).mark_line(
            point=alt.OverlayMarkDef(size=110), interpolate='monotone', strokeWidth=3
        ).encode(
            x=alt.X('Período', sort=desired_order, title=''),
            y=alt.Y('Valor:Q', title="% da Receita Líquida"),
            color=alt.Color('Diretoria:N', legend=alt.Legend(
                title='Diretoria', labelFontWeight='bold', titleFontWeight='bold',
                labelFontSize=13, titleFontSize=14, symbolSize=300, symbolStrokeWidth=2))
        )
        labels = alt.Chart(chart_df).mark_text(
            align='left', dx=8, dy=-8, fontWeight='bold', fontSize=13
        ).encode(x='Período', y='Valor:Q', text=alt.Text('Label'), color='Diretoria:N')
        st.markdown(f"**{name} – %RL**")
        st.altair_chart(base + labels, use_container_width=True)

# ==================== TELA (ABAS) ====================
if tab_choice == "Visão Geral":
    # ---- Filtro de KPI (filtra a tabela para 1 KPI quando escolhido) ----
    kpi_opts = ["(todos)"] + kpi_list_ordered(df_all_dirs)
    kpi_filter = st.selectbox("Filtrar KPI (linha):", options=kpi_opts, index=0)
    if kpi_filter != "(todos)":
        # filtra m para o KPI escolhido
        m_show = m[(m["KPI_COMPACT"]==kpi_filter) | (m["KPI"]==kpi_filter)].copy()
        if m_show.empty:
            st.info("KPI sem dados para os filtros."); st.stop()
    else:
        m_show = m

    # tabela
    st.markdown(render_table_general(m_show, df), unsafe_allow_html=True)

    # Highlights — regra: mostrar "puxado por..." SOMENTE quando diretoria está filtrada em Consolidado e sem setor selecionado
    st.markdown("### 🔎 Highlights do mês")

    def _is_consolidado_selected_only():
        if len(diretoria_sel_keys) != 1:
            return False
        k = diretoria_sel_keys[0]
        return k in ["", "CONSOLIDADO", "TOTAL"]

    _show_sector_breakdown = _is_consolidado_selected_only() and (len(setor_sel) == 0)

    def _excluded_kpi(name: str) -> bool:
        up = (name or "").upper()
        if any(x in up for x in ["IMPOST", "TRIBUT", "TAXA", "ICMS", "PIS", "COFINS", "ISS"]):
            return True
        if "APOS MARGEM 4" in up or "APÓS MARGEM 4" in up:
            return True
        if re.search(r"MARGEM\s*[5-9]", up):
            return True
        return False

    work = m[(m["AGREG"].str.lower()=="pai")].copy()
    work["gap_m1"] = work["d_m1"]
    work = work[(work["gap_m1"].notna()) & (work["gap_m1"] < 0) & (work["gap_m1"].abs() >= 100_000)]
    work = work[~work["KPI_COMPACT"].apply(_excluded_kpi)]

    if work.empty:
        st.markdown("_Sem quedas ≥ R$ 100 mil nos KPIs elegíveis._")
    else:
        st.markdown("**Comparativo vs M-1 – Maiores quedas (gap ≥ R$ 100 mil)**")
        for _, r in work.sort_values("gap_m1").iterrows():
            kpi_name = r["KPI_COMPACT"]
            delta_rs = abs(float(r["gap_m1"]))
            setores_txt = ""
            if _show_sector_breakdown:
                contr = sector_contribution_delta_m1(df_all_dirs, kpi_name, p0, p_m1)
                if not contr.empty:
                    # preferir negativos; se não houver, maiores absolutos
                    if (contr < 0).any():
                        contr_use = contr[contr < 0].sort_values().head(2)
                    else:
                        contr_use = contr.abs().sort_values(ascending=False).head(2)
                    parts = [f"setor {str(n).title()} ({fmt_brl(abs(float(v)))})" for n, v in contr_use.items()]
                    if parts:
                        setores_txt = ", puxado pelo " + (" e ".join(parts) if len(parts) <= 2 else ", ".join(parts[:2]))
            st.markdown(
                f"<div class='hl-card'><div class='hl-sub'>"
                f"O <b>{kpi_name.upper()}</b> apresenta uma <span class='hl-bad'>queda</span> de "
                f"<b>{fmt_brl(delta_rs)}</b> em comparação ao M-1{setores_txt}."
                f"</div></div>", unsafe_allow_html=True
            )

elif tab_choice == "Visão Diretoria":
    # ---- Filtro de KPI (filtra a tabela para 1 KPI quando escolhido) ----
    kpi_opts = ["(todos)"] + kpi_list_ordered(df_all_dirs)
    kpi_filter = st.selectbox("Filtrar KPI (linha):", options=kpi_opts, index=0)
    if kpi_filter != "(todos)":
        mD_show = mD[(mD["KPI_COMPACT"]==kpi_filter) | (mD["KPI"]==kpi_filter)].copy()
        if mD_show.empty:
            st.info("KPI sem dados para os filtros."); st.stop()
    else:
        mD_show = mD

    html_dir = render_table_diretoria(mD_show, table_id="pnltbl_dir")
    st.markdown(html_dir, unsafe_allow_html=True)

else:  # Gráficos
    # KPIs ordenados por ORDEM
    kpi_options_ordered = kpi_list_ordered(df_all_dirs)
    kpi_sel = st.multiselect("KPI(s) (opcional):", options=kpi_options_ordered, default=[])

    if kpi_sel:
        for kpi_name in kpi_sel:
            draw_kpi_evolution(mD, df_all_dirs, kpi_name, diretoria_sel_keys)
    else:
        draw_margin_block(mD, df_all_dirs, diretoria_sel_keys)

# ==================== EXPORT XLSX ====================
def to_xlsx_bytes(df_export: pd.DataFrame) -> bytes:
    m_consol = m.copy()
    tbl = dedup_kpi(m_consol)
    tbl["_PAI"] = (tbl["AGREG"].str.lower()=="pai").astype(int)
    tbl["DRE"] = np.where(tbl["_PAI"]==1, "**"+tbl["KPI_COMPACT"]+"**", tbl["KPI"])

    out = io.BytesIO()
    wb = Workbook(); ws = wb.active; ws.title = "P&L"

    headers = ["DRE"]
    if show_money:   headers += ["Real M-3","Real M-2","Real M-1","Projeção","Δ vs M-1","Δ vs M-12","Δ vs Forecast"]
    if show_percent: headers += ["Projeção %RL","Δ vs M-1 %RL","Δ vs M-12 %RL","Δ vs Forecast %RL"]
    ws.append(headers)

    tipo_lookup = (
        df[["AGREG","KPI_COMPACT","KPI","TIPO"]]
        .drop_duplicates()
        .assign(AGREG=lambda d: d["AGREG"].astype(str).str.lower())
        .rename(columns={"TIPO":"TIPO_SRC"})
    )
    dfv = tbl.merge(tipo_lookup, on=["AGREG","KPI_COMPACT","KPI"], how="left")
    dfv["TIPO"] = dfv["TIPO_SRC"].fillna("VALOR")

    for _, r in dfv.iterrows():
        row_vals = [r["DRE"]]
        if show_money:
            row_vals += [r.get("real_m3",""), r.get("real_m2",""),
                         r.get("real_m1",""), r.get("proj",""),
                         r.get("d_m1",""), r.get("d_m12",""), r.get("d_fc","")]
        if show_percent:
            row_vals += [r.get("p_proj",""), r.get("pd_m1",""), r.get("pd_m12",""), r.get("pd_fc","")]
        ws.append(row_vals)
        lx = ws.max_row
        if int(r.get("_PAI",0))==1:
            for c in range(1, len(headers)+1):
                ws.cell(lx, c).fill = PatternFill("solid", fgColor="F7F7F7")
                ws.cell(lx, c).font = Font(bold=True)

    for i in range(1, len(headers)+1):
        ws.cell(1, i).fill = PatternFill("solid", fgColor="F1F3F5")
        ws.cell(1, i).font = Font(bold=True)
        ws.cell(1, i).alignment = Alignment(horizontal="left")

    widths = [44] + [16]*(len(headers)-1)
    for i,w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64+i)].width = w

    wb.save(out); out.seek(0)
    return out.read()

st.download_button(
    "⬇️ Baixar XLSX",
    data=to_xlsx_bytes(m),
    file_name=f"pnl_{p0}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
