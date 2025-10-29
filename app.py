# app.py
# -----------------------------------------------------------
# P&L – Projeção, Realizados, Comparativos e Highlights
# -----------------------------------------------------------

import io
import os
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
        padding-bottom: 1.25rem !important;
        overflow: visible !important;
    }}
    header[data-testid="stHeader"] {{ background: {CB["bg"]}; }}

    .brand-section {{
        width: 100%; display: block; box-sizing: border-box;
        background: linear-gradient(90deg, {CB["blue"]} 0%, {CB["red"]} 100%);
        color: #fff; padding: 18px 20px; border-radius: 14px;
        font-weight: 800; letter-spacing: .2px; margin: 8px 0 12px 0;
        text-align: center; line-height: 1.25; font-size: clamp(18px, 2.0vw, 26px);
        box-shadow: 0 2px 10px rgba(0,0,0,.06);
    }}

    /* Tabelas responsivas */
    .table-wrap {{
        max-height: 70vh; overflow-y: auto; overflow-x: auto;
        border: 1px solid #e9edf4; border-radius: 10px;
        white-space: nowrap;  /* evita quebra de linha no mobile */
    }}
    table.pnltbl {{ border-collapse: collapse; width: 100%; }}

    /* Sticky headers */
    table.pnltbl thead tr:nth-child(1) th {{
        position: sticky; top: 0; z-index: 3;
        background: {CB["blue"]}; color: #fff; font-weight: 700; padding: 8px;
        border-bottom: 1px solid #d0d7de;
    }}
    table.pnltbl thead tr:nth-child(2) th {{
        position: sticky; top: 38px; z-index: 3;
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

    /* Touch targets maiores no mobile */
    @media (max-width: 640px) {{
      .stMultiSelect, .stSelectbox, .stCheckbox, .stRadio, button[kind="secondary"] {{
        font-size: 16px !important;
      }}
    }}
    </style>
    """, unsafe_allow_html=True)

def section(title: str):
    inject_css()
    st.markdown(f'<div class="brand-section">{title}</div>', unsafe_allow_html=True)

# ==================== SETUP ====================
st.set_page_config(page_title="P&L – Projeção e Comparativos", layout="wide")
section("📊 P&L – Projeção do mês e comparativos")

# ==================== KPI MASTER (completa filtros) ====================
def remove_accents(s: str) -> str:
    return ''.join(ch for ch in unicodedata.normalize('NFKD', s) if not unicodedata.combining(ch))
def _norm_key(s: str) -> str:
    return remove_accents(str(s)).strip().upper()

KPI_MASTER_LIST = [
    # (lista fornecida por você – mantida para completar filtros quando faltar algo na base)
    "GMV TOTAL","GMV RECEITA DE MERCADORIA","GMV RECEITA DE FRETE","GMV RECEITA DE SERVIÇO",
    "GMV REPASSE SERVIÇOS","GMV COMISSÃO 3P","GMV RECEITA DE CREDIÁRIO","GMV RECEITA DE CARTÕES",
    "GMV RECEITA DE MONTAGEM","GMV RECEITA DE MERCADORIA","GMV RECEITA DE ADS","GMV WRITE OFF CREDIÁRIO",
    "GMV CANCELAMENTO","GMV CUT-OFF","GMV JUROS DE CARTÃO","GMV DEMAIS EMPRESAS","RB TOTAL","RB MERCADORIA",
    "RB FRETE","RB SERVIÇOS","RB COMISSÃO 3P","RB CREDIÁRIO","RB CARTÕES","RB MONTAGEM","RB ADS","RB JUROS DE CARTÃO",
    "IMPOSTOS TOTAL","IMPOSTOS MERCADORIA","IMPOSTOS FRETE","IMPOSTOS SERVIÇOS","IMPOSTOS COMISSÃO 3P",
    "IMPOSTOS CREDIÁRIO","IMPOSTOS CARTÕES","IMPOSTOS MONTAGEM","IMPOSTOS ADS","IMPOSTOS JUROS DE CARTÃO",
    "RL TOTAL (MERCADORIA + SERVIÇOS)","CUSTO MERCADORIA TOTAL","CUSTO MERCADORIA GCB","CUSTO MERCADORIA GCB CUT-OFF",
    "CUSTO MERCADORIA GCB RESSARCIMENTO TRANSPORTADORA","CUSTO MERCADORIA GCB MARGEM NEGATIVA E OBSOLETO",
    "CUSTO MERCADORIA GCB FECHAMENTO DE ESTOQUE","CUSTO MERCADORIA GCB CRÉDITO STF","CUSTO MERCADORIA GCB CRÉDITO PRESUMIDO",
    "CUSTO MERCADORIA GCB FUNDO EQUILIBRIO FISCAL RJ","CUSTO MERCADORIA GCB OUTROS","CUSTO MERCADORIA BARTIRA MATÉRIA PRIMA NACIONAL",
    "CUSTO MERCADORIA BARTIRA MATÉRIA PRIMA EXPORTAÇÃO","CUSTO MERCADORIA BARTIRA PERDAS DE PRODUÇÃO",
    "CUSTO MERCADORIA BARTIRA CRÉDITO OUTORGADO","CUSTO MERCADORIA BARTIRA GCB NO ESTOQUE","CUSTO MERCADORIA CNT",
    "CASH 1P MERCADORIA","% MARGEM 1P","BONIFICAÇÃO TOTAL","BONIFICAÇÃO","BONIFICAÇÃO CUSTO DE CESSÃO E PAGAMENTO",
    "BONIFICAÇÃO GIRO DE ESTOQUE","BONIFICAÇÃO DESPESA FINANCEIRA","BONIFICAÇÃO DESPESAS MARKETING","BONIFICAÇÃO PROVISÕES",
    "BONIFICAÇÃO % MERCADORIA","CASH 1P MERCADORIA E BONIFICAÇÃO","% MARGEM 1P MERCADORIA E BONIFICAÇÃO",
    "CUSTO ADS TOTAL","CASH MARGIN ADS","% CASH ADS","DEMAIS RL TOTAL","MARGEM #1 CASHMERC + BONIFICAÇÃO + DEMAIS REC",
    "DESPESAS VARIÁVEIS TOTAL","DESPESAS VARIÁVEIS MARKETING","DESPESAS VARIAVEIS MARKETING PERFORMANCE",
    "DESPESAS VARIAVEIS MARKETING MIDIAS TOP MID FUNNEL","DESPESAS VARIAVEIS MARKETING APP","DESPESAS VARIAVEIS MARKETING AQUISIÇÃO",
    "DESPESAS VARIAVEIS MARKETING PRODUÇÃO","DESPESAS VARIAVEIS MARKETING REGIONAL","DESPESAS VARIAVEIS MARKETING SOCIAL BRANDING",
    "DESPESAS VARIAVEIS MARKETING CRM","DESPESAS VARIAVEIS MARKETING FERRAMENTAS","DESPESAS VARIÁVEIS ADS",
    "DESPESAS VARIÁVEIS FRETE","DESPESAS VARIÁVEIS ESCOLTA","DESPESAS VARIÁVEIS ENTREGA AO CLIENTE",
    "DESPESAS VARIÁVEIS COMISSÃO VENDEDORES","DESPESAS VARIÁVEIS DESCONTOS E COMISSÕES","DESPESAS VARIAVEIS AGENDA NEGATIVA",
    "DESPESAS VARIAVEIS DESCONTO DE INCENTIVO","DESPESAS VARIAVEIS DESPESAS COMERCIAIS","DESPESAS VARIAVEIS COMISSÃO E FIDELIDADE",
    "DESPESAS VARIAVEIS PERDA COM PARCEIROS","DESPESAS VARIÁVEIS CONSERTO MERCADORIAS","DESPESAS VARIÁVEIS ABASTECIMENTO LOJAS",
    "DESPESAS VARIÁVEIS DEMAIS CUSTOS LOGÍSTICOS","DESPESAS VARIAVEIS TRANSFERENCIA ENTRE CDS","DESPESAS VARIÁVEIS MONTADORES TERCEIRIZADOS",
    "DESPESAS VARIAVEIS TERCEIRIZAÇÃO DE CARGA E DESCARGA","DESPESAS VARIAVEIS TERCEIRIZAÇÃO DO APOIO LOGÍSTICO",
    "DESPESAS VARIAVEIS ALUGUEL DE EQUIPAMENTOS","DESPESAS VARIAVEIS MANUTENÇÃO DE EQUIPAMENTOS","DESPESAS VARIÁVEIS CUSTO DE CREDIÁRIO",
    "DESPESAS VARIAVEIS CUSTO DE CREDIÁRIO FUNDING","DESPESAS VARIAVEIS CUSTO DE CREDIÁRIO BOLETO CDCI",
    "DESPESAS VARIAVEIS CUSTO DE CREDIÁRIO COMISSÃO BANQI","DESPESAS VARIÁVEIS PRODUTOS DE CRÉDITO",
    "DESPESAS VARIAVEIS PERDA EFETIVA","DESPESAS VARIAVEIS IOF CANCELAMENTO","DESPESAS VARIAVEIS CRÉDITO PROVISÃO",
    "DESPESAS VARIAVEIS RENDAS DE RECUPERAÇÃO","DESPESAS VARIÁVEIS FRAUDES COM CARNÊ","DESPESAS VARIÁVEIS CUSTO DE CRÉDITO E COBRANÇA",
    "DESPESAS VARIAVEIS ANÁLISE DE CRÉDITO","DESPESAS VARIÁVEIS COBRANÇA","MARGEM #2","CUSTO CARREGAMENTO ESTOQUE TOTAL",
    "CUSTO CARREGAMENTO ESTOQUE","CUSTO CARREGAMENTO ESTOQUE FINANCIAMENTO FORNECEDOR","CUSTO CARREGAMENTO ESTOQUE FORNECEDOR CONVÊNIO",
    "MARGEM #3","CFC TOTAL","CFC FRAUDE E CHARGEBACK TOTAL","CFC  FRAUDE E CHARGEBACK PERDA COM FRAUDE",
    "CFC  FRAUDE E CHARGEBACK PERDA COM DESACORDO","CFC FRAUDE E CHARGEBACK FERRAMENTAS ANTIFRAUDE","CFC  FRAUDE E CHARGEBACK PROVISÃO",
    "CFC VENDAS NO CARTÃO COM JUROS","CFC MEIOS DE PAGAMENTO","CFC  MEIOS DE PAGAMENTO CARTÃO DE CRÉDITO",
    "CFC  MEIOS DE PAGAMENTO NUMERÁRIO","CFC  MEIOS DE PAGAMENTO PIX","CFC  MEIOS DE PAGAMENTO GATEWAY","CFC FINANCIAMENTO AO CONSUMIDOR",
    "MARGEM #4","DESPESAS SEMI VARIÁVEIS TOTAL","DESPESAS SEMI VARIÁVEIS FROTA PRÓPRIA","DESPESAS SEMI VARIÁVEIS CTO",
    "DESPESAS SEMI VARIÁVEIS BANCÁRIAS E FIANÇA","DESPESAS SEMI VARIÁVEIS COM PESSOAL","MARGEM #5","DEMAIS DESPESAS DIRETAS TOTAL",
    "DEMAIS DESPESAS DIRETAS JURÍDICO","DEMAIS DESPESAS DIRETAS TECNOLOGIA","DEMAIS DESPESAS DIRETAS GERAIS E ADMINISTRATIVAS",
    "DEMAIS DESPESAS DIRETAS ASSESSORIA","DEMAIS DESPESAS DIRETAS COMUNICAÇÃO E SUSTENTABILIDADE","DEMAIS DESPESAS DIRETAS RESULTADO FINANCEIRO",
    "DEMAIS DESPESAS DIRETAS ATENDIMENTO E TELEVENDAS","DEMAIS DESPESAS DIRETAS PERDAS OPERACIONAIS","DEMAIS DESPESAS DIRETAS DEPRECIAÇÃO",
    "DEMAIS DESPESAS DIRETAS OUTROS","DEMAIS DESPESAS DIRETAS EQUIVALÊNCIA","MBL","DESPESAS INDIRETAS TOTAL","LAIR TOTAL"
]
KPI_MASTER_NORM = {_norm_key(k): k for k in KPI_MASTER_LIST}

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

def normalize_key(x: str) -> str:
    x = re.sub(r"\s+", " ", str(x)).strip().upper()
    return remove_accents(x)

# ===== Ordenação de KPI pelo campo ORDEM da base =====
def kpi_order_map(df_src: pd.DataFrame) -> dict:
    if "KPI_COMPACT" not in df_src.columns:
        return {}
    tmp = df_src[["KPI_COMPACT","ORDEM"]].dropna(subset=["KPI_COMPACT"]).copy()
    tmp["ORDEM"] = pd.to_numeric(tmp["ORDEM"], errors="coerce")
    # pega a menor ORDEM por KPI (como seu exemplo)
    m = tmp.groupby("KPI_COMPACT", dropna=True)["ORDEM"].min().to_dict()
    return m

def kpi_filter_options_from_base(df_src: pd.DataFrame) -> list[str]:
    """
    Retorna a lista de KPIs para filtros respeitando a ORDEM da base:
    - usa EXATAMENTE os nomes de df_src['KPI_COMPACT']
    - ordena por menor ORDEM por KPI (min)
    - adiciona KPIs da lista-mestra que não estão na base no final (A–Z)
    """
    if "KPI_COMPACT" not in df_src.columns:
        base_names = []
        ord_map = {}
    else:
        ord_map = kpi_order_map(df_src)
        # mantém ordem de primeira ocorrência e evita set (que desordena)
        base_names = (
            df_src["KPI_COMPACT"]
            .dropna().astype(str)
            .pipe(lambda s: list(dict.fromkeys(s)))
        )
    base_sorted = sorted(base_names, key=lambda n: (ord_map.get(n, 9999), n))
    base_norm = {_norm_key(n) for n in base_sorted}
    extras = [k for k in KPI_MASTER_LIST if _norm_key(k) not in base_norm]
    extras_sorted = sorted(extras)
    return base_sorted + extras_sorted

# ==================== CARGA / NORMALIZAÇÃO ====================
@st.cache_data(show_spinner=False)
def load_normalize(file_bytes: bytes, filename: str) -> pd.DataFrame:
    base = pd.read_csv(io.BytesIO(file_bytes)) if filename.lower().endswith(".csv") else pd.read_excel(io.BytesIO(file_bytes))
    base = _to_upper(base)

    if "KPI_COMPACT" not in base.columns:
        for alt in ["KPI_COMPACTO", "KPI COMPACTO", "KPI COMPACT", "KPI_COMPACTO "]:
            if alt in base.columns:
                base.rename(columns={alt:"KPI_COMPACT"}, inplace=True)
                break

    for c in ["$", "PCT"]:
        if c in base.columns:
            base[c] = pd.to_numeric(base[c], errors="coerce")

    if "METRICA" in base.columns: base["METRICA"] = _norm_metric(base["METRICA"])
    if "PERIODO" in base.columns: base["PERIODO"] = _norm_period(base["PERIODO"])

    if "PCT" in base.columns:
        base["PCT"] = np.where(base["PCT"].abs() > 1, base["PCT"]/100.0, base["PCT"])

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

    # DIRETORIA_KEY super permissivo (garante CAUDA/INFO)
    base["DIRETORIA_KEY"] = base["DIRETORIA"].apply(normalize_key)
    def _norm_dirkey(x: str) -> str:
        s = normalize_key(x)
        if s in {"", "CONSOLIDADO", "TOTAL", "GERAL", "CONSOLIDADO ECOM", "ECOM CONSOLIDADO"}:
            return ""  # Consolidado
        if "LINHA BRANCA" in s: return "LINHA BRANCA"
        if "MOVEIS" in s or "MÓVEIS" in s: return "MOVEIS"
        if "TELAS" in s or "TV" in s: return "TELAS"
        if "TELEFONIA" in s or "CELULAR" in s or "MOBILE" in s: return "TELEFONIA"
        if "LINHA LEVE" in s or "SAZONAL" in s or "SAZONAIS" in s: return "LINHA LEVE E SAZONAL"
        if any(k in s for k in ["INFO","INFORMATI","PERIFERIC","PERIFÉRIC","INFORMÁTICA","INFORMATICA","INFO/PERIF"]):
            return "INFO"
        if any(k in s for k in ["CAUDA","LONG TAIL","CAUDA LONGA","LONGA"]):
            return "CAUDA"
        return s
    base["DIRETORIA_KEY"] = base["DIRETORIA_KEY"].apply(_norm_dirkey)

    return base

# === Fonte de dados: arquivo do repositório por padrão (sidebar apenas p/ fonte + versão) ===
DEFAULT_DATA_PATH = os.path.join(os.path.dirname(__file__), "BASE_PNL.xlsx")
with st.sidebar:
    st.markdown("### Fonte de dados")
    use_repo_file = st.checkbox("Usar BASE_PNL.xlsx do repositório", value=True)
    uploaded = None
    if not use_repo_file:
        uploaded = st.file_uploader("Carregue uma base (XLSX/CSV)", type=["xlsx","xls","csv"], key="upl1")
    st.markdown("---")
    st.markdown("<div style='color:#8A94A6;font-size:12px;'>Versão: <b>V1</b></div>", unsafe_allow_html=True)

# leitura bytes + nome
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

base = load_normalize(file_bytes, filename)

# ==================== Abas (topo) ====================
tab1, tab2, tab3, tab4 = st.tabs(["Visão Geral", "Visão Diretoria", "Gráficos", "Versionamento"])

# ==================== Filtros (expander comum – melhor mobile) ====================
# Diretoria: ordem fixa
DIR_FIXED_ORDER = ["", "LINHA BRANCA", "MOVEIS", "TELAS", "TELEFONIA", "LINHA LEVE E SAZONAL", "INFO", "CAUDA"]
def order_diretorias(opts):
    seen, out = set(), []
    for k in DIR_FIXED_ORDER:
        if k in opts and k not in seen:
            out.append(k); seen.add(k)
    for k in sorted(opts):
        if k not in seen:
            out.append(k); seen.add(k)
    return out

# opções para filtros
dir_keys_raw = [x for x in base["DIRETORIA_KEY"].dropna().astype(str).unique().tolist()]
dir_keys_options = order_diretorias(dir_keys_raw)
def _fmt_dir(k): return ("Consolidado" if k=="" else k.title())
default_dir_key = "" if "" in dir_keys_options else (dir_keys_options[0] if dir_keys_options else "")

bu_vals = sorted([x for x in base["BU"].dropna().unique() if x])
if "ORDEM SETOR" in base.columns:
    setores = base[["CATEGORIA","ORDEM SETOR"]].drop_duplicates().sort_values("ORDEM SETOR")
    setor_list = setores["CATEGORIA"].tolist()
else:
    setor_list = sorted([x for x in base["CATEGORIA"].dropna().unique() if x])

periods = sorted(base["PERIODO"].dropna().unique().tolist())
p0_default_index = len(periods)-1 if periods else 0

with st.expander("🧩 Filtros", expanded=True):
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        diretoria_sel_keys = st.multiselect("Diretoria", dir_keys_options,
                                            default=[default_dir_key] if default_dir_key in dir_keys_options else [],
                                            format_func=_fmt_dir, key="flt_diretoria")
    with c2:
        bu_sel = st.multiselect("BU", bu_vals, default=[], key="flt_bu")
    with c3:
        setor_sel = st.multiselect("Setor", setor_list, default=[], key="flt_setor")
    with c4:
        p0 = st.selectbox("Mês vigente (P0)", periods, index=p0_default_index, key="flt_p0")

    t1, t2, t3, t4 = st.columns(4)
    with t1:
        show_totais = st.checkbox("Mostrar apenas totais", value=False, key="opt_totais")
    with t2:
        show_money = st.checkbox("Exibir colunas $", value=True, key="opt_money")
    with t3:
        show_percent = st.checkbox("Exibir colunas %RL", value=False, key="opt_pct")  # default desmarcado
    with t4:
        show_principais = st.checkbox("Exibir apenas KPIs principais", value=False, key="opt_main")

    only_margins = st.checkbox("Exibir apenas linhas de Margem (MC#1 a MC#4)", value=False, key="opt_margins")

# --- aplica filtros base comuns ---
def apply_common_filters(df0: pd.DataFrame):
    d = df0.copy()
    if diretoria_sel_keys:
        d = d[d["DIRETORIA_KEY"].isin(diretoria_sel_keys)]
    if bu_sel:
        d = d[d["BU"].isin(bu_sel)]
    if setor_sel:
        d = d[d["CATEGORIA"].isin(setor_sel)]
    if show_principais:
        d = d[d["PRINCIPAL"]=="SIM"]
    if show_totais:
        d = d[d["AGREG"].str.lower()=="pai"]
    return d

df = apply_common_filters(base)
df_all_dirs = base.copy()
df_all_dirs = apply_common_filters(df_all_dirs.assign(DIRETORIA_KEY=base["DIRETORIA_KEY"]))  # aplica BU/Setor/etc., mas mantém todas diretorias caso precise

if df.empty:
    st.info("Sem dados para os filtros selecionados."); st.stop()

# --- P0 efetivo ---
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

# ==================== PIVOT ====================
@st.cache_data(show_spinner=False)
def pivotize(df_in: pd.DataFrame, p0, p_m1, p_m2, p_m3, p_m12):
    index_cols = ["AGREG","KPI_COMPACT","KPI","SINAL","FAMILIA","ORDEM","CATEGORIA","TIPO","DIRETORIA","DIRETORIA_KEY"]
    df_key_sorted = df_in.sort_values(["AGREG","KPI_COMPACT","KPI","PERIODO","METRICA"])
    df_dedup = df_key_sorted.drop_duplicates(subset=index_cols + ["PERIODO","METRICA"], keep="first")
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

    # deltas (somente calculadas aqui)
    m["d_m1"]   = m["proj"] - m["real_m1"]
    m["d_m12"]  = m["proj"] - m["real_m12"]
    m["d_fc"]   = m["proj"] - m["fcst"]
    m["pd_m1"]  = m["p_proj"] - m["p_m1v"]
    m["pd_m12"] = m["p_proj"] - m["p_m12v"]
    m["pd_fc"]  = m["p_proj"] - m["p_fcst"]
    return m

m  = pivotize(df,          p0, p_m1, p_m2, p_m3, p_m12)
mD = pivotize(df_all_dirs, p0, p_m1, p_m2, p_m3, p_m12)

def dedup_kpi(df_in: pd.DataFrame) -> pd.DataFrame:
    pais   = df_in[df_in["AGREG"].str.lower()=="pai"].sort_values("ORDEM")
    filhos = df_in[df_in["AGREG"].str.lower()=="filho"].sort_values("ORDEM")
    pais   = pais.drop_duplicates(subset=["KPI_COMPACT"], keep="first")
    filhos = filhos.drop_duplicates(subset=["KPI"], keep="first")
    out = pd.concat([pais, filhos], ignore_index=True).sort_values("ORDEM")
    return out

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

# ==================== CONTRIBUIÇÃO POR SETOR (robusto) ====================
def sector_contribution_delta_m1(df_raw: pd.DataFrame, kpi_compact: str, p0: str, p_m1: str) -> pd.Series:
    """
    Retorna contribuição por CATEGORIA do gap (Projeção - Real M-1) do KPI pai.
    Estratégia:
      1) Consolidado (DIR in {"", CONSOLIDADO, TOTAL}) + filhos
      2) Consolidado + pais
      3) Todas diretorias + filhos; se vazio, todas diretorias + pais
    """
    if df_raw.empty:
        return pd.Series(dtype=float)

    def _cleanup(d):
        d = d.copy()
        d["DIR_KEY_N"] = d["DIRETORIA_KEY"].astype(str).str.strip().str.upper().fillna("")
        d["AGREG_N"]   = d["AGREG"].astype(str).str.strip().str.lower()
        d["CAT_UP"]    = d["CATEGORIA"].astype(str).str.strip().str.upper()
        return d

    def _sum_rs(g):
        vals  = pd.to_numeric(g["$"], errors="coerce").fillna(0.0)
        sinal = pd.to_numeric(g["SINAL"], errors="coerce").fillna(1.0)
        return (vals * sinal).sum(min_count=1)

    def _make_series(dsub: pd.DataFrame) -> pd.Series:
        proj = dsub[(dsub["PERIODO"]==p0)  & (dsub["METRICA"]=="projecao")].groupby("CAT_UP").apply(_sum_rs)
        m1   = dsub[(dsub["PERIODO"]==p_m1) & (dsub["METRICA"]=="realizado")].groupby("CAT_UP").apply(_sum_rs)
        s = (proj - m1)
        if s is None or s.empty:
            return pd.Series(dtype=float)
        s = s.dropna()
        s = s[~s.index.isin(["", "TOTAL", "CONSOLIDADO"])]
        s = s[s != 0]
        return s

    d0 = _cleanup(df_raw)
    d0 = d0[d0["KPI_COMPACT"]==kpi_compact]

    cons_keys = {"", "CONSOLIDADO", "TOTAL"}
    have_cons = d0["DIR_KEY_N"].isin(cons_keys).any()

    # 1) consolidado + filhos
    if have_cons:
        s = _make_series(d0[d0["DIR_KEY_N"].isin(cons_keys) & (d0["AGREG_N"]=="filho")])
        if not s.empty:
            return s.sort_values(ascending=True) if (s<0).any() else s.reindex(s.abs().sort_values(ascending=False).index)

    # 2) consolidado + pais
    if have_cons:
        s = _make_series(d0[d0["DIR_KEY_N"].isin(cons_keys) & (d0["AGREG_N"]=="pai")])
        if not s.empty:
            return s.sort_values(ascending=True) if (s<0).any() else s.reindex(s.abs().sort_values(ascending=False).index)

    # 3) todas diretorias + filhos (ou pais)
    s = _make_series(d0[d0["AGREG_N"]=="filho"])
    if s.empty:
        s = _make_series(d0[d0["AGREG_N"]=="pai"])
    if not s.empty:
        return s.sort_values(ascending=True) if (s<0).any() else s.reindex(s.abs().sort_values(ascending=False).index)

    return pd.Series(dtype=float)

# ==================== RENDERIZAÇÃO TABELAS ====================
def render_table_general(m_df: pd.DataFrame, df_raw: pd.DataFrame, table_id="pnltbl_general") -> str:
    m_consol = m_df.copy()
    tbl = dedup_kpi(m_consol)
    tbl["_PAI"] = (tbl["AGREG"].str.lower()=="pai").astype(int)
    tbl["DRE"] = np.where(tbl["_PAI"]==1, "**"+tbl["KPI_COMPACT"]+"**", tbl["KPI"])

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
    if show_money:   cols += ["real_m3","real_m2","real_m1","proj","d_m1","d_m12","d_fc"]
    if show_percent: cols += ["p_proj","pd_m1","pd_m12","pd_fc"]

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
                if str(t).upper()!="PP": return fmt_brl(v)
                if pd.isna(p): p = _fallback_pct_real_m1(df_raw, a, kc, k, p_m1)
                return "" if pd.isna(p) else fmt_pp_value(p)
            df_show["real_m1"] = [
                _fmt_real_m1_pp(v, p, t, a, kc, k)
                for v,p,t,a,kc,k in zip(df_show["real_m1"], pm1_vals, df_show["TIPO"], dfv["AGREG"], dfv["KPI_COMPACT"], dfv["KPI"])
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
    all_keys = [x for x in m_df["DIRETORIA_KEY"].fillna("").astype(str).unique().tolist()]
    # ordem fixa
    def order_diretorias_local(opts):
        seen, out = set(), []
        for k in DIR_FIXED_ORDER:
            if k in opts and k not in seen:
                out.append(k); seen.add(k)
        for k in sorted(opts):
            if k not in seen:
                out.append(k); seen.add(k)
        return out
    dir_order = order_diretorias_local(all_keys)

    key_to_label = (
        m_df[["DIRETORIA_KEY","DIRETORIA"]]
        .drop_duplicates()
        .set_index("DIRETORIA_KEY")["DIRETORIA"]
        .to_dict()
    )
    def disp_label(k):
        return "Consolidado" if k=="" else (key_to_label.get(k) or k).title()

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
    desired_order = ["M-12","M-3","M-2","M-1","Projeção"]
    # diretorias a exibir
    keys_want = list(diretoria_sel_keys or [])
    if not keys_want:
        for k in DIR_FIXED_ORDER:
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
        if sdir.empty: continue
        sdir = dedup_kpi(sdir); r = sdir.iloc[0]
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
                rows.append({"Diretoria": ("Consolidado" if k=="" else k.title()), "Período": lab, "Valor": y, "YLabel": y_label, "Label": label})
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
        align='left', dx=8, dy=-8, fontWeight='bold', fontSize=13
    ).encode(
        x='Período', y='Valor:Q', text=alt.Text('Label'), color='Diretoria:N'
    )
    st.markdown(f"**{kpi_name} – Evolução**")
    st.altair_chart(base + labels, use_container_width=True)

def draw_margin_block(m_df: pd.DataFrame, keys_source_df: pd.DataFrame, diretoria_sel_keys: list[str]):
    desired_order = ["M-12","M-3","M-2","M-1","Projeção"]
    margin_names = [
        "Margem Contribuição #1 (CashMerc + Bonif. + Demais Rec)",
        "Margem Contribuição #2",
        "Margem Contribuição #3",
        "Margem Contribuição #4",
    ]
    keys_want = list(diretoria_sel_keys or [])
    if not keys_want:
        for k in DIR_FIXED_ORDER:
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
                    rows.append({"Diretoria": ("Consolidado" if k=="" else k.title()), "Período": lab,
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

# ==================== Telas ====================
with tab1:  # Visão Geral
    # Filtro de KPI respeitando ORDEM da base
    kpi_opts_all = kpi_filter_options_from_base(df_all_dirs)
    kpi_opts = ["(todos)"] + kpi_opts_all
    kpi_filter = st.selectbox("Filtrar KPI (linha):", options=kpi_opts, index=0, key="kpi_vg")

    # aplica filtro (somente linha específica) se selecionado
    if kpi_filter != "(todos)":
        m_show = m[(m["KPI_COMPACT"]==kpi_filter) | (m["KPI"]==kpi_filter)].copy()
        if m_show.empty:
            st.info("KPI sem dados para os filtros."); st.stop()
    else:
        m_show = m

    st.markdown(render_table_general(m_show, df), unsafe_allow_html=True)

    # Highlights
    st.markdown("### 🔎 Highlights do mês")

    def _is_consolidado_selected_only():
        return (len(diretoria_sel_keys)==1 and diretoria_sel_keys[0]=="" and len(setor_sel)==0)

    _show_sector_breakdown = _is_consolidado_selected_only()

    def _excluded_kpi(name: str) -> bool:
        up = (name or "").upper()
        if any(x in up for x in ["IMPOST", "TRIBUT", "TAXA", "ICMS", "PIS", "COFINS", "ISS"]): return True
        if re.search(r"MARGEM\s*[5-9]", up): return True
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
            kpi_name = r["KPI_COMPACT"]; delta_rs = abs(float(r["gap_m1"]))
            setores_txt = ""
            if _show_sector_breakdown:
                contr = sector_contribution_delta_m1(df_all_dirs, kpi_name, p0, p_m1)
                if contr.empty:
                    contr = sector_contribution_delta_m1(df, kpi_name, p0, p_m1)
                if not contr.empty:
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

with tab2:  # Visão Diretoria
    kpi_opts_all = kpi_filter_options_from_base(df_all_dirs)
    kpi_opts = ["(todos)"] + kpi_opts_all
    kpi_filter = st.selectbox("Filtrar KPI (linha):", options=kpi_opts, index=0, key="kpi_vd")

    if kpi_filter != "(todos)":
        mD_show = mD[(mD["KPI_COMPACT"]==kpi_filter) | (mD["KPI"]==kpi_filter)].copy()
        if mD_show.empty:
            st.info("KPI sem dados para os filtros."); st.stop()
    else:
        mD_show = mD

    html_dir = render_table_diretoria(mD_show, table_id="pnltbl_dir")
    st.markdown(html_dir, unsafe_allow_html=True)

with tab3:  # Gráficos
    kpi_options_ordered = kpi_filter_options_from_base(df_all_dirs)
    kpi_sel = st.multiselect("KPI(s) (opcional):", options=kpi_options_ordered, default=[], key="kpi_gfx")

    if kpi_sel:
        for kpi_name in kpi_sel:
            draw_kpi_evolution(mD, df_all_dirs, kpi_name, diretoria_sel_keys)
    else:
        draw_margin_block(mD, df_all_dirs, diretoria_sel_keys)

with tab4:  # Versionamento
    st.markdown("## 📌 Roadmap / Próximas entregas")
    st.markdown("""
- **Filtro B2B**
- **Visão canal B2C**
- **Visão parceiro B2B**
- **Simulador**
    """)
    with st.expander("🛠 Diagnóstico (para suporte)"):
        st.write("Diretorias disponíveis (KEY → count):")
        st.write(df_all_dirs["DIRETORIA_KEY"].value_counts())
        st.write("Diretorias originais:")
        st.write(df_all_dirs["DIRETORIA"].value_counts())
        st.write("KPIs (KPI_COMPACT) exemplos:")
        st.write(df_all_dirs["KPI_COMPACT"].dropna().unique()[:50])

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
