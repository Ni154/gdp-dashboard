# streamlit_app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import zipfile, io
import plotly.express as px
from streamlit_plotly_events import plotly_events
from pathlib import Path
import unicodedata
from datetime import date

st.set_page_config(page_title="Balancete (clicável)", page_icon="📘", layout="wide")
st.title("📘 Painel de Balancete — com clique para filtrar")
st.caption("Envie .xlsx (ou .zip com .xlsx) com as abas **Balancete** e **Mapa_Classificacao**. "
           "Também aceita planilhas em formato **matriz** (Conta, Descrição, centros e Total). "
           "Clique nos gráficos para filtrar KPIs e Tabela.")

# ---------- util
def _strip_accents(s: str) -> str:
    if s is None: return ""
    s = unicodedata.normalize("NFKD", str(s))
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def _norm_token(s: str) -> str:
    s = _strip_accents(str(s)).lower()
    for ch in [" ", "_", "-", ".", "/"]:
        s = s.replace(ch, "")
    return s

# ---------- normalização de colunas “balancete”
def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # aliases básicos
    aliases = {
        "Conta": "ContaCodigo", "Conta Código": "ContaCodigo", "Conta Contábil": "ContaCodigo",
        "Descrição": "ContaDescricao", "Descricao": "ContaDescricao", "Histórico": "ContaDescricao",
        "DataCompetencia": "Competencia", "Competência": "Competencia",
        "Centro de Custo": "CentroCusto", "CC": "CentroCusto",
    }
    for a, b in aliases.items():
        if a in df.columns and b not in df.columns:
            df.rename(columns={a: b}, inplace=True)
    return df

# ---------- conversor: planilha MATRIZ -> balancete-like
def _is_matrix_format(df: pd.DataFrame) -> bool:
    cols = set(df.columns.astype(str))
    if not {"Conta", "Descrição"}.issubset(cols):  # cabeçalhos originais
        return False
    centers = [c for c in df.columns if c not in ["Conta", "Descrição"]]  # vários centros (e possivelmente Total)
    return len(centers) >= 2

def _matrix_to_balancete(df_matrix: pd.DataFrame, empresa_default="Empresa") -> pd.DataFrame:
    dfm = df_matrix.copy()
    center_cols = [c for c in dfm.columns if c not in ["Conta", "Descrição", "Total"]]

    # Converte números pt-BR -> float
    for c in center_cols + ["Total"]:
        if c in dfm.columns:
            dfm[c] = (dfm[c].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
            dfm[c] = pd.to_numeric(dfm[c], errors="coerce")

    # Longo
    mlong = dfm.melt(id_vars=["Conta", "Descrição"], value_vars=center_cols,
                     var_name="CentroCusto", value_name="Valor").dropna(subset=["Valor"])
    mlong = mlong[mlong["CentroCusto"].astype(str).str.len() > 0]

    # Monta balancete-like
    bal = pd.DataFrame({
        "Empresa": empresa_default,
        "Competencia": pd.Timestamp(date.today().replace(day=1)),
        "ContaCodigo": mlong["Conta"].astype(str).str.strip(),
        "ContaDescricao": mlong["Descrição"].astype(str).str.strip(),
        "CentroCusto": mlong["CentroCusto"].astype(str).str.strip(),
    })

    conta_str = bal["ContaCodigo"].fillna("").astype(str)
    is_receita = conta_str.str.strip().str.startswith("3")
    is_despesa = conta_str.str.strip().str.startswith("4")
    valor = mlong["Valor"].fillna(0.0)

    bal["Devedor"] = np.where(is_despesa, valor, 0.0)
    bal["Credor"]  = np.where(is_receita, valor, 0.0)

    other = ~(is_receita | is_despesa)
    bal.loc[other, "Devedor"] = np.where(other, np.where(valor >= 0, valor, 0.0), bal["Devedor"])
    bal.loc[other, "Credor"]  = np.where(other, np.where(valor < 0, -valor, 0.0), bal["Credor"])

    bal["Devedor"] = pd.to_numeric(bal["Devedor"], errors="coerce").fillna(0.0)
    bal["Credor"]  = pd.to_numeric(bal["Credor"], errors="coerce").fillna(0.0)

    return _norm_cols(bal)

# ---------- mapa gerado a partir da “Conta”
def _auto_mapa_from_conta_prefix(bal: pd.DataFrame) -> pd.DataFrame:
    def split_prefix(code, n):
        if pd.isna(code): return None
        parts = [p for p in str(code).split(".") if p]
        if not parts: return None
        return ".".join(parts[:min(n, len(parts))])

    prefixes = sorted(set(bal["ContaCodigo"].apply(lambda x: split_prefix(x, 3)).dropna()))
    rows = []
    for p in prefixes:
        first = str(p).split(".")[0]
        if first == "3":
            natureza, sinal = "Receita", -1
            grupo = "Receitas Operacionais"
        elif first == "4":
            natureza, sinal = "Despesa", 1
            grupo = "Despesas Operacionais"
        else:
            natureza, sinal = "Outros", 1
            grupo = "Outros"
        subgrupo = p
        rows.append([p, natureza, grupo, subgrupo, sinal, "Operacional"])
    mapa = pd.DataFrame(rows, columns=["ContaPrefixo","Natureza","GrupoGerencial","Subgrupo","Sinal","TipoOperacional"])
    return mapa

# ---------- detecção de abas por nome/conteúdo
def _norm_sheet_name(s: str) -> str:
    return _norm_token(s)

def _resolve_sheet(xls: pd.ExcelFile, desired: str) -> str | None:
    names = list(xls.sheet_names)
    norm_map = {name: _norm_sheet_name(name) for name in names}
    target = _norm_sheet_name(desired)
    for name, normed in norm_map.items():
        if normed == target:
            return name
    for name, normed in norm_map.items():
        if target in normed or normed in target:
            return name
    if "balancete" in target:
        for name, normed in norm_map.items():
            if "balancete" in normed:
                return name
    if any(k in target for k in ["mapa", "classificacao", "classific"]):
        for name, normed in norm_map.items():
            if "mapa" in normed or "classific" in normed:
                return name
    return None

# ---------- leitura flexível
def read_excel_like(uploaded, sheet_bal="Balancete", sheet_map="Mapa_Classificacao"):
    def _read_xlsx(flike):
        xls = pd.ExcelFile(flike)

        # Tenta modo balancete clássico
        found_bal = _resolve_sheet(xls, sheet_bal)
        found_map = _resolve_sheet(xls, sheet_map)

        bal, mapa = None, None

        if found_bal:
            bal_raw = pd.read_excel(xls, sheet_name=found_bal)
            mapa_raw = pd.read_excel(xls, sheet_name=found_map) if found_map else pd.DataFrame()
            bal = _norm_cols(bal_raw)
            mapa = _norm_cols(mapa_raw)

        # Se não achar, tenta formato MATRIZ
        if bal is None:
            for nm in xls.sheet_names:
                probe = pd.read_excel(xls, sheet_name=nm, nrows=50)
                if _is_matrix_format(probe):
                    bal = _matrix_to_balancete(pd.read_excel(xls, sheet_name=nm))
                    mapa = _auto_mapa_from_conta_prefix(bal)
                    break

        if bal is None:
            first = xls.sheet_names[0]
            bal = _norm_cols(pd.read_excel(xls, sheet_name=first))
            if len(xls.sheet_names) > 1:
                second = xls.sheet_names[1]
                mapa = _norm_cols(pd.read_excel(xls, sheet_name=second))
            else:
                mapa = pd.DataFrame()

        # Normalizações finais mínimas
        if "Empresa" not in bal.columns:
            bal["Empresa"] = "Empresa"
        if "Competencia" not in bal.columns:
            bal["Competencia"] = pd.Timestamp(date.today().replace(day=1))
        if "ContaCodigo" not in bal.columns and "Conta" in bal.columns:
            bal.rename(columns={"Conta": "ContaCodigo"}, inplace=True)
        if "ContaDescricao" not in bal.columns and "Descrição" in bal.columns:
            bal.rename(columns={"Descrição": "ContaDescricao"}, inplace=True)

        if "Devedor" not in bal.columns or "Credor" not in bal.columns:
            cand_val = None
            for c in ["Saldo", "Valor", "Total"]:
                if c in bal.columns:
                    cand_val = c
                    break
            if cand_val:
                v = (bal[cand_val].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
                v = pd.to_numeric(v, errors="coerce").fillna(0.0)
                if "Devedor" not in bal.columns:
                    bal["Devedor"] = np.where(v >= 0, v, 0.0)
                if "Credor" not in bal.columns:
                    bal["Credor"] = np.where(v < 0, -v, 0.0)
            else:
                bal["Devedor"] = 0.0
                bal["Credor"] = 0.0

        bal["Competencia"] = pd.to_datetime(bal["Competencia"], errors="coerce")
        bal["Devedor"] = pd.to_numeric(bal["Devedor"], errors="coerce").fillna(0.0)
        bal["Credor"]  = pd.to_numeric(bal["Credor"], errors="coerce").fillna(0.0)

        if mapa is None or mapa.empty:
            mapa = _auto_mapa_from_conta_prefix(bal)
        for c in ["ContaPrefixo","Natureza","GrupoGerencial","Subgrupo","Sinal","TipoOperacional"]:
            if c not in mapa.columns:
                mapa[c] = np.nan
        mapa["Sinal"] = pd.to_numeric(mapa["Sinal"], errors="coerce").fillna(1.0)

        return bal, mapa

    if hasattr(uploaded, "name") and str(uploaded.name).lower().endswith(".zip"):
        with zipfile.ZipFile(uploaded) as z:
            xlsx_names = [n for n in z.namelist() if n.lower().endswith(".xlsx")]
            if not xlsx_names:
                raise ValueError("ZIP sem arquivo .xlsx dentro.")
            with z.open(xlsx_names[0]) as xf:
                data = xf.read()
            bal, mapa = _read_xlsx(io.BytesIO(data))
    else:
        bal, mapa = _read_xlsx(uploaded)

    return bal, mapa

# ---------- merge + classificação
def merge_classify(bal, mapa):
    df = bal.copy()

    def split_prefix(code, n):
        if pd.isna(code): return None
        parts = [p for p in str(code).split(".") if p]
        if not parts: return None
        return ".".join(parts[:min(n, len(parts))])

    df["prefix3"] = df["ContaCodigo"].apply(lambda x: split_prefix(x, 3))
    df["prefix2"] = df["ContaCodigo"].apply(lambda x: split_prefix(x, 2))
    df["prefix1"] = df["ContaCodigo"].apply(lambda x: split_prefix(x, 1))

    m3 = df.merge(mapa.add_prefix("m3_"), left_on="prefix3", right_on="m3_ContaPrefixo", how="left")
    m2 = df.merge(mapa.add_prefix("m2_"), left_on="prefix2", right_on="m2_ContaPrefixo", how="left")
    m1 = df.merge(mapa.add_prefix("m1_"), left_on="prefix1", right_on="m1_ContaPrefixo", how="left")

    def coalesce(*cols):
        out = cols[0].copy()
        for c in cols[1:]: out = out.where(~out.isna(), c)
        return out

    out = df.copy()
    out["Natureza"]        = coalesce(m3.get("m3_Natureza"),        m2.get("m2_Natureza"),        m1.get("m1_Natureza"))
    out["GrupoGerencial"]  = coalesce(m3.get("m3_GrupoGerencial"),  m2.get("m2_GrupoGerencial"),  m1.get("m1_GrupoGerencial"))
    out["Subgrupo"]        = coalesce(m3.get("m3_Subgrupo"),        m2.get("m2_Subgrupo"),        m1.get("m1_Subgrupo"))
    out["Sinal"]           = coalesce(m3.get("m3_Sinal"),           m2.get("m2_Sinal"),           m1.get("m1_Sinal")).fillna(1.0)
    out["TipoOperacional"] = coalesce(m3.get("m3_TipoOperacional"), m2.get("m2_TipoOperacional"), m1.get("m1_TipoOperacional"))

    out["Saldo"] = out["Devedor"] - out["Credor"]
    out["Sinal"] = pd.to_numeric(out["Sinal"], errors="coerce").fillna(1.0)
    out["SaldoGerencial"] = out["Saldo"] * out["Sinal"]
    out["Competencia"] = pd.to_datetime(out["Competencia"], errors="coerce")
    out["AnoMes"] = out["Competencia"].dt.strftime("%Y-%m")
    return out

def metric_fmt(v):
    try: return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    except: return str(v)

def to_excel_bytes(dfs: dict) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        for name, d in dfs.items():
            d.to_excel(w, index=False, sheet_name=name[:31] or "Dados")
    out.seek(0)
    return out

@st.cache_data
def sample_excel_bytes() -> bytes:
    months = pd.period_range("2025-01", "2025-08", freq="M")
    rows = []
    for comp in [pd.Timestamp(m.start_time.date()) for m in months]:
        for cod, desc, cc, dev, cre in [
            ("3.1.1.01","Receita de Locação","Geral",0,1500000),
            ("3.1.1.02","Receita de Fretes","Geral",0,200000),
            ("3.1.1.03","Serviços Acessórios","Geral",0,50000),
            ("4.1.1.01","Despesas com Manutenção","Operação",74000,0),
            ("4.1.2.01","Despesas com Pessoal","RH",2200000,0),
            ("4.1.3.01","Despesas Administrativas","ADM",90000,0),
            ("1.1.1.01","Caixa","Geral",120000,0),
            ("2.1.1.01","Fornecedores","Geral",0,60000),
        ]:
            rows.append(["Empresa Modelo", comp.date().isoformat(), cod, desc, cc, float(dev), float(cre)])
    df_bal = pd.DataFrame(rows, columns=["Empresa","Competencia","ContaCodigo","ContaDescricao","CentroCusto","Devedor","Credor"])
    df_map = pd.DataFrame([
        ["3.1.1","Receita","Receitas Operacionais","Locação/Fretes",-1,"Operacional",""],
        ["4.1.1","Despesa","Despesas Operacionais","Manutenção",1,"Operacional",""],
        ["4.1.2","Despesa","Despesas Operacionais","Pessoal",1,"Operacional",""],
        ["4.1.3","Despesa","Despesas Operacionais","Administrativas",1,"Operacional",""],
        ["1.1.1","Ativo","Ativo Circulante","Disponibilidades",1,"Não Operacional",""],
        ["2.1.1","Passivo","Passivo Circulante","Fornecedores",-1,"Não Operacional",""],
    ], columns=["ContaPrefixo","Natureza","GrupoGerencial","Subgrupo","Sinal","TipoOperacional","Observacao"])
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        df_bal.to_excel(w, index=False, sheet_name="Balancete")
        df_map.to_excel(w, index=False, sheet_name="Mapa_Classificacao")
    out.seek(0)
    return out.read()

# ---------- sidebar
with st.sidebar:
    st.header("⚙️ Entrada")
    use_sample = st.toggle("Usar dados de exemplo", value=False, key="use_sample")
    up = None
    if not use_sample:
        up = st.file_uploader("Envie .xlsx ou .zip com .xlsx", type=["xlsx","zip"], key="uploader")
    sheet_bal = st.text_input("Aba do Balancete", "Balancete", key="sheet_bal")
    sheet_map = st.text_input("Aba do Mapa", "Mapa_Classificacao", key="sheet_map")

    x_default = Path("Exemplo_Balancete.xlsx")
    if x_default.exists() and not use_sample and not up:
        st.info("Nenhum arquivo enviado — carregando arquivo local **Exemplo_Balancete.xlsx** automaticamente.")

    st.download_button("⬇️ Baixar Excel de Exemplo", data=sample_excel_bytes(),
                       file_name="Exemplo_Balancete.xlsx", key="dl_sample")

# ---------- load
if st.session_state.get("use_sample", False):
    bal, mapa = read_excel_like(io.BytesIO(sample_excel_bytes()), sheet_bal, sheet_map)
else:
    if up:
        bal, mapa = read_excel_like(up, sheet_bal, sheet_map)
    else:
        default_path = Path("Exemplo_Balancete.xlsx")
        if default_path.exists():
            with default_path.open("rb") as f:
                bal, mapa = read_excel_like(f, sheet_bal, sheet_map)
        else:
            st.info("Envie um arquivo, ative **Usar dados de exemplo** ou inclua **Exemplo_Balancete.xlsx** junto do app.")
            st.stop()

df = merge_classify(bal, mapa)

# ---------- filtros base
empresas = sorted(df["Empresa"].dropna().unique().tolist())
naturezas = sorted(df["Natureza"].dropna().unique().tolist()) if "Natureza" in df.columns else []
cc_list  = sorted(df["CentroCusto"].dropna().unique().tolist()) if "CentroCusto" in df.columns else []

colf1, colf2, colf3 = st.columns(3)
with colf1:
    f_emp = st.multiselect("Empresa", empresas, default=empresas, key="f_emp")
with colf2:
    f_nat = st.multiselect("Natureza", naturezas or ["Receita","Despesa"],
                           default=naturezas or ["Receita","Despesa"], key="f_nat")
with colf3:
    f_cc = st.multiselect("Centro de Custo", cc_list, default=cc_list, key="f_cc") if cc_list else None

# ---------- slider robusto (corrige erro quando min=max ou quando datas são NaT)
comp_series = pd.to_datetime(df["Competencia"], errors="coerce")
if comp_series.notna().any():
    min_date = comp_series.min().date()
    max_date = comp_series.max().date()
else:
    today = date.today()
    min_date = today.replace(day=1)
    max_date = min_date

if min_date == max_date:
    f_date = st.slider("Competência (período)",
                       min_value=min_date, max_value=max_date,
                       value=min_date, key="f_periodo")
    start_dt = pd.to_datetime(f_date)
    end_dt   = pd.to_datetime(f_date)
else:
    f_date = st.slider("Competência (período)",
                       min_value=min_date, max_value=max_date,
                       value=(min_date, max_date), key="f_periodo")
    start_dt = pd.to_datetime(f_date[0])
    end_dt   = pd.to_datetime(f_date[1])

mask = (
    df["Empresa"].isin(f_emp)
    & df["Competencia"].between(start_dt, end_dt)
    & ((df["Natureza"].isin(f_nat)) if df["Natureza"].notna().any() else True)
)
if f_cc is not None:
    mask &= df["CentroCusto"].isin(f_cc)
df_base = df.loc[mask].copy()

# ---------- estado de clique
if "click_filters" not in st.session_state:
    st.session_state.click_filters = {"GrupoGerencial": None, "Subgrupo": None, "AnoMes": None}

colr1, colr2, colr3, colr4 = st.columns(4)
with colr1:
    if st.button("🔄 Limpar Grupo", key="btn_reset_grp"):
        st.session_state.click_filters["GrupoGerencial"] = None
with colr2:
    if st.button("🔄 Limpar Subgrupo", key="btn_reset_sub"):
        st.session_state.click_filters["Subgrupo"] = None
with colr3:
    if st.button("🔄 Limpar Mês", key="btn_reset_mes"):
        st.session_state.click_filters["AnoMes"] = None
with colr4:
    if st.button("🧹 Limpar TODOS", key="btn_reset_all"):
        st.session_state.click_filters = {"GrupoGerencial": None, "Subgrupo": None, "AnoMes": None}

# aplica filtros de clique
df_f = df_base.copy()
cf = st.session_state.click_filters
if cf["GrupoGerencial"]:
    df_f = df_f[df_f["GrupoGerencial"] == cf["GrupoGerencial"]]
if cf["Subgrupo"]:
    df_f = df_f[df_f["Subgrupo"] == cf["Subgrupo"]]
if cf["AnoMes"]:
    df_f = df_f[df_f["AnoMes"] == cf["AnoMes"]]

# ---------- KPIs
def metric_fmt(v):
    try: return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    except: return str(v)

colA, colB, colC, colD = st.columns(4)
receita = df_f.loc[df_f["Natureza"]=="Receita","SaldoGerencial"].sum()
despesa = df_f.loc[df_f["Natureza"]=="Despesa","SaldoGerencial"].sum()
resultado = receita + despesa
margem = (resultado / receita) if receita else np.nan
with colA:
    st.metric("Receita", metric_fmt(receita))
with colB:
    st.metric("Despesa", metric_fmt(despesa))
with colC:
    st.metric("Resultado", metric_fmt(resultado))
with colD:
    st.metric("Margem %", metric_fmt((margem * 100) if pd.notna(margem) else 0))
st.markdown("---")

# ---------- containers fixos (DOM estável)
c_resultado = st.container()
c_despesas  = st.container()
c_receitas  = st.container()
c_tabela    = st.container()
c_export    = st.container()

with c_resultado:
    st.subheader("📈 Resultado por Mês (clique para filtrar)")
    series = df_f.groupby("AnoMes", as_index=False)["SaldoGerencial"].sum().sort_values("AnoMes")
    if not series.empty:
        fig_line = px.line(series, x="AnoMes", y="SaldoGerencial", markers=True)
        sel = plotly_events(fig_line, click_event=True, hover_event=False, select_event=False,
                            override_height=420, override_width="100%", key="ev_resultado_mes")
        if sel:
            x_val = sel[0].get("x")
            if x_val:
                st.session_state.click_filters["AnoMes"] = str(x_val)
                st.success(f"Filtro aplicado: AnoMes = {st.session_state.click_filters['AnoMes']}")
    else:
        st.info("Sem dados no período.")

with c_despesas:
    st.subheader("📊 Despesas por Grupo (clique para filtrar)")
    dep = df_f[df_f["Natureza"]=="Despesa"].groupby("GrupoGerencial", as_index=False)["SaldoGerencial"].sum()
    if not dep.empty:
        dep = dep.sort_values("SaldoGerencial")
        fig_bar = px.bar(dep, x="SaldoGerencial", y="GrupoGerencial",
                         orientation="h", color="GrupoGerencial")  # cores por grupo
        sel = plotly_events(fig_bar, click_event=True, hover_event=False, select_event=False,
                            override_height=420, override_width="100%", key="ev_despesas_grp")
        if sel:
            y_val = sel[0].get("y")
            if y_val:
                st.session_state.click_filters["GrupoGerencial"] = str(y_val)
                st.success(f"Filtro aplicado: GrupoGerencial = {st.session_state.click_filters['GrupoGerencial']}")
    else:
        st.info("Sem dados de despesa nos filtros atuais.")

with c_receitas:
    st.subheader("🏆 Top 10 Receitas por Subgrupo (clique para filtrar)")
    rec = df_f[df_f["Natureza"]=="Receita"].groupby("Subgrupo", as_index=False)["SaldoGerencial"].sum()
    if not rec.empty:
        rec = rec.sort_values("SaldoGerencial", ascending=False).head(10)
        fig_rec = px.bar(rec, x="Subgrupo", y="SaldoGerencial", color="Subgrupo")  # cores por subgrupo
        sel = plotly_events(fig_rec, click_event=True, hover_event=False, select_event=False,
                            override_height=420, override_width="100%", key="ev_receitas_sub")
        if sel:
            x_val = sel[0].get("x")
            if x_val:
                st.session_state.click_filters["Subgrupo"] = str(x_val)
                st.success(f"Filtro aplicado: Subgrupo = {st.session_state.click_filters['Subgrupo']}")
    else:
        st.info("Sem dados de receita nos filtros atuais.")

with c_tabela:
    st.subheader("Tabela detalhada (após filtros por clique)")
    show_cols = ["Empresa","Competencia","ContaCodigo","ContaDescricao","CentroCusto",
                 "Natureza","GrupoGerencial","Subgrupo","Devedor","Credor",
                 "Saldo","Sinal","SaldoGerencial","AnoMes"]
    show_cols = [c for c in show_cols if c in df_f.columns]
    st.dataframe(df_f[show_cols].sort_values(["Competencia","ContaCodigo"]).reset_index(drop=True),
                 use_container_width=True, key="grid_detalhe")

with c_export:
    st.subheader("Exportações")
    pivot_mes = df_f.pivot_table(index="AnoMes", columns="Natureza", values="SaldoGerencial",
                                 aggfunc="sum", fill_value=0).reset_index()
    by_grupo = df_f.groupby(["Natureza","GrupoGerencial"], as_index=False)["SaldoGerencial"] \
                   .sum().sort_values(["Natureza","SaldoGerencial"], ascending=[True, False])

    excel_bytes = to_excel_bytes({"Detalhado": df_f[show_cols], "Resumo_Mensal": pivot_mes, "Por_Grupo": by_grupo})
    st.download_button("⬇️ Excel (Detalhado + Resumos)", data=excel_bytes,
                       file_name="analise_balancete.xlsx", key="dl_excel")
    st.download_button("⬇️ CSV Detalhado", data=df_f.to_csv(index=False).encode("utf-8"),
                       file_name="balancete_detalhado.csv", key="dl_csv")
