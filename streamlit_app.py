# streamlit_app.py
import sys
from datetime import date
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit_plotly_events import plotly_events

st.set_page_config(page_title="Balancete (seu modelo)", page_icon="📘", layout="wide")
st.title("📘 Painel de Balancete — seu modelo (1 aba)")
st.caption("Importe .xlsx com colunas: Empresa, Competencia, ContaCodigo, ContaDescricao, CentroCusto, Devedor, Credor.")

# ================= Helpers =================
REQUIRED = {"Empresa","Competencia","ContaCodigo","ContaDescricao","CentroCusto","Devedor","Credor"}

ALIASES = {
    "Conta Código":"ContaCodigo", "Conta Contábil":"ContaCodigo", "Conta":"ContaCodigo", "ContaCod":"ContaCodigo",
    "Descrição":"ContaDescricao", "Descricao":"ContaDescricao", "Historico":"ContaDescricao", "Histórico":"ContaDescricao",
    "Centro de Custo":"CentroCusto", "CC":"CentroCusto",
    "DataCompetencia":"Competencia", "Competência":"Competencia",
}

def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for a,b in ALIASES.items():
        if a in df.columns and b not in df.columns:
            df.rename(columns={a:b}, inplace=True)
    if "Conta" in df.columns and "ContaCodigo" not in df.columns:
        df.rename(columns={"Conta":"ContaCodigo"}, inplace=True)
    if "Descrição" in df.columns and "ContaDescricao" not in df.columns:
        df.rename(columns={"Descrição":"ContaDescricao"}, inplace=True)
    return df

def read_sheet(xfile, wanted_sheet: str|None=None) -> pd.DataFrame:
    xls = pd.ExcelFile(xfile)
    sheet = wanted_sheet if (wanted_sheet and wanted_sheet in xls.sheet_names) else xls.sheet_names[0]
    return norm_cols(pd.read_excel(xls, sheet_name=sheet))

def ensure_required(df: pd.DataFrame):
    miss = REQUIRED - set(df.columns)
    if miss:
        raise ValueError(f"Planilha faltando colunas obrigatórias: {sorted(miss)}")

def to_excel_bytes(dfs: dict) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        for name, d in dfs.items():
            d.to_excel(w, index=False, sheet_name=name[:31] or "Dados")
    out.seek(0)
    return out

def fmt(v):
    try: return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    except: return str(v)

# ================= Sidebar =================
with st.sidebar:
    st.header("⚙️ Entrada")
    up = st.file_uploader("Envie .xlsx", type=["xlsx"], key="uploader")
    sheet_name = st.text_input("(Opcional) Nome da aba", "", key="sheet")

if not up:
    st.info("Envie sua planilha .xlsx no formato indicado.")
    st.stop()

# ================= Leitura + preparação =================
df = read_sheet(up, sheet_name if sheet_name.strip() else None)
ensure_required(df)

# Tipos
df["Competencia"] = pd.to_datetime(df["Competencia"], errors="coerce")
for c in ["Devedor","Credor"]:
    df[c] = (df[c].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

# Natureza/Sinal por prefixo (3=Receita, 4=Despesa, outros=Outros)
conta_str = df["ContaCodigo"].astype(str).str.strip()
df["Natureza"] = np.select(
    [conta_str.str.startswith("3"), conta_str.str.startswith("4")],
    ["Receita","Despesa"], default="Outros"
)
df["Sinal"] = np.select(
    [df["Natureza"].eq("Receita"), df["Natureza"].eq("Despesa")],
    [-1, 1], default=1
)

# Saldos
df["Saldo"] = df["Devedor"] - df["Credor"]
df["SaldoGerencial"] = df["Saldo"] * df["Sinal"]
df["AnoMes"] = df["Competencia"].dt.strftime("%Y-%m")

# Se não houver datas válidas, força um mês único para não quebrar
if df["Competencia"].isna().all():
    today = date.today().replace(day=1)
    df["Competencia"] = pd.Timestamp(today)
    df["AnoMes"] = df["Competencia"].dt.strftime("%Y-%m")

# ================= Filtros =================
empresas = sorted(df["Empresa"].dropna().unique().tolist())
naturezas = sorted(df["Natureza"].dropna().unique().tolist()) or ["Receita","Despesa"]
ccs = sorted(df["CentroCusto"].dropna().unique().tolist())

colf1,colf2,colf3 = st.columns(3)
with colf1: f_emp = st.multiselect("Empresa", empresas, default=empresas, key="f_emp")
with colf2: f_nat = st.multiselect("Natureza", naturezas, default=naturezas, key="f_nat")
with colf3: f_cc  = st.multiselect("Centro de Custo", ccs, default=ccs, key="f_cc")

# ================= Slider (índice de mês – robusto) =================
meses = sorted(df["AnoMes"].dropna().unique().tolist())
min_idx, max_idx = 0, len(meses)-1
if max_idx < 0:
    st.warning("Sem competências válidas.")
    st.stop()

file_sig = getattr(up, "name", "arquivo_sem_nome")
slider_key = f"mes_idx::{file_sig}::{len(meses)}"

for k in list(st.session_state.keys()):
    if k.startswith("mes_idx::") and k != slider_key:
        del st.session_state[k]

def make_month_idx_slider(lo, hi, rng, key):
    if rng:
        return st.slider("Competência (período)", min_value=lo, max_value=hi, value=(lo, hi), key=key)
    return st.slider("Competência (período)", min_value=lo, max_value=hi, value=lo, key=key)

is_range = max_idx > min_idx
try:
    sel = make_month_idx_slider(min_idx, max_idx, is_range, slider_key)
except Exception:
    if slider_key in st.session_state: del st.session_state[slider_key]
    sel = make_month_idx_slider(min_idx, max_idx, is_range, slider_key)

start_idx, end_idx = (sel if is_range else (sel, sel))
start_ym, end_ym = meses[start_idx], meses[end_idx]

# Aplica filtros
mask = (
    df["Empresa"].isin(f_emp)
    & df["AnoMes"].between(start_ym, end_ym)
    & df["Natureza"].isin(f_nat)
    & df["CentroCusto"].isin(f_cc)
)
df_f = df.loc[mask].copy()

# ================= KPIs =================
colA,colB,colC,colD = st.columns(4)
receita   = df_f.loc[df_f["Natureza"]=="Receita","SaldoGerencial"].sum()
despesa   = df_f.loc[df_f["Natureza"]=="Despesa","SaldoGerencial"].sum()
resultado = receita + despesa
margem    = (resultado / receita) if receita else np.nan
with colA: st.metric("Receita",   fmt(receita))
with colB: st.metric("Despesa",   fmt(despesa))
with colC: st.metric("Resultado", fmt(resultado))
with colD: st.metric("Margem %",  fmt((margem*100) if np.isfinite(margem) else 0))
st.markdown("---")

# ================= Gráficos =================
c1 = st.container(); c2 = st.container(); c3 = st.container(); c4 = st.container()

with c1:
    st.subheader("📈 Resultado por Mês")
    serie = df_f.groupby("AnoMes", as_index=False)["SaldoGerencial"].sum().sort_values("AnoMes")
    if not serie.empty:
        st.plotly_chart(px.line(serie, x="AnoMes", y="SaldoGerencial", markers=True),
                        use_container_width=True, theme="plotly")
    else:
        st.info("Sem dados no período.")

with c2:
    st.subheader("📊 Despesas por Centro de Custo")
    dep = df_f[df_f["Natureza"]=="Despesa"].groupby("CentroCusto", as_index=False)["SaldoGerencial"].sum()
    if not dep.empty:
        dep = dep.sort_values("SaldoGerencial")
        st.plotly_chart(px.bar(dep, x="SaldoGerencial", y="CentroCusto", orientation="h", color="CentroCusto"),
                        use_container_width=True, theme="plotly")
    else:
        st.info("Sem despesas nos filtros.")

with c3:
    st.subheader("🏆 Top 10 Receitas (por ContaDescricao)")
    rec = df_f[df_f["Natureza"]=="Receita"].groupby("ContaDescricao", as_index=False)["SaldoGerencial"].sum()
    if not rec.empty:
        rec = rec.sort_values("SaldoGerencial", ascending=False).head(10)
        st.plotly_chart(px.bar(rec, x="ContaDescricao", y="SaldoGerencial", color="ContaDescricao"),
                        use_container_width=True, theme="plotly")
    else:
        st.info("Sem receitas nos filtros.")

with c4:
    st.subheader("Tabela detalhada")
    cols = ["Empresa","Competencia","AnoMes","CentroCusto","ContaCodigo","ContaDescricao",
            "Natureza","Devedor","Credor","Saldo","Sinal","SaldoGerencial"]
    cols = [c for c in cols if c in df_f.columns]
    st.dataframe(df_f[cols].sort_values(["Competencia","ContaCodigo"]).reset_index(drop=True),
                 use_container_width=True, height=420)

# ================= Exportações =================
st.markdown("---")
st.subheader("Exportações")
pivot_mes = df_f.pivot_table(index="AnoMes", columns="Natureza", values="SaldoGerencial",
                             aggfunc="sum", fill_value=0).reset_index()
by_cc = df_f.groupby(["Natureza","CentroCusto"], as_index=False)["SaldoGerencial"] \
            .sum().sort_values(["Natureza","SaldoGerencial"], ascending=[True, False])

excel_bytes = to_excel_bytes({
    "Detalhado": df_f[cols],
    "Resumo_Mensal": pivot_mes,
    "Por_CentroCusto": by_cc
})
st.download_button("⬇️ Excel (Detalhado + Resumos)", data=excel_bytes,
                   file_name="analise_balancete.xlsx", key="dl_excel")

st.download_button("⬇️ CSV Detalhado", data=df_f.to_csv(index=False).encode("utf-8"),
                   file_name="balancete_detalhado.csv", key="dl_csv")
