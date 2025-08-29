# streamlit_app.py
from datetime import date
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Balancete (seu modelo)", page_icon="📘", layout="wide")
st.title("📘 Painel de Balancete — seu modelo (1 aba)")
st.caption("Importe .xlsx com colunas (ou equivalentes): Empresa, Competencia, ContaCodigo, ContaDescricao, CentroCusto, Devedor, Credor.")

# ======== util: normalização forte de cabeçalhos ========
import unicodedata, re
def _norm_token(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^a-zA-Z0-9]+", "", s).lower()
    return s

CANON = {
    "empresacnpj":"Empresa","empresa":"Empresa",
    "competencia":"Competencia","datacompetencia":"Competencia","mescompetencia":"Competencia","mesref":"Competencia",
    "contacodigo":"ContaCodigo","conta":"ContaCodigo","contacontabil":"ContaCodigo","codigoconta":"ContaCodigo",
    "contadescricao":"ContaDescricao","descricao":"ContaDescricao","historico":"ContaDescricao","descricaocta":"ContaDescricao",
    "centrocusto":"CentroCusto","centrodecusto":"CentroCusto","cc":"CentroCusto","setor":"CentroCusto",
    "devedor":"Devedor","debito":"Devedor","debitos":"Devedor","valordebito":"Devedor",
    "credor":"Credor","credito":"Credor","creditos":"Credor","valorcredito":"Credor",
    "saldo":"Saldo","valor":"Valor","total":"Total",
}

REQUIRED = ["Empresa","Competencia","ContaCodigo","ContaDescricao","CentroCusto","Devedor","Credor"]

def strong_rename(df: pd.DataFrame) -> pd.DataFrame:
    m = {}
    used = set()
    for c in df.columns:
        key = _norm_token(c)
        tgt = CANON.get(key)
        if tgt and tgt not in used:
            m[c] = tgt
            used.add(tgt)
    out = df.rename(columns=m).copy()
    # acertos simples
    if "Conta" in out.columns and "ContaCodigo" not in out.columns:
        out.rename(columns={"Conta":"ContaCodigo"}, inplace=True)
    if "Descrição" in out.columns and "ContaDescricao" not in out.columns:
        out.rename(columns={"Descrição":"ContaDescricao"}, inplace=True)
    return out

def coerce_required_or_fill(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Garante colunas mínimas; cria defaults quando necessário. Retorna (df, avisos)."""
    df = strong_rename(df)
    notes = []

    # Empresa
    if "Empresa" not in df.columns:
        df["Empresa"] = "Empresa"
        notes.append("Empresa criada como 'Empresa' (default).")

    # Competencia
    if "Competencia" not in df.columns:
        from datetime import date
        comp = date.today().replace(day=1)
        df["Competencia"] = pd.Timestamp(comp)
        notes.append("Competencia ausente: usado 1º dia do mês atual.")
    df["Competencia"] = pd.to_datetime(df["Competencia"], errors="coerce")
    if df["Competencia"].isna().all():
        comp = date.today().replace(day=1)
        df["Competencia"] = pd.Timestamp(comp)
        notes.append("Competencia inválida: normalizada para 1º dia do mês atual.")

    # ContaCodigo/ContaDescricao
    if "ContaCodigo" not in df.columns:
        raise ValueError("Não encontrei coluna de conta (Conta, ContaCódigo, ContaContábil...).")
    if "ContaDescricao" not in df.columns:
        df["ContaDescricao"] = df["ContaCodigo"].astype(str)
        notes.append("ContaDescricao ausente: copiado de ContaCodigo.")

    # CentroCusto
    if "CentroCusto" not in df.columns:
        df["CentroCusto"] = "Geral"
        notes.append("CentroCusto ausente: definido como 'Geral'.")

    # Devedor/Credor
    def _to_num(series):
        s = series.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        return pd.to_numeric(s, errors="coerce")

    if "Devedor" not in df.columns and "Credor" not in df.columns:
        cand = None
        for c in ["Saldo","Valor","Total"]:
            if c in df.columns:
                cand = c; break
        if cand is not None:
            v = _to_num(df[cand]).fillna(0.0)
            df["Devedor"] = np.where(v >= 0, v, 0.0)
            df["Credor"]  = np.where(v < 0, -v, 0.0)
            notes.append(f"Sem Devedor/Credor: derivado de '{cand}' (>=0 → Devedor; <0 → Credor).")
        else:
            df["Devedor"] = 0.0
            df["Credor"]  = 0.0
            notes.append("Sem Devedor/Credor/Saldo/Valor: criado Devedor=0 e Credor=0.")
    else:
        if "Devedor" not in df.columns: df["Devedor"] = 0.0; notes.append("Devedor ausente: definido 0.")
        if "Credor"  not in df.columns: df["Credor"]  = 0.0; notes.append("Credor ausente: definido 0.")

    # numéricos
    df["Devedor"] = _to_num(df["Devedor"]).fillna(0.0)
    df["Credor"]  = _to_num(df["Credor"]).fillna(0.0)

    # strings base
    for c in ["Empresa","ContaCodigo","ContaDescricao","CentroCusto"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    return df, notes

def to_excel_bytes(dfs: dict) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        for name, d in dfs.items():
            d.to_excel(w, index=False, sheet_name=name[:31] or "Dados")
    out.seek(0)
    return out

def money(v):
    try: return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    except: return str(v)

# ======== Sidebar (upload) ========
with st.sidebar:
    st.header("⚙️ Entrada")
    up = st.file_uploader("Envie .xlsx (1 aba)", type=["xlsx"], key="uploader")

if not up:
    st.info("Envie sua planilha .xlsx.")
    st.stop()

# ======== Leitura ========
xls = pd.ExcelFile(up)
sheet = xls.sheet_names[0]
raw = pd.read_excel(xls, sheet_name=sheet)

try:
    df, notes = coerce_required_or_fill(raw)
except Exception as e:
    st.error(f"Arquivo não pôde ser interpretado: {e}")
    st.write("Colunas encontradas:", list(raw.columns))
    st.stop()

if notes:
    st.warning("Ajustes aplicados automaticamente:\n- " + "\n- ".join(notes))

# ======== Cálculos ========
conta_str = df["ContaCodigo"].astype(str).str.strip()
df["Natureza"] = np.select(
    [conta_str.str.startswith("3"), conta_str.str.startswith("4")],
    ["Receita","Despesa"], default="Outros"
)
df["Sinal"] = np.select(
    [df["Natureza"].eq("Receita"), df["Natureza"].eq("Despesa")],
    [-1, 1], default=1
)

df["Saldo"] = df["Devedor"] - df["Credor"]
df["SaldoGerencial"] = df["Saldo"] * df["Sinal"]
df["AnoMes"] = pd.to_datetime(df["Competencia"], errors="coerce").dt.strftime("%Y-%m")

if df["AnoMes"].isna().all():
    today = date.today().replace(day=1)
    df["AnoMes"] = pd.Timestamp(today).strftime("%Y-%m")

# ======== Filtros ========
empresas = sorted(df["Empresa"].dropna().unique().tolist())
naturezas = sorted(df["Natureza"].dropna().unique().tolist()) or ["Receita","Despesa"]
ccs = sorted(df["CentroCusto"].dropna().unique().tolist())

colf1, colf2, colf3 = st.columns(3)
with colf1: f_emp = st.multiselect("Empresa", empresas, default=empresas)
with colf2: f_nat = st.multiselect("Natureza", naturezas, default=naturezas)
with colf3: f_cc  = st.multiselect("Centro de Custo", ccs, default=ccs)

# Slider por índice de mês (robusto a troca de arquivo)
meses = sorted(df["AnoMes"].dropna().unique().tolist())
min_idx, max_idx = 0, len(meses)-1
file_sig = getattr(up, "name", "arquivo")
slider_key = f"mes_idx::{file_sig}::{len(meses)}"

for k in list(st.session_state.keys()):
    if k.startswith("mes_idx::") and k != slider_key:
        del st.session_state[k]

def month_idx_slider(lo, hi, rng, key):
    return st.slider("Competência (período)", min_value=lo, max_value=hi,
                     value=(lo, hi) if rng else lo, key=key)

rng = max_idx > min_idx
try:
    sel = month_idx_slider(min_idx, max_idx, rng, slider_key)
except Exception:
    if slider_key in st.session_state: del st.session_state[slider_key]
    sel = month_idx_slider(min_idx, max_idx, rng, slider_key)

start_idx, end_idx = (sel if rng else (sel, sel))
start_ym, end_ym = meses[start_idx], meses[end_idx]

mask = (
    df["Empresa"].isin(f_emp)
    & df["AnoMes"].between(start_ym, end_ym)
    & df["Natureza"].isin(f_nat)
    & df["CentroCusto"].isin(f_cc)
)
df_f = df.loc[mask].copy()

# ======== KPIs ========
colA, colB, colC, colD = st.columns(4)
receita = df_f.loc[df_f["Natureza"]=="Receita","SaldoGerencial"].sum()
despesa = df_f.loc[df_f["Natureza"]=="Despesa","SaldoGerencial"].sum()
resultado = receita + despesa
margem = (resultado / receita) if receita else np.nan
with colA: st.metric("Receita", money(receita))
with colB: st.metric("Despesa", money(despesa))
with colC: st.metric("Resultado", money(resultado))
with colD: st.metric("Margem %", money((margem*100) if np.isfinite(margem) else 0))
st.markdown("---")

# ======== Gráficos ========
c1 = st.container(); c2 = st.container(); c3 = st.container(); c4 = st.container()

with c1:
    st.subheader("📈 Resultado por Mês")
    serie = df_f.groupby("AnoMes", as_index=False)["SaldoGerencial"].sum().sort_values("AnoMes")
    if not serie.empty:
        st.plotly_chart(px.line(serie, x="AnoMes", y="SaldoGerencial", markers=True),
                        use_container_width=True)
    else:
        st.info("Sem dados no período.")

with c2:
    st.subheader("📊 Despesas por Centro de Custo")
    dep = df_f[df_f["Natureza"]=="Despesa"].groupby("CentroCusto", as_index=False)["SaldoGerencial"].sum()
    if not dep.empty:
        dep = dep.sort_values("SaldoGerencial")
        st.plotly_chart(px.bar(dep, x="SaldoGerencial", y="CentroCusto", orientation="h", color="CentroCusto"),
                        use_container_width=True)
    else:
        st.info("Sem despesas nos filtros.")

with c3:
    st.subheader("🏆 Top 10 Receitas (por ContaDescricao)")
    rec = df_f[df_f["Natureza"]=="Receita"].groupby("ContaDescricao", as_index=False)["SaldoGerencial"].sum()
    if not rec.empty:
        rec = rec.sort_values("SaldoGerencial", ascending=False).head(10)
        st.plotly_chart(px.bar(rec, x="ContaDescricao", y="SaldoGerencial", color="ContaDescricao"),
                        use_container_width=True)
    else:
        st.info("Sem receitas nos filtros.")

with c4:
    st.subheader("Tabela detalhada")
    cols = ["Empresa","Competencia","AnoMes","CentroCusto","ContaCodigo","ContaDescricao",
            "Natureza","Devedor","Credor","Saldo","Sinal","SaldoGerencial"]
    cols = [c for c in cols if c in df_f.columns]
    st.dataframe(df_f[cols].sort_values(["Competencia","ContaCodigo"]).reset_index(drop=True),
                 use_container_width=True, height=420)

# ======== Exportações ========
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
