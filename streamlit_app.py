# app.py
# -----------------------------------------------------------------------------
# Painel de Balancete (arquivo único)
# Rodar:  pip install streamlit pandas numpy plotly xlsxwriter openpyxl
#         streamlit run app.py
#
# Entrada: .xlsx OU .zip contendo um .xlsx
# Abas necessárias no Excel:
#   - Balancete: Empresa, Competencia, ContaCodigo, ContaDescricao, CentroCusto, Devedor, Credor
#   - Mapa_Classificacao: ContaPrefixo, Natureza, GrupoGerencial, Subgrupo, Sinal, TipoOperacional
#
# O app:
#  - Normaliza nomes de colunas (alias), corrige vírgula decimal pt-BR.
#  - Faz "merge" por prefixo do código contábil (3 → 2 → 1 níveis).
#  - KPIs: Receita, Despesa, Resultado, Margem %.
#  - Gráficos: Resultado por mês, Despesas por Grupo, Top 10 Receitas por Subgrupo.
#  - Exporta Excel (Detalhado + Resumos) e CSV.
# -----------------------------------------------------------------------------

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import zipfile, io
import plotly.express as px

# -----------------------------
# Config da página
# -----------------------------
st.set_page_config(page_title="Painel Balancete (Único)", page_icon="📘", layout="wide")
st.title("📘 Painel de Balancete — arquivo único")
st.caption("Envie um .xlsx (ou .zip com .xlsx) contendo as abas **Balancete** e **Mapa_Classificacao**. Opcionalmente, use dados de exemplo para testar.")

# -----------------------------
# Helpers
# -----------------------------
def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza nomes de colunas e aplica alguns aliases comuns."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    aliases = {
        "Conta": "ContaCodigo",
        "Conta Código": "ContaCodigo",
        "Descrição": "ContaDescricao",
        "Descricao": "ContaDescricao",
        "DataCompetencia": "Competencia",
        "Competência": "Competencia",
        "Centro de Custo": "CentroCusto",
    }
    for a, b in aliases.items():
        if a in df.columns and b not in df.columns:
            df.rename(columns={a: b}, inplace=True)
    return df

def split_prefix(code, n_segments: int):
    """Pega os primeiros N níveis do código contábil. Ex.: 4.1.2.01 -> n=3 -> 4.1.2"""
    if pd.isna(code):
        return None
    parts = [p for p in str(code).split(".") if p != ""]
    if not parts:
        return None
    return ".".join(parts[: min(n_segments, len(parts))])

def read_excel_like(uploaded, sheet_bal="Balancete", sheet_map="Mapa_Classificacao"):
    """Lê .xlsx direto ou .zip contendo um .xlsx. Retorna (balancete, mapa)."""
    def _read_xlsx(flike):
        xls = pd.ExcelFile(flike)
        bal = pd.read_excel(xls, sheet_name=sheet_bal)
        mapa = pd.read_excel(xls, sheet_name=sheet_map)
        return bal, mapa

    # aceita zip com xlsx dentro
    if hasattr(uploaded, "name") and str(uploaded.name).lower().endswith(".zip"):
        with zipfile.ZipFile(uploaded) as z:
            xlsx_names = [n for n in z.namelist() if n.lower().endswith(".xlsx")]
            if not xlsx_names:
                raise ValueError("ZIP sem arquivo .xlsx dentro.")
            with z.open(xlsx_names[0]) as xf:
                data = xf.read()
            bal, mapa = _read_xlsx(io.BytesIO(data))
    else:
        # xlsx direto
        bal, mapa = _read_xlsx(uploaded)

    # normaliza
    bal = _norm_cols(bal)
    mapa = _norm_cols(mapa)

    # datas e números pt-BR
    if "Competencia" in bal.columns:
        bal["Competencia"] = pd.to_datetime(bal["Competencia"], errors="coerce")

    for col in ["Devedor", "Credor"]:
        if col in bal.columns:
            bal[col] = (
                bal[col].astype(str)
                .str.replace(".", "", regex=False)     # remove milhar
                .str.replace(",", ".", regex=False)    # vírgula -> ponto
            )
            bal[col] = pd.to_numeric(bal[col], errors="coerce").fillna(0.0)

    if "Sinal" in mapa.columns:
        mapa["Sinal"] = (
            mapa["Sinal"].astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        mapa["Sinal"] = pd.to_numeric(mapa["Sinal"], errors="coerce").fillna(1.0)
    else:
        mapa["Sinal"] = 1.0

    # colunas mínimas
    need_bal = {"Empresa","Competencia","ContaCodigo","ContaDescricao","Devedor","Credor"}
    miss = need_bal - set(bal.columns)
    if miss:
        raise ValueError(f"Planilha Balancete faltando colunas obrigatórias: {miss}")

    # mapa mínimo
    for c in ["ContaPrefixo","Natureza","GrupoGerencial","Subgrupo","Sinal","TipoOperacional"]:
        if c not in mapa.columns:
            mapa[c] = np.nan
    mapa["Sinal"] = mapa["Sinal"].fillna(1.0)

    return bal, mapa

def merge_classify(bal: pd.DataFrame, mapa: pd.DataFrame) -> pd.DataFrame:
    """Casa o balancete com o mapa por prefixo (3 -> 2 -> 1) e cria métricas."""
    df = bal.copy()
    df["prefix3"] = df["ContaCodigo"].apply(lambda x: split_prefix(x, 3))
    df["prefix2"] = df["ContaCodigo"].apply(lambda x: split_prefix(x, 2))
    df["prefix1"] = df["ContaCodigo"].apply(lambda x: split_prefix(x, 1))

    m3 = df.merge(mapa.add_prefix("m3_"), left_on="prefix3", right_on="m3_ContaPrefixo", how="left")
    m2 = df.merge(mapa.add_prefix("m2_"), left_on="prefix2", right_on="m2_ContaPrefixo", how="left")
    m1 = df.merge(mapa.add_prefix("m1_"), left_on="prefix1", right_on="m1_ContaPrefixo", how="left")

    def coalesce(*cols):
        out = cols[0].copy()
        for c in cols[1:]:
            out = out.where(~out.isna(), c)
        return out

    out = df.copy()
    out["Natureza"]        = coalesce(m3.get("m3_Natureza"),        m2.get("m2_Natureza"),        m1.get("m1_Natureza"))
    out["GrupoGerencial"]  = coalesce(m3.get("m3_GrupoGerencial"),  m2.get("m2_GrupoGerencial"),  m1.get("m1_GrupoGerencial"))
    out["Subgrupo"]        = coalesce(m3.get("m3_Subgrupo"),        m2.get("m2_Subgrupo"),        m1.get("m1_Subgrupo"))
    out["Sinal"]           = coalesce(m3.get("m3_Sinal"),           m2.get("m2_Sinal"),           m1.get("m1_Sinal")).fillna(1.0)
    out["TipoOperacional"] = coalesce(m3.get("m3_TipoOperacional"), m2.get("m2_TipoOperacional"), m1.get("m1_TipoOperacional"))

    # métricas
    out["Saldo"] = out["Devedor"] - out["Credor"]
    out["Sinal"] = pd.to_numeric(out["Sinal"], errors="coerce").fillna(1.0)
    out["SaldoGerencial"] = out["Saldo"] * out["Sinal"]
    out["Competencia"] = pd.to_datetime(out["Competencia"], errors="coerce")
    out["AnoMes"] = out["Competencia"].dt.strftime("%Y-%m")

    return out

def metric_fmt(v):
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    except:
        return str(v)

def to_excel_bytes(dfs: dict) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as w:
        for name, d in dfs.items():
            d.to_excel(w, index=False, sheet_name=name[:31] or "Dados")
    output.seek(0)
    return output

@st.cache_data
def sample_excel_bytes() -> bytes:
    """Gera um Excel de exemplo em memória (para teste rápido)."""
    months = pd.period_range("2025-01", "2025-08", freq="M")
    competencias = [pd.Timestamp(m.start_time.date()) for m in months]
    rows = []
    for comp in competencias:
        entries = [
            ("3.1.1.01","Receita de Locação","Geral",0,1500000),
            ("3.1.1.02","Receita de Fretes","Geral",0,200000),
            ("3.1.1.03","Serviços Acessórios","Geral",0,50000),
            ("4.1.1.01","Despesas com Manutenção","Operação",74000,0),
            ("4.1.2.01","Despesas com Pessoal","RH",2200000,0),
            ("4.1.3.01","Despesas Administrativas","ADM",90000,0),
            ("1.1.1.01","Caixa","Geral",120000,0),
            ("2.1.1.01","Fornecedores","Geral",0,60000),
        ]
        for cod, desc, cc, dev, cre in entries:
            rows.append(["Empresa Modelo", comp.date().isoformat(), cod, desc, cc, float(dev), float(cre)])

    df_bal = pd.DataFrame(rows, columns=["Empresa","Competencia","ContaCodigo","ContaDescricao","CentroCusto","Devedor","Credor"])
    df_map = pd.DataFrame([
        ["3.1.1", "Receita", "Receitas Operacionais", "Locação/Fretes", -1, "Operacional", ""],
        ["4.1.1", "Despesa", "Despesas Operacionais", "Manutenção", 1, "Operacional", ""],
        ["4.1.2", "Despesa", "Despesas Operacionais", "Pessoal", 1, "Operacional", ""],
        ["4.1.3", "Despesa", "Despesas Operacionais", "Administrativas", 1, "Operacional", ""],
        ["1.1.1", "Ativo", "Ativo Circulante", "Disponibilidades", 1, "Não Operacional", ""],
        ["2.1.1", "Passivo", "Passivo Circulante", "Fornecedores", -1, "Não Operacional", ""],
    ], columns=["ContaPrefixo","Natureza","GrupoGerencial","Subgrupo","Sinal","TipoOperacional","Observacao"])

    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        df_bal.to_excel(w, index=False, sheet_name="Balancete")
        df_map.to_excel(w, index=False, sheet_name="Mapa_Classificacao")
    out.seek(0)
    return out.read()

# -----------------------------
# Sidebar (upload e opções)
# -----------------------------
with st.sidebar:
    st.header("⚙️ Entrada")
    use_sample = st.toggle("Usar dados de exemplo (sem enviar arquivo)", value=False)
    up = None
    if not use_sample:
        up = st.file_uploader("Envie .xlsx ou .zip com .xlsx", type=["xlsx","zip"])
    sheet_bal = st.text_input("Aba do Balancete", "Balancete")
    sheet_map = st.text_input("Aba do Mapa", "Mapa_Classificacao")
    st.download_button("⬇️ Baixar Excel de Exemplo", data=sample_excel_bytes(), file_name="Exemplo_Balancete.xlsx")

# -----------------------------
# Carregar dados
# -----------------------------
if use_sample:
    try:
        bal, mapa = read_excel_like(io.BytesIO(sample_excel_bytes()), sheet_bal, sheet_map)
    except Exception as e:
        st.exception(e)
        st.stop()
else:
    if not up:
        st.info("Envie um arquivo ou ative **Usar dados de exemplo** na barra lateral.")
        st.stop()
    try:
        bal, mapa = read_excel_like(up, sheet_bal, sheet_map)
    except Exception as e:
        st.exception(e)
        st.stop()

df = merge_classify(bal, mapa)

# -----------------------------
# Filtros
# -----------------------------
empresas = sorted(df["Empresa"].dropna().unique().tolist())
naturezas = sorted(df["Natureza"].dropna().unique().tolist()) if "Natureza" in df.columns else []
cc_list = sorted(df["CentroCusto"].dropna().unique().tolist()) if "CentroCusto" in df.columns else []

colf1, colf2, colf3 = st.columns(3)
with colf1:
    f_emp = st.multiselect("Empresa", empresas, default=empresas)
with colf2:
    f_nat = st.multiselect("Natureza", naturezas or ["Receita","Despesa"], default=naturezas or ["Receita","Despesa"])
with colf3:
    f_cc = st.multiselect("Centro de Custo", cc_list, default=cc_list) if cc_list else None

min_date = df["Competencia"].min().date()
max_date = df["Competencia"].max().date()
f_date = st.slider("Competência (período)", min_value=min_date, max_value=max_date, value=(min_date, max_date))

mask = (
    df["Empresa"].isin(f_emp) &
    df["Competencia"].between(pd.to_datetime(f_date[0]), pd.to_datetime(f_date[1])) &
    ((df["Natureza"].isin(f_nat)) if df["Natureza"].notna().any() else True)
)
if f_cc is not None:
    mask &= df["CentroCusto"].isin(f_cc)

df_f = df.loc[mask].copy()

# -----------------------------
# KPIs
# -----------------------------
colA, colB, colC, colD = st.columns(4)
receita = df_f.loc[df_f["Natureza"]=="Receita","SaldoGerencial"].sum()
despesa = df_f.loc[df_f["Natureza"]=="Despesa","SaldoGerencial"].sum()
resultado = receita + despesa
margem = (resultado / receita) if receita else np.nan
with colA: st.metric("Receita", metric_fmt(receita))
with colB: st.metric("Despesa", metric_fmt(despesa))
with colC: st.metric("Resultado", metric_fmt(resultado))
with colD: st.metric("Margem %", metric_fmt((margem*100) if pd.notna(margem) else 0))

st.markdown("---")

# -----------------------------
# Gráficos
# -----------------------------
st.subheader("📈 Resultado por mês (Saldo Gerencial)")
series = df_f.groupby("AnoMes", as_index=False)["SaldoGerencial"].sum().sort_values("AnoMes")
if not series.empty:
    st.plotly_chart(px.line(series, x="AnoMes", y="SaldoGerencial", markers=True), use_container_width=True)
else:
    st.info("Sem dados no período.")

st.subheader("📊 Despesas por Grupo")
dep = df_f[df_f["Natureza"]=="Despesa"].groupby("GrupoGerencial", as_index=False)["SaldoGerencial"].sum()
if not dep.empty:
    dep = dep.sort_values("SaldoGerencial")
    st.plotly_chart(px.bar(dep, x="SaldoGerencial", y="GrupoGerencial", orientation="h"), use_container_width=True)
else:
    st.info("Sem dados de despesa nos filtros atuais.")

st.subheader("🏆 Top 10 Receitas por Subgrupo")
rec = df_f[df_f["Natureza"]=="Receita"].groupby("Subgrupo", as_index=False)["SaldoGerencial"].sum()
if not rec.empty:
    rec = rec.sort_values("SaldoGerencial", ascending=False).head(10)
    st.plotly_chart(px.bar(rec, x="Subgrupo", y="SaldoGerencial"), use_container_width=True)
else:
    st.info("Sem dados de receita nos filtros atuais.")

st.markdown("---")

# -----------------------------
# Tabela + Exportações
# -----------------------------
st.subheader("Tabela detalhada")
show_cols = [
    "Empresa","Competencia","ContaCodigo","ContaDescricao","CentroCusto",
    "Natureza","GrupoGerencial","Subgrupo","Devedor","Credor",
    "Saldo","Sinal","SaldoGerencial","AnoMes"
]
show_cols = [c for c in show_cols if c in df_f.columns]
st.dataframe(df_f[show_cols].sort_values(["Competencia","ContaCodigo"]).reset_index(drop=True), use_container_width=True)

st.subheader("Exportações")
pivot_mes = df_f.pivot_table(index="AnoMes", columns="Natureza", values="SaldoGerencial", aggfunc="sum", fill_value=0).reset_index()
by_grupo = df_f.groupby(["Natureza","GrupoGerencial"], as_index=False)["SaldoGerencial"].sum().sort_values(["Natureza","SaldoGerencial"], ascending=[True, False])

excel_bytes = to_excel_bytes({
    "Detalhado": df_f[show_cols],
    "Resumo_Mensal": pivot_mes,
    "Por_Grupo": by_grupo
})
st.download_button("⬇️ Excel (Detalhado + Resumos)", data=excel_bytes, file_name="analise_balancete.xlsx")
st.download_button("⬇️ CSV Detalhado", data=df_f.to_csv(index=False).encode("utf-8"), file_name="balancete_detalhado.csv")
