# streamlit_app.py
from datetime import date
from io import BytesIO
import unicodedata, re, os

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# 1ª chamada do Streamlit
st.set_page_config(page_title="Balancete (seu modelo)", page_icon="📘", layout="wide")
st.title("📘 Painel de Balancete — seu modelo (1 aba)")
st.caption("Importe .xlsx com colunas (ou equivalentes): Empresa, Competencia, ContaCodigo, ContaDescricao, CentroCusto, Devedor, Credor.")

# ===== Normalização de cabeçalhos (case-insensitive, sem acentos) =====
def _norm_token(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^a-zA-Z0-9]+", "", s).lower()
    return s

CANON = {
    "empresacnpj":"Empresa","empresa":"Empresa",
    "competencia":"Competencia","datacompetencia":"Competencia","mescompetencia":"Competencia","mesref":"Competencia",
    "mes":"Mes","ano":"Ano",
    "contacodigo":"ContaCodigo","conta":"ContaCodigo","contacontabil":"ContaCodigo","codigoconta":"ContaCodigo",
    "contadescricao":"ContaDescricao","descricao":"ContaDescricao","historico":"ContaDescricao","descricaocta":"ContaDescricao",
    "centrocusto":"CentroCusto","centrodecusto":"CentroCusto","cc":"CentroCusto","setor":"CentroCusto",
    "devedor":"Devedor","debito":"Devedor","debitos":"Devedor","valordebito":"Devedor",
    "credor":"Credor","credito":"Credor","creditos":"Credor","valorcredito":"Credor",
    "saldo":"Saldo","valor":"Valor","total":"Total",
}

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
    if "Conta" in out.columns and "ContaCodigo" not in out.columns:
        out.rename(columns={"Conta":"ContaCodigo"}, inplace=True)
    if "Descrição" in out.columns and "ContaDescricao" not in out.columns:
        out.rename(columns={"Descrição":"ContaDescricao"}, inplace=True)
    return out

# ===== Conversão numérica robusta (BR/US) — sem explodir valores =====
def to_num_safe(series: pd.Series) -> pd.Series:
    # Se já é numérico, não mexe.
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    s = series.astype(str)

    # Remove símbolos não numéricos exceto . , - 
    s = s.str.replace(r"[^\d,.\-]", "", regex=True)

    # Heurística de separador decimal:
    def _parse_one(x: str) -> float:
        if x == "" or x == "-" or x == "--":
            return np.nan
        # Casos com vírgula como decimal: 1.234.567,89
        if "," in x and (x.rfind(",") > x.rfind(".")):
            x2 = x.replace(".", "").replace(",", ".")
            try: return float(x2)
            except: pass
        # Só vírgula (decimal BR): 123,45
        if "," in x and "." not in x:
            try: return float(x.replace(",", "."))
            except: pass
        # Padrão US (123,456.78 ou 1234.56)
        try:
            return float(x.replace(",", ""))
        except:
            return np.nan

    return s.map(_parse_one)

# ===== Inferência de Competencia quando ausente =====
def infer_competencia(df: pd.DataFrame, up_name: str|None) -> pd.Series:
    # 1) Se já existe Competencia, converte e devolve
    if "Competencia" in df.columns:
        comp = pd.to_datetime(df["Competencia"], errors="coerce")
        if comp.notna().any():
            return comp

    # 2) Mes/Ano -> primeiro dia do mês
    if {"Mes","Ano"}.issubset(df.columns):
        try:
            mes = pd.to_numeric(df["Mes"], errors="coerce").fillna(1).astype(int).clip(1,12)
            ano = pd.to_numeric(df["Ano"], errors="coerce").fillna(date.today().year).astype(int)
            comp = pd.to_datetime(dict(year=ano, month=mes, day=1), errors="coerce")
            if comp.notna().any():
                return comp
        except:
            pass

    # 3) Nome do arquivo: 08-2025 ou 2025-08
    if up_name:
        m = re.search(r"(?:(\d{2})[-_\.](\d{4}))|(?:(\d{4})[-_\.](\d{2}))", up_name)
        if m:
            if m.group(1) and m.group(2):   # MM-YYYY
                mm, yy = int(m.group(1)), int(m.group(2))
            else:                            # YYYY-MM
                yy, mm = int(m.group(3)), int(m.group(4))
            try:
                return pd.Series(pd.Timestamp(year=yy, month=mm, day=1), index=df.index)
            except:
                pass

    # 4) Fallback → mês atual (1º dia)
    return pd.Series(pd.Timestamp(date.today().replace(day=1)), index=df.index)

# ===== Preenchimento das colunas obrigatórias =====
def coerce_required_or_fill(df: pd.DataFrame, up_name: str|None) -> tuple[pd.DataFrame, list[str]]:
    df = strong_rename(df)
    notes = []

    if "Empresa" not in df.columns:
        df["Empresa"] = "Empresa"
        notes.append("Empresa criada como 'Empresa' (default).")

    # Competencia (com inferência por Mes/Ano/Arquivo)
    comp = infer_competencia(df, up_name)
    if "Competencia" not in df.columns:
        notes.append("Competencia ausente: inferida (Mes/Ano ou nome do arquivo; senão mês atual).")
    elif comp.isna().all():
        notes.append("Competencia inválida: normalizada (Mes/Ano ou nome do arquivo; senão mês atual).")
    df["Competencia"] = comp

    if "ContaCodigo" not in df.columns:
        raise ValueError("Não encontrei coluna de conta (Conta, ContaCódigo, ContaContábil...).")
    if "ContaDescricao" not in df.columns:
        df["ContaDescricao"] = df["ContaCodigo"].astype(str)
        notes.append("ContaDescricao ausente: copiado de ContaCodigo.")
    if "CentroCusto" not in df.columns:
        df["CentroCusto"] = "Geral"
        notes.append("CentroCusto ausente: definido como 'Geral'.")

    # Devedor/Credor com derivação de Saldo/Valor/Total
    if "Devedor" not in df.columns and "Credor" not in df.columns:
        cand = next((c for c in ["Saldo","Valor","Total"] if c in df.columns), None)
        if cand is not None:
            v = to_num_safe(df[cand]).fillna(0.0)
            df["Devedor"] = np.where(v >= 0, v, 0.0)
            df["Credor"]  = np.where(v < 0, -v, 0.0)
            notes.append(f"Sem Devedor/Credor: derivado de '{cand}' (>=0 → Devedor; <0 → Credor).")
        else:
            df["Devedor"] = 0.0
            df["Credor"]  = 0.0
            notes.append("Sem Devedor/Credor/Saldo/Valor: criado Devedor=0 e Credor=0.")
    else:
        if "Devedor" not in df.columns:
            df["Devedor"] = 0.0; notes.append("Devedor ausente: definido 0.")
        if "Credor"  not in df.columns:
            df["Credor"]  = 0.0; notes.append("Credor ausente: definido 0.")

    # Tipagem numérica segura
    df["Devedor"] = to_num_safe(df["Devedor"]).fillna(0.0)
    df["Credor"]  = to_num_safe(df["Credor"]).fillna(0.0)

    # Strings base
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

# ===== Sidebar (upload) =====
with st.sidebar:
    st.header("⚙️ Entrada")
    up = st.file_uploader("Envie .xlsx (1 aba)", type=["xlsx"], key="uploader")

if not up:
    st.info("Envie sua planilha .xlsx.")
    st.stop()

# ===== Leitura =====
xls = pd.ExcelFile(up)
sheet = xls.sheet_names[0]
raw = pd.read_excel(xls, sheet_name=sheet)

try:
    df, notes = coerce_required_or_fill(raw, getattr(up, "name", None))
except Exception as e:
    st.error(f"Arquivo não pôde ser interpretado: {e}")
    st.write("Colunas encontradas:", list(raw.columns))
    st.stop()

if notes:
    st.warning("Ajustes aplicados automaticamente:\n- " + "\n- ".join(notes))

# ===== Cálculos =====
# Natureza: prefixo 3/4, senão palavras-chave, senão pelo sinal
conta_str = df["ContaCodigo"].astype(str).str.strip()
desc_str  = df["ContaDescricao"].astype(str).str.lower()

natureza = np.select(
    [conta_str.str.startswith("3"), conta_str.str.startswith("4")],
    ["Receita", "Despesa"],
    default="Outros"
)

# Palavras-chave se ainda ficou "Outros"
mask_out = natureza == "Outros"
if mask_out.any():
    kw_rec = desc_str.str.contains(r"receit|fatur|venda|renda|loca", regex=True)
    kw_desp = desc_str.str.contains(r"despes|custo|impost|taxa|encargo|manuten|pessoal|administr", regex=True)
    natureza = np.where(mask_out & kw_rec, "Receita", natureza)
    natureza = np.where((natureza == "Outros") & kw_desp, "Despesa", natureza)

# Se mesmo assim não classificou nada como Receita/Despesa, usa o sinal do valor
if not (np.isin(natureza, ["Receita","Despesa"]).any()):
    valor = df["Devedor"] - df["Credor"]
    natureza = np.where(valor < 0, "Receita", "Despesa")

df["Natureza"] = natureza
df["Sinal"] = np.select([df["Natureza"].eq("Receita"), df["Natureza"].eq("Despesa")], [-1, 1], default=1)

df["Saldo"] = df["Devedor"] - df["Credor"]
df["SaldoGerencial"] = df["Saldo"] * df["Sinal"]

df["Competencia"] = pd.to_datetime(df["Competencia"], errors="coerce")
df["AnoMes"] = df["Competencia"].dt.strftime("%Y-%m")
if df["AnoMes"].isna().all():
    df["AnoMes"] = pd.Timestamp(date.today().replace(day=1)).strftime("%Y-%m")

# ===== Filtros =====
empresas = sorted(df["Empresa"].dropna().unique().tolist())
naturezas = sorted(df["Natureza"].dropna().unique().tolist()) or ["Receita","Despesa"]
ccs = sorted(df["CentroCusto"].dropna().unique().tolist())

colf1, colf2, colf3 = st.columns(3)
with colf1: f_emp = st.multiselect("Empresa", empresas, default=empresas)
with colf2: f_nat = st.multiselect("Natureza", naturezas, default=[n for n in naturezas if n!="Outros"] or naturezas)
with colf3: f_cc  = st.multiselect("Centro de Custo", ccs, default=ccs)

# ===== Competência: 1 mês → sem slider; 2+ meses → slider por índice =====
meses = sorted(df["AnoMes"].dropna().unique().tolist())
if not meses:
    st.warning("Sem competências válidas."); st.stop()

if len(meses) == 1:
    start_ym = end_ym = meses[0]
    st.info(f"Competência única: **{start_ym}**")
else:
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

# ===== Aplica filtros =====
mask = (
    df["Empresa"].isin(f_emp)
    & df["AnoMes"].between(start_ym, end_ym)
    & df["Natureza"].isin(f_nat)
    & df["CentroCusto"].isin(f_cc)
)
df_f = df.loc[mask].copy()

# ===== KPIs =====
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

# ===== Gráficos =====
c1 = st.container(); c2 = st.container(); c3 = st.container(); c4 = st.container()

with c1:
    st.subheader("📈 Resultado por Mês")
    serie = df_f.groupby("AnoMes", as_index=False)["SaldoGerencial"].sum().sort_values("AnoMes")
    if not serie.empty:
        st.plotly_chart(px.line(serie, x="AnoMes", y="SaldoGerencial", markers=True), use_container_width=True)
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
    # Formatação amigável (sem notação científica)
    styled = df_f[cols].sort_values(["Competencia","ContaCodigo"]).reset_index(drop=True) \
                       .style.format({c: "{:,.2f}".format for c in ["Devedor","Credor","Saldo","SaldoGerencial"] if c in cols})
    st.dataframe(styled, use_container_width=True, height=420)

# ===== Exportações =====
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
