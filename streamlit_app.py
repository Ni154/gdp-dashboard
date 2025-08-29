# streamlit_app.py
from datetime import date
from io import BytesIO
import unicodedata, re

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ───────────────────────────── Page config ─────────────────────────────
st.set_page_config(page_title="Análise de Balancete — Dashboard", page_icon="📊", layout="wide")
st.title("📊 Análise de Balancete — Dashboard")
st.caption(
    "Importe seu arquivo e vamos fazer a análise para melhor tomada de decisão. "
    "Formato esperado (ou equivalentes): **Empresa, Competencia, ContaCodigo, ContaDescricao, CentroCusto, Devedor, Credor**."
)

# ───────────────────── helpers: header, números, datas ─────────────────
def _norm_token(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^a-zA-Z0-9]+", "", s).lower()
    return s

def strip_accents_upper(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.upper().strip()

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
    m, used = {}, set()
    for c in df.columns:
        key = _norm_token(c)
        tgt = CANON.get(key)
        if tgt and tgt not in used:
            m[c] = tgt; used.add(tgt)
    out = df.rename(columns=m).copy()
    if "Conta" in out.columns and "ContaCodigo" not in out.columns:
        out.rename(columns={"Conta":"ContaCodigo"}, inplace=True)
    if "Descrição" in out.columns and "ContaDescricao" not in out.columns:
        out.rename(columns={"Descrição":"ContaDescricao"}, inplace=True)
    return out

def to_num_safe(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    s = series.astype(str)
    s = s.str.replace(r"[^\d,.\-]", "", regex=True)
    def _parse(x: str) -> float:
        if x in ("", "-", "--"): return np.nan
        if "," in x and (x.rfind(",") > x.rfind(".")):  # 1.234.567,89
            x2 = x.replace(".", "").replace(",", ".")
            try: return float(x2)
            except: pass
        if "," in x and "." not in x:  # 123,45
            try: return float(x.replace(",", "."))
            except: pass
        try:
            return float(x.replace(",", ""))  # 123,456.78 ou 1234.56
        except: return np.nan
    return s.map(_parse)

def infer_competencia(df: pd.DataFrame, up_name: str|None) -> pd.Series:
    if "Competencia" in df.columns:
        comp = pd.to_datetime(df["Competencia"], errors="coerce")
        if comp.notna().any(): return comp
    if {"Mes","Ano"}.issubset(df.columns):
        try:
            mes = pd.to_numeric(df["Mes"], errors="coerce").fillna(1).astype(int).clip(1,12)
            ano = pd.to_numeric(df["Ano"], errors="coerce").fillna(date.today().year).astype(int)
            comp = pd.to_datetime(dict(year=ano, month=mes, day=1), errors="coerce")
            if comp.notna().any(): return comp
        except: pass
    if up_name:
        m = re.search(r"(?:(\d{2})[-_\.](\d{4}))|(?:(\d{4})[-_\.](\d{2}))", up_name)
        if m:
            if m.group(1): mm, yy = int(m.group(1)), int(m.group(2))
            else:          yy, mm = int(m.group(3)), int(m.group(4))
            try:
                return pd.Series(pd.Timestamp(year=yy, month=mm, day=1), index=df.index)
            except: pass
    return pd.Series(pd.Timestamp(date.today().replace(day=1)), index=df.index)

def money(v):
    try: return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    except: return str(v)

def to_excel_bytes(dfs: dict) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        for name, d in dfs.items():
            d.to_excel(w, index=False, sheet_name=name[:31] or "Dados")
    out.seek(0); return out

# ───────────────────────── Persistência do upload ──────────────────────
with st.sidebar:
    st.header("📥 Importação")
    st.caption("**Importe seu arquivo aqui** e vamos fazer a análise para melhor tomada de decisão.")
    up_new = st.file_uploader("Arraste e solte o .xlsx (1 aba)", type=["xlsx"], key="uploader")
    colb1, colb2 = st.columns(2)
    with colb1:
        clear_btn = st.button("🧹 Trocar arquivo", use_container_width=True)
    with colb2:
        analyze_btn = st.button("🔎 Recalcular", use_container_width=True)

if clear_btn:
    for k in ["file_bytes","file_name"]: st.session_state.pop(k, None)

if up_new is not None:
    st.session_state["file_bytes"] = up_new.read()
    st.session_state["file_name"] = getattr(up_new, "name", "arquivo.xlsx")

if "file_bytes" not in st.session_state:
    st.info("Envie sua planilha .xlsx.")
    st.stop()

file_bytes = st.session_state["file_bytes"]
file_name  = st.session_state.get("file_name", "arquivo.xlsx")

# ─────────────────────────── Leitura + preparo ─────────────────────────
@st.cache_data(show_spinner=True)
def load_dataframe_from_bytes(b: bytes, fname: str):
    xls = pd.ExcelFile(BytesIO(b))
    sheet = xls.sheet_names[0]
    raw = pd.read_excel(xls, sheet_name=sheet)
    df = strong_rename(raw)

    notes = []
    if "Empresa" not in df.columns:
        df["Empresa"] = "Empresa"; notes.append("Empresa criada como 'Empresa' (default).")

    comp = infer_competencia(df, fname)
    if "Competencia" not in df.columns:
        notes.append("Competencia ausente: inferida (Mes/Ano ou nome do arquivo; senão mês atual).")
    elif comp.isna().all():
        notes.append("Competencia inválida: normalizada (Mes/Ano ou nome do arquivo; senão mês atual).")
    df["Competencia"] = comp

    if "ContaCodigo" not in df.columns:
        raise ValueError("Não encontrei coluna de conta (Conta/ContaCódigo/ContaContábil).")
    if "ContaDescricao" not in df.columns:
        df["ContaDescricao"] = df["ContaCodigo"].astype(str); notes.append("ContaDescricao copiada de ContaCodigo.")
    if "CentroCusto" not in df.columns:
        df["CentroCusto"] = "Geral"; notes.append("CentroCusto ausente: 'Geral'.")

    # valores
    if "Devedor" not in df.columns and "Credor" not in df.columns:
        cand = next((c for c in ["Saldo","Valor","Total"] if c in df.columns), None)
        if cand:
            v = to_num_safe(df[cand]).fillna(0.0)
            df["Devedor"] = np.where(v >= 0, v, 0.0)
            df["Credor"]  = np.where(v < 0, -v, 0.0)
            notes.append(f"Sem Devedor/Credor: derivado de '{cand}'.")
        else:
            df["Devedor"] = 0.0; df["Credor"] = 0.0
            notes.append("Sem Devedor/Credor/Saldo/Valor: criado Devedor=0 e Credor=0.")
    else:
        if "Devedor" not in df.columns: df["Devedor"] = 0.0; notes.append("Devedor ausente: 0.")
        if "Credor"  not in df.columns: df["Credor"]  = 0.0; notes.append("Credor ausente: 0.")

    df["Devedor"] = to_num_safe(df["Devedor"]).fillna(0.0)
    df["Credor"]  = to_num_safe(df["Credor"]).fillna(0.0)

    # Natureza
    conta_str = df["ContaCodigo"].astype(str).str.strip()
    desc_str  = df["ContaDescricao"].astype(str).str.lower()
    natureza = np.select(
        [conta_str.str.startswith("3"), conta_str.str.startswith("4")],
        ["Receita","Despesa"], default="Outros"
    )
    mask_out = natureza == "Outros"
    if mask_out.any():
        kw_rec  = desc_str.str.contains(r"receit|fatur|venda|renda|loca", regex=True)
        kw_desp = desc_str.str.contains(r"despes|custo|impost|taxa|encargo|manuten|pessoal|administr", regex=True)
        natureza = np.where(mask_out & kw_rec,  "Receita", natureza)
        natureza = np.where((natureza=="Outros") & kw_desp, "Despesa", natureza)
    if not (np.isin(natureza, ["Receita","Despesa"]).any()):  # fallback pelo sinal
        valor = (df["Devedor"] - df["Credor"])
        natureza = np.where(valor < 0, "Receita", "Despesa")

    df["Natureza"] = natureza
    df["Sinal"]    = np.select([df["Natureza"].eq("Receita"), df["Natureza"].eq("Despesa")], [-1, 1], default=1)
    df["Saldo"]    = df["Devedor"] - df["Credor"]
    df["SaldoGerencial"] = df["Saldo"] * df["Sinal"]

    df["Competencia"] = pd.to_datetime(df["Competencia"], errors="coerce")
    df["AnoMes"] = df["Competencia"].dt.strftime("%Y-%m")
    if df["AnoMes"].isna().all():
        df["AnoMes"] = pd.Timestamp(date.today().replace(day=1)).strftime("%Y-%m")

    # Normalização de CentroCusto p/ comparação segura (resolve LOGÍSTICA vs Logistica)
    df["CentroCusto"] = df["CentroCusto"].astype(str).str.strip()
    df["CentroCustoNorm"] = df["CentroCusto"].map(strip_accents_upper)

    # strings
    for c in ["Empresa","ContaCodigo","ContaDescricao"]:
        if c in df.columns: df[c] = df[c].astype(str).str.strip()

    return df, notes

with st.spinner("Processando seu arquivo..."):
    df, notes = load_dataframe_from_bytes(file_bytes, file_name)

if notes:
    st.warning("Ajustes aplicados automaticamente:\n- " + "\n- ".join(notes))

# ─────────────────────────────── Filtros ───────────────────────────────
empresas  = sorted(df["Empresa"].dropna().unique().tolist())
naturezas = sorted(df["Natureza"].dropna().unique().tolist())
# Mostra os nomes originais, mas filtra usando a versão normalizada
centros_display = sorted(df["CentroCusto"].dropna().unique().tolist())
centros_norm_map = {c: strip_accents_upper(c) for c in centros_display}

colf1, colf2, colf3 = st.columns(3)
with colf1: f_emp = st.multiselect("Empresa", empresas, default=empresas)
with colf2: f_nat = st.multiselect("Natureza", [n for n in naturezas if n!="Outros"] or naturezas,
                                   default=[n for n in naturezas if n!="Outros"] or naturezas)
with colf3: f_cc_display  = st.multiselect("Centro de Custo", centros_display, default=centros_display)

f_cc_norm = [centros_norm_map[c] for c in f_cc_display]

# Competência por índice (1 mês → sem slider)
meses = sorted(df["AnoMes"].dropna().unique().tolist())
if len(meses) == 0:
    st.warning("Sem competências válidas."); st.stop()
elif len(meses) == 1:
    start_ym = end_ym = meses[0]
    st.info(f"Competência única: **{start_ym}**")
else:
    min_idx, max_idx = 0, len(meses)-1
    slider_key = f"mes_idx::{file_name}::{len(meses)}"
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
    & df["CentroCustoNorm"].isin(f_cc_norm)
)
df_f = df.loc[mask].copy()

# ─────────────────────────────── KPIs ──────────────────────────────────
# Receita/Despesa positivas para leitura; Caixa = Receita - Despesa
receita_pos = -df_f.loc[df_f["Natureza"]=="Receita","SaldoGerencial"].sum()
despesa_pos =  df_f.loc[df_f["Natureza"]=="Despesa","SaldoGerencial"].sum()
caixa = receita_pos - despesa_pos
margem = (caixa / receita_pos) if receita_pos else np.nan

colA, colB, colC, colD = st.columns(4)
with colA: st.metric("Receita",  money(receita_pos))
with colB: st.metric("Despesa",  money(despesa_pos))
with colC: st.metric("Caixa (Receita − Despesa)", money(caixa))
with colD: st.metric("Margem %", money((margem*100) if np.isfinite(margem) else 0))

st.markdown("---")

# ───────────────────────── Abas de Análise (inclui comparador) ───────────────────
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "1) Receita x Despesa por C.C.", 
    "2) Deep‑dive por Centro", 
    "3) Comparativo de Centros (x − x = valor)", 
    "4) Departamentos só com Despesa", 
    "5) Margem no Tempo", 
    "6) Top Receitas & Despesas", 
    "7) Tabela / Exportar"
])

with tab1:
    st.subheader("Receita x Despesa por Centro de Custo")
    por_cc = df_f.groupby(["CentroCusto","CentroCustoNorm","Natureza"], as_index=False)["SaldoGerencial"].sum()
    if por_cc.empty:
        st.info("Sem dados nos filtros.")
    else:
        por_cc["ValorPos"] = np.where(por_cc["Natureza"].eq("Receita"), -por_cc["SaldoGerencial"], por_cc["SaldoGerencial"])
        pivot = por_cc.pivot(index="CentroCusto", columns="Natureza", values="ValorPos").fillna(0)
        for col in ["Receita","Despesa"]:
            if col not in pivot.columns: pivot[col] = 0.0
        pivot = pivot.reset_index()
        fig = px.bar(pivot.sort_values("Receita", ascending=False), x="CentroCusto", y=["Receita","Despesa"], barmode="group")
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.subheader("Deep‑dive por Centro de Custo")
    cc_sel_display = st.selectbox("Escolha um Centro de Custo", options=sorted(centros_display))
    cc_sel_norm = strip_accents_upper(cc_sel_display)
    df_cc = df_f[df_f["CentroCustoNorm"]==cc_sel_norm].copy()
    if df_cc.empty:
        st.info("Sem dados para o centro selecionado dentro dos filtros atuais.")
    else:
        st.markdown("**Top 15 contas por valor absoluto (Receita positiva / Despesa positiva)**")
        ag = df_cc.groupby(["ContaCodigo","ContaDescricao","Natureza"], as_index=False)["SaldoGerencial"].sum()
        ag["ValorPos"] = np.where(ag["Natureza"].eq("Receita"), -ag["SaldoGerencial"], ag["SaldoGerencial"])
        ag["Abs"] = ag["ValorPos"].abs()
        top = ag.sort_values("Abs", ascending=False).head(15)
        fig = px.bar(top, x="ContaDescricao", y="ValorPos", color="Natureza")
        st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.subheader("Comparativo de Centros (x − x = valor)")
    colc1, colc2, colc3 = st.columns([1,1,1])
    with colc1:
        a_display = st.selectbox("Centro A", options=sorted(centros_display), key="cmp_a")
    with colc2:
        b_display = st.selectbox("Centro B", options=sorted(centros_display), key="cmp_b")
    with colc3:
        nat_cmp = st.selectbox("Natureza", options=["Receita","Despesa","Ambos"], index=2)

    a_norm, b_norm = strip_accents_upper(a_display), strip_accents_upper(b_display)
    df_cmp = df_f[df_f["CentroCustoNorm"].isin([a_norm, b_norm])].copy()
    if nat_cmp != "Ambos":
        df_cmp = df_cmp[df_cmp["Natureza"]==nat_cmp]
    # Receita positiva para leitura
    df_cmp["ValorPos"] = np.where(df_cmp["Natureza"].eq("Receita"), -df_cmp["SaldoGerencial"], df_cmp["SaldoGerencial"])
    somas = df_cmp.groupby("CentroCustoNorm", as_index=False)["ValorPos"].sum()
    va = float(somas.loc[somas["CentroCustoNorm"]==a_norm, "ValorPos"].sum()) if (somas["CentroCustoNorm"]==a_norm).any() else 0.0
    vb = float(somas.loc[somas["CentroCustoNorm"]==b_norm, "ValorPos"].sum()) if (somas["CentroCustoNorm"]==b_norm).any() else 0.0
    diff = va - vb
    m1, m2, m3 = st.columns(3)
    with m1: st.metric(f"{a_display} (A)", money(va))
    with m2: st.metric(f"{b_display} (B)", money(vb))
    with m3: st.metric("A − B", money(diff))
    base = pd.DataFrame({"Centro":[a_display,b_display], "Valor":[va,vb]})
    st.plotly_chart(px.bar(base, x="Centro", y="Valor"), use_container_width=True)

with tab4:
    st.subheader("Departamentos com apenas Despesas (sem Receita)")
    # Usa pivot para garantir robustez
    pv = df_f.pivot_table(index=["CentroCusto","CentroCustoNorm"], columns="Natureza",
                          values="SaldoGerencial", aggfunc="sum", fill_value=0).reset_index()
    # Receita positiva p/ leitura
    if "Receita" in pv.columns:
        pv["ReceitaPos"] = -pv["Receita"]
    else:
        pv["ReceitaPos"] = 0.0
    pv["DespesaPos"] = pv["Despesa"] if "Despesa" in pv.columns else 0.0
    only_cost = pv[(pv["ReceitaPos"]==0) & (pv["DespesaPos"]>0)].copy()
    if only_cost.empty:
        st.success("Todos os centros possuem alguma receita nos filtros.")
    else:
        only_cost = only_cost.sort_values("DespesaPos", ascending=True)
        st.plotly_chart(px.bar(only_cost, x="DespesaPos", y="CentroCusto", orientation="h"), use_container_width=True)

with tab5:
    st.subheader("Margem (Caixa/Receita) ao longo do tempo")
    # Evita lambdas com índice externo: usa pivot
    pv_mens = df_f.pivot_table(index="AnoMes", columns="Natureza", values="SaldoGerencial",
                               aggfunc="sum", fill_value=0).reset_index()
    if pv_mens.empty:
        st.info("Sem dados para calcular margem.")
    else:
        pv_mens["ReceitaPos"] = -pv_mens["Receita"] if "Receita" in pv_mens.columns else 0.0
        pv_mens["DespesaPos"] = pv_mens["Despesa"] if "Despesa" in pv_mens.columns else 0.0
        pv_mens["Caixa"] = pv_mens["ReceitaPos"] - pv_mens["DespesaPos"]
        pv_mens["Margem%"] = np.where(pv_mens["ReceitaPos"]>0, 100*pv_mens["Caixa"]/pv_mens["ReceitaPos"], np.nan)
        left, right = st.columns([2,1])
        with left:
            st.plotly_chart(px.line(pv_mens.sort_values("AnoMes"), x="AnoMes", y=["ReceitaPos","DespesaPos","Caixa"], markers=True),
                            use_container_width=True)
        with right:
            st.plotly_chart(px.line(pv_mens.sort_values("AnoMes"), x="AnoMes", y="Margem%", markers=True),
                            use_container_width=True)

with tab6:
    st.subheader("Top 10 Receitas & Despesas")
    # Receitas
    rec = df_f[df_f["Natureza"]=="Receita"].groupby("ContaDescricao", as_index=False)["SaldoGerencial"].sum()
    if not rec.empty:
        rec["ReceitaPos"] = -rec["SaldoGerencial"]
        rec = rec.sort_values("ReceitaPos", ascending=False).head(10)
        st.plotly_chart(px.bar(rec, x="ContaDescricao", y="ReceitaPos"), use_container_width=True)
    else:
        st.info("Sem receitas nos filtros.")
    # Despesas
    dep = df_f[df_f["Natureza"]=="Despesa"].groupby("ContaDescricao", as_index=False)["SaldoGerencial"].sum()
    if not dep.empty:
        dep = dep.sort_values("SaldoGerencial", ascending=False).head(10)
        st.plotly_chart(px.bar(dep, x="ContaDescricao", y="SaldoGerencial"), use_container_width=True)
    else:
        st.info("Sem despesas nos filtros.")

with tab7:
    st.subheader("Tabela detalhada e Exportações")
    cols = ["Empresa","Competencia","AnoMes","CentroCusto","ContaCodigo","ContaDescricao",
            "Natureza","Devedor","Credor","Saldo","Sinal","SaldoGerencial"]
    cols = [c for c in cols if c in df_f.columns]
    styled = (df_f[cols]
              .sort_values(["Competencia","ContaCodigo"])
              .reset_index(drop=True)
              .style
              .format({c: "{:,.2f}".format for c in ["Devedor","Credor","Saldo","SaldoGerencial"] if c in cols}))
    st.dataframe(styled, use_container_width=True, height=420)
    st.markdown("---")
    pivot_mes = df_f.pivot_table(index="AnoMes", columns="Natureza", values="SaldoGerencial",
                                 aggfunc="sum", fill_value=0).reset_index()
    if "Receita" in pivot_mes.columns:
        pivot_mes["Receita"] = -pivot_mes["Receita"]  # positiva no export
    by_cc_exp = df_f.groupby(["Natureza","CentroCusto"], as_index=False)["SaldoGerencial"] \
                    .sum().sort_values(["Natureza","SaldoGerencial"], ascending=[True, False])
    excel_bytes = to_excel_bytes({
        "Detalhado": df_f[cols],
        "Resumo_Mensal": pivot_mes,
        "Por_CentroCusto": by_cc_exp
    })
    st.download_button("⬇️ Excel (Detalhado + Resumos)", data=excel_bytes,
                       file_name="analise_balancete.xlsx", key="dl_excel")
    st.download_button("⬇️ CSV Detalhado", data=df_f.to_csv(index=False).encode("utf-8"),
                       file_name="balancete_detalhado.csv", key="dl_csv")
