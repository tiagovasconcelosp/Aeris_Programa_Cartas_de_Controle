import os
import sys

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ===============================
# CONFIGURAÇÃO DA PÁGINA
# ===============================
st.set_page_config(layout="wide")
st.markdown(""" """, unsafe_allow_html=True)

def resource_path(relative_path: str) -> str:
    """
    Retorna o caminho absoluto de um recurso, funcionando no PyInstaller e no dev.
    (Use para assets empacotados, como LOGO; não para o Excel externo.)
    """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# Logo empacotada
st.sidebar.image(resource_path("LOGO AERIS.png"), width=180)

# ===============================
# UPLOAD / FONTE DE DADOS
# ===============================
def exe_dir() -> str:
    """
    Retorna a pasta onde o executável está (no PyInstaller onefile/onedir)
    ou, em dev, a pasta do arquivo .py.
    """
    if getattr(sys, 'frozen', False):  # Executando empacotado (PyInstaller)
        return os.path.dirname(sys.executable)
    # Execução em dev (python main.py / streamlit run)
    return os.path.dirname(os.path.abspath(__file__))

# Nome do Excel esperado ao lado do .exe
EXCEL_NAME = "V163_CTQ_Blank Template_R3-piloto.xlsx"
excel_external_path = os.path.join(exe_dir(), EXCEL_NAME)

st.sidebar.markdown("### 🗂️ Fonte de dados")
st.sidebar.divider()
use_external = True  # padrão: usar arquivo externo ao lado do .exe

# (Opcional) permitir upload também:
uploaded = st.sidebar.file_uploader("Carregar Excel (.xlsx)", type=["xlsx"])

# Aviso amigável se o externo não existir
if not os.path.exists(excel_external_path):
    st.sidebar.warning(
        f"Coloque o arquivo **{EXCEL_NAME}** na mesma pasta do executável "
        f"e reinicie o aplicativo."
    )

# --- Seleção da fonte e criação do 'xls' (OBJETO CORRETO A SER USADO ABAIXO) ---
if uploaded is not None:
    # Prioriza o upload se o usuário enviar
    xls = pd.ExcelFile(uploaded, engine="openpyxl")
    st.sidebar.success(f"Usando upload: {uploaded.name}")
elif use_external:
    if os.path.exists(excel_external_path):
        xls = pd.ExcelFile(excel_external_path, engine="openpyxl")
        st.sidebar.success(f"Usando arquivo externo: {EXCEL_NAME}")
    else:
        st.sidebar.error(
            f"Arquivo não encontrado ao lado do .exe:\n{excel_external_path}"
        )
        st.stop()
else:
    st.sidebar.error("Nenhuma fonte de dados selecionada.")
    st.stop()

# ===============================
# FUNÇÕES DE LEITURA
# ===============================
def ler_main_mold(df_raw):
    dados = []
    maquinas = [
        {"ctq_col": 1, "val_col": 2},
        {"ctq_col": 4, "val_col": 5},
    ]
    for i, m in enumerate(maquinas, start=1):
        ctq_nome = df_raw.iloc[1, m["ctq_col"]]
        lsl = pd.to_numeric(df_raw.iloc[2, m["val_col"]], errors="coerce")
        usl = pd.to_numeric(df_raw.iloc[4, m["val_col"]], errors="coerce")
        pas = df_raw.iloc[7:361, m["ctq_col"]].values
        valores = df_raw.iloc[7:361, m["val_col"]].values
        for p, v in zip(pas, valores):
            dados.append({
                "CTQ": ctq_nome,
                "Radius": f"Máquina {i}",
                "PA": p,
                "Valor": v,
                "LSL": lsl,
                "USL": usl
            })
    return pd.DataFrame(dados)

def ler_finish_fb(df_raw):
    dados = []
    total_cols = df_raw.shape[1]
    col = 0
    while col < total_cols:
        try:
            ctq_nome = df_raw.iloc[1, col]
            if pd.isna(ctq_nome):
                col += 1
                continue
            lsl = pd.to_numeric(str(df_raw.iloc[2, col + 1]).replace(",", "."), errors="coerce")
            usl = pd.to_numeric(str(df_raw.iloc[3, col + 1]).replace(",", "."), errors="coerce")
            pas = df_raw.iloc[5:300, col].values
            valores = df_raw.iloc[5:300, col + 1].values
            for p, v in zip(pas, valores):
                if pd.notna(p):
                    dados.append({
                        "CTQ": ctq_nome,
                        "Radius": "Finish FB",
                        "PA": p,
                        "Valor": v,
                        "LSL": lsl,
                        "USL": usl
                    })
            col += 3
        except:
            col += 1
    return pd.DataFrame(dados)

def ler_resistance_measurement(df_raw):
    dados = []
    total_cols = df_raw.shape[1]
    col = 0
    while col < total_cols - 1:
        ctq_nome = df_raw.iloc[1, col]
        if pd.isna(ctq_nome) or str(ctq_nome).strip() == "":
            col += 2
            continue
        usl_raw = df_raw.iloc[3, col + 1]
        usl = pd.to_numeric(str(usl_raw).replace(",", "."), errors="coerce")
        lsl = np.nan  # LSL é NA nessa planilha
        linha = 4
        while linha < df_raw.shape[0]:
            pa = df_raw.iloc[linha, col]
            valor = df_raw.iloc[linha, col + 1]
            if pd.isna(pa) or str(pa).strip() == "":
                break
            dados.append({
                "CTQ": str(ctq_nome).strip(),
                "Radius": str(ctq_nome).strip(),
                "PA": pa,
                "Valor": valor,
                "LSL": lsl,
                "USL": usl
            })
            linha += 1
        col += 2
    df_final = pd.DataFrame(dados)
    return df_final

def ler_padrao(df):
    df.columns = df.columns.astype(str).str.strip().str.upper()
    obrigatorias = ["CTQ", "RADIUS", "LSL", "USL"]
    for col in obrigatorias:
        if col not in df.columns:
            st.error("Obrigatório: Selecionar Tecnologia na Coluna na esquerda")
            st.stop()
    df["CTQ"] = df["CTQ"].ffill()
    pas_cols = [c for c in df.columns if c.isdigit()]
    df_long = df.melt(
        id_vars=["CTQ", "RADIUS", "LSL", "USL"],
        value_vars=pas_cols,
        var_name="PA",
        value_name="Valor"
    )
    df_long.rename(columns={"RADIUS": "Radius"}, inplace=True)
    return df_long

# ===============================
# SELEÇÃO DA ABA
# ===============================
sheet_escolhida = st.sidebar.selectbox(
    "📄 Selecionar Tecnologia",
    options=xls.sheet_names
)

df_raw = pd.read_excel(xls, sheet_name=sheet_escolhida, header=None)

mapa_abas = {
    "Main Mold (MB)": ler_main_mold,
    "Finish (FB)": ler_finish_fb,
    "Resistance Measurement": ler_resistance_measurement
}

if sheet_escolhida in mapa_abas:
    df_long = mapa_abas[sheet_escolhida](df_raw)
else:
    df_normal = pd.read_excel(xls, sheet_name=sheet_escolhida)
    df_long = ler_padrao(df_normal)

# ===============================
# LIMPEZA
# ===============================
df_long["Valor"] = (
    df_long["Valor"].astype(str)
    .str.replace(",", ".", regex=False)
    .str.strip()
)
df_long["Valor"] = pd.to_numeric(df_long["Valor"], errors="coerce")
df_long["PA"] = pd.to_numeric(df_long["PA"], errors="coerce")
df_long = df_long.dropna(subset=["PA"])
df_long["PA"] = df_long["PA"].astype(int)

# ===============================
# PIVOT
# ===============================
df_pivot = df_long.pivot_table(
    index=["CTQ", "Radius", "LSL", "USL"],
    columns="PA",
    values="Valor",
    aggfunc="mean"
).reset_index()
df = df_pivot
df.columns = [
    str(int(float(c))) if str(c).replace('.', '', 1).isdigit() else str(c)
    for c in df.columns
]

df["LSL"] = pd.to_numeric(df["LSL"], errors="coerce")
df["USL"] = pd.to_numeric(df["USL"], errors="coerce")

# ===============================
# FILTROS
# ===============================
st.sidebar.markdown("### 🔎 Filtros")
selecao_ctq = st.sidebar.multiselect(
    "Selecionar CTQ:",
    options=df['CTQ'].dropna().unique(),
    default=df['CTQ'].dropna().unique()
)
selecao_raios = st.sidebar.multiselect(
    "Selecionar Raios:",
    options=df['Radius'].dropna().unique(),
    default=df['Radius'].dropna().unique()
)

colunas_pas = []
for c in df.columns:
    try:
        numero = int(float(c))
        colunas_pas.append(numero)
    except:
        pass
colunas_pas = sorted(colunas_pas)
if not colunas_pas:
    st.error("Nenhuma coluna de PÁS foi encontrada nessa aba.")
    st.stop()

pa_min, pa_max = st.sidebar.slider(
    "Selecionar intervalo de PÁS:",
    min_value=min(colunas_pas),
    max_value=max(colunas_pas),
    value=(min(colunas_pas), min(max(colunas_pas), min(colunas_pas) + 29))
)
intervalo_pas = list(range(pa_min, pa_max + 1))

# ===============================
# SELEÇÃO MANUAL
# ===============================
pas_selecionadas = st.sidebar.multiselect(
    "Selecionar PÁS dentro do intervalo:", options=intervalo_pas,
    default=intervalo_pas
)
pas_selecionadas = [str(i) for i in pas_selecionadas]

# segurança caso nenhuma pá seja escolhida
if not pas_selecionadas:
    st.warning("Selecione pelo menos uma PA.")
    st.stop()

df_select = df[
    (df['CTQ'].isin(selecao_ctq)) &
    (df['Radius'].isin(selecao_raios))
].copy()

with st.expander("📋 Visualizar dados filtrados"):
    st.dataframe(
        df_select[['CTQ', 'Radius', 'LSL', 'USL'] + pas_selecionadas],
        width="stretch"
    )

# ===============================
# FUNÇÃO COR CPK
# ===============================
def cor_cpk(valor):
    if np.isnan(valor):
        return "#6B7280"
    if valor >= 1.67:
        return "#16A34A"
    elif valor >= 1.33:
        return "#84CC16"
    elif valor >= 1.00:
        return "#F59E0B"
    else:
        return "#DC2626"

# ===============================
# TÍTULO
# ===============================
st.markdown("## 📈 Controle Estatístico do Processo – V163")

# ===============================
# GRÁFICOS
# ===============================
for raio in selecao_raios:
    col_grafico, col_kpi = st.columns([4, 1])
    fig = go.Figure()

    df_r = df_select[df_select['Radius'] == raio]
    if df_r.empty:
        continue

    ctq_nome = df_r['CTQ'].iloc[0]
    y_vals = df_r[pas_selecionadas].mean()
    lsl = df_r['LSL'].mean()
    usl = df_r['USL'].mean()

    valores = y_vals.dropna().values
    if len(valores) > 1:
        media = np.mean(valores)
        std = np.std(valores, ddof=1)
    else:
        media = np.nan
        std = np.nan

    ucl = media + 3 * std if not np.isnan(std) else np.nan
    lcl = media - 3 * std if not np.isnan(std) else np.nan
    cp = (usl - lsl) / (6 * std) if not np.isnan(std) and std > 0 else np.nan
    cpk = min((usl - media) / (3 * std), (media - lsl) / (3 * std)) if not np.isnan(std) and std > 0 else np.nan

    # MÉDIA
    fig.add_trace(go.Scatter(
        x=pas_selecionadas, y=y_vals,
        mode='lines+markers', name="Média",
        line=dict(color="#0F766E", width=3)
    ))
    # LSL / USL
    fig.add_trace(go.Scatter(
        x=pas_selecionadas, y=[lsl] * len(pas_selecionadas),
        line=dict(color="#DC2626", dash="dash", width=2), name="LSL"
    ))
    fig.add_trace(go.Scatter(
        x=pas_selecionadas, y=[usl] * len(pas_selecionadas),
        line=dict(color="#DC2626", dash="dash", width=2), name="USL"
    ))
    # LCL / UCL
    fig.add_trace(go.Scatter(
        x=pas_selecionadas, y=[lcl] * len(pas_selecionadas),
        line=dict(color="#F97316", dash="dot", width=2), name="LCL"
    ))
    fig.add_trace(go.Scatter(
        x=pas_selecionadas, y=[ucl] * len(pas_selecionadas),
        line=dict(color="#F97316", dash="dot", width=2), name="UCL"
    ))

    # Layout
    if len(valores) > 0:
        y_min = min(lcl, lsl, min(valores))
        y_max = max(ucl, usl, max(valores))
        margem = (y_max - y_min) * 0.1  # reservado, caso queira usar

    fig.update_layout(
        title=dict(text=f"{ctq_nome} — {raio}", x=0.5, xanchor="center"),
        margin=dict(b=80),
        xaxis=dict(
            tickmode="array",
            tickvals=pas_selecionadas,
            ticktext=pas_selecionadas,
            tickangle=90,
            type="category"
        )
    )

    # Render
    with col_grafico:
        st.plotly_chart(fig, width="stretch")

    with col_kpi:
        st.markdown(f"""
##### Sumário do {raio}

Cp: **{cp:.2f}** <br>
Cpk: **{cpk:.2f}** <br>
Desvio padrão σ: **{std:.2f}**
""", unsafe_allow_html=True)

    st.divider()