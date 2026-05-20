import streamlit as st
import pandas as pd
import plotly.express as px
import locale

# ==================================================
# CONFIGURAÇÃO REGIONAL
# ==================================================

try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
except:
    pass

# ==================================================
# CONFIGURAÇÃO DA PÁGINA
# ==================================================

st.set_page_config(
    page_title="Painel de Produtividade",
    layout="wide"
)

# ==================================================
# CORES
# ==================================================

COR_LARANJA = "#F97316"
COR_CINZA = "#6B7280"
COR_BRANCO = "#FFFFFF"
FUNDO = "#F5F5F5"

# ==================================================
# CSS
# ==================================================

st.markdown(f"""
<style>

.main {{
    background-color: {FUNDO};
}}

[data-testid="stMetric"] {{
    background-color: {COR_BRANCO};
    border: 1px solid #E5E7EB;
    padding: 15px;
    border-radius: 12px;
}}

.status-card {{
    padding: 20px;
    border-radius: 15px;
    color: white;
    text-align: center;
    min-height: 300px;
}}

</style>
""", unsafe_allow_html=True)

# ==================================================
# TÍTULO
# ==================================================

st.title("📊 Painel de Produtividade")

# ==================================================
# LEITURA DOS ARQUIVOS
# ==================================================

df = pd.read_excel("produtividade.xlsx")

usuarios = pd.read_excel("usuarios.xlsx")

# ==================================================
# AJUSTAR COLUNAS
# ==================================================

df.columns = df.columns.str.strip()

usuarios.columns = (
    usuarios.columns
    .str.strip()
    .str.lower()
)

# ==================================================
# PADRONIZAR DADOS
# ==================================================

df["Técnico"] = (
    df["Técnico"]
    .astype(str)
    .str.lower()
    .str.strip()
)

usuarios["usuario"] = (
    usuarios["usuario"]
    .astype(str)
    .str.lower()
    .str.strip()
)

usuarios["senha"] = (
    usuarios["senha"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)

usuarios["tecnico"] = (
    usuarios["tecnico"]
    .astype(str)
    .str.lower()
    .str.strip()
)

# ==================================================
# DATA
# ==================================================

df["Data"] = pd.to_datetime(
    df["Data"],
    dayfirst=True,
    errors="coerce"
)

df = df.dropna(subset=["Data"])

df["Data Formatada"] = (
    df["Data"]
    .dt.strftime("%d/%m/%Y")
)

# ==================================================
# LOGIN
# ==================================================

st.sidebar.title("🔐 Login")

usuario_input = st.sidebar.text_input(
    "Usuário"
)

senha_input = st.sidebar.text_input(
    "Senha",
    type="password"
)

if usuario_input == "" or senha_input == "":
    st.warning("Digite usuário e senha.")
    st.stop()

usuario_digitado = (
    str(usuario_input)
    .lower()
    .strip()
)

senha_digitada = (
    str(senha_input)
    .replace(".0", "")
    .strip()
)

# ==================================================
# LOGIN GESTÃO
# ==================================================

if (
    usuario_digitado == "gestao"
    and
    senha_digitada == "30071997"
):

    modo_gestao = True

    st.sidebar.success(
        "Bem-vinda Gestão"
    )

else:

    modo_gestao = False

    login = usuarios[
        (usuarios["usuario"] == usuario_digitado)
        &
        (usuarios["senha"] == senha_digitada)
    ]

    if login.empty:

        st.error(
            "Usuário ou senha inválidos."
        )

        st.stop()

    tecnico = login.iloc[0]["tecnico"]

    st.sidebar.success(
        f"Bem-vindo(a), {tecnico.title()}"
    )

# ==================================================
# MODO GESTÃO
# ==================================================

if modo_gestao:

    tecnico = st.selectbox(
        "Selecione o Técnico",
        sorted(df["Técnico"].unique())
    )

# ==================================================
# DADOS TÉCNICO
# ==================================================

dados_tecnico = df[
    df["Técnico"] == tecnico
]

if dados_tecnico.empty:

    st.error(
        "Nenhum dado encontrado."
    )

    st.stop()

# ==================================================
# NÍVEL TÉCNICO
# ==================================================

nivel_tecnico = (
    dados_tecnico["Nível"]
    .iloc[0]
)

# ==================================================
# ÚLTIMO DIA COM MOVIMENTAÇÃO
# ==================================================

ultima_data = df["Data"].max()

df_ultimo_dia = df[
    df["Data"] == ultima_data
]

# ==================================================
# CLASSIFICAÇÃO DIÁRIA
# ==================================================

def definir_status(percentual):

    if percentual < 70:
        return "CRÍTICO"

    elif percentual < 90:
        return "ATENÇÃO"

    elif percentual < 100:
        return "BOM"

    else:
        return "EXCELENTE"

# ==================================================
# QUARTIL DIÁRIO GESTÃO
# ==================================================

if modo_gestao:

    st.divider()

    st.header("📌 Status Diário da Operação")

    resumo_dia = (
        df_ultimo_dia.groupby("Técnico")
        .agg({
            "Realizado": "sum",
            "Esperado": "sum"
        })
        .reset_index()
    )

    resumo_dia["Percentual"] = (
        resumo_dia["Realizado"]
        /
        resumo_dia["Esperado"]
    ) * 100

    resumo_dia["Status"] = (
        resumo_dia["Percentual"]
        .apply(definir_status)
    )

    col1, col2, col3, col4 = st.columns(4)

    status_cores = {
        "CRÍTICO": "#DC2626",
        "ATENÇÃO": "#F97316",
        "BOM": "#2563EB",
        "EXCELENTE": "#16A34A"
    }

    colunas = {
        "CRÍTICO": col1,
        "ATENÇÃO": col2,
        "BOM": col3,
        "EXCELENTE": col4
    }

    for status in status_cores.keys():

        tecnicos_status = resumo_dia[
            resumo_dia["Status"] == status
        ]["Técnico"].tolist()

        nomes = "<br>".join(
            [x.title() for x in tecnicos_status]
        )

        with colunas[status]:

            st.markdown(f"""
            <div class="status-card"
            style="background-color:{status_cores[status]}">

            <h2>{status}</h2>

            <p>{nomes}</p>

            </div>
            """, unsafe_allow_html=True)

# ==================================================
# QUARTIL MENSAL GESTÃO
# ==================================================

if modo_gestao:

    st.divider()

    st.header("📅 Status Mensal da Operação")

    mes_atual = df["Data"].dt.month.max()

    ano_atual = df["Data"].dt.year.max()

    df_mes = df[
        (df["Data"].dt.month == mes_atual)
        &
        (df["Data"].dt.year == ano_atual)
    ]

    resumo_mes = (
        df_mes.groupby("Técnico")
        .agg({
            "Realizado": "sum",
            "Esperado": "sum"
        })
        .reset_index()
    )

    resumo_mes["Percentual"] = (
        resumo_mes["Realizado"]
        /
        resumo_mes["Esperado"]
    ) * 100

    resumo_mes["Status"] = (
        resumo_mes["Percentual"]
        .apply(definir_status)
    )

    col1, col2, col3, col4 = st.columns(4)

    status_cores = {
        "CRÍTICO": "#DC2626",
        "ATENÇÃO": "#F97316",
        "BOM": "#2563EB",
        "EXCELENTE": "#16A34A"
    }

    colunas = {
        "CRÍTICO": col1,
        "ATENÇÃO": col2,
        "BOM": col3,
        "EXCELENTE": col4
    }

    for status in status_cores.keys():

        tecnicos_status = resumo_mes[
            resumo_mes["Status"] == status
        ]["Técnico"].tolist()

        nomes = "<br>".join(
            [x.title() for x in tecnicos_status]
        )

        with colunas[status]:

            st.markdown(f"""
            <div class="status-card"
            style="background-color:{status_cores[status]}">

            <h2>{status}</h2>

            <p>{nomes}</p>

            </div>
            """, unsafe_allow_html=True)

# ==================================================
# RESULTADOS INDIVIDUAIS
# ==================================================

st.divider()

st.header(
    f"📌 Resultados Individuais - {tecnico.title()}"
)

col1, col2, col3, col4, col5 = st.columns(5)

with col1:

    st.metric(
        "Realizado Total",
        int(dados_tecnico["Realizado"].sum())
    )

with col2:

    st.metric(
        "SSC Total",
        int(dados_tecnico["SSC"].sum())
    )

with col3:

    st.metric(
        "RO Total",
        int(dados_tecnico["RO"].sum())
    )

with col4:

    st.metric(
        "Votação Média",
        f"{round(dados_tecnico['Votação'].mean(),2)}%"
    )

with col5:

    classificacao_mode = (
        dados_tecnico["Classificação"]
        .dropna()
        .mode()
    )

    if classificacao_mode.empty:

        classificacao = "Sem classificação"

    else:

        classificacao = classificacao_mode.iloc[0]

    st.metric(
        "Classificação",
        classificacao
    )

# ==================================================
# REALIZADO X ESPERADO
# ==================================================

st.divider()

st.subheader(
    "📈 Realizado x Esperado"
)

comparativo = (
    dados_tecnico.groupby(
        ["Data", "Data Formatada"]
    )
    .agg({
        "Realizado": "sum",
        "Esperado": "sum"
    })
    .reset_index()
)

comparativo = comparativo.sort_values(
    by="Data"
)

comparativo_long = comparativo.melt(
    id_vars=["Data", "Data Formatada"],
    value_vars=["Realizado", "Esperado"],
    var_name="Indicador",
    value_name="Quantidade"
)

grafico = px.bar(
    comparativo_long,
    x="Data Formatada",
    y="Quantidade",
    color="Indicador",
    barmode="group",
    text="Quantidade",
    color_discrete_map={
        "Realizado": COR_LARANJA,
        "Esperado": COR_CINZA
    }
)

grafico.update_traces(
    textposition="outside"
)

grafico.update_layout(
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis_title="Data",
    yaxis_title="Quantidade",
    xaxis=dict(type="category")
)

st.plotly_chart(
    grafico,
    use_container_width=True
)

# ==================================================
# RANKING DIÁRIO
# ==================================================

st.divider()

st.header(
    f"🏆 Ranking Diário - {nivel_tecnico}"
)

df_ranking_dia = df[
    (df["Data"] == ultima_data)
    &
    (df["Nível"] == nivel_tecnico)
]

ranking_dia = (
    df_ranking_dia.groupby("Técnico")
    ["Realizado"]
    .sum()
    .reset_index()
)

ranking_dia = ranking_dia.sort_values(
    by="Realizado",
    ascending=False
)

grafico_rank_dia = px.bar(
    ranking_dia,
    x="Técnico",
    y="Realizado",
    text="Realizado",
    color_discrete_sequence=[COR_LARANJA]
)

grafico_rank_dia.update_traces(
    textposition="outside"
)

grafico_rank_dia.update_layout(
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis_title="Técnico",
    yaxis_title="Realizado"
)

st.plotly_chart(
    grafico_rank_dia,
    use_container_width=True
)

# ==================================================
# RANKING MENSAL
# ==================================================

st.divider()

st.header(
    f"🏆 Ranking Mensal - {nivel_tecnico}"
)

mes_atual = df["Data"].dt.month.max()

ano_atual = df["Data"].dt.year.max()

df_ranking_mes = df[
    (df["Data"].dt.month == mes_atual)
    &
    (df["Data"].dt.year == ano_atual)
    &
    (df["Nível"] == nivel_tecnico)
]

ranking_mes = (
    df_ranking_mes.groupby("Técnico")
    ["Realizado"]
    .sum()
    .reset_index()
)

ranking_mes = ranking_mes.sort_values(
    by="Realizado",
    ascending=False
)

grafico_rank_mes = px.bar(
    ranking_mes,
    x="Técnico",
    y="Realizado",
    text="Realizado",
    color_discrete_sequence=[COR_CINZA]
)

grafico_rank_mes.update_traces(
    textposition="outside"
)

grafico_rank_mes.update_layout(
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis_title="Técnico",
    yaxis_title="Realizado"
)

st.plotly_chart(
    grafico_rank_mes,
    use_container_width=True
)