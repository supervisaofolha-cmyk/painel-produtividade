import streamlit as st
import pandas as pd
import plotly.express as px
import locale

# ==================================================
# CONFIGURAÇÃO REGIONAL BRASIL
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
# CSS PERSONALIZADO
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
    box-shadow: 0px 2px 5px rgba(0,0,0,0.05);
}}

h1, h2, h3 {{
    color: {COR_CINZA};
}}

</style>
""", unsafe_allow_html=True)

# ==================================================
# TÍTULO
# ==================================================

st.title("📊 Painel de Produtividade")

# ==================================================
# LEITURA PLANILHA
# ==================================================

df = pd.read_excel("produtividade.xlsx")

# ==================================================
# AJUSTAR COLUNAS
# ==================================================

df.columns = df.columns.str.strip()

# ==================================================
# PADRONIZAR TÉCNICO
# ==================================================

df["Técnico"] = (
    df["Técnico"]
    .astype(str)
    .str.lower()
    .str.strip()
)

# ==================================================
# CONVERTER DATA
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

usuarios = pd.read_excel("usuarios.xlsx")

usuarios.columns = (
    usuarios.columns
    .str.strip()
    .str.lower()
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
# SIDEBAR LOGIN
# ==================================================

st.sidebar.title("🔐 Login")

usuario_input = st.sidebar.text_input(
    "Usuário"
)

senha_input = st.sidebar.text_input(
    "Senha",
    type="password"
)

# ==================================================
# VALIDAR LOGIN
# ==================================================

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

        st.error("Usuário ou senha inválidos.")
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
# FILTRAR DADOS
# ==================================================

dados_tecnico = df[
    df["Técnico"] == tecnico
]

# ==================================================
# VALIDAR DADOS
# ==================================================

if dados_tecnico.empty:

    st.error(
        "Nenhum dado encontrado."
    )

    st.stop()

# ==================================================
# PEGAR NÍVEL
# ==================================================

nivel_tecnico = (
    dados_tecnico["Nível"]
    .iloc[0]
)

# ==================================================
# VISÃO GERAL GESTÃO
# ==================================================

if modo_gestao:

    st.divider()

    st.header("📊 Visão Geral da Operação")

    resumo_nivel = (
        df.groupby("Nível")
        .agg({
            "Realizado": "sum",
            "SSC": "sum",
            "RO": "sum",
            "Votação": "mean",
            "Satisfação": "mean"
        })
        .reset_index()
    )

    resumo_nivel["Votação"] = (
        resumo_nivel["Votação"]
        .round(2)
    )

    resumo_nivel["Satisfação"] = (
        resumo_nivel["Satisfação"]
        .round(2)
    )

    st.dataframe(
        resumo_nivel,
        use_container_width=True
    )

    grafico_geral = px.bar(
        resumo_nivel,
        x="Nível",
        y="Realizado",
        color="Nível",
        text="Realizado",
        title="Realizado Geral por Nível",
        color_discrete_sequence=[
            COR_LARANJA,
            COR_CINZA,
            "#D1D5DB"
        ]
    )

    grafico_geral.update_traces(
        textposition="outside"
    )

    grafico_geral.update_layout(
        plot_bgcolor=COR_BRANCO,
        paper_bgcolor=COR_BRANCO,
        font_color=COR_CINZA,
        xaxis_title="Nível",
        yaxis_title="Realizado"
    )

    st.plotly_chart(
        grafico_geral,
        use_container_width=True
    )

# ==================================================
# RESULTADOS INDIVIDUAIS
# ==================================================

st.divider()

st.subheader(
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
# GRÁFICO REALIZADO X ESPERADO
# ==================================================

st.divider()

st.subheader(
    "📈 Realizado x Esperado por Dia"
)

comparativo_dia = (
    dados_tecnico.groupby(
        ["Data", "Data Formatada"]
    )
    .agg({
        "Realizado": "sum",
        "Esperado": "sum"
    })
    .reset_index()
)

comparativo_dia = comparativo_dia.sort_values(
    by="Data"
)

comparativo_long = comparativo_dia.melt(
    id_vars=["Data", "Data Formatada"],
    value_vars=["Realizado", "Esperado"],
    var_name="Indicador",
    value_name="Quantidade"
)

grafico_comparativo = px.bar(
    comparativo_long,
    x="Data Formatada",
    y="Quantidade",
    color="Indicador",
    barmode="group",
    text="Quantidade",
    title="Comparativo Diário",
    labels={
        "Data Formatada": "Data",
        "Quantidade": "Quantidade",
        "Indicador": "Indicador"
    },
    color_discrete_map={
        "Realizado": COR_LARANJA,
        "Esperado": COR_CINZA
    }
)

grafico_comparativo.update_traces(
    textposition="outside"
)

grafico_comparativo.update_layout(
    xaxis_title="Data",
    yaxis_title="Quantidade",
    legend_title="Indicadores",
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis=dict(type="category")
)

st.plotly_chart(
    grafico_comparativo,
    use_container_width=True
)

# ==================================================
# VOTAÇÃO
# ==================================================

st.divider()

st.subheader(
    "🗳️ Votação Média por Dia"
)

votacao_dia = (
    dados_tecnico.groupby(
        ["Data", "Data Formatada"]
    )["Votação"]
    .mean()
    .reset_index()
)

votacao_dia = votacao_dia.sort_values(
    by="Data"
)

grafico_votacao = px.line(
    votacao_dia,
    x="Data Formatada",
    y="Votação",
    markers=True,
    title="Votação Média",
    labels={
        "Data Formatada": "Data",
        "Votação": "Votação (%)"
    }
)

grafico_votacao.update_traces(
    line_color=COR_LARANJA
)

grafico_votacao.update_layout(
    xaxis_title="Data",
    yaxis_title="Votação (%)",
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis=dict(type="category")
)

st.plotly_chart(
    grafico_votacao,
    use_container_width=True
)

# ==================================================
# SATISFAÇÃO
# ==================================================

st.divider()

st.subheader(
    "😊 Satisfação Média por Dia"
)

satisfacao_dia = (
    dados_tecnico.groupby(
        ["Data", "Data Formatada"]
    )["Satisfação"]
    .mean()
    .reset_index()
)

satisfacao_dia = satisfacao_dia.sort_values(
    by="Data"
)

grafico_satisfacao = px.line(
    satisfacao_dia,
    x="Data Formatada",
    y="Satisfação",
    markers=True,
    title="Satisfação Média",
    labels={
        "Data Formatada": "Data",
        "Satisfação": "Satisfação (%)"
    }
)

grafico_satisfacao.update_traces(
    line_color=COR_CINZA
)

grafico_satisfacao.update_layout(
    xaxis_title="Data",
    yaxis_title="Satisfação (%)",
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis=dict(type="category")
)

st.plotly_chart(
    grafico_satisfacao,
    use_container_width=True
)

# ==================================================
# RESULTADO MENSAL
# ==================================================

st.divider()

st.subheader("📅 Resultado Mensal")

mes_atual = df["Data"].dt.month.max()

ano_atual = df["Data"].dt.year.max()

dados_mes = dados_tecnico[
    (dados_tecnico["Data"].dt.month == mes_atual)
    &
    (dados_tecnico["Data"].dt.year == ano_atual)
]

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        "Realizado Mensal",
        int(dados_mes["Realizado"].sum())
    )

with col2:
    st.metric(
        "SSC Mensal",
        int(dados_mes["SSC"].sum())
    )

with col3:
    st.metric(
        "RO Mensal",
        int(dados_mes["RO"].sum())
    )

with col4:
    st.metric(
        "Votação Mensal",
        f"{round(dados_mes['Votação'].mean(),2)}%"
    )

# ==================================================
# RANKING DIÁRIO
# ==================================================

st.divider()

st.subheader(
    f"🏆 Ranking Diário - {nivel_tecnico}"
)

ultima_data = df["Data"].max()

df_dia = df[
    (df["Data"] == ultima_data)
    &
    (df["Nível"] == nivel_tecnico)
]

ranking_dia = (
    df_dia.groupby("Técnico")
    ["Realizado"]
    .sum()
    .reset_index()
)

ranking_dia = ranking_dia.sort_values(
    by="Realizado",
    ascending=False
)

grafico_ranking_dia = px.bar(
    ranking_dia,
    x="Técnico",
    y="Realizado",
    text="Realizado",
    title=f"Ranking Diário - {nivel_tecnico}",
    color_discrete_sequence=[COR_LARANJA]
)

grafico_ranking_dia.update_traces(
    textposition="outside"
)

grafico_ranking_dia.update_layout(
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis_title="Técnico",
    yaxis_title="Realizado Diário"
)

st.plotly_chart(
    grafico_ranking_dia,
    use_container_width=True
)

# ==================================================
# RANKING MENSAL
# ==================================================

st.divider()

st.subheader(
    f"🏆 Ranking Mensal - {nivel_tecnico}"
)

df_mes = df[
    (df["Data"].dt.month == mes_atual)
    &
    (df["Data"].dt.year == ano_atual)
    &
    (df["Nível"] == nivel_tecnico)
]

ranking_mes = (
    df_mes.groupby("Técnico")
    ["Realizado"]
    .sum()
    .reset_index()
)

ranking_mes = ranking_mes.sort_values(
    by="Realizado",
    ascending=False
)

grafico_ranking_mes = px.bar(
    ranking_mes,
    x="Técnico",
    y="Realizado",
    text="Realizado",
    title=f"Ranking Mensal - {nivel_tecnico}",
    color_discrete_sequence=[COR_CINZA]
)

grafico_ranking_mes.update_traces(
    textposition="outside"
)

grafico_ranking_mes.update_layout(
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis_title="Técnico",
    yaxis_title="Realizado Mensal"
)

st.plotly_chart(
    grafico_ranking_mes,
    use_container_width=True
)