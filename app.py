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
# LEITURA PLANILHAS
# ==================================================

df = pd.read_excel("produtividade.xlsx")
usuarios = pd.read_excel("usuarios.xlsx")

# ==================================================
# AJUSTAR COLUNAS
# ==================================================

df.columns = df.columns.str.strip()
usuarios.columns = usuarios.columns.str.strip().str.lower()

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

usuario_input = st.sidebar.text_input("Usuário")

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
# MODO GESTÃO
# ==================================================

modo_gestao = False

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
# SELEÇÃO TÉCNICO GESTÃO
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

if dados_tecnico.empty:
    st.error("Nenhum dado encontrado.")
    st.stop()

nivel_tecnico = (
    dados_tecnico["Nível"]
    .iloc[0]
)

# ==================================================
# FUNÇÃO STATUS
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
# QUARTIL DIÁRIO
# ==================================================

if modo_gestao:

    st.divider()

    st.header("📊 Status Diário da Operação")

    ultima_data = df["Data"].max()

    dados_dia = df[
        df["Data"] == ultima_data
    ]

    status_dia = (
        dados_dia.groupby("Técnico")
        .agg({
            "Realizado": "sum",
            "Esperado": "sum"
        })
        .reset_index()
    )

    status_dia["Percentual"] = (
        status_dia["Realizado"]
        /
        status_dia["Esperado"]
    ) * 100

    status_dia["Status"] = status_dia[
        "Percentual"
    ].apply(definir_status)

    st.subheader(
        f"📅 Diário - {ultima_data.strftime('%d/%m/%Y')}"
    )

    col1, col2, col3, col4 = st.columns(4)

    cores = {
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

    for status in [
        "CRÍTICO",
        "ATENÇÃO",
        "BOM",
        "EXCELENTE"
    ]:

        tecnicos = status_dia[
            status_dia["Status"] == status
        ]["Técnico"].tolist()

        nomes = "<br>".join(
            [t.title() for t in tecnicos]
        )

        with colunas[status]:

            st.markdown(
                f"""
                <div style="
                    background-color:{cores[status]};
                    padding:20px;
                    border-radius:15px;
                    min-height:400px;
                    color:white;
                    text-align:center;
                ">
                    <h1>{status}</h1>
                    <p>{nomes}</p>
                </div>
                """,
                unsafe_allow_html=True
            )

# ==================================================
# QUARTIL MENSAL
# ==================================================

if modo_gestao:

    st.divider()

    st.header("📊 Status Mensal da Operação")

    mes_atual = df["Data"].dt.month.max()

    ano_atual = df["Data"].dt.year.max()

    dados_mes_quartil = df[
        (df["Data"].dt.month == mes_atual)
        &
        (df["Data"].dt.year == ano_atual)
    ]

    status_mes = (
        dados_mes_quartil.groupby("Técnico")
        .agg({
            "Realizado": "sum",
            "Esperado": "sum"
        })
        .reset_index()
    )

    status_mes["Percentual"] = (
        status_mes["Realizado"]
        /
        status_mes["Esperado"]
    ) * 100

    status_mes["Status"] = status_mes[
        "Percentual"
    ].apply(definir_status)

    st.subheader(
        f"📅 Mensal - {mes_atual:02d}/{ano_atual}"
    )

    col1, col2, col3, col4 = st.columns(4)

    colunas = {
        "CRÍTICO": col1,
        "ATENÇÃO": col2,
        "BOM": col3,
        "EXCELENTE": col4
    }

    for status in [
        "CRÍTICO",
        "ATENÇÃO",
        "BOM",
        "EXCELENTE"
    ]:

        tecnicos = status_mes[
            status_mes["Status"] == status
        ]["Técnico"].tolist()

        nomes = "<br>".join(
            [t.title() for t in tecnicos]
        )

        with colunas[status]:

            st.markdown(
                f"""
                <div style="
                    background-color:{cores[status]};
                    padding:20px;
                    border-radius:15px;
                    min-height:400px;
                    color:white;
                    text-align:center;
                ">
                    <h1>{status}</h1>
                    <p>{nomes}</p>
                </div>
                """,
                unsafe_allow_html=True
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
# RANKING DIÁRIO
# ==================================================

st.divider()

if modo_gestao:

    st.subheader("🏆 Ranking Diário Geral por Nível")

    ultima_data = df["Data"].max()

    niveis = sorted(df["Nível"].unique())

    for nivel in niveis:

        st.markdown(f"## 🔹 {nivel}")

        df_dia = df[
            (df["Data"] == ultima_data)
            &
            (df["Nível"] == nivel)
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
            title=f"Ranking Diário - {nivel}",
            color_discrete_sequence=[COR_LARANJA]
        )

        st.plotly_chart(
            grafico_ranking_dia,
            use_container_width=True
        )

else:

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

    st.plotly_chart(
        grafico_ranking_dia,
        use_container_width=True
    )

# ==================================================
# RANKING MENSAL
# ==================================================

st.divider()

if modo_gestao:

    st.subheader("🏆 Ranking Mensal Geral por Nível")

    mes_atual = df["Data"].dt.month.max()
    ano_atual = df["Data"].dt.year.max()

    niveis = sorted(df["Nível"].unique())

    for nivel in niveis:

        st.markdown(f"## 🔹 {nivel}")

        df_mes = df[
            (df["Data"].dt.month == mes_atual)
            &
            (df["Data"].dt.year == ano_atual)
            &
            (df["Nível"] == nivel)
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
            title=f"Ranking Mensal - {nivel}",
            color_discrete_sequence=[COR_CINZA]
        )

        st.plotly_chart(
            grafico_ranking_mes,
            use_container_width=True
        )

else:

    st.subheader(
        f"🏆 Ranking Mensal - {nivel_tecnico}"
    )

    mes_atual = df["Data"].dt.month.max()
    ano_atual = df["Data"].dt.year.max()

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

    st.plotly_chart(
        grafico_ranking_mes,
        use_container_width=True
    )