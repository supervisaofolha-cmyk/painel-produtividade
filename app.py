import streamlit as st
import pandas as pd
import plotly.express as px
import locale
import json
import os
import re
import uuid
import unicodedata
from html.parser import HTMLParser
from io import BytesIO
from copy import deepcopy
from datetime import date, timedelta
from difflib import SequenceMatcher

import openpyxl
import requests

ARQUIVO_PRODUTIVIDADE = "produtividade.xlsx"
ABA_PRODUTIVIDADE = "Produtividade"
ABA_ALIASES = "Aliases"

POWERBI_RESOURCE_KEY = "6b54dc9f-c2f8-4ee5-bbd2-e2ca5781ab06"
POWERBI_API_BASE = "https://wabi-brazil-south-b-primary-api.analysis.windows.net"

SGD_BASE_URL = "https://sgd.dominiosistemas.com.br"
SGD_LOGIN_URL = f"{SGD_BASE_URL}/login"
SGD_RELATORIO_URL = f"{SGD_BASE_URL}/sgsc/faces/rel-satisfacao.html"

MESES_POWERBI = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}


class FormularioSGDParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.no_formulario = False
        self.inputs = []
        self.selects = []
        self._select_atual = None
        self._opcoes_select = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "form" and attrs.get("id") == "formFiltroRelatorio":
            self.no_formulario = True

        if not self.no_formulario:
            return

        if tag == "input":
            self.inputs.append(attrs)
        elif tag == "select":
            self._select_atual = attrs
            self._opcoes_select = []
        elif tag == "option" and self._select_atual is not None:
            self._opcoes_select.append(attrs)

    def handle_endtag(self, tag):
        if self.no_formulario and tag == "select":
            opcao = next(
                (item for item in self._opcoes_select if "selected" in item),
                self._opcoes_select[0] if self._opcoes_select else {"value": ""},
            )
            self.selects.append((self._select_atual, opcao))
            self._select_atual = None
            self._opcoes_select = []

        if tag == "form" and self.no_formulario:
            self.no_formulario = False


def carregar_env_local():
    valores = {}
    if not os.path.exists(".env"):
        return valores

    with open(".env", "r", encoding="utf-8") as arquivo:
        for linha in arquivo:
            linha = linha.strip()
            if not linha or linha.startswith("#") or "=" not in linha:
                continue
            chave, valor = linha.split("=", 1)
            valores[chave.strip()] = valor.strip().strip('"').strip("'")

    return valores


def normalizar_nome(valor):
    texto = str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"\([^)]*\)", " ", texto)
    texto = re.sub(r"\b(de|da|do|das|dos|e)\b", " ", texto)
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def cabecalhos(ws):
    return {
        str(cell.value).strip(): cell.column
        for cell in ws[1]
        if cell.value is not None
    }


def valor_repetido_ou_nulo(linha, indice_coluna, valores, indice_valor, anterior):
    repetidos = int(linha.get("R", 0))
    nulos = int(linha.get("Ø", 0))

    if nulos & (1 << indice_coluna):
        return None, indice_valor

    if repetidos & (1 << indice_coluna):
        return anterior[indice_coluna], indice_valor

    return valores[indice_valor], indice_valor + 1


def decodificar_linhas_powerbi(linhas):
    registros = []
    anterior = [None] * 8

    for linha in linhas:
        valores = linha.get("C", [])
        atual = []
        indice_valor = 0

        for indice_coluna in range(8):
            valor, indice_valor = valor_repetido_ou_nulo(
                linha,
                indice_coluna,
                valores,
                indice_valor,
                anterior,
            )
            atual.append(valor)

        anterior = atual

        if atual[0]:
            registros.append(
                {
                    "agente": atual[0],
                    "atendidas": int(atual[1] or 0),
                    "maior_2min": int(atual[2] or 0),
                    "tma": segundos_para_tma_planilha(atual[7] or 0),
                }
            )

    return registros


def segundos_para_tma_planilha(valor):
    segundos = int(round(float(valor or 0)))
    minutos, resto = divmod(segundos, 60)
    return round(minutos + (resto / 100), 2)


def trimestre_powerbi(data_referencia):
    return f"Trim {((data_referencia.month - 1) // 3) + 1}"


def headers_powerbi():
    return {
        "Accept": "application/json",
        "ActivityId": str(uuid.uuid4()),
        "RequestId": str(uuid.uuid4()),
        "X-PowerBI-ResourceKey": POWERBI_RESOURCE_KEY,
        "User-Agent": "Mozilla/5.0",
        "Origin": "https://app.powerbi.com",
        "Referer": "https://app.powerbi.com/",
        "Content-Type": "application/json",
    }


@st.cache_data(ttl=300, show_spinner=False)
def carregar_metadados_powerbi():
    url = (
        f"{POWERBI_API_BASE}/public/reports/{POWERBI_RESOURCE_KEY}"
        "/modelsAndExploration?preferReadOnlySession=true"
    )
    resposta = requests.get(url, headers=headers_powerbi(), timeout=30)
    resposta.raise_for_status()
    return resposta.json()


def montar_consulta_powerbi(data_referencia):
    metadados = carregar_metadados_powerbi()
    secao_mapa = next(
        secao
        for secao in metadados["exploration"]["sections"]
        if secao.get("displayName") == "Mapa"
    )
    visual = secao_mapa["visualContainers"][0]
    consulta = json.loads(visual["query"])
    comando = consulta["Commands"][0]["SemanticQueryDataShapeCommand"]
    valores_data = comando["Query"]["Where"][0]["Condition"]["In"]["Values"][0]

    valores_data[0]["Literal"]["Value"] = f"{data_referencia.year}L"
    valores_data[1]["Literal"]["Value"] = f"'{trimestre_powerbi(data_referencia)}'"
    valores_data[2]["Literal"]["Value"] = f"'{MESES_POWERBI[data_referencia.month]}'"
    valores_data[3]["Literal"]["Value"] = f"{data_referencia.day}L"

    return {
        "version": "1.0.0",
        "queries": [
            {
                "Query": consulta,
                "ApplicationContext": {
                    "DatasetId": str(metadados["models"][0]["id"]),
                    "Sources": [
                        {
                            "ReportId": metadados["exploration"]["reportId"],
                            "VisualId": json.loads(visual["config"])["name"],
                        }
                    ],
                },
            }
        ],
        "cancelQueries": [],
        "modelId": metadados["models"][0]["id"],
    }


def buscar_produtividade_powerbi(data_referencia):
    corpo = montar_consulta_powerbi(data_referencia)
    resposta = requests.post(
        f"{POWERBI_API_BASE}/public/reports/querydata?synchronous=true",
        headers=headers_powerbi(),
        json=corpo,
        timeout=45,
    )
    resposta.raise_for_status()
    dados = resposta.json()
    linhas = (
        dados["results"][0]["result"]["data"]["dsr"]["DS"][0]["PH"][1]["DM1"]
    )
    return decodificar_linhas_powerbi(linhas)


def melhor_agente(nome_tecnico, agentes_bi):
    nome_normalizado = normalizar_nome(nome_tecnico)
    melhor = None
    melhor_score = 0

    for agente in agentes_bi:
        score = SequenceMatcher(
            None,
            nome_normalizado,
            normalizar_nome(agente),
        ).ratio()
        if score > melhor_score:
            melhor = agente
            melhor_score = score

    if melhor_score >= 0.72:
        return melhor

    return ""


def garantir_estrutura_aliases(ws_aliases):
    headers = ["Técnico Planilha", "Agente BI", "Agente SGD"]
    for indice, header in enumerate(headers, start=1):
        if ws_aliases.cell(row=1, column=indice).value != header:
            ws_aliases.cell(row=1, column=indice).value = header


def garantir_aba_aliases(wb, tecnicos, agentes_bi):
    if ABA_ALIASES not in wb.sheetnames:
        ws_aliases = wb.create_sheet(ABA_ALIASES)
    else:
        ws_aliases = wb[ABA_ALIASES]

    garantir_estrutura_aliases(ws_aliases)

    existentes = {
        normalizar_nome(ws_aliases.cell(row=row, column=1).value)
        for row in range(2, ws_aliases.max_row + 1)
    }

    for tecnico in tecnicos:
        if normalizar_nome(tecnico) not in existentes:
            ws_aliases.append([tecnico, melhor_agente(tecnico, agentes_bi), ""])
            continue

        for row in range(2, ws_aliases.max_row + 1):
            if normalizar_nome(ws_aliases.cell(row=row, column=1).value) == normalizar_nome(tecnico):
                if not ws_aliases.cell(row=row, column=2).value:
                    ws_aliases.cell(row=row, column=2).value = melhor_agente(
                        tecnico,
                        agentes_bi,
                    )
                break

    return ws_aliases


def garantir_aba_aliases_sgd(wb, tecnicos, agentes_sgd):
    if ABA_ALIASES not in wb.sheetnames:
        ws_aliases = wb.create_sheet(ABA_ALIASES)
    else:
        ws_aliases = wb[ABA_ALIASES]

    garantir_estrutura_aliases(ws_aliases)

    existentes = {
        normalizar_nome(ws_aliases.cell(row=row, column=1).value)
        for row in range(2, ws_aliases.max_row + 1)
    }

    for tecnico in tecnicos:
        tecnico_normalizado = normalizar_nome(tecnico)
        if tecnico_normalizado not in existentes:
            ws_aliases.append([tecnico, "", melhor_agente(tecnico, agentes_sgd)])
            continue

        for row in range(2, ws_aliases.max_row + 1):
            if normalizar_nome(ws_aliases.cell(row=row, column=1).value) == tecnico_normalizado:
                if not ws_aliases.cell(row=row, column=3).value:
                    ws_aliases.cell(row=row, column=3).value = melhor_agente(
                        tecnico,
                        agentes_sgd,
                    )
                break

    return ws_aliases


def ler_aliases(ws_aliases, coluna_alias=2):
    aliases = {}
    for row in range(2, ws_aliases.max_row + 1):
        tecnico = ws_aliases.cell(row=row, column=1).value
        agente = ws_aliases.cell(row=row, column=coluna_alias).value
        if tecnico and agente:
            aliases[normalizar_nome(tecnico)] = normalizar_nome(agente)
    return aliases


def atualizar_planilha_com_bi(data_referencia):
    registros_bi = buscar_produtividade_powerbi(data_referencia)
    registros_por_agente = {
        normalizar_nome(registro["agente"]): registro
        for registro in registros_bi
    }

    wb = openpyxl.load_workbook(ARQUIVO_PRODUTIVIDADE)
    ws = wb[ABA_PRODUTIVIDADE]
    colunas = cabecalhos(ws)
    tecnicos = sorted(
        {
            ws.cell(row=row, column=colunas["Técnico"]).value
            for row in range(2, ws.max_row + 1)
            if ws.cell(row=row, column=colunas["Técnico"]).value
        }
    )
    ws_aliases = garantir_aba_aliases(
        wb,
        tecnicos,
        [registro["agente"] for registro in registros_bi],
    )
    aliases = ler_aliases(ws_aliases)

    atualizados = []
    sem_alias = []

    for row in range(2, ws.max_row + 1):
        data_linha = ws.cell(row=row, column=colunas["Data"]).value
        if not data_linha:
            continue

        data_linha = (
            data_linha.date()
            if hasattr(data_linha, "date")
            else data_linha
        )

        if data_linha != data_referencia:
            continue

        tecnico = ws.cell(row=row, column=colunas["Técnico"]).value
        chave_tecnico = normalizar_nome(tecnico)
        chave_agente = aliases.get(chave_tecnico, chave_tecnico)
        registro = registros_por_agente.get(chave_agente)

        if not registro:
            sem_alias.append(tecnico)
            continue

        ws.cell(row=row, column=colunas["Atendidas"]).value = registro["atendidas"]
        ws.cell(row=row, column=colunas[" > 2min"]).value = registro["maior_2min"]
        ws.cell(row=row, column=colunas["TMA"]).value = registro["tma"]
        atualizados.append(tecnico)

    wb.save(ARQUIVO_PRODUTIVIDADE)

    return {
        "data": data_referencia.strftime("%d/%m/%Y"),
        "bi": len(registros_bi),
        "atualizados": len(atualizados),
        "sem_alias": sorted(set(sem_alias)),
    }


def montar_campos_formulario_sgd(html, data_referencia, extensao="EXCEL_XLSX"):
    parser = FormularioSGDParser()
    parser.feed(html)
    dados = []
    submit_name = "formFiltroRelatorio:j_id_jsp_1813582124_61"

    for attrs in parser.inputs:
        nome = attrs.get("name")
        tipo = (attrs.get("type") or "").lower()

        if tipo == "submit" and attrs.get("value", "").startswith("Gerar"):
            submit_name = nome

        if not nome or tipo in {"submit", "button", "image", "reset"}:
            continue

        if tipo == "radio":
            if "checked" in attrs:
                dados.append((nome, attrs.get("value", "")))
        elif tipo == "checkbox":
            if nome not in {
                "formFiltroRelatorio:desconsiderarSscAnexadaSaNe",
                "formFiltroRelatorio:considerarUnidadesAtendidas",
            }:
                dados.append((nome, attrs.get("value", "on")))
        else:
            dados.append((nome, attrs.get("value", "")))

    for select, opcao in parser.selects:
        nome = select.get("name")
        if nome:
            dados.append((nome, opcao.get("value", "")))

    def definir_unico(nome, valor):
        nonlocal dados
        dados = [(chave, item) for chave, item in dados if chave != nome]
        dados.append((nome, valor))

    data_formatada = data_referencia.strftime("%d/%m/%y")
    definir_unico("formFiltroRelatorio:dataInicial", data_formatada)
    definir_unico("formFiltroRelatorio:dataFinal", data_formatada)
    definir_unico("formFiltroRelatorio:extensao", extensao)
    dados.append((submit_name, "Gerar Relatório"))
    return dados


def login_sgd(usuario, senha):
    sessao = requests.Session()
    resposta = sessao.post(
        SGD_LOGIN_URL,
        data={
            "p_chat_plug": "",
            "j_username": usuario,
            "j_password": senha,
            "loginBtn": "Entrar",
        },
        timeout=30,
    )
    resposta.raise_for_status()

    if "login.html" in resposta.url or "invalido.html" in resposta.url:
        raise ValueError("Login do SGD inválido.")

    return sessao


def gerar_relatorio_sgd(sessao, data_referencia):
    resposta = sessao.get(SGD_RELATORIO_URL, timeout=30)
    resposta.raise_for_status()
    html = resposta.text

    link_download = None
    for _ in range(5):
        dados = montar_campos_formulario_sgd(html, data_referencia)
        resposta = sessao.post(SGD_RELATORIO_URL, data=dados, timeout=180)
        resposta.raise_for_status()

        tipo = resposta.headers.get("content-type", "")
        if "spreadsheet" in tipo:
            return resposta.content

        html = resposta.text
        match = re.search(r'href="([^"]+\.xlsx)"', html)
        if match:
            link_download = match.group(1)
            break

    if not link_download:
        raise ValueError("O SGD não gerou o arquivo Excel do relatório.")

    resposta = sessao.get(f"{SGD_BASE_URL}{link_download}", timeout=60)
    resposta.raise_for_status()
    return resposta.content


def extrair_registros_sgd(conteudo_xlsx):
    wb = openpyxl.load_workbook(BytesIO(conteudo_xlsx), data_only=True)
    ws = wb.active
    registros = []

    for row in range(10, ws.max_row + 1):
        tecnico = ws.cell(row=row, column=2).value
        if not tecnico or str(tecnico).strip().lower().startswith("total"):
            continue

        registros.append(
            {
                "tecnico": tecnico,
                "satisfacao": float(ws.cell(row=row, column=5).value or 0),
                "votacao": float(ws.cell(row=row, column=9).value or 0),
                "ssc": int(ws.cell(row=row, column=10).value or 0),
            }
        )

    return registros


def buscar_satisfacao_sgd(data_referencia, usuario, senha):
    sessao = login_sgd(usuario, senha)
    conteudo = gerar_relatorio_sgd(sessao, data_referencia)
    return extrair_registros_sgd(conteudo)


def atualizar_planilha_com_sgd(data_referencia, usuario, senha):
    registros_sgd = buscar_satisfacao_sgd(data_referencia, usuario, senha)
    registros_por_tecnico = {
        normalizar_nome(registro["tecnico"]): registro
        for registro in registros_sgd
    }

    wb = openpyxl.load_workbook(ARQUIVO_PRODUTIVIDADE)
    ws = wb[ABA_PRODUTIVIDADE]
    colunas = cabecalhos(ws)
    tecnicos = sorted(
        {
            ws.cell(row=row, column=colunas["Técnico"]).value
            for row in range(2, ws.max_row + 1)
            if ws.cell(row=row, column=colunas["Técnico"]).value
        }
    )
    ws_aliases = garantir_aba_aliases_sgd(
        wb,
        tecnicos,
        [registro["tecnico"] for registro in registros_sgd],
    )
    aliases = ler_aliases(ws_aliases, coluna_alias=3)

    atualizados = []
    sem_alias = []

    for row in range(2, ws.max_row + 1):
        data_linha = ws.cell(row=row, column=colunas["Data"]).value
        if not data_linha:
            continue

        data_linha = (
            data_linha.date()
            if hasattr(data_linha, "date")
            else data_linha
        )

        if data_linha != data_referencia:
            continue

        tecnico = ws.cell(row=row, column=colunas["Técnico"]).value
        chave_tecnico = normalizar_nome(tecnico)
        chave_sgd = aliases.get(chave_tecnico, chave_tecnico)
        registro = registros_por_tecnico.get(chave_sgd)

        if not registro:
            sem_alias.append(tecnico)
            continue

        ws.cell(row=row, column=colunas["SSC"]).value = registro["ssc"]
        ws.cell(row=row, column=colunas["Satisfação"]).value = registro["satisfacao"]
        ws.cell(row=row, column=colunas["Votação"]).value = registro["votacao"]
        atualizados.append(tecnico)

    wb.save(ARQUIVO_PRODUTIVIDADE)

    return {
        "data": data_referencia.strftime("%d/%m/%Y"),
        "sgd": len(registros_sgd),
        "atualizados": len(atualizados),
        "sem_alias": sorted(set(sem_alias)),
    }


def recalcular_colunas_derivadas(df):
    esperado_por_nivel = {
        "Técnico III": 35,
        "Técnico II": 28,
        "Técnico I": 23,
        "JR": 20,
        "Estágio": 10,
    }

    def status_local(desvio):
        if desvio < -5:
            return "CRÍTICO"
        if desvio < 0:
            return "ATENÇÃO"
        if desvio <= 5:
            return "BOM"
        return "EXCELENTE"

    coluna_2min = " > 2min" if " > 2min" in df.columns else "> 2min"

    for coluna in [coluna_2min, "RO", "CHAT"]:
        if coluna in df.columns:
            df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0)

    df["Realizado"] = df[coluna_2min] + df["RO"] + df["CHAT"]
    df["Esperado"] = df["Nível"].map(esperado_por_nivel).fillna(0)
    df["Desvio"] = df["Realizado"] - df["Esperado"]
    df["Classificação"] = df["Desvio"].apply(status_local)
    return df

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

cores = {
    "CRÍTICO": "#DC2626",
    "ATENÇÃO": "#F97316",
    "BOM": "#2563EB",
    "EXCELENTE": "#16A34A"
}

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
    box-shadow: 0px 2px 5px rgba(0,0,0,0.05);
}}

.card-tecnico {{
    padding:12px;
    border-radius:10px;
    margin-bottom:10px;
    color:white;
    font-weight:bold;
    font-size:16px;
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

df = pd.read_excel(ARQUIVO_PRODUTIVIDADE)
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

df = recalcular_colunas_derivadas(df)

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

    st.sidebar.success("Bem-vinda Gestão")

    data_bi = date.today() - timedelta(days=1)
    env_local = carregar_env_local()
    usuario_sgd_padrao = env_local.get("SGD_USUARIO", os.getenv("SGD_USUARIO", ""))
    senha_sgd_padrao = env_local.get("SGD_SENHA", os.getenv("SGD_SENHA", ""))

    st.sidebar.divider()
    st.sidebar.caption("Atualizações automáticas")

    if st.sidebar.button("Atualizar BI do dia anterior"):
        with st.spinner("Buscando dados no PowerBI e atualizando a planilha..."):
            try:
                resultado_bi = atualizar_planilha_com_bi(data_bi)
            except PermissionError:
                st.sidebar.error(
                    "Feche a planilha produtividade.xlsx no Excel e tente novamente."
                )
            except Exception as erro:
                st.sidebar.error(f"Não foi possível atualizar o BI: {erro}")
            else:
                st.sidebar.success(
                    f"{resultado_bi['atualizados']} técnicos atualizados em "
                    f"{resultado_bi['data']}."
                )
                if resultado_bi["sem_alias"]:
                    st.sidebar.warning(
                        "Revise a aba Aliases para: "
                        + ", ".join(resultado_bi["sem_alias"][:8])
                    )
                st.cache_data.clear()
                st.rerun()

    usuario_sgd = st.sidebar.text_input(
        "Usuário SGD",
        value=usuario_sgd_padrao,
    )
    senha_sgd = st.sidebar.text_input(
        "Senha SGD",
        value=senha_sgd_padrao,
        type="password",
    )

    if st.sidebar.button("Atualizar SGD do dia anterior"):
        if not usuario_sgd or not senha_sgd:
            st.sidebar.error("Informe usuário e senha do SGD.")
        else:
            with st.spinner("Gerando relatório no SGD e atualizando a planilha..."):
                try:
                    resultado_sgd = atualizar_planilha_com_sgd(
                        data_bi,
                        usuario_sgd,
                        senha_sgd,
                    )
                except PermissionError:
                    st.sidebar.error(
                        "Feche a planilha produtividade.xlsx no Excel e tente novamente."
                    )
                except Exception as erro:
                    st.sidebar.error(f"Não foi possível atualizar o SGD: {erro}")
                else:
                    st.sidebar.success(
                        f"{resultado_sgd['atualizados']} técnicos atualizados em "
                        f"{resultado_sgd['data']}."
                    )
                    if resultado_sgd["sem_alias"]:
                        st.sidebar.warning(
                            "Revise a coluna Agente SGD da aba Aliases para: "
                            + ", ".join(resultado_sgd["sem_alias"][:8])
                        )
                    st.cache_data.clear()
                    st.rerun()

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

def definir_status(desvio):

    if desvio < -5:
        return "CRÍTICO"

    elif desvio < 0:
        return "ATENÇÃO"

    elif desvio <= 5:
        return "BOM"

    else:
        return "EXCELENTE"

# ==================================================
# LEGENDA STATUS
# ==================================================

if modo_gestao:

    st.divider()

    st.subheader("📌 Legenda dos Status")

    col1, col2, col3, col4 = st.columns(4)

    legenda = {
        "CRÍTICO": "Desvio menor que -5",
        "ATENÇÃO": "Desvio entre -5 e menor que 0",
        "BOM": "Desvio entre 0 e 5",
        "EXCELENTE": "Desvio acima de 5"
    }

    for status, coluna in zip(
        ["CRÍTICO", "ATENÇÃO", "BOM", "EXCELENTE"],
        [col1, col2, col3, col4]
    ):

        with coluna:

            st.markdown(
                f"""
                <div style="
                    background-color:{cores[status]};
                    padding:15px;
                    border-radius:12px;
                    color:white;
                    text-align:center;
                    min-height:120px;
                ">
                    <h2>{status}</h2>
                    <p style="
                        font-size:18px;
                        font-weight:bold;
                    ">
                        {legenda[status]}
                    </p>
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

# ==================================================
# PEGAR ÚLTIMO DIA COM MOVIMENTAÇÃO DO MÊS
# ==================================================

mes_atual = df["Data"].dt.month.max()
ano_atual = df["Data"].dt.year.max()

dados_mes_atual = dados_tecnico[
    (dados_tecnico["Data"].dt.month == mes_atual)
    &
    (dados_tecnico["Data"].dt.year == ano_atual)
]

ultima_data_mes = dados_mes_atual["Data"].max()

dados_ultimo_dia = dados_mes_atual[
    dados_mes_atual["Data"] == ultima_data_mes
]

# ==================================================
# MÉTRICAS
# ==================================================

col1, col2, col3, col4, col5, col6 = st.columns(6)

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

    votacao_ultimo_dia = round(
        dados_ultimo_dia["Votação"].mean(),
        2
    )

    st.metric(
        "Votação Média",
        f"{votacao_ultimo_dia}%"
    )

with col5:

    satisfacao_ultimo_dia = round(
        dados_ultimo_dia["Satisfação"].mean(),
        2
    )

    st.metric(
        "Satisfação",
        f"{satisfacao_ultimo_dia}%"
    )

with col6:

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
# GRÁFICO PRODUTIVIDADE MENSAL
# ==================================================

st.divider()

st.subheader("📊 Produtividade Mensal")

dados_produtividade = dados_tecnico[
    (dados_tecnico["Data"].dt.month == mes_atual)
    &
    (dados_tecnico["Data"].dt.year == ano_atual)
]

dias_disponiveis = sorted(
    dados_produtividade["Data Formatada"].unique()
)

dias_selecionados = st.multiselect(
    "Selecione os dias",
    dias_disponiveis,
    default=dias_disponiveis
)

dados_produtividade = dados_produtividade[
    dados_produtividade["Data Formatada"]
    .isin(dias_selecionados)
]

produtividade = (
    dados_produtividade.groupby(
        ["Data", "Data Formatada"]
    )
    .agg({
        "Realizado": "sum",
        "Esperado": "sum"
    })
    .reset_index()
)

produtividade = produtividade.sort_values(
    by="Data"
)

produtividade_long = produtividade.melt(
    id_vars=["Data", "Data Formatada"],
    value_vars=["Realizado", "Esperado"],
    var_name="Indicador",
    value_name="Quantidade"
)

grafico_produtividade = px.bar(
    produtividade_long,
    x="Data Formatada",
    y="Quantidade",
    color="Indicador",
    barmode="group",
    text="Quantidade",
    labels={
        "Data Formatada": "Dia",
        "Quantidade": "Quantidade",
        "Indicador": "Indicador"
    },
    title="Produtividade do Mês",
    color_discrete_map={
        "Realizado": COR_LARANJA,
        "Esperado": COR_CINZA
    }
)

grafico_produtividade.update_traces(
    textposition="outside"
)

grafico_produtividade.update_layout(
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis_title="Dia",
    yaxis_title="Quantidade",
    legend_title="Indicadores",
    xaxis=dict(type="category")
)

st.plotly_chart(
    grafico_produtividade,
    use_container_width=True
)
