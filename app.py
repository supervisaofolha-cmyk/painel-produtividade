import csv
import json
import locale
import os
import re
import unicodedata
import uuid
import builtins
from datetime import date, timedelta
from difflib import SequenceMatcher
from html import unescape
from html.parser import HTMLParser
from io import BytesIO, StringIO

import openpyxl
import pandas as pd
import plotly.express as px
import requests
import streamlit as st


ARQUIVO_PRODUTIVIDADE = "produtividade.xlsx"
ARQUIVO_USUARIOS = "usuarios.xlsx"
ABA_PRODUTIVIDADE = "Produtividade"
ABA_ALIASES = "Aliases"

POWERBI_RESOURCE_KEY = "6b54dc9f-c2f8-4ee5-bbd2-e2ca5781ab06"
POWERBI_API_BASE = "https://wabi-brazil-south-b-primary-api.analysis.windows.net"

SGD_BASE_URL = "https://sgd.dominiosistemas.com.br"
SGD_LOGIN_URL = f"{SGD_BASE_URL}/login"
SGD_RELATORIO_URL = f"{SGD_BASE_URL}/sgsc/faces/rel-satisfacao.html"
RO_DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1O-1uJ6D9al9piHgOv_Ju0fpL-3Xbz3zK"

COR_LARANJA = "#F97316"
COR_CINZA = "#6B7280"
COR_BRANCO = "#FFFFFF"
FUNDO = "#F5F5F5"

CORES_STATUS = {
    "CRÍTICO": "#DC2626",
    "ATENÇÃO": "#F97316",
    "BOM": "#2563EB",
    "EXCELENTE": "#16A34A",
}

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

MESES_RO = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Marco",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}

MOTIVOS_RO_CONTAM = {
    "Ligação com duração inferior a 2 minutos/Retorno solicitado pela supervisão",
    "Ligação transferida para outro técnico ou setor (Transferência)",
}

TECNICOS_DESCONSIDERADOS_ESPERADO = {
    "patricia karla de sousa araujo",
    "lorena dias de araujo",
    "lucas luiz romero",
}

PERCENTUAL_ABSORCAO_POR_NIVEL = {
    "Técnico III": 3.23,
    "Técnico II": 2.59,
    "Técnico I": 2.12,
    "JR": 1.85,
    "Estágio": 0.92,
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
    texto = builtins.str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"\([^)]*\)", " ", texto)
    texto = re.sub(r"\b(de|da|do|das|dos|e)\b", " ", texto)
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def cabecalhos(ws):
    return {
        builtins.str(cell.value).strip(): cell.column
        for cell in ws[1]
        if cell.value is not None
    }


def normalizar_cabecalho(valor):
    texto = builtins.str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"\s+", " ", texto)
    return texto


def coluna(colunas, *nomes):
    colunas_normalizadas = {
        normalizar_cabecalho(nome): indice
        for nome, indice in colunas.items()
    }

    for nome in nomes:
        indice = colunas_normalizadas.get(normalizar_cabecalho(nome))
        if indice:
            return indice

    raise KeyError(nomes[0])


def status_por_desvio(desvio):
    if desvio < -5:
        return "CRÍTICO"
    if desvio < 0:
        return "ATENÇÃO"
    if desvio <= 5:
        return "BOM"
    return "EXCELENTE"


def melhor_alias(nome_tecnico, candidatos):
    nome_normalizado = normalizar_nome(nome_tecnico)
    melhor = ""
    melhor_score = 0

    for candidato in candidatos:
        score = SequenceMatcher(
            None,
            nome_normalizado,
            normalizar_nome(candidato),
        ).ratio()
        if score > melhor_score:
            melhor = candidato
            melhor_score = score

    return melhor if melhor_score >= 0.72 else ""


def garantir_estrutura_aliases(ws_aliases):
    headers = ["Técnico Planilha", "Agente BI", "Agente SGD"]
    for indice, header in enumerate(headers, start=1):
        if ws_aliases.cell(row=1, column=indice).value != header:
            ws_aliases.cell(row=1, column=indice).value = header


def tecnicos_da_planilha(ws, colunas):
    return sorted(
        {
            ws.cell(row=row, column=colunas["Técnico"]).value
            for row in range(2, ws.max_row + 1)
            if ws.cell(row=row, column=colunas["Técnico"]).value
        }
    )


def garantir_linhas_da_data(ws, colunas, data_referencia):
    col_data = coluna(colunas, "Data")
    col_tecnico = coluna(colunas, "Técnico")
    col_nivel = coluna(colunas, "Nível")

    linhas_existentes = []
    datas_existentes = set()
    linhas_por_data = {}

    for row in range(2, ws.max_row + 1):
        data_linha = ws.cell(row=row, column=col_data).value
        if not data_linha:
            continue
        data_linha = data_linha.date() if hasattr(data_linha, "date") else data_linha
        datas_existentes.add(data_linha)
        linhas_por_data.setdefault(data_linha, []).append(row)
        if data_linha == data_referencia:
            linhas_existentes.append(row)

    if linhas_existentes:
        return 0

    datas_anteriores = [data for data in datas_existentes if data < data_referencia]
    if not datas_anteriores:
        return 0

    data_base = max(datas_anteriores)
    criadas = 0

    for row_base in linhas_por_data[data_base]:
        tecnico = ws.cell(row=row_base, column=col_tecnico).value
        nivel = ws.cell(row=row_base, column=col_nivel).value
        ws.append(
            [
                data_referencia,
                tecnico,
                nivel,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            ]
        )
        criadas += 1

    return criadas


def garantir_aba_aliases(wb, tecnicos, candidatos, coluna_alias):
    if ABA_ALIASES not in wb.sheetnames:
        ws_aliases = wb.create_sheet(ABA_ALIASES)
    else:
        ws_aliases = wb[ABA_ALIASES]

    garantir_estrutura_aliases(ws_aliases)

    existentes = {
        normalizar_nome(ws_aliases.cell(row=row, column=1).value): row
        for row in range(2, ws_aliases.max_row + 1)
        if ws_aliases.cell(row=row, column=1).value
    }

    for tecnico in tecnicos:
        tecnico_normalizado = normalizar_nome(tecnico)
        if tecnico_normalizado not in existentes:
            ws_aliases.append([tecnico, "", ""])
            row = ws_aliases.max_row
            existentes[tecnico_normalizado] = row
        else:
            row = existentes[tecnico_normalizado]

        if not ws_aliases.cell(row=row, column=coluna_alias).value:
            ws_aliases.cell(row=row, column=coluna_alias).value = melhor_alias(
                tecnico,
                candidatos,
            )

    return ws_aliases


def ler_aliases(ws_aliases, coluna_alias):
    aliases = {}
    for row in range(2, ws_aliases.max_row + 1):
        tecnico = ws_aliases.cell(row=row, column=1).value
        alias = ws_aliases.cell(row=row, column=coluna_alias).value
        if tecnico and alias:
            aliases[normalizar_nome(tecnico)] = normalizar_nome(alias)
    return aliases


def data_dia_anterior():
    data_referencia = date.today() - timedelta(days=1)
    while data_referencia.weekday() >= 5:
        data_referencia -= timedelta(days=1)
    return data_referencia


def inicio_mes(data_referencia):
    return date(data_referencia.year, data_referencia.month, 1)


def datas_no_periodo(data_inicial, data_final):
    atual = data_inicial
    while atual <= data_final:
        yield atual
        atual += timedelta(days=1)


def criar_sessao_http():
    sessao = requests.Session()
    sessao.trust_env = False
    return sessao


@st.cache_data(ttl=300, show_spinner=False)
def listar_arquivos_ro():
    sessao = criar_sessao_http()
    resposta = sessao.get(RO_DRIVE_FOLDER_URL, timeout=30)
    resposta.raise_for_status()

    arquivos = []
    vistos = set()
    for arquivo_id, titulo in re.findall(
        r'data-id="([^"]+)"[^>]*data-tooltip="([^"]+)"',
        resposta.text,
    ):
        if arquivo_id in vistos:
            continue
        vistos.add(arquivo_id)
        arquivos.append(
            {
                "id": arquivo_id,
                "titulo": unescape(titulo),
            }
        )

    return arquivos


def planilha_ro_do_mes(data_referencia):
    esperado = normalizar_cabecalho(
        f"RO {MESES_RO[data_referencia.month]} {data_referencia.year} (respostas)"
    )

    for arquivo in listar_arquivos_ro():
        if esperado in normalizar_cabecalho(arquivo["titulo"]):
            return arquivo

    raise ValueError("Não encontrei a planilha de respostas do RO para esse mês.")


def buscar_ro_forms(data_referencia):
    arquivo = planilha_ro_do_mes(data_referencia)
    sessao = criar_sessao_http()
    resposta = sessao.get(
        f"https://docs.google.com/spreadsheets/d/{arquivo['id']}/export?format=csv",
        timeout=60,
    )
    resposta.raise_for_status()

    texto_csv = resposta.content.decode("utf-8-sig", errors="replace")
    leitor = csv.DictReader(StringIO(texto_csv))
    contagens = {}
    motivos_validos = {
        normalizar_cabecalho(motivo) for motivo in MOTIVOS_RO_CONTAM
    }

    for linha in leitor:
        linha_normalizada = {
            normalizar_cabecalho(chave): (valor or "").strip()
            for chave, valor in linha.items()
            if chave
        }

        tecnico = linha_normalizada.get("nome tecnico", "")
        data_texto = linha_normalizada.get("data", "")
        motivo = linha_normalizada.get("ligacao inferior a 2 min", "")

        if (
            not tecnico
            or not data_texto
            or normalizar_cabecalho(motivo) not in motivos_validos
        ):
            continue

        data_linha = pd.to_datetime(data_texto, dayfirst=True, errors="coerce")
        if pd.isna(data_linha) or data_linha.date() != data_referencia:
            continue

        chave = normalizar_nome(tecnico)
        contagens[chave] = contagens.get(chave, 0) + 1

    return {
        "arquivo": arquivo["titulo"],
        "contagens": contagens,
    }


def headers_powerbi():
    return {
        "Accept": "application/json",
        "ActivityId": builtins.str(uuid.uuid4()),
        "RequestId": builtins.str(uuid.uuid4()),
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
    sessao = criar_sessao_http()
    resposta = sessao.get(url, headers=headers_powerbi(), timeout=30)
    resposta.raise_for_status()
    return resposta.json()


def trimestre_powerbi(data_referencia):
    return f"Trim {((data_referencia.month - 1) // 3) + 1}"


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
                    "DatasetId": builtins.str(metadados["models"][0]["id"]),
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


def valor_powerbi(linha, indice_coluna, valores, indice_valor, anterior):
    repetidos = int(linha.get("R", 0))
    nulos = int(linha.get("Ø", 0))

    if nulos & (1 << indice_coluna):
        return None, indice_valor
    if repetidos & (1 << indice_coluna):
        return anterior[indice_coluna], indice_valor
    return valores[indice_valor], indice_valor + 1


def segundos_para_tma_planilha(valor):
    segundos = int(round(float(valor or 0)))
    minutos, resto = divmod(segundos, 60)
    return round(minutos + (resto / 100), 2)


def decodificar_linhas_powerbi(linhas):
    registros = []
    anterior = [None] * 8

    for linha in linhas:
        valores = linha.get("C", [])
        atual = []
        indice_valor = 0

        for indice_coluna in range(8):
            valor, indice_valor = valor_powerbi(
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


def buscar_produtividade_powerbi(data_referencia):
    sessao = criar_sessao_http()
    resposta = sessao.post(
        f"{POWERBI_API_BASE}/public/reports/querydata?synchronous=true",
        headers=headers_powerbi(),
        json=montar_consulta_powerbi(data_referencia),
        timeout=45,
    )
    resposta.raise_for_status()
    dados = resposta.json()
    linhas = (
        dados["results"][0]["result"]["data"]["dsr"]["DS"][0]["PH"][1]["DM1"]
    )
    return decodificar_linhas_powerbi(linhas)


def atualizar_planilha_com_bi(data_referencia):
    registros_bi = buscar_produtividade_powerbi(data_referencia)
    registros_por_agente = {
        normalizar_nome(registro["agente"]): registro
        for registro in registros_bi
    }

    wb = openpyxl.load_workbook(ARQUIVO_PRODUTIVIDADE)
    ws = wb[ABA_PRODUTIVIDADE]
    colunas = cabecalhos(ws)
    col_data = coluna(colunas, "Data")
    col_tecnico = coluna(colunas, "Técnico")
    col_atendidas = coluna(colunas, "Atendidas")
    col_2min = coluna(colunas, " > 2min", "> 2min", ">2min")
    col_tma = coluna(colunas, "TMA")
    linhas_criadas = garantir_linhas_da_data(ws, colunas, data_referencia)
    tecnicos = tecnicos_da_planilha(ws, colunas)
    ws_aliases = garantir_aba_aliases(
        wb,
        tecnicos,
        [registro["agente"] for registro in registros_bi],
        coluna_alias=2,
    )
    aliases = ler_aliases(ws_aliases, coluna_alias=2)

    atualizados = []
    sem_alias = []

    for row in range(2, ws.max_row + 1):
        data_linha = ws.cell(row=row, column=col_data).value
        if not data_linha:
            continue
        data_linha = data_linha.date() if hasattr(data_linha, "date") else data_linha
        if data_linha != data_referencia:
            continue

        tecnico = ws.cell(row=row, column=col_tecnico).value
        ws.cell(row=row, column=col_atendidas).value = 0
        ws.cell(row=row, column=col_2min).value = 0
        ws.cell(row=row, column=col_tma).value = 0
        chave_agente = aliases.get(normalizar_nome(tecnico), normalizar_nome(tecnico))
        registro = registros_por_agente.get(chave_agente)

        if not registro:
            sem_alias.append(tecnico)
            continue

        ws.cell(row=row, column=col_atendidas).value = registro["atendidas"]
        ws.cell(row=row, column=col_2min).value = registro["maior_2min"]
        ws.cell(row=row, column=col_tma).value = registro["tma"]
        atualizados.append(tecnico)

    wb.save(ARQUIVO_PRODUTIVIDADE)

    return {
        "data": data_referencia.strftime("%d/%m/%Y"),
        "fonte": len(registros_bi),
        "linhas_criadas": linhas_criadas,
        "atualizados": len(atualizados),
        "sem_alias": sorted(set(sem_alias)),
    }


def login_sgd(usuario, senha):
    sessao = criar_sessao_http()
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


def montar_campos_sgd(html, data_inicial, data_final, extensao="EXCEL_XLSX"):
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

    definir_unico("formFiltroRelatorio:dataInicial", data_inicial.strftime("%d/%m/%y"))
    definir_unico("formFiltroRelatorio:dataFinal", data_final.strftime("%d/%m/%y"))
    definir_unico("formFiltroRelatorio:extensao", extensao)
    dados.append((submit_name, "Gerar Relatório"))
    return dados


def gerar_relatorio_sgd(sessao, data_inicial, data_final):
    resposta = sessao.get(SGD_RELATORIO_URL, timeout=30)
    resposta.raise_for_status()
    html = resposta.text

    link_download = None
    for _ in range(10):
        dados = montar_campos_sgd(html, data_inicial, data_final)
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
        if not tecnico or builtins.str(tecnico).strip().lower().startswith("total"):
            continue

        registros.append(
            {
                "tecnico": tecnico,
                "satisfacao": float(ws.cell(row=row, column=5).value or 0),
                "votacao": float(ws.cell(row=row, column=11).value or 0),
                "ssc": int(ws.cell(row=row, column=10).value or 0),
            }
        )

    return registros


def buscar_satisfacao_sgd(data_inicial, data_final, usuario, senha):
    sessao = login_sgd(usuario, senha)
    conteudo = gerar_relatorio_sgd(sessao, data_inicial, data_final)
    return extrair_registros_sgd(conteudo)


def buscar_satisfacao_sgd_diaria(data_referencia, usuario, senha):
    sessao = login_sgd(usuario, senha)
    registros = []

    if data_referencia.weekday() >= 5:
        return registros

    try:
        conteudo = gerar_relatorio_sgd(sessao, data_referencia, data_referencia)
    except ValueError:
        return registros

    for registro in extrair_registros_sgd(conteudo):
        registro["data"] = data_referencia
        registros.append(registro)

    return registros


def atualizar_planilha_com_sgd(data_referencia, usuario, senha):
    data_inicial = inicio_mes(data_referencia)
    registros_periodo = buscar_satisfacao_sgd(
        data_inicial,
        data_referencia,
        usuario,
        senha,
    )
    registros_sgd = buscar_satisfacao_sgd_diaria(
        data_referencia,
        usuario,
        senha,
    )
    registros_periodo_por_tecnico = {
        normalizar_nome(registro["tecnico"]): registro
        for registro in registros_periodo
    }
    registros_por_chave = {
        (normalizar_nome(registro["tecnico"]), registro["data"]): registro
        for registro in registros_sgd
    }

    wb = openpyxl.load_workbook(ARQUIVO_PRODUTIVIDADE)
    ws = wb[ABA_PRODUTIVIDADE]
    colunas = cabecalhos(ws)
    col_data = coluna(colunas, "Data")
    col_tecnico = coluna(colunas, "Técnico")
    col_ssc = coluna(colunas, "SSC")
    col_satisfacao = coluna(colunas, "Satisfação")
    col_votacao = coluna(colunas, "Votação")
    linhas_criadas = garantir_linhas_da_data(ws, colunas, data_referencia)
    tecnicos = tecnicos_da_planilha(ws, colunas)
    ws_aliases = garantir_aba_aliases(
        wb,
        tecnicos,
        [registro["tecnico"] for registro in registros_sgd],
        coluna_alias=3,
    )
    aliases = ler_aliases(ws_aliases, coluna_alias=3)

    atualizados = []
    sem_alias = []

    for row in range(2, ws.max_row + 1):
        data_linha = ws.cell(row=row, column=col_data).value
        if not data_linha:
            continue
        data_linha = data_linha.date() if hasattr(data_linha, "date") else data_linha
        if data_linha != data_referencia:
            continue

        tecnico = ws.cell(row=row, column=col_tecnico).value
        chave_sgd = aliases.get(normalizar_nome(tecnico), normalizar_nome(tecnico))
        registro_periodo = registros_periodo_por_tecnico.get(chave_sgd)
        registro = registros_por_chave.get((chave_sgd, data_linha))

        if not registro_periodo and not registro:
            ws.cell(row=row, column=col_ssc).value = 0
            ws.cell(row=row, column=col_satisfacao).value = 0
            ws.cell(row=row, column=col_votacao).value = 0
            sem_alias.append(tecnico)
            continue

        ws.cell(row=row, column=col_ssc).value = 0 if not registro else registro["ssc"]
        ws.cell(row=row, column=col_satisfacao).value = (
            0 if not registro_periodo else registro_periodo["satisfacao"]
        )
        ws.cell(row=row, column=col_votacao).value = (
            0 if not registro_periodo else registro_periodo["votacao"]
        )
        atualizados.append(f"{tecnico} - {data_linha}")

    wb.save(ARQUIVO_PRODUTIVIDADE)

    return {
        "data": data_referencia.strftime("%d/%m/%Y"),
        "periodo": f"{data_inicial.strftime('%d/%m/%Y')} a {data_referencia.strftime('%d/%m/%Y')}",
        "fonte": len(registros_sgd),
        "dias_processados": 1,
        "linhas_criadas": linhas_criadas,
        "atualizados": len(atualizados),
        "sem_alias": sorted(set(sem_alias)),
    }


def atualizar_planilha_com_ro(data_referencia):
    dados_ro = buscar_ro_forms(data_referencia)
    contagens = dados_ro["contagens"]

    wb = openpyxl.load_workbook(ARQUIVO_PRODUTIVIDADE)
    ws = wb[ABA_PRODUTIVIDADE]
    colunas = cabecalhos(ws)
    col_data = coluna(colunas, "Data")
    col_tecnico = coluna(colunas, "Técnico")
    col_ro = coluna(colunas, "RO")
    linhas_criadas = garantir_linhas_da_data(ws, colunas, data_referencia)

    atualizados = 0

    for row in range(2, ws.max_row + 1):
        data_linha = ws.cell(row=row, column=col_data).value
        if not data_linha:
            continue
        data_linha = data_linha.date() if hasattr(data_linha, "date") else data_linha
        if data_linha != data_referencia:
            continue

        tecnico = ws.cell(row=row, column=col_tecnico).value
        ws.cell(row=row, column=col_ro).value = contagens.get(
            normalizar_nome(tecnico),
            0,
        )
        atualizados += 1

    wb.save(ARQUIVO_PRODUTIVIDADE)

    return {
        "data": data_referencia.strftime("%d/%m/%Y"),
        "arquivo": dados_ro["arquivo"],
        "fonte": sum(contagens.values()),
        "linhas_criadas": linhas_criadas,
        "atualizados": atualizados,
    }


def recalcular_colunas_derivadas(df):
    esperado_por_nivel = {
        "Técnico III": 35,
        "Técnico II": 28,
        "Técnico I": 23,
        "JR": 20,
        "Estágio": 10,
    }

    coluna_2min = " > 2min" if " > 2min" in df.columns else "> 2min"
    for coluna in [coluna_2min, "RO", "CHAT"]:
        if coluna in df.columns:
            df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0)

    if "Atendidas" in df.columns:
        df["Atendidas"] = pd.to_numeric(df["Atendidas"], errors="coerce").fillna(0)

    df["Realizado"] = df[coluna_2min] + df["RO"] + df["CHAT"]
    meta_por_linha = df["Nível"].map(esperado_por_nivel).fillna(0)
    tecnicos_normalizados = df["Técnico"].apply(normalizar_nome)
    atendidas_validas = df["Atendidas"].where(
        ~tecnicos_normalizados.isin(TECNICOS_DESCONSIDERADOS_ESPERADO),
        0,
    )
    total_atendidas_dia = atendidas_validas.groupby(df["Data"]).transform("sum")
    proporcional = total_atendidas_dia * (meta_por_linha / 1083)
    df["Esperado"] = proporcional.clip(upper=meta_por_linha).round(0)
    df["Desvio"] = df["Realizado"] - df["Esperado"]
    df["Classificação"] = df["Desvio"].apply(status_por_desvio)
    return df


def precisa_atualizar_colunas(dataframe, data_referencia, colunas_verificacao):
    dados_dia = dataframe[dataframe["Data"] == pd.Timestamp(data_referencia)]
    if dados_dia.empty:
        return True

    for nome_coluna in colunas_verificacao:
        if nome_coluna not in dados_dia.columns:
            return True
        if dados_dia[nome_coluna].isna().any():
            return True

    return False


def atualizar_dados_automaticamente():
    if st.session_state.get("auto_refresh_executado"):
        return

    st.session_state["auto_refresh_executado"] = True
    data_referencia = data_dia_anterior()
    env_local = carregar_env_local()
    usuario_sgd = env_local.get("SGD_USUARIO", os.getenv("SGD_USUARIO", ""))
    senha_sgd = env_local.get("SGD_SENHA", os.getenv("SGD_SENHA", ""))

    try:
        df_atual = pd.read_excel(ARQUIVO_PRODUTIVIDADE)
        df_atual.columns = df_atual.columns.str.strip()
        df_atual["Data"] = pd.to_datetime(df_atual["Data"], dayfirst=True, errors="coerce")
        df_atual = df_atual.dropna(subset=["Data"])
    except Exception:
        return

    try:
        if precisa_atualizar_colunas(
            df_atual,
            data_referencia,
            ["Atendidas", "> 2min" if "> 2min" in df_atual.columns else " > 2min", "TMA"],
        ):
            atualizar_planilha_com_bi(data_referencia)
    except Exception:
        pass

    try:
        if precisa_atualizar_colunas(df_atual, data_referencia, ["RO"]):
            atualizar_planilha_com_ro(data_referencia)
    except Exception:
        pass

    try:
        if usuario_sgd and senha_sgd and precisa_atualizar_colunas(
            df_atual,
            data_referencia,
            ["SSC", "Satisfação", "Votação"],
        ):
            atualizar_planilha_com_sgd(data_referencia, usuario_sgd, senha_sgd)
    except Exception:
        pass


try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
except Exception:
    pass

st.set_page_config(page_title="Painel de Produtividade", layout="wide")
atualizar_dados_automaticamente()

st.markdown(
    f"""
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
""",
    unsafe_allow_html=True,
)

st.title("📊 Painel de Produtividade")

df = pd.read_excel(ARQUIVO_PRODUTIVIDADE)
usuarios = pd.read_excel(ARQUIVO_USUARIOS)

df.columns = df.columns.str.strip()
usuarios.columns = usuarios.columns.str.strip().str.lower()

df["Técnico"] = df["Técnico"].astype(str).str.lower().str.strip()
usuarios["usuario"] = usuarios["usuario"].astype(str).str.lower().str.strip()
usuarios["senha"] = (
    usuarios["senha"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)
usuarios["tecnico"] = usuarios["tecnico"].astype(str).str.lower().str.strip()

df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Data"])
df["Data Formatada"] = df["Data"].dt.strftime("%d/%m/%Y")
for coluna_numerica in ["SSC", "Satisfação", "Votação"]:
    if coluna_numerica in df.columns:
        df[coluna_numerica] = pd.to_numeric(
            df[coluna_numerica], errors="coerce"
        ).fillna(0)
df = recalcular_colunas_derivadas(df)

st.sidebar.title("🔐 Login")
usuario_input = st.sidebar.text_input("Usuário")
senha_input = st.sidebar.text_input("Senha", type="password")

if usuario_input == "" or senha_input == "":
    st.warning("Digite usuário e senha.")
    st.stop()

usuario_digitado = builtins.str(usuario_input).lower().strip()
senha_digitada = builtins.str(senha_input).replace(".0", "").strip()

modo_gestao = False

if usuario_digitado == "gestao" and senha_digitada == "30071997":
    modo_gestao = True
    st.sidebar.success("Bem-vinda Gestão")

    data_referencia = data_dia_anterior()
    env_local = carregar_env_local()
    usuario_sgd_padrao = env_local.get("SGD_USUARIO", os.getenv("SGD_USUARIO", ""))
    senha_sgd_padrao = env_local.get("SGD_SENHA", os.getenv("SGD_SENHA", ""))

    st.sidebar.divider()
    st.sidebar.caption(f"Atualização para {data_referencia.strftime('%d/%m/%Y')}")

    if st.sidebar.button("Atualizar BI do dia anterior"):
        with st.spinner("Buscando dados no PowerBI e atualizando a planilha..."):
            try:
                resultado = atualizar_planilha_com_bi(data_referencia)
            except PermissionError:
                st.sidebar.error("Feche a produtividade.xlsx no Excel e tente novamente.")
            except Exception as erro:
                st.sidebar.error(f"Não foi possível atualizar o BI: {erro}")
            else:
                st.sidebar.success(
                    f"{resultado['atualizados']} técnicos atualizados em {resultado['data']}."
                )
                if resultado["linhas_criadas"]:
                    st.sidebar.info(
                        f"{resultado['linhas_criadas']} linhas foram criadas para essa data."
                    )
                if resultado["sem_alias"]:
                    st.sidebar.warning(
                        "Revise a coluna Agente BI da aba Aliases para: "
                        + ", ".join(resultado["sem_alias"][:8])
                    )
                st.cache_data.clear()
                st.rerun()

    if st.sidebar.button("Atualizar RO do dia anterior"):
        with st.spinner("Buscando dados do RO e atualizando a planilha..."):
            try:
                resultado = atualizar_planilha_com_ro(data_referencia)
            except PermissionError:
                st.sidebar.error("Feche a produtividade.xlsx no Excel e tente novamente.")
            except Exception as erro:
                st.sidebar.error(f"Não foi possível atualizar o RO: {erro}")
            else:
                st.sidebar.success(
                    f"{resultado['atualizados']} técnicos atualizados em {resultado['data']}."
                )
                st.sidebar.info(
                    f"Origem: {resultado['arquivo']}. Total de RO válidos: {resultado['fonte']}."
                )
                if resultado["linhas_criadas"]:
                    st.sidebar.info(
                        f"{resultado['linhas_criadas']} linhas foram criadas para essa data."
                    )
                st.cache_data.clear()
                st.rerun()

    usuario_sgd = st.sidebar.text_input("Usuário SGD", value=usuario_sgd_padrao)
    senha_sgd = st.sidebar.text_input(
        "Senha SGD",
        value=senha_sgd_padrao,
        type="password",
    )

    if st.sidebar.button("Atualizar SGD do mês até ontem"):
        if not usuario_sgd or not senha_sgd:
            st.sidebar.error("Informe usuário e senha do SGD.")
        else:
            with st.spinner("Gerando relatório no SGD e atualizando a planilha..."):
                try:
                    resultado = atualizar_planilha_com_sgd(
                        data_referencia,
                        usuario_sgd,
                        senha_sgd,
                    )
                except PermissionError:
                    st.sidebar.error("Feche a produtividade.xlsx no Excel e tente novamente.")
                except Exception as erro:
                    st.sidebar.error(f"Não foi possível atualizar o SGD: {erro}")
                else:
                    st.sidebar.success(
                        f"{resultado['atualizados']} registros atualizados em "
                        f"{resultado['dias_processados']} dias. "
                        f"Período: {resultado['periodo']}."
                    )
                    if resultado["linhas_criadas"]:
                        st.sidebar.info(
                            f"{resultado['linhas_criadas']} linhas foram criadas para essa data."
                        )
                    if resultado["sem_alias"]:
                        st.sidebar.warning(
                            "Revise a coluna Agente SGD da aba Aliases para: "
                            + ", ".join(resultado["sem_alias"][:8])
                        )
                    st.cache_data.clear()
                    st.rerun()

else:
    login = usuarios[
        (usuarios["usuario"] == usuario_digitado)
        & (usuarios["senha"] == senha_digitada)
    ]

    if login.empty:
        st.error("Usuário ou senha inválidos.")
        st.stop()

    tecnico = login.iloc[0]["tecnico"]
    st.sidebar.success(f"Bem-vindo(a), {tecnico.title()}")

if modo_gestao:
    tecnico = st.selectbox("Selecione o Técnico", sorted(df["Técnico"].unique()))

dados_tecnico = df[df["Técnico"] == tecnico]

if dados_tecnico.empty:
    st.error("Nenhum dado encontrado.")
    st.stop()

mes_atual = df["Data"].dt.month.max()
ano_atual = df["Data"].dt.year.max()

dados_mes_atual = dados_tecnico[
    (dados_tecnico["Data"].dt.month == mes_atual)
    & (dados_tecnico["Data"].dt.year == ano_atual)
]

ultima_data_mes = dados_mes_atual["Data"].max()
dados_ultimo_dia = dados_mes_atual[dados_mes_atual["Data"] == ultima_data_mes]


def montar_grafico_ranking(ranking, titulo, eixo_x):
    ranking = ranking.copy()
    ranking["Posição"] = ranking.index + 1
    ranking["Técnico Exibição"] = ranking["Técnico"].str.title()
    ranking["Cor"] = ranking["Técnico"].apply(
        lambda nome: "Selecionado" if nome == tecnico else "Mesmo nível"
    )

    grafico = px.bar(
        ranking,
        x="Técnico Exibição",
        y="Realizado",
        color="Cor",
        text="Realizado",
        labels={
            "Realizado": eixo_x,
            "Técnico Exibição": "Técnico",
            "Cor": "",
        },
        title=titulo,
        color_discrete_map={
            "Selecionado": COR_LARANJA,
            "Mesmo nível": COR_CINZA,
        },
        category_orders={
            "Técnico Exibição": ranking["Técnico Exibição"].tolist()
        },
    )

    grafico.update_traces(textposition="outside")
    grafico.update_layout(
        plot_bgcolor=COR_BRANCO,
        paper_bgcolor=COR_BRANCO,
        font_color=COR_CINZA,
        xaxis_title="Técnico",
        yaxis_title=eixo_x,
        legend_title="",
    )
    return grafico

if modo_gestao:
    st.divider()
    st.subheader("📌 Legenda dos Status")
    col1, col2, col3, col4 = st.columns(4)
    legenda = {
        "CRÍTICO": "Desvio menor que -5",
        "ATENÇÃO": "Desvio entre -5 e menor que 0",
        "BOM": "Desvio entre 0 e 5",
        "EXCELENTE": "Desvio acima de 5",
    }

    for status, coluna in zip(
        ["CRÍTICO", "ATENÇÃO", "BOM", "EXCELENTE"],
        [col1, col2, col3, col4],
    ):
        with coluna:
            st.markdown(
                f"""
                <div style="
                    background-color:{CORES_STATUS[status]};
                    padding:15px;
                    border-radius:12px;
                    color:white;
                    text-align:center;
                    min-height:120px;
                ">
                    <h2>{status}</h2>
                    <p style="font-size:18px;font-weight:bold;">
                        {legenda[status]}
                    </p>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.divider()
    st.subheader("Percentual de Absorção por Nível")
    col1, col2, col3, col4, col5 = st.columns(5)
    for (nivel, percentual), coluna in zip(
        PERCENTUAL_ABSORCAO_POR_NIVEL.items(),
        [col1, col2, col3, col4, col5],
    ):
        with coluna:
            st.metric(nivel, f"{percentual:.2f}%")

    st.divider()
    data_status = df["Data"].max()
    st.subheader(
        f"Técnicos por Status em {data_status.strftime('%d/%m/%Y')}"
    )

    status_atual = (
        df[df["Data"] == data_status][["Técnico", "Nível", "Classificação"]]
        .dropna(subset=["Técnico", "Classificação"])
        .sort_values(by=["Classificação", "Técnico"])
    )

    col1, col2, col3, col4 = st.columns(4)
    for status, coluna in zip(
        ["CRÍTICO", "ATENÇÃO", "BOM", "EXCELENTE"],
        [col1, col2, col3, col4],
    ):
        grupo_status = status_atual[status_atual["Classificação"] == status]
        with coluna:
            conteudo = [f"<strong>{status}</strong>"]
            if grupo_status.empty:
                conteudo.append("Nenhum técnico")
            else:
                for _, linha in grupo_status.iterrows():
                    conteudo.append(
                        f"{builtins.str(linha['Técnico']).title()} - {linha['Nível']}"
                    )

            st.markdown(
                f"""
                <div style="
                    background-color:{CORES_STATUS[status]};
                    padding:14px;
                    border-radius:12px;
                    color:white;
                    min-height:240px;
                ">
                    {'<br>'.join(conteudo)}
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.divider()
    st.subheader("Ranking Geral de Todos os Níveis")

    niveis_disponiveis = sorted(
        [nivel for nivel in df["Nível"].dropna().unique() if builtins.str(nivel).strip()]
    )
    nivel_ranking_geral = st.selectbox(
        "Selecione o nível do ranking geral",
        ["Todos os níveis"] + niveis_disponiveis,
    )

    base_ranking_geral = df[
        (df["Data"].dt.month == mes_atual)
        & (df["Data"].dt.year == ano_atual)
    ]

    if nivel_ranking_geral != "Todos os níveis":
        base_ranking_geral = base_ranking_geral[
            base_ranking_geral["Nível"] == nivel_ranking_geral
        ]

    ranking_geral_diario = (
        base_ranking_geral[base_ranking_geral["Data"] == ultima_data_mes]
        .groupby("Técnico", as_index=False)["Realizado"]
        .sum()
        .sort_values(by="Realizado", ascending=False)
        .reset_index(drop=True)
    )

    ranking_geral_mensal = (
        base_ranking_geral.groupby("Técnico", as_index=False)["Realizado"]
        .sum()
        .sort_values(by="Realizado", ascending=False)
        .reset_index(drop=True)
    )

    aba_geral_diario, aba_geral_mensal = st.tabs(["Diário", "Mensal"])

    with aba_geral_diario:
        st.plotly_chart(
            montar_grafico_ranking(
                ranking_geral_diario,
                f"Ranking Diário - {nivel_ranking_geral}",
                "Produtividade do dia",
            ),
            use_container_width=True,
        )

    with aba_geral_mensal:
        st.plotly_chart(
            montar_grafico_ranking(
                ranking_geral_mensal,
                f"Ranking Mensal - {nivel_ranking_geral}",
                "Produtividade do mês",
            ),
            use_container_width=True,
        )

st.divider()
st.subheader(f"📌 Resultados Individuais - {tecnico.title()}")

nivel_mode = dados_mes_atual["Nível"].dropna().mode()
nivel_tecnico = None if nivel_mode.empty else nivel_mode.iloc[0]

if modo_gestao:
    col1, col2, col3, col4, col5, col6 = st.columns(6)
else:
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)

with col1:
    st.metric("Realizado Total", int(dados_mes_atual["Realizado"].sum()))

with col2:
    st.metric("SSC Total", int(dados_mes_atual["SSC"].sum()))

with col3:
    st.metric("RO Total", int(dados_mes_atual["RO"].sum()))

with col4:
    votacao_ultimo_dia = round(dados_ultimo_dia["Votação"].mean(), 2)
    st.metric("Votação Média", f"{votacao_ultimo_dia}%")

with col5:
    satisfacao_ultimo_dia = round(dados_ultimo_dia["Satisfação"].mean(), 2)
    st.metric("Satisfação", f"{satisfacao_ultimo_dia}%")

with col6:
    classificacao_mode = dados_mes_atual["Classificação"].dropna().mode()
    classificacao = (
        "Sem classificação"
        if classificacao_mode.empty
        else classificacao_mode.iloc[0]
    )
    st.metric("Classificação", classificacao)

if not modo_gestao:
    percentual_absorcao = PERCENTUAL_ABSORCAO_POR_NIVEL.get(nivel_tecnico, 0)
    with col7:
        st.metric("Percentual de Absorção", f"{percentual_absorcao:.2f}%")

if not modo_gestao:
    if nivel_tecnico:
        base_ranking_nivel = df[
            (df["Data"].dt.month == mes_atual)
            & (df["Data"].dt.year == ano_atual)
            & (df["Nível"] == nivel_tecnico)
        ]

        ranking_nivel_diario = (
            base_ranking_nivel[base_ranking_nivel["Data"] == ultima_data_mes]
            .groupby("Técnico", as_index=False)["Realizado"]
            .sum()
            .sort_values(by="Realizado", ascending=False)
            .reset_index(drop=True)
        )

        ranking_nivel_mensal = (
            base_ranking_nivel.groupby("Técnico", as_index=False)["Realizado"]
            .sum()
            .sort_values(by="Realizado", ascending=False)
            .reset_index(drop=True)
        )

        st.divider()
        st.subheader(f"Ranking do Seu Nível - {nivel_tecnico}")
        aba_nivel_diario, aba_nivel_mensal = st.tabs(["Diário", "Mensal"])

        with aba_nivel_diario:
            st.plotly_chart(
                montar_grafico_ranking(
                    ranking_nivel_diario,
                    f"Ranking Diário - {nivel_tecnico}",
                    "Produtividade do dia",
                ),
                use_container_width=True,
            )

        with aba_nivel_mensal:
            st.plotly_chart(
                montar_grafico_ranking(
                    ranking_nivel_mensal,
                    f"Ranking Mensal - {nivel_tecnico}",
                    "Produtividade do mês",
                ),
                use_container_width=True,
            )

st.divider()
st.subheader("📊 Produtividade Mensal")

dados_produtividade = dados_tecnico[
    (dados_tecnico["Data"].dt.month == mes_atual)
    & (dados_tecnico["Data"].dt.year == ano_atual)
]

dias_disponiveis = sorted(dados_produtividade["Data Formatada"].unique())
dias_selecionados = st.multiselect(
    "Selecione os dias",
    dias_disponiveis,
    default=dias_disponiveis,
)

dados_produtividade = dados_produtividade[
    dados_produtividade["Data Formatada"].isin(dias_selecionados)
]

produtividade = (
    dados_produtividade.groupby(["Data", "Data Formatada"])
    .agg({"Realizado": "sum", "Esperado": "sum"})
    .reset_index()
    .sort_values(by="Data")
)

produtividade_long = produtividade.melt(
    id_vars=["Data", "Data Formatada"],
    value_vars=["Realizado", "Esperado"],
    var_name="Indicador",
    value_name="Quantidade",
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
        "Indicador": "Indicador",
    },
    title="Produtividade do Mês",
    color_discrete_map={
        "Realizado": COR_LARANJA,
        "Esperado": COR_CINZA,
    },
)

grafico_produtividade.update_traces(textposition="outside")
grafico_produtividade.update_layout(
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis_title="Dia",
    yaxis_title="Quantidade",
    legend_title="Indicadores",
    xaxis=dict(type="category"),
)

st.plotly_chart(grafico_produtividade, use_container_width=True)

st.divider()
st.subheader("SSC Atendido por Dia")

ssc_diario = (
    dados_produtividade.groupby(["Data", "Data Formatada"])["SSC"]
    .sum()
    .reset_index()
    .sort_values(by="Data")
)

grafico_ssc = px.line(
    ssc_diario,
    x="Data Formatada",
    y="SSC",
    markers=True,
    labels={
        "Data Formatada": "Dia",
        "SSC": "SSC Atendido",
    },
    title="SSC Atendido no Mês",
)

grafico_ssc.update_traces(line_color=COR_LARANJA, marker_color=COR_LARANJA)
grafico_ssc.update_layout(
    plot_bgcolor=COR_BRANCO,
    paper_bgcolor=COR_BRANCO,
    font_color=COR_CINZA,
    xaxis_title="Dia",
    yaxis_title="SSC Atendido",
    xaxis=dict(type="category"),
)

st.plotly_chart(grafico_ssc, use_container_width=True)
