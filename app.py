import csv
import base64
import json
import locale
import os
import pathlib
import re
import shutil
import sqlite3
import tempfile
import time
import unicodedata
import uuid
import zipfile
import builtins
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher
from html import unescape
from html.parser import HTMLParser
from io import BytesIO, StringIO
import calendar
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

import openpyxl
import pandas as pd
import plotly.express as px
import requests
import streamlit as st

try:
    import psycopg
    from psycopg.rows import dict_row
except Exception:
    psycopg = None
    dict_row = None


ARQUIVO_PRODUTIVIDADE = "produtividade.xlsx"
ARQUIVO_PRODUTIVIDADE_BACKUP = "produtividade_backup.xlsx"
ARQUIVO_USUARIOS = "usuarios.xlsx"
ARQUIVO_LISTA_APOIO = "lista_apoio.xlsx"
ARQUIVO_LISTA_APOIO_BACKUP = "lista_apoio_backup.xlsx"
ARQUIVO_LISTA_APOIO_DB = "lista_apoio.db"
ENV_LISTA_APOIO_DATABASE_URL = "LISTA_APOIO_DATABASE_URL"
LISTA_APOIO_DATABASE_URL_PADRAO = "postgresql://postgres.pogiuykdubgvcjjzjihs:TgKyzWZhAXJjczXd@aws-1-us-west-2.pooler.supabase.com:6543/postgres?sslmode=require"
ARQUIVO_LISTA_APOIO_HISTORICO = "lista_apoio_historico.jsonl"
ARQUIVO_META_BACKUP_LISTA_APOIO = "lista_apoio_backup_meta.json"
PASTA_BACKUP_LISTA_APOIO = "backups_lista_apoio"
ARQUIVO_LOG_BACKUP_PAINEL = "painel_backup_db_erro.log"
ABA_PRODUTIVIDADE = "Produtividade"
ABA_ALIASES = "Aliases"
COLUNA_DIAS_META = "Dias Meta"
COLUNA_ABANDONADAS = "Abandonadas"
CABECALHOS_LISTA_APOIO = [
    "Carimbo de data/hora",
    "Nome do Técnico",
    "Selecione o tópico",
    "Descreva o problema/pedido em poucas palavras",
    "Situação",
    "Responsável/Apoio",
]
TOPICOS_LISTA_APOIO = [
    "13° Salário",
    "Integração",
    "Provisão de Férias",
    "Provisão de 13°",
    "Crédito em conta",
    "Fórmula",
    "Configuração de Rubrica",
    "Médias",
    "Afastamentos",
    "Rescisão",
    "Férias",
    "Folha Mensal",
    "Folha Complementar",
    "Alteração Cadastral",
    "Alteração Salarial",
    "DCTFWEB",
    "FGTS Digital",
    "Tranferencia",
    "Informativos",
    "Cadastro",
    "e-Social",
    "Outros",
]
USUARIOS_APOIO = {"subbrenda", "subluma"}
FUSO_HORARIO_APP = ZoneInfo("America/Araguaina")
PLUG_CHATBOX_URL = "https://tr.plugsocial.com.br/#/app/chatbox"
PLUG_PERFIL_DIR = pathlib.Path(".playwright-plug")
PLUG_GRUPO_FOLHA = "Folha de Pagamento"
PLUG_STATUS_FECHADO = "Fechado"

POWERBI_RESOURCE_KEY = "6b54dc9f-c2f8-4ee5-bbd2-e2ca5781ab06"
POWERBI_API_BASE = "https://wabi-brazil-south-b-primary-api.analysis.windows.net"
X2_CONTROLLER_BASE_URL = "http://192.168.1.252/x2-controller"
POWERBI_FILA_FALLBACK = "Folha - FGTS/DCTF/ESOCIAL"

SGD_BASE_URL = "https://sgd.dominiosistemas.com.br"
SGD_LOGIN_URL = f"{SGD_BASE_URL}/login"
SGD_RELATORIO_URL = f"{SGD_BASE_URL}/sgsc/faces/rel-satisfacao.html"
SGD_RELATORIO_TIMEOUT = 300
SGD_RELATORIO_TENTATIVAS = 3
RO_DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1O-1uJ6D9al9piHgOv_Ju0fpL-3Xbz3zK"
RO_PLANILHAS_FIXAS = {
    (2026, 6): "https://docs.google.com/spreadsheets/d/1AZKhjWnIS_hys9B9oaeazcSUP0wS6ExX8U9wRdrktRs/edit?resourcekey#gid=148357412",
}
PONTOWEB_BASE_URL = "https://pontoweb.secullum.com.br/"
PONTOWEB_CLIENT_ID = "3001"
PONTOWEB_REDIRECT_URI = f"{PONTOWEB_BASE_URL}Auth"
PONTOWEB_AUTH_URL = (
    "https://autenticador.secullum.com.br/Authorization"
    f"?response_type=code&client_id={PONTOWEB_CLIENT_ID}"
    f"&redirect_uri={PONTOWEB_REDIRECT_URI}"
)
PONTOWEB_TOKEN_URL = "https://autenticador.secullum.com.br/Token"
SERVICO_PLANILHA_LOCAL_URL = "http://127.0.0.1:8765"

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
    "Retorno solicitado pela Supervisão",
    "Ligação transferida para outro técnico ou setor (Transferência)",
}

TECNICOS_DESCONSIDERADOS_ESPERADO = {
    "patricia karla sousa araujo",
    "lorena dias araujo",
    "lucas luiz romero",
}

META_ESPERADA_POR_NIVEL = {
    "Técnico III": 35,
    "Técnico II": 28,
    "Técnico I": 23,
    "JR": 20,
    "Estágio": 10,
}

FERIADOS_FEDERAIS_FIXOS = {
    (1, 1),
    (4, 21),
    (5, 1),
    (9, 7),
    (10, 12),
    (11, 2),
    (11, 15),
    (11, 20),
    (12, 25),
}

FERIADOS_ADICIONAIS = {
    date(2026, 6, 4),
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


def arquivo_excel_integro(caminho):
    if not os.path.exists(caminho):
        return False

    try:
        with zipfile.ZipFile(caminho) as arquivo_zip:
            return arquivo_zip.testzip() is None
    except Exception:
        return False


def garantir_planilha_produtividade_integra():
    principal_ok = arquivo_excel_integro(ARQUIVO_PRODUTIVIDADE)
    backup_ok = arquivo_excel_integro(ARQUIVO_PRODUTIVIDADE_BACKUP)

    if principal_ok:
        if not backup_ok:
            shutil.copyfile(ARQUIVO_PRODUTIVIDADE, ARQUIVO_PRODUTIVIDADE_BACKUP)
        return

    if backup_ok:
        shutil.copyfile(ARQUIVO_PRODUTIVIDADE_BACKUP, ARQUIVO_PRODUTIVIDADE)
        return

    raise zipfile.BadZipFile(
        "Nenhuma cópia íntegra da produtividade.xlsx foi encontrada."
    )


def ler_dataframe_produtividade():
    garantir_planilha_produtividade_integra()
    return pd.read_excel(ARQUIVO_PRODUTIVIDADE)


def carregar_workbook_produtividade(**kwargs):
    garantir_planilha_produtividade_integra()
    return openpyxl.load_workbook(ARQUIVO_PRODUTIVIDADE, **kwargs)


def salvar_workbook_produtividade(wb):
    garantir_planilha_produtividade_integra()
    diretorio = os.path.dirname(os.path.abspath(ARQUIVO_PRODUTIVIDADE)) or "."
    descritor, caminho_temporario = tempfile.mkstemp(
        prefix="produtividade_",
        suffix=".xlsx",
        dir=diretorio,
    )
    os.close(descritor)
    try:
        wb.save(caminho_temporario)
        with zipfile.ZipFile(caminho_temporario) as arquivo_zip:
            if arquivo_zip.testzip() is not None:
                raise zipfile.BadZipFile("Arquivo temporário gerado com corrupção.")
        os.replace(caminho_temporario, ARQUIVO_PRODUTIVIDADE)
        shutil.copyfile(ARQUIVO_PRODUTIVIDADE, ARQUIVO_PRODUTIVIDADE_BACKUP)
        tentar_salvar_backup_arquivo_painel_no_banco(ARQUIVO_PRODUTIVIDADE)
    finally:
        if os.path.exists(caminho_temporario):
            os.remove(caminho_temporario)


def garantir_planilha_lista_apoio_integra():
    principal_existe = os.path.exists(ARQUIVO_LISTA_APOIO)
    principal_ok = arquivo_excel_integro(ARQUIVO_LISTA_APOIO)
    backup_ok = arquivo_excel_integro(ARQUIVO_LISTA_APOIO_BACKUP)

    if principal_ok:
        if not backup_ok:
            shutil.copyfile(ARQUIVO_LISTA_APOIO, ARQUIVO_LISTA_APOIO_BACKUP)
        return

    if backup_ok:
        shutil.copyfile(ARQUIVO_LISTA_APOIO_BACKUP, ARQUIVO_LISTA_APOIO)
        return

    if principal_existe:
        raise zipfile.BadZipFile(
            "Nenhuma cópia íntegra da lista_apoio.xlsx foi encontrada."
        )


def chave_registro_lista_apoio(registro):
    coluna_carimbo = CABECALHOS_LISTA_APOIO[0]
    coluna_tecnico = CABECALHOS_LISTA_APOIO[1]
    coluna_topico = CABECALHOS_LISTA_APOIO[2]
    coluna_descricao = CABECALHOS_LISTA_APOIO[3]
    return "||".join(
        [
            texto_lista_apoio(valor_campo_registro_lista_apoio(registro, coluna_carimbo)),
            texto_lista_apoio(valor_campo_registro_lista_apoio(registro, coluna_tecnico)),
            texto_lista_apoio(valor_campo_registro_lista_apoio(registro, coluna_topico)),
            texto_lista_apoio(
                valor_campo_registro_lista_apoio(
                    registro,
                    coluna_descricao,
                )
            ),
        ]
    )


def dataframe_lista_apoio_vazio():
    return pd.DataFrame(columns=CABECALHOS_LISTA_APOIO)


def valor_campo_registro_lista_apoio(registro, coluna_esperada):
    if coluna_esperada in registro:
        return registro.get(coluna_esperada, "")

    chave_esperada = normalizar_cabecalho(coluna_esperada)
    for chave, valor in registro.items():
        if normalizar_cabecalho(chave) == chave_esperada:
            return valor
    return ""


def padronizar_dataframe_lista_apoio(dataframe):
    if dataframe is None or dataframe.empty:
        return dataframe_lista_apoio_vazio()

    dataframe = dataframe.copy()
    for coluna_lista in CABECALHOS_LISTA_APOIO:
        if coluna_lista not in dataframe.columns:
            chave_esperada = normalizar_cabecalho(coluna_lista)
            coluna_equivalente = next(
                (
                    nome_coluna
                    for nome_coluna in dataframe.columns
                    if normalizar_cabecalho(nome_coluna) == chave_esperada
                ),
                None,
            )
            if coluna_equivalente is not None:
                dataframe[coluna_lista] = dataframe[coluna_equivalente]
            else:
                dataframe[coluna_lista] = ""
        dataframe[coluna_lista] = dataframe[coluna_lista].astype("object")

    dataframe = dataframe[CABECALHOS_LISTA_APOIO]
    colunas_obrigatorias = [
        "Carimbo de data/hora",
        "Nome do Técnico",
        "Selecione o tópico",
        "Descreva o problema/pedido em poucas palavras",
    ]
    for coluna_obrigatoria in colunas_obrigatorias:
        dataframe[coluna_obrigatoria] = dataframe[coluna_obrigatoria].map(
            texto_lista_apoio
        )

    filtro_valido = dataframe["Carimbo de data/hora"].ne("")
    filtro_valido &= dataframe["Nome do Técnico"].ne("")
    filtro_valido &= dataframe["Selecione o tópico"].ne("")
    filtro_valido &= dataframe["Descreva o problema/pedido em poucas palavras"].ne("")
    dataframe = dataframe.loc[filtro_valido].reset_index(drop=True)

    return dataframe


def ler_dataframe_lista_apoio_arquivo(caminho):
    if not os.path.exists(caminho) or not arquivo_excel_integro(caminho):
        return dataframe_lista_apoio_vazio()

    try:
        dataframe = pd.read_excel(caminho)
    except Exception:
        return dataframe_lista_apoio_vazio()

    return padronizar_dataframe_lista_apoio(dataframe)


def registrar_historico_lista_apoio(acao, registro):
    os.makedirs(PASTA_BACKUP_LISTA_APOIO, exist_ok=True)
    payload = {
        "timestamp": datetime.now(FUSO_HORARIO_APP).strftime("%Y-%m-%d %H:%M:%S"),
        "acao": acao,
        "registro": {
            coluna: texto_lista_apoio(registro.get(coluna, ""))
            for coluna in CABECALHOS_LISTA_APOIO
        },
    }
    with open(ARQUIVO_LISTA_APOIO_HISTORICO, "a", encoding="utf-8") as arquivo:
        arquivo.write(json.dumps(payload, ensure_ascii=False) + "\n")


def restaurar_lista_apoio_do_historico():
    if not os.path.exists(ARQUIVO_LISTA_APOIO_HISTORICO):
        return dataframe_lista_apoio_vazio()

    registros = {}
    with open(ARQUIVO_LISTA_APOIO_HISTORICO, "r", encoding="utf-8") as arquivo:
        for linha in arquivo:
            linha = linha.strip()
            if not linha:
                continue
            try:
                evento = json.loads(linha)
            except Exception:
                continue
            registro = evento.get("registro") or {}
            chave = chave_registro_lista_apoio(registro)
            if not chave:
                continue
            registros[chave] = {
                coluna: texto_lista_apoio(
                    valor_campo_registro_lista_apoio(registro, coluna)
                )
                for coluna in CABECALHOS_LISTA_APOIO
            }

    if not registros:
        return dataframe_lista_apoio_vazio()

    dataframe = pd.DataFrame(list(registros.values()))
    return padronizar_dataframe_lista_apoio(dataframe)


def mesclar_dataframes_lista_apoio(dataframes):
    registros = {}
    for dataframe in dataframes:
        dataframe = padronizar_dataframe_lista_apoio(dataframe)
        if dataframe.empty:
            continue
        for _, linha in dataframe.iterrows():
            registro = {
                coluna: texto_lista_apoio(linha.get(coluna, ""))
                for coluna in CABECALHOS_LISTA_APOIO
            }
            chave = chave_registro_lista_apoio(registro)
            if not chave:
                continue
            registros[chave] = registro

    if not registros:
        return dataframe_lista_apoio_vazio()

    return padronizar_dataframe_lista_apoio(pd.DataFrame(list(registros.values())))


def database_url_lista_apoio():
    env_local = carregar_env_local()
    segredo = ""
    try:
        segredo = builtins.str(st.secrets.get(ENV_LISTA_APOIO_DATABASE_URL, "") or "").strip()
    except Exception:
        segredo = ""
    url = builtins.str(
        segredo
        or os.getenv(
            ENV_LISTA_APOIO_DATABASE_URL,
            env_local.get(ENV_LISTA_APOIO_DATABASE_URL, ""),
        )
        or LISTA_APOIO_DATABASE_URL_PADRAO
        or ""
    ).strip()
    return url


def backend_lista_apoio():
    url = database_url_lista_apoio()
    if not url:
        return "sqlite"

    esquema = urlparse(url).scheme.lower()
    if esquema.startswith("postgres"):
        return "postgres"
    return "sqlite"


def conexao_lista_apoio_db():
    backend = backend_lista_apoio()
    if backend == "postgres":
        if psycopg is None:
            raise RuntimeError(
                "Banco online da Lista de Apoio configurado, mas a dependência psycopg não está instalada."
            )
        return psycopg.connect(
            database_url_lista_apoio(),
            row_factory=dict_row,
            prepare_threshold=None,
        )

    conexao = sqlite3.connect(ARQUIVO_LISTA_APOIO_DB)
    conexao.row_factory = sqlite3.Row
    return conexao


def valor_backup_painel(valor):
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    return valor


def texto_chave_backup_painel(valor):
    valor = valor_backup_painel(valor)
    return builtins.str(valor or "").strip()


def garantir_banco_backup_painel():
    with conexao_lista_apoio_db() as conexao:
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS painel_dados_backup (
                id TEXT PRIMARY KEY,
                origem TEXT NOT NULL,
                aba TEXT NOT NULL,
                chave TEXT NOT NULL,
                payload TEXT NOT NULL,
                criado_em TEXT NOT NULL,
                atualizado_em TEXT NOT NULL,
                UNIQUE(origem, aba, chave)
            )
            """
        )
        conexao.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_painel_dados_backup_origem_aba
            ON painel_dados_backup (origem, aba)
            """
        )
        conexao.commit()


def registros_backup_painel_do_workbook(wb):
    registros = []
    abas_permitidas = {ABA_PRODUTIVIDADE, ABA_ALIASES}

    for nome_aba in wb.sheetnames:
        if nome_aba not in abas_permitidas:
            continue

        ws = wb[nome_aba]
        cabecalhos_aba = [
            texto_chave_backup_painel(celula.value)
            for celula in ws[1]
        ]

        for row in range(2, ws.max_row + 1):
            valores = [
                valor_backup_painel(ws.cell(row=row, column=coluna).value)
                for coluna in range(1, len(cabecalhos_aba) + 1)
            ]
            if not any(texto_chave_backup_painel(valor) for valor in valores):
                continue

            payload = {
                cabecalho: valor
                for cabecalho, valor in zip(cabecalhos_aba, valores)
                if cabecalho
            }

            if nome_aba == ABA_PRODUTIVIDADE:
                chave = "||".join(
                    [
                        texto_chave_backup_painel(payload.get("Data")),
                        texto_chave_backup_painel(payload.get("Técnico")),
                    ]
                )
            elif nome_aba == ABA_ALIASES:
                chave = texto_chave_backup_painel(
                    payload.get("Técnico Planilha")
                    or payload.get("T??cnico Planilha")
                    or payload.get("TÃ©cnico Planilha")
                )
            else:
                chave = ""

            if not chave or chave == "||":
                chave = f"linha_{row}"

            registros.append(
                {
                    "origem": ARQUIVO_PRODUTIVIDADE,
                    "aba": nome_aba,
                    "chave": chave,
                    "payload": json.dumps(payload, ensure_ascii=False, default=builtins.str),
                }
            )

    return registros


def salvar_backup_painel_no_banco(wb):
    registros = registros_backup_painel_do_workbook(wb)
    if not registros:
        return 0

    garantir_banco_backup_painel()
    backend = backend_lista_apoio()
    agora = datetime.now(FUSO_HORARIO_APP).strftime("%Y-%m-%d %H:%M:%S")
    consulta = """
        INSERT INTO painel_dados_backup (
            id,
            origem,
            aba,
            chave,
            payload,
            criado_em,
            atualizado_em
        ) VALUES (%s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT(origem, aba, chave) DO UPDATE SET
            payload = excluded.payload,
            atualizado_em = excluded.atualizado_em
    """
    if backend != "postgres":
        consulta = """
            INSERT INTO painel_dados_backup (
                id,
                origem,
                aba,
                chave,
                payload,
                criado_em,
                atualizado_em
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(origem, aba, chave) DO UPDATE SET
                payload = excluded.payload,
                atualizado_em = excluded.atualizado_em
        """

    with conexao_lista_apoio_db() as conexao:
        for registro in registros:
            identificador = builtins.str(
                uuid.uuid5(
                    uuid.NAMESPACE_URL,
                    f"{registro['origem']}|{registro['aba']}|{registro['chave']}",
                )
            )
            conexao.execute(
                consulta,
                (
                    identificador,
                    registro["origem"],
                    registro["aba"],
                    registro["chave"],
                    registro["payload"],
                    agora,
                    agora,
                ),
            )
        conexao.commit()

    return len(registros)


def registrar_erro_backup_painel(erro):
    try:
        with open(ARQUIVO_LOG_BACKUP_PAINEL, "a", encoding="utf-8") as arquivo:
            timestamp = datetime.now(FUSO_HORARIO_APP).strftime("%Y-%m-%d %H:%M:%S")
            arquivo.write(f"{timestamp} - {erro}\n")
    except Exception:
        pass


def tentar_salvar_backup_painel_no_banco(wb):
    try:
        return salvar_backup_painel_no_banco(wb)
    except Exception as erro:
        registrar_erro_backup_painel(erro)
        return 0


def garantir_banco_arquivo_backup_painel():
    with conexao_lista_apoio_db() as conexao:
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS painel_arquivo_backup (
                origem TEXT PRIMARY KEY,
                conteudo_base64 TEXT NOT NULL,
                tamanho_bytes INTEGER NOT NULL,
                atualizado_em TEXT NOT NULL
            )
            """
        )
        conexao.commit()


def salvar_backup_arquivo_painel_no_banco(caminho_arquivo):
    if not os.path.exists(caminho_arquivo):
        return 0

    garantir_banco_arquivo_backup_painel()
    origem = os.path.basename(caminho_arquivo)
    with open(caminho_arquivo, "rb") as arquivo:
        conteudo = arquivo.read()

    conteudo_base64 = base64.b64encode(conteudo).decode("ascii")
    agora = datetime.now(FUSO_HORARIO_APP).strftime("%Y-%m-%d %H:%M:%S")
    backend = backend_lista_apoio()
    consulta = """
        INSERT INTO painel_arquivo_backup (
            origem,
            conteudo_base64,
            tamanho_bytes,
            atualizado_em
        ) VALUES (%s, %s, %s, %s)
        ON CONFLICT(origem) DO UPDATE SET
            conteudo_base64 = excluded.conteudo_base64,
            tamanho_bytes = excluded.tamanho_bytes,
            atualizado_em = excluded.atualizado_em
    """
    if backend != "postgres":
        consulta = """
            INSERT INTO painel_arquivo_backup (
                origem,
                conteudo_base64,
                tamanho_bytes,
                atualizado_em
            ) VALUES (?, ?, ?, ?)
            ON CONFLICT(origem) DO UPDATE SET
                conteudo_base64 = excluded.conteudo_base64,
                tamanho_bytes = excluded.tamanho_bytes,
                atualizado_em = excluded.atualizado_em
        """

    with conexao_lista_apoio_db() as conexao:
        conexao.execute(
            consulta,
            (origem, conteudo_base64, len(conteudo), agora),
        )
        conexao.commit()

    return len(conteudo)


def tentar_salvar_backup_arquivo_painel_no_banco(caminho_arquivo):
    try:
        return salvar_backup_arquivo_painel_no_banco(caminho_arquivo)
    except Exception as erro:
        registrar_erro_backup_painel(erro)
        return 0


def garantir_banco_ro_backup():
    with conexao_lista_apoio_db() as conexao:
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS ro_backup (
                id TEXT PRIMARY KEY,
                data_referencia TEXT NOT NULL,
                arquivo TEXT NOT NULL,
                arquivo_id TEXT NOT NULL,
                origem TEXT NOT NULL,
                url TEXT NOT NULL,
                csv_base64 TEXT NOT NULL,
                contagens_json TEXT NOT NULL,
                tecnicos_origem_json TEXT NOT NULL,
                criado_em TEXT NOT NULL,
                atualizado_em TEXT NOT NULL,
                UNIQUE(data_referencia, arquivo_id)
            )
            """
        )
        conexao.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_ro_backup_data
            ON ro_backup (data_referencia)
            """
        )
        conexao.commit()


def salvar_backup_ro_no_banco(data_referencia, arquivo, texto_csv, contagens, tecnicos_origem):
    garantir_banco_ro_backup()
    backend = backend_lista_apoio()
    data_texto = data_referencia.strftime("%Y-%m-%d")
    agora = datetime.now(FUSO_HORARIO_APP).strftime("%Y-%m-%d %H:%M:%S")
    identificador = builtins.str(
        uuid.uuid5(
            uuid.NAMESPACE_URL,
            f"ro|{data_texto}|{arquivo.get('id', '')}",
        )
    )
    csv_base64 = base64.b64encode(texto_csv.encode("utf-8")).decode("ascii")
    consulta = """
        INSERT INTO ro_backup (
            id,
            data_referencia,
            arquivo,
            arquivo_id,
            origem,
            url,
            csv_base64,
            contagens_json,
            tecnicos_origem_json,
            criado_em,
            atualizado_em
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT(data_referencia, arquivo_id) DO UPDATE SET
            arquivo = excluded.arquivo,
            origem = excluded.origem,
            url = excluded.url,
            csv_base64 = excluded.csv_base64,
            contagens_json = excluded.contagens_json,
            tecnicos_origem_json = excluded.tecnicos_origem_json,
            atualizado_em = excluded.atualizado_em
    """
    if backend != "postgres":
        consulta = """
            INSERT INTO ro_backup (
                id,
                data_referencia,
                arquivo,
                arquivo_id,
                origem,
                url,
                csv_base64,
                contagens_json,
                tecnicos_origem_json,
                criado_em,
                atualizado_em
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(data_referencia, arquivo_id) DO UPDATE SET
                arquivo = excluded.arquivo,
                origem = excluded.origem,
                url = excluded.url,
                csv_base64 = excluded.csv_base64,
                contagens_json = excluded.contagens_json,
                tecnicos_origem_json = excluded.tecnicos_origem_json,
                atualizado_em = excluded.atualizado_em
        """

    with conexao_lista_apoio_db() as conexao:
        conexao.execute(
            consulta,
            (
                identificador,
                data_texto,
                arquivo.get("titulo", ""),
                arquivo.get("id", ""),
                arquivo.get("origem", ""),
                arquivo.get("url", ""),
                csv_base64,
                json.dumps(contagens, ensure_ascii=False, sort_keys=True),
                json.dumps(tecnicos_origem, ensure_ascii=False),
                agora,
                agora,
            ),
        )
        conexao.commit()


def tentar_salvar_backup_ro_no_banco(data_referencia, arquivo, texto_csv, contagens, tecnicos_origem):
    try:
        salvar_backup_ro_no_banco(
            data_referencia,
            arquivo,
            texto_csv,
            contagens,
            tecnicos_origem,
        )
        return True
    except Exception as erro:
        registrar_erro_backup_painel(erro)
        return False


def garantir_banco_lista_apoio():
    with conexao_lista_apoio_db() as conexao:
        conexao.execute(
            """
            CREATE TABLE IF NOT EXISTS lista_apoio (
                id TEXT PRIMARY KEY,
                chave TEXT NOT NULL UNIQUE,
                carimbo_data_hora TEXT NOT NULL,
                nome_tecnico TEXT NOT NULL,
                topico TEXT NOT NULL,
                descricao TEXT NOT NULL,
                situacao TEXT NOT NULL,
                responsavel_apoio TEXT NOT NULL,
                criado_em TEXT NOT NULL,
                atualizado_em TEXT NOT NULL
            )
            """
        )
        conexao.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_lista_apoio_carimbo
            ON lista_apoio (carimbo_data_hora)
            """
        )
        conexao.commit()


def dataframe_para_registros_lista_apoio(dataframe):
    dataframe = padronizar_dataframe_lista_apoio(dataframe)
    registros = []
    for _, linha in dataframe.iterrows():
        registro = {
            coluna: texto_lista_apoio(linha.get(coluna, ""))
            for coluna in CABECALHOS_LISTA_APOIO
        }
        if chave_registro_lista_apoio(registro):
            registros.append(registro)
    return registros


def deduplicar_dataframe_lista_apoio(dataframe):
    dataframe = padronizar_dataframe_lista_apoio(dataframe)
    if dataframe.empty:
        return dataframe

    registros_por_chave = {}
    for _, linha in dataframe.iterrows():
        registro = {
            coluna: texto_lista_apoio(linha.get(coluna, ""))
            for coluna in CABECALHOS_LISTA_APOIO
        }
        chave = chave_registro_lista_apoio(registro)
        if not chave:
            continue
        registros_por_chave[chave] = registro

    if not registros_por_chave:
        return dataframe_lista_apoio_vazio()

    return padronizar_dataframe_lista_apoio(
        pd.DataFrame(list(registros_por_chave.values()))
    )


def espelhar_lista_apoio_para_arquivos(dataframe):
    dataframe = padronizar_dataframe_lista_apoio(dataframe)
    salvar_dataframe_lista_apoio(dataframe)


def tentar_espelhar_lista_apoio_para_arquivos(dataframe):
    try:
        espelhar_lista_apoio_para_arquivos(dataframe)
    except Exception:
        return False
    return True


def ler_lista_apoio_do_banco():
    garantir_banco_lista_apoio()
    backend = backend_lista_apoio()
    coluna_carimbo = CABECALHOS_LISTA_APOIO[0]
    coluna_tecnico = CABECALHOS_LISTA_APOIO[1]
    coluna_topico = CABECALHOS_LISTA_APOIO[2]
    coluna_descricao = CABECALHOS_LISTA_APOIO[3]
    coluna_situacao = CABECALHOS_LISTA_APOIO[4]
    coluna_responsavel = CABECALHOS_LISTA_APOIO[5]
    consulta = """
        SELECT
            carimbo_data_hora,
            nome_tecnico,
            topico,
            descricao,
            situacao,
            responsavel_apoio
        FROM lista_apoio
    """
    if backend == "postgres":
        consulta += """
        ORDER BY to_timestamp(carimbo_data_hora, 'DD/MM/YYYY HH24:MI:SS') ASC,
                 criado_em ASC
        """
    else:
        consulta += """
        ORDER BY datetime(substr(carimbo_data_hora, 7, 4) || '-' || substr(carimbo_data_hora, 4, 2) || '-' || substr(carimbo_data_hora, 1, 2) || ' ' || substr(carimbo_data_hora, 12, 8)) ASC,
                 rowid ASC
        """
    with conexao_lista_apoio_db() as conexao:
        linhas = conexao.execute(consulta).fetchall()

    if not linhas:
        return dataframe_lista_apoio_vazio()

    registros = []
    for linha in linhas:
        registros.append(
            {
                coluna_carimbo: texto_lista_apoio(linha["carimbo_data_hora"]),
                coluna_tecnico: texto_lista_apoio(linha["nome_tecnico"]),
                coluna_topico: texto_lista_apoio(linha["topico"]),
                coluna_descricao: texto_lista_apoio(linha["descricao"]),
                coluna_situacao: texto_lista_apoio(linha["situacao"]),
                coluna_responsavel: texto_lista_apoio(linha["responsavel_apoio"]),
            }
        )
    return deduplicar_dataframe_lista_apoio(pd.DataFrame(registros))
def salvar_registros_lista_apoio_no_banco(registros):
    garantir_banco_lista_apoio()
    backend = backend_lista_apoio()
    agora = datetime.now(FUSO_HORARIO_APP).strftime("%Y-%m-%d %H:%M:%S")
    coluna_carimbo = CABECALHOS_LISTA_APOIO[0]
    coluna_tecnico = CABECALHOS_LISTA_APOIO[1]
    coluna_topico = CABECALHOS_LISTA_APOIO[2]
    coluna_descricao = CABECALHOS_LISTA_APOIO[3]
    coluna_situacao = CABECALHOS_LISTA_APOIO[4]
    coluna_responsavel = CABECALHOS_LISTA_APOIO[5]
    consulta_busca = "SELECT id FROM lista_apoio WHERE chave = %s"
    consulta_upsert = """
                INSERT INTO lista_apoio (
                    id,
                    chave,
                    carimbo_data_hora,
                    nome_tecnico,
                    topico,
                    descricao,
                    situacao,
                    responsavel_apoio,
                    criado_em,
                    atualizado_em
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT(chave) DO UPDATE SET
                    carimbo_data_hora = excluded.carimbo_data_hora,
                    nome_tecnico = excluded.nome_tecnico,
                    topico = excluded.topico,
                    descricao = excluded.descricao,
                    situacao = excluded.situacao,
                    responsavel_apoio = excluded.responsavel_apoio,
                    atualizado_em = excluded.atualizado_em
                """
    if backend != "postgres":
        consulta_busca = "SELECT id FROM lista_apoio WHERE chave = ?"
        consulta_upsert = """
                INSERT INTO lista_apoio (
                    id,
                    chave,
                    carimbo_data_hora,
                    nome_tecnico,
                    topico,
                    descricao,
                    situacao,
                    responsavel_apoio,
                    criado_em,
                    atualizado_em
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(chave) DO UPDATE SET
                    carimbo_data_hora = excluded.carimbo_data_hora,
                    nome_tecnico = excluded.nome_tecnico,
                    topico = excluded.topico,
                    descricao = excluded.descricao,
                    situacao = excluded.situacao,
                    responsavel_apoio = excluded.responsavel_apoio,
                    atualizado_em = excluded.atualizado_em
                """
    with conexao_lista_apoio_db() as conexao:
        for registro in registros:
            chave = chave_registro_lista_apoio(registro)
            if not chave:
                continue
            identificador_existente = conexao.execute(
                consulta_busca,
                (chave,),
            ).fetchone()
            identificador = (
                identificador_existente["id"] if identificador_existente else builtins.str(uuid.uuid4())
            )
            conexao.execute(
                consulta_upsert,
                (
                    identificador,
                    chave,
                    texto_lista_apoio(valor_campo_registro_lista_apoio(registro, coluna_carimbo)),
                    texto_lista_apoio(valor_campo_registro_lista_apoio(registro, coluna_tecnico)),
                    texto_lista_apoio(valor_campo_registro_lista_apoio(registro, coluna_topico)),
                    texto_lista_apoio(
                        valor_campo_registro_lista_apoio(
                            registro,
                            coluna_descricao,
                        )
                    ),
                    texto_lista_apoio(valor_campo_registro_lista_apoio(registro, coluna_situacao)) or "Aberto",
                    texto_lista_apoio(valor_campo_registro_lista_apoio(registro, coluna_responsavel)),
                    agora,
                    agora,
                ),
            )
        conexao.commit()
    return


def deduplicar_lista_apoio_no_banco():
    garantir_banco_lista_apoio()
    backend = backend_lista_apoio()
    consulta = """
        SELECT
            id,
            chave,
            carimbo_data_hora,
            nome_tecnico,
            topico,
            descricao,
            situacao,
            responsavel_apoio,
            criado_em,
            atualizado_em
        FROM lista_apoio
    """
    with conexao_lista_apoio_db() as conexao:
        linhas = conexao.execute(consulta).fetchall()
        if not linhas:
            return 0

        grupos = {}
        for linha in linhas:
            registro = {
                CABECALHOS_LISTA_APOIO[0]: texto_lista_apoio(linha["carimbo_data_hora"]),
                CABECALHOS_LISTA_APOIO[1]: texto_lista_apoio(linha["nome_tecnico"]),
                CABECALHOS_LISTA_APOIO[2]: texto_lista_apoio(linha["topico"]),
                CABECALHOS_LISTA_APOIO[3]: texto_lista_apoio(linha["descricao"]),
                CABECALHOS_LISTA_APOIO[4]: texto_lista_apoio(linha["situacao"]),
                CABECALHOS_LISTA_APOIO[5]: texto_lista_apoio(linha["responsavel_apoio"]),
            }
            chave_natural = chave_registro_lista_apoio(registro)
            if not chave_natural:
                continue
            grupos.setdefault(chave_natural, []).append(
                {
                    "id": linha["id"],
                    "chave_atual": texto_lista_apoio(linha["chave"]),
                    "chave_natural": chave_natural,
                    "atualizado_em": texto_lista_apoio(linha["atualizado_em"]),
                    "criado_em": texto_lista_apoio(linha["criado_em"]),
                }
            )

        ids_para_excluir = []
        atualizacoes = []
        for chave_natural, grupo in grupos.items():
            grupo_ordenado = sorted(
                grupo,
                key=lambda item: (
                    item["chave_atual"] == chave_natural,
                    item["atualizado_em"],
                    item["criado_em"],
                    item["id"],
                ),
                reverse=True,
            )
            manter = grupo_ordenado[0]
            for duplicado in grupo_ordenado[1:]:
                ids_para_excluir.append(duplicado["id"])
            if manter["chave_atual"] != chave_natural:
                atualizacoes.append((chave_natural, manter["id"]))

        if ids_para_excluir:
            marcador = "%s" if backend == "postgres" else "?"
            consulta_delete = (
                f"DELETE FROM lista_apoio WHERE id IN ({', '.join([marcador] * len(ids_para_excluir))})"
            )
            conexao.execute(consulta_delete, tuple(ids_para_excluir))

        consulta_update = (
            "UPDATE lista_apoio SET chave = %s WHERE id = %s"
            if backend == "postgres"
            else "UPDATE lista_apoio SET chave = ? WHERE id = ?"
        )
        for chave_natural, identificador in atualizacoes:
            conexao.execute(consulta_update, (chave_natural, identificador))

        conexao.commit()
        return len(ids_para_excluir) + len(atualizacoes)


def migrar_lista_apoio_para_banco_se_necessario():
    garantir_banco_lista_apoio()
    with conexao_lista_apoio_db() as conexao:
        linha_total = conexao.execute("SELECT COUNT(*) AS total FROM lista_apoio").fetchone()
        total_banco = linha_total["total"] if linha_total else 0

    if total_banco > 0:
        return

    garantir_planilha_lista_apoio_integra()
    dataframe_principal = ler_dataframe_lista_apoio_arquivo(ARQUIVO_LISTA_APOIO)
    dataframe_backup = ler_dataframe_lista_apoio_arquivo(ARQUIVO_LISTA_APOIO_BACKUP)
    dataframe_historico = restaurar_lista_apoio_do_historico()
    dataframe_mesclado = mesclar_dataframes_lista_apoio(
        [dataframe_principal, dataframe_backup, dataframe_historico]
    )

    registros = dataframe_para_registros_lista_apoio(dataframe_mesclado)
    if registros:
        salvar_registros_lista_apoio_no_banco(registros)


def salvar_snapshot_lista_apoio(caminho_origem, prefixo="lista_apoio"):
    if not os.path.exists(caminho_origem):
        return
    os.makedirs(PASTA_BACKUP_LISTA_APOIO, exist_ok=True)
    timestamp = datetime.now(FUSO_HORARIO_APP).strftime("%Y%m%d_%H%M%S")
    destino = os.path.join(
        PASTA_BACKUP_LISTA_APOIO,
        f"{prefixo}_{timestamp}.xlsx",
    )
    shutil.copyfile(caminho_origem, destino)


def salvar_snapshot_banco_lista_apoio(prefixo="lista_apoio_db"):
    if backend_lista_apoio() != "sqlite":
        return
    if not os.path.exists(ARQUIVO_LISTA_APOIO_DB):
        return
    os.makedirs(PASTA_BACKUP_LISTA_APOIO, exist_ok=True)
    timestamp = datetime.now(FUSO_HORARIO_APP).strftime("%Y%m%d_%H%M%S")
    destino = os.path.join(
        PASTA_BACKUP_LISTA_APOIO,
        f"{prefixo}_{timestamp}.db",
    )
    shutil.copyfile(ARQUIVO_LISTA_APOIO_DB, destino)


def bytes_dataframe_excel(dataframe):
    buffer = BytesIO()
    dataframe = padronizar_dataframe_lista_apoio(dataframe)
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False)
    return buffer.getvalue()


def garantir_backup_lista_apoio_por_versao_app():
    if not os.path.exists(ARQUIVO_LISTA_APOIO) and not os.path.exists(ARQUIVO_LISTA_APOIO_DB):
        return

    os.makedirs(PASTA_BACKUP_LISTA_APOIO, exist_ok=True)
    caminho_app = os.path.abspath("app.py")
    assinatura_atual = builtins.str(int(os.path.getmtime(caminho_app)))
    caminho_meta = os.path.join(PASTA_BACKUP_LISTA_APOIO, ARQUIVO_META_BACKUP_LISTA_APOIO)

    assinatura_registrada = ""
    if os.path.exists(caminho_meta):
        try:
            with open(caminho_meta, "r", encoding="utf-8") as arquivo_meta:
                assinatura_registrada = builtins.str(
                    json.load(arquivo_meta).get("assinatura_app", "")
                ).strip()
        except Exception:
            assinatura_registrada = ""

    if assinatura_atual == assinatura_registrada:
        return

    salvar_snapshot_lista_apoio(ARQUIVO_LISTA_APOIO, prefixo="lista_apoio_pre_alteracao_app")
    salvar_snapshot_banco_lista_apoio(prefixo="lista_apoio_db_pre_alteracao_app")

    with open(caminho_meta, "w", encoding="utf-8") as arquivo_meta:
        json.dump(
            {
                "assinatura_app": assinatura_atual,
                "atualizado_em": datetime.now(FUSO_HORARIO_APP).strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
            },
            arquivo_meta,
            ensure_ascii=False,
            indent=2,
        )


def salvar_workbook_lista_apoio(wb):
    diretorio = os.path.dirname(os.path.abspath(ARQUIVO_LISTA_APOIO)) or "."
    descritor, caminho_temporario = tempfile.mkstemp(
        prefix="lista_apoio_",
        suffix=".xlsx",
        dir=diretorio,
    )
    os.close(descritor)
    try:
        wb.save(caminho_temporario)
        with zipfile.ZipFile(caminho_temporario) as arquivo_zip:
            if arquivo_zip.testzip() is not None:
                raise zipfile.BadZipFile("Arquivo temporário gerado com corrupção.")
        os.replace(caminho_temporario, ARQUIVO_LISTA_APOIO)
        shutil.copyfile(ARQUIVO_LISTA_APOIO, ARQUIVO_LISTA_APOIO_BACKUP)
        salvar_snapshot_lista_apoio(ARQUIVO_LISTA_APOIO)
    finally:
        if os.path.exists(caminho_temporario):
            os.remove(caminho_temporario)


def salvar_dataframe_lista_apoio(dataframe):
    diretorio = os.path.dirname(os.path.abspath(ARQUIVO_LISTA_APOIO)) or "."
    descritor, caminho_temporario = tempfile.mkstemp(
        prefix="lista_apoio_",
        suffix=".xlsx",
        dir=diretorio,
    )
    os.close(descritor)
    try:
        dataframe.to_excel(caminho_temporario, index=False)
        with zipfile.ZipFile(caminho_temporario) as arquivo_zip:
            if arquivo_zip.testzip() is not None:
                raise zipfile.BadZipFile("Arquivo temporário gerado com corrupção.")
        os.replace(caminho_temporario, ARQUIVO_LISTA_APOIO)
        shutil.copyfile(ARQUIVO_LISTA_APOIO, ARQUIVO_LISTA_APOIO_BACKUP)
        salvar_snapshot_lista_apoio(ARQUIVO_LISTA_APOIO)
    finally:
        if os.path.exists(caminho_temporario):
            os.remove(caminho_temporario)


def garantir_lista_apoio():
    garantir_banco_lista_apoio()
    migrar_lista_apoio_para_banco_se_necessario()
    deduplicar_lista_apoio_no_banco()
    return None, None


def registrar_duvida_apoio(tecnico, topico, resumo):
    garantir_lista_apoio()
    coluna_carimbo = CABECALHOS_LISTA_APOIO[0]
    coluna_tecnico = CABECALHOS_LISTA_APOIO[1]
    coluna_topico = CABECALHOS_LISTA_APOIO[2]
    coluna_descricao = CABECALHOS_LISTA_APOIO[3]
    coluna_situacao = CABECALHOS_LISTA_APOIO[4]
    coluna_responsavel = CABECALHOS_LISTA_APOIO[5]
    registro = {
        coluna_carimbo: datetime.now(FUSO_HORARIO_APP).strftime(
            "%d/%m/%Y %H:%M:%S"
        ),
        coluna_tecnico: builtins.str(tecnico or "").title(),
        coluna_topico: builtins.str(topico or "").strip(),
        coluna_descricao: builtins.str(resumo or "").strip(),
        coluna_situacao: "Aberto",
        coluna_responsavel: "",
    }
    salvar_registros_lista_apoio_no_banco([registro])
    registrar_historico_lista_apoio("create", registro)


def ler_lista_apoio():
    garantir_lista_apoio()
    dataframe_banco = ler_lista_apoio_do_banco()
    if not dataframe_banco.empty:
        return dataframe_banco
    dataframe_historico = restaurar_lista_apoio_do_historico()
    if not dataframe_historico.empty:
        salvar_registros_lista_apoio_no_banco(
            dataframe_para_registros_lista_apoio(dataframe_historico)
        )
        return dataframe_historico
    return dataframe_lista_apoio_vazio()


def texto_lista_apoio(valor):
    if pd.isna(valor):
        return ""
    return builtins.str(valor or "").strip()


def normalizar_valor_lista_apoio(coluna, valor, valor_atual=""):
    texto = texto_lista_apoio(valor)
    texto_atual = texto_lista_apoio(valor_atual)

    if coluna == "Situação":
        if texto == "Respondido":
            texto = "Resolvido pelo técnico"
        if texto_atual == "Respondido":
            texto_atual = "Resolvido pelo técnico"

        opcoes = ["Aberto", "Em análise", "Resolvido pelo técnico", "Finalizado"]
        if texto in opcoes:
            return texto
        if texto_atual in opcoes:
            return texto_atual
        return "Aberto"

    if coluna == "Responsável/Apoio":
        opcoes = ["Brenda", "Luma"]
        if texto in opcoes:
            return texto
        return texto_atual

    return texto or texto_atual


def salvar_lista_apoio(dataframe):
    existente = ler_lista_apoio()
    existente_antes = existente.copy()
    dataframe = dataframe.copy()
    for coluna_lista in CABECALHOS_LISTA_APOIO:
        if coluna_lista not in dataframe.columns:
            dataframe[coluna_lista] = ""
        dataframe[coluna_lista] = dataframe[coluna_lista].astype("object")
        existente[coluna_lista] = existente[coluna_lista].astype("object")
    dataframe = dataframe[CABECALHOS_LISTA_APOIO]

    colunas_editaveis = [CABECALHOS_LISTA_APOIO[4], CABECALHOS_LISTA_APOIO[5]]
    colunas_chave = [
        CABECALHOS_LISTA_APOIO[0],
        CABECALHOS_LISTA_APOIO[1],
        CABECALHOS_LISTA_APOIO[2],
        CABECALHOS_LISTA_APOIO[3],
    ]

    for coluna_lista in colunas_editaveis:
        if coluna_lista not in existente.columns:
            existente[coluna_lista] = ""

    for indice, linha in dataframe.iterrows():
        if indice not in existente.index:
            continue
        for coluna in colunas_editaveis:
            existente.at[indice, coluna] = normalizar_valor_lista_apoio(
                coluna,
                linha[coluna],
                existente.at[indice, coluna],
            )

    edits_por_chave = {}
    for _, linha in dataframe.iterrows():
        chave = tuple(texto_lista_apoio(linha[coluna]) for coluna in colunas_chave)
        edits_por_chave[chave] = {
            coluna: linha[coluna]
            for coluna in colunas_editaveis
        }

    for indice, linha in existente.iterrows():
        chave = tuple(texto_lista_apoio(linha[coluna]) for coluna in colunas_chave)
        if chave not in edits_por_chave:
            continue
        for coluna in colunas_editaveis:
            existente.at[indice, coluna] = normalizar_valor_lista_apoio(
                coluna,
                edits_por_chave[chave][coluna],
                existente.at[indice, coluna],
            )

    for indice, linha in existente.iterrows():
        if indice >= len(existente_antes):
            continue
        registro_antes = {
            coluna: texto_lista_apoio(existente_antes.iloc[indice][coluna])
            for coluna in CABECALHOS_LISTA_APOIO
        }
        registro_depois = {
            coluna: texto_lista_apoio(linha[coluna])
            for coluna in CABECALHOS_LISTA_APOIO
        }
        if registro_antes != registro_depois:
            registrar_historico_lista_apoio("update", registro_depois)

    dataframe_final = existente[CABECALHOS_LISTA_APOIO]
    salvar_registros_lista_apoio_no_banco(
        dataframe_para_registros_lista_apoio(dataframe_final)
    )


def bytes_lista_apoio():
    garantir_lista_apoio()
    dataframe = ler_lista_apoio_do_banco()
    return bytes_dataframe_excel(dataframe)


def versao_lista_apoio():
    if backend_lista_apoio() == "postgres":
        try:
            with conexao_lista_apoio_db() as conexao:
                linha = conexao.execute(
                    "SELECT COALESCE(MAX(atualizado_em), 'sem_arquivo') AS versao, COUNT(*) AS total FROM lista_apoio"
                ).fetchone()
                if linha:
                    return f"{linha['versao']}|{linha['total']}"
        except Exception:
            return "sem_arquivo"
    if os.path.exists(ARQUIVO_LISTA_APOIO_DB):
        return builtins.str(int(os.path.getmtime(ARQUIVO_LISTA_APOIO_DB)))
    return "sem_arquivo"


def mostrar_lista_apoio_gestao():
    st.divider()
    st.subheader("Lista de Apoio")

    if "filtro_data_lista_apoio_persistido" not in st.session_state:
        st.session_state["filtro_data_lista_apoio_persistido"] = None
    if (
        "filtro_data_lista_apoio_widget" not in st.session_state
        and st.session_state["filtro_data_lista_apoio_persistido"] is not None
    ):
        st.session_state["filtro_data_lista_apoio_widget"] = st.session_state[
            "filtro_data_lista_apoio_persistido"
        ]

    col_atualizar_apoio, col_status_apoio = st.columns([1, 3])
    with col_atualizar_apoio:
        if st.button("Atualizar lista de ajuda", key="atualizar_lista_apoio"):
            st.session_state["versao_editor_lista_apoio"] = versao_lista_apoio()
            st.rerun()
    with col_status_apoio:
        st.caption("Use o botão para atualizar a lista sem derrubar o login.")

    try:
        lista_apoio = ler_lista_apoio()
    except Exception as erro_lista_apoio:
        st.error(f"Não foi possível carregar a lista de apoio: {erro_lista_apoio}")
        lista_apoio = pd.DataFrame(columns=CABECALHOS_LISTA_APOIO)

    if lista_apoio.empty:
        st.info("Ainda não há dúvidas registradas pelos técnicos.")
        return

    col_filtro_data, col_limpar_filtro = st.columns([3, 1])
    with col_filtro_data:
        filtro_data_apoio = st.date_input(
            "Filtrar ajudas por data",
            value=st.session_state["filtro_data_lista_apoio_persistido"],
            format="DD/MM/YYYY",
            key="filtro_data_lista_apoio_widget",
        )
        st.session_state["filtro_data_lista_apoio_persistido"] = filtro_data_apoio
    with col_limpar_filtro:
        st.write("")
        st.write("")
        if st.button("Limpar filtro", key="limpar_filtro_lista_apoio"):
            st.session_state["filtro_data_lista_apoio_persistido"] = None
            st.session_state["filtro_data_lista_apoio_widget"] = None
            st.rerun()

    lista_exibida = lista_apoio.copy()
    if filtro_data_apoio:
        datas_apoio = pd.to_datetime(
            lista_exibida["Carimbo de data/hora"].astype(str).str.strip(),
            dayfirst=True,
            errors="coerce",
        ).dt.date
        lista_exibida = lista_exibida[datas_apoio == filtro_data_apoio]

    st.caption(f"{len(lista_exibida)} registro(s) exibido(s).")

    if lista_exibida.empty:
        st.info("Nenhuma ajuda registrada para a data informada.")
        st.download_button(
            "Baixar lista em Excel",
            data=bytes_lista_apoio(),
            file_name=ARQUIVO_LISTA_APOIO,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="baixar_lista_apoio_vazia",
        )
        return

    if "versao_editor_lista_apoio" not in st.session_state:
        st.session_state["versao_editor_lista_apoio"] = versao_lista_apoio()

    chave_editor_apoio = (
        "editor_lista_apoio_"
        + re.sub(
            r"[^0-9a-zA-Z]+",
            "_",
            filtro_data_apoio.strftime("%d_%m_%Y") if filtro_data_apoio else "todas",
        )
        + "_"
        + st.session_state["versao_editor_lista_apoio"]
    )

    lista_editada = st.data_editor(
        lista_exibida,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        disabled=[
            "Carimbo de data/hora",
            "Nome do Técnico",
            "Selecione o tópico",
            "Descreva o problema/pedido em poucas palavras",
        ],
        column_config={
            "Situação": st.column_config.SelectboxColumn(
                "Situação",
                options=["Aberto", "Em análise", "Resolvido pelo técnico", "Finalizado"],
            ),
            "Responsável/Apoio": st.column_config.SelectboxColumn(
                "Responsável/Apoio",
                options=["Brenda", "Luma"],
            )
        },
        key=chave_editor_apoio,
    )

    col_apoio_1, col_apoio_2 = st.columns([1, 1])
    with col_apoio_1:
        if st.button("Salvar lista de apoio", key="salvar_lista_apoio"):
            try:
                salvar_lista_apoio(lista_editada)
            except PermissionError:
                st.error("Feche a lista_apoio.xlsx e tente salvar novamente.")
            except Exception as erro_lista_apoio:
                st.error(f"Não foi possível salvar a lista: {erro_lista_apoio}")
            else:
                st.success("Lista de apoio atualizada.")
                st.session_state["versao_editor_lista_apoio"] = versao_lista_apoio()
                st.rerun()
    with col_apoio_2:
        st.download_button(
            "Baixar lista em Excel",
            data=bytes_lista_apoio(),
            file_name=ARQUIVO_LISTA_APOIO,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="baixar_lista_apoio",
        )


def salvar_env_local(atualizacoes):
    valores = carregar_env_local()
    for chave, valor in atualizacoes.items():
        valores[chave] = builtins.str(valor or "").strip()

    linhas = [f"{chave}={valor}" for chave, valor in sorted(valores.items())]
    with open(".env", "w", encoding="utf-8") as arquivo:
        arquivo.write("\n".join(linhas) + "\n")


def normalizar_nome(valor):
    texto = builtins.str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"\([^)]*\)", " ", texto)
    texto = re.sub(r"\b(de|da|do|das|dos|e)\b", " ", texto)
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def nivel_canonico(valor):
    nivel = normalizar_nome(valor)
    niveis = {
        "tecnico iii": "Técnico III",
        "tecnico ii": "Técnico II",
        "tecnico i": "Técnico I",
        "jr": "JR",
        "estagio": "Estágio",
    }
    return niveis.get(nivel, builtins.str(valor or "").strip())


def meta_esperada_nivel(valor):
    return META_ESPERADA_POR_NIVEL.get(nivel_canonico(valor), 0)


def arredondar_esperado(valor):
    if pd.isna(valor):
        return 0
    return int(
        Decimal(builtins.str(float(valor))).quantize(
            Decimal("1"),
            rounding=ROUND_HALF_UP,
        )
    )


def percentuais_absorcao_por_mes(dataframe, ano, mes):
    dados_mes = dataframe[
        (dataframe["Data"].dt.year == ano)
        & (dataframe["Data"].dt.month == mes)
    ].copy()
    if dados_mes.empty:
        return {nivel: 0 for nivel in META_ESPERADA_POR_NIVEL}

    dados_mes = dados_mes[
        ~dados_mes["Técnico"]
        .map(normalizar_nome)
        .isin(TECNICOS_DESCONSIDERADOS_ESPERADO)
    ].copy()
    dados_mes["Nivel Canonico"] = dados_mes["Nível"].map(nivel_canonico)
    dados_mes = dados_mes.sort_values("Data")
    ultimo_nivel_tecnico = dados_mes.drop_duplicates(
        subset=["Técnico"],
        keep="last",
    )
    total_meta_mes = ultimo_nivel_tecnico["Nivel Canonico"].map(
        META_ESPERADA_POR_NIVEL
    ).fillna(0).sum()

    if not total_meta_mes:
        return {nivel: 0 for nivel in META_ESPERADA_POR_NIVEL}

    return {
        nivel: (meta / total_meta_mes) * 100
        for nivel, meta in META_ESPERADA_POR_NIVEL.items()
    }


def percentual_absorcao_tecnico(dataframe, ano, mes, nivel):
    return percentuais_absorcao_por_mes(dataframe, ano, mes).get(
        nivel_canonico(nivel),
        0,
    )


def total_meta_mensal_por_linha(dataframe):
    totais_por_mes = {}
    for periodo, dados_mes in dataframe.groupby(dataframe["Data"].dt.to_period("M")):
        dados_validos = dados_mes[
            ~dados_mes["Técnico"]
            .map(normalizar_nome)
            .isin(TECNICOS_DESCONSIDERADOS_ESPERADO)
        ].copy()
        if dados_validos.empty:
            totais_por_mes[periodo] = 0
            continue

        dados_validos = dados_validos.sort_values("Data")
        ultimo_nivel_tecnico = dados_validos.drop_duplicates(
            subset=["Técnico"],
            keep="last",
        )
        totais_por_mes[periodo] = ultimo_nivel_tecnico["Nível"].map(
            meta_esperada_nivel
        ).fillna(0).sum()

    return dataframe["Data"].dt.to_period("M").map(totais_por_mes).fillna(0)


def cabecalhos(ws):
    return {
        builtins.str(cell.value).strip(): cell.column
        for cell in ws[1]
        if cell.value is not None
    }


def garantir_coluna(ws, nome_coluna):
    colunas = cabecalhos(ws)
    if nome_coluna in colunas:
        return colunas[nome_coluna]

    indice = ws.max_column + 1
    ws.cell(row=1, column=indice).value = nome_coluna
    return indice


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


def nome_coluna_dataframe(dataframe, *nomes):
    colunas_normalizadas = {
        normalizar_cabecalho(nome): nome for nome in dataframe.columns
    }

    for nome in nomes:
        coluna_encontrada = colunas_normalizadas.get(normalizar_cabecalho(nome))
        if coluna_encontrada:
            return coluna_encontrada

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
    headers = [
        "T??cnico Planilha",
        "Agente BI",
        "Agente SGD",
        "Agente Ponto",
        "Agente Chat",
        "Agente RO",
    ]
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
        nova_linha = [None] * ws.max_column
        nova_linha[col_data - 1] = data_referencia
        nova_linha[col_tecnico - 1] = tecnico
        nova_linha[col_nivel - 1] = nivel
        ws.append(nova_linha)
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
            ws_aliases.append([tecnico, "", "", "", "", ""])
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


def localizar_funcionario_por_nome(funcionarios, nome_referencia):
    nome_normalizado = normalizar_nome(nome_referencia)
    if not nome_normalizado:
        return None

    tokens_referencia = set(nome_normalizado.split())
    melhor_funcionario = None
    melhor_score = 0.0

    for funcionario in funcionarios:
        nome_funcionario = funcionario.get("Nome", "")
        nome_funcionario_normalizado = normalizar_nome(nome_funcionario)
        if not nome_funcionario_normalizado:
            continue
        if nome_funcionario_normalizado == nome_normalizado:
            return funcionario

        tokens_funcionario = set(nome_funcionario_normalizado.split())
        intersecao = len(tokens_referencia & tokens_funcionario)
        uniao = len(tokens_referencia | tokens_funcionario) or 1
        score_tokens = intersecao / uniao
        score_texto = SequenceMatcher(
            None,
            nome_normalizado,
            nome_funcionario_normalizado,
        ).ratio()
        score = max(score_texto, (score_tokens * 0.7) + (score_texto * 0.3))

        if score > melhor_score:
            melhor_score = score
            melhor_funcionario = funcionario

    return melhor_funcionario if melhor_score >= 0.6 else None


def obter_alias_pontoweb(tecnico):
    try:
        wb = carregar_workbook_produtividade(read_only=True, data_only=True)
    except Exception:
        return ""

    if ABA_ALIASES not in wb.sheetnames:
        return ""

    ws_aliases = wb[ABA_ALIASES]
    for row in range(2, ws_aliases.max_row + 1):
        tecnico_planilha = ws_aliases.cell(row=row, column=1).value
        if normalizar_nome(tecnico_planilha) != normalizar_nome(tecnico):
            continue
        return builtins.str(ws_aliases.cell(row=row, column=4).value or "").strip()
    return ""


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


def dias_uteis_para_meta(ano, mes):
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    total = 0
    for dia in range(1, ultimo_dia + 1):
        data_atual = date(ano, mes, dia)
        if data_atual.weekday() >= 5:
            continue
        if eh_feriado_federal(data_atual):
            continue
        total += 1
    return total


def eh_feriado_federal(data_atual):
    return (
        (data_atual.month, data_atual.day) in FERIADOS_FEDERAIS_FIXOS
        or data_atual in FERIADOS_ADICIONAIS
    )


def hhmm_para_minutos(valor):
    texto = builtins.str(valor or "").strip()
    if not texto or ":" not in texto:
        return 0

    negativo = texto.startswith("-")
    texto = texto.lstrip("+-")
    partes = texto.split(":")
    if len(partes) != 2 or not all(parte.isdigit() for parte in partes):
        return 0

    minutos = int(partes[0]) * 60 + int(partes[1])
    return -minutos if negativo else minutos


def extrair_valor_input(html, nome_campo):
    padrao = re.compile(
        rf'<input[^>]+name="{re.escape(nome_campo)}"[^>]+value="([^"]*)"',
        re.IGNORECASE,
    )
    match = padrao.search(html)
    return unescape(match.group(1)) if match else ""


def extrair_erro_login_pontoweb(html):
    match = re.search(
        r'<span[^>]*class="[^"]*field-validation-error[^"]*"[^>]*>(.*?)</span>',
        html,
        re.IGNORECASE | re.DOTALL,
    )
    if not match:
        return ""
    mensagem = re.sub(r"<[^>]+>", " ", match.group(1))
    return re.sub(r"\s+", " ", unescape(mensagem)).strip()


def login_pontoweb(email, senha):
    sessao = criar_sessao_http()
    resposta_login = sessao.get(PONTOWEB_AUTH_URL, timeout=30)
    resposta_login.raise_for_status()

    html_login = resposta_login.text
    token_verificacao = extrair_valor_input(html_login, "__RequestVerificationToken")
    cliente_id = extrair_valor_input(html_login, "ClienteId") or PONTOWEB_CLIENT_ID
    redirect_uri = extrair_valor_input(html_login, "RedirectUri") or PONTOWEB_REDIRECT_URI

    payload = {
        "Email": email,
        "Senha": senha,
        "ContinuarConectado": "true",
        "ClienteId": cliente_id,
        "RedirectUri": redirect_uri,
        "__RequestVerificationToken": token_verificacao,
        "action:Login": "Login",
    }

    resposta_autorizacao = sessao.post(
        PONTOWEB_AUTH_URL,
        data=payload,
        headers={"Referer": PONTOWEB_AUTH_URL},
        allow_redirects=False,
        timeout=30,
    )

    if resposta_autorizacao.status_code not in {302, 303}:
        mensagem = extrair_erro_login_pontoweb(resposta_autorizacao.text)
        if mensagem:
            raise ValueError(mensagem)
        raise ValueError("Login do PontoWeb inválido.")

    location = resposta_autorizacao.headers.get("Location", "")
    if not location:
        raise ValueError("O PontoWeb não retornou o código de autorização.")

    match_code = re.search(r"[?&]code=([^&]+)", location)
    if not match_code:
        raise ValueError("Não encontrei o código de autorização do PontoWeb.")

    resposta_token = sessao.post(
        PONTOWEB_TOKEN_URL,
        data={
            "client_id": cliente_id,
            "redirect_uri": redirect_uri,
            "grant_type": "authorization_code",
            "code": match_code.group(1),
        },
        headers={"Content-type": "application/x-www-form-urlencoded; charset=UTF-8"},
        timeout=30,
    )
    resposta_token.raise_for_status()

    dados_token = resposta_token.json()
    return {
        "sessao": sessao,
        "access_token": dados_token["access_token"],
        "refresh_token": dados_token.get("refresh_token", ""),
    }


def cabecalhos_pontoweb(access_token, identificador_banco=""):
    headers = {"Authorization": f"Bearer {access_token}"}
    if identificador_banco:
        headers["secullumbancoselecionado"] = identificador_banco
    return headers


def buscar_toolbar_pontoweb(sessao, access_token):
    resposta = sessao.get(
        f"{PONTOWEB_BASE_URL}Toolbar",
        headers=cabecalhos_pontoweb(access_token),
        timeout=30,
    )
    resposta.raise_for_status()
    return resposta.json()


def localizar_lista_bancos(objeto):
    if isinstance(objeto, list):
        if objeto and all(
            isinstance(item, dict)
            and ("BancoId" in item or "bancoId" in item)
            and ("Identificador" in item or "identificador" in item)
            for item in objeto
        ):
            return objeto
        for item in objeto:
            lista = localizar_lista_bancos(item)
            if lista:
                return lista
        return []

    if isinstance(objeto, dict):
        for chave in ("listaBancos", "ListaBancos", "bancos", "Bancos"):
            valor = objeto.get(chave)
            if isinstance(valor, list) and valor:
                return valor
        for valor in objeto.values():
            lista = localizar_lista_bancos(valor)
            if lista:
                return lista

    return []


def selecionar_banco_pontoweb(lista_bancos, banco_id="", identificador=""):
    if not lista_bancos:
        return {}

    banco_id = builtins.str(banco_id or "").strip()
    identificador = builtins.str(identificador or "").strip()

    for banco in lista_bancos:
        if banco_id and builtins.str(
            banco.get("BancoId", banco.get("bancoId", ""))
        ) == banco_id:
            return banco
        if identificador and builtins.str(
            banco.get("Identificador", banco.get("identificador", ""))
        ) == identificador:
            return banco

    return lista_bancos[0]


def buscar_funcionarios_pontoweb(sessao, access_token, identificador_banco):
    resposta = sessao.get(
        f"{PONTOWEB_BASE_URL}Funcionarios",
        headers=cabecalhos_pontoweb(access_token, identificador_banco),
        timeout=60,
    )
    resposta.raise_for_status()
    return resposta.json()


def buscar_calculos_pontoweb(
    sessao,
    access_token,
    identificador_banco,
    funcionario_id,
    data_inicial,
    data_final,
):
    endpoint = (
        f"{PONTOWEB_BASE_URL}Calculos/"
        f"{funcionario_id}/{data_inicial.strftime('%Y-%m-%d')}/"
        f"{data_final.strftime('%Y-%m-%d')}"
    )
    resposta = sessao.get(
        endpoint,
        headers=cabecalhos_pontoweb(access_token, identificador_banco),
        timeout=60,
    )
    resposta.raise_for_status()
    return resposta.json()


def indice_coluna_calculo(colunas, *nomes):
    nomes_normalizados = {normalizar_nome(nome) for nome in nomes}
    for indice, coluna_calculo in enumerate(colunas):
        nome_coluna = coluna_calculo.get("NomeExibicao") or coluna_calculo.get("Nome")
        if normalizar_nome(nome_coluna) in nomes_normalizados:
            return indice
    return -1


def valor_linha_calculo(linha, indice):
    if indice < 0 or indice >= len(linha):
        return ""
    return linha[indice]


def calcular_resumo_meta_pontoweb_contexto(
    sessao,
    access_token,
    identificador_banco,
    funcionarios,
    tecnico,
    ano,
    mes,
):
    alias_pontoweb = obter_alias_pontoweb(tecnico)
    nomes_funcionarios = [
        funcionario.get("Nome", "")
        for funcionario in funcionarios
        if funcionario.get("Nome")
    ]
    funcionario = None

    for nome_referencia in [alias_pontoweb, tecnico, melhor_alias(tecnico, nomes_funcionarios)]:
        if not nome_referencia:
            continue
        funcionario = localizar_funcionario_por_nome(funcionarios, nome_referencia)
        if funcionario:
            break

    if not funcionario:
        raise ValueError(f"Não encontrei o técnico {tecnico} no PontoWeb.")

    data_inicial = date(ano, mes, 1)
    data_final = date(ano, mes, calendar.monthrange(ano, mes)[1])
    dados_calculo = buscar_calculos_pontoweb(
        sessao,
        access_token,
        identificador_banco,
        funcionario.get("Id"),
        data_inicial,
        data_final,
    )

    colunas = dados_calculo.get("Colunas", [])
    linhas = dados_calculo.get("Linhas", [])
    situacoes = dados_calculo.get("SituacaoDias", [])

    idx_data = indice_coluna_calculo(colunas, "Data")
    idx_faltas = indice_coluna_calculo(colunas, "Faltas")
    idx_carga = indice_coluna_calculo(colunas, "Carga")
    idx_bdeb = indice_coluna_calculo(colunas, "BDeb.", "BDeb")

    abatimento_total = 0.0
    dias_base = dias_uteis_para_meta(ano, mes)

    for indice_linha, linha in enumerate(linhas):
        data_texto = builtins.str(valor_linha_calculo(linha, idx_data)).strip()
        match_data = re.search(r"(\d{2}/\d{2}/\d{4})", data_texto)
        if not match_data:
            continue

        data_linha = pd.to_datetime(match_data.group(1), dayfirst=True, errors="coerce")
        if pd.isna(data_linha):
            continue

        data_linha = data_linha.date()
        if data_linha.year != ano or data_linha.month != mes:
            continue
        if data_linha.weekday() >= 5 or eh_feriado_federal(data_linha):
            continue

        valores_linha = [builtins.str(valor or "").strip().upper() for valor in linha]
        faltas_minutos = max(hhmm_para_minutos(valor_linha_calculo(linha, idx_faltas)), 0)
        carga_minutos = max(hhmm_para_minutos(valor_linha_calculo(linha, idx_carga)), 0)
        bdeb_minutos = max(hhmm_para_minutos(valor_linha_calculo(linha, idx_bdeb)), 0)
        situacao = situacoes[indice_linha] if indice_linha < len(situacoes) else None

        if any("FERIAS" in normalizar_nome(valor) for valor in valores_linha):
            abatimento_total += 1
            continue

        if any("ATESTAD" in normalizar_nome(valor) for valor in valores_linha):
            abatimento_total += 1
            continue

        if carga_minutos <= 0:
            continue

        if situacao == 2 and faltas_minutos >= carga_minutos:
            abatimento_total += 1
            continue

        minutos_ausentes = max(faltas_minutos, bdeb_minutos)
        if 0 < minutos_ausentes < carga_minutos:
            abatimento_total += minutos_ausentes / carga_minutos

    dias_considerados = max(dias_base - abatimento_total, 0)
    return {
        "dias_base": dias_base,
        "abatimento": round(abatimento_total, 2),
        "dias_considerados": round(dias_considerados, 2),
    }


def calcular_resumo_meta_pontoweb(
    tecnico,
    ano,
    mes,
    email,
    senha,
    banco_id="",
    banco_identificador="",
):
    login = login_pontoweb(email, senha)
    sessao = login["sessao"]
    access_token = login["access_token"]

    toolbar = buscar_toolbar_pontoweb(sessao, access_token)
    lista_bancos = localizar_lista_bancos(toolbar)
    banco = selecionar_banco_pontoweb(lista_bancos, banco_id, banco_identificador)
    identificador_banco = builtins.str(
        banco.get("Identificador", banco.get("identificador", ""))
    )
    funcionarios = buscar_funcionarios_pontoweb(sessao, access_token, identificador_banco)
    return calcular_resumo_meta_pontoweb_contexto(
        sessao,
        access_token,
        identificador_banco,
        funcionarios,
        tecnico,
        ano,
        mes,
    )


@st.cache_data(ttl=1800, show_spinner=False)
def obter_resumo_meta_pontoweb(
    tecnico,
    ano,
    mes,
    email,
    senha,
    banco_id="",
    banco_identificador="",
):
    return calcular_resumo_meta_pontoweb(
        tecnico,
        ano,
        mes,
        email,
        senha,
        banco_id,
        banco_identificador,
    )


def criar_sessao_http():
    sessao = requests.Session()
    sessao.trust_env = False
    return sessao


def atualizar_via_servico_local(fonte, data_referencia):
    sessao = criar_sessao_http()
    resposta = sessao.post(
        f"{SERVICO_PLANILHA_LOCAL_URL}/atualizar",
        json={
            "fonte": fonte,
            "data": data_referencia.strftime("%d/%m/%Y"),
        },
        timeout=600,
    )
    resposta.raise_for_status()
    payload = resposta.json()
    if not payload.get("ok"):
        raise RuntimeError(payload.get("erro", "Falha na atualização local."))
    return payload.get("resultados", {})


def servico_local_disponivel():
    try:
        sessao = criar_sessao_http()
        resposta = sessao.get(
            f"{SERVICO_PLANILHA_LOCAL_URL}/health",
            timeout=5,
        )
        resposta.raise_for_status()
        payload = resposta.json()
        return bool(payload.get("ok"))
    except Exception:
        return False


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
    chave_mes = (data_referencia.year, data_referencia.month)
    link_fixo = RO_PLANILHAS_FIXAS.get(chave_mes)
    if link_fixo:
        correspondencia = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", link_fixo)
        if correspondencia:
            return {
                "id": correspondencia.group(1),
                "titulo": f"RO {MESES_RO[data_referencia.month]} {data_referencia.year} (respostas)",
                "origem": "link_fixo",
                "url": link_fixo,
            }

    esperado = normalizar_cabecalho(
        f"RO {MESES_RO[data_referencia.month]} {data_referencia.year} (respostas)"
    )

    for arquivo in listar_arquivos_ro():
        if esperado in normalizar_cabecalho(arquivo["titulo"]):
            return {
                **arquivo,
                "origem": "drive",
                "url": f"https://docs.google.com/spreadsheets/d/{arquivo['id']}/edit",
            }

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
    tecnicos_origem = []
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
        tecnicos_origem.append(tecnico)

    tentar_salvar_backup_ro_no_banco(
        data_referencia,
        arquivo,
        texto_csv,
        contagens,
        tecnicos_origem,
    )

    return {
        "arquivo": arquivo["titulo"],
        "contagens": contagens,
        "tecnicos_origem": tecnicos_origem,
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


@st.cache_data(ttl=300, show_spinner=False)
def carregar_filas_folha_bi():
    try:
        sessao = criar_sessao_http()
        resposta = sessao.get(
            f"{X2_CONTROLLER_BASE_URL}/api/ami/queues",
            timeout=30,
        )
        resposta.raise_for_status()
        dados = resposta.json()
        filas = []

        for fila in dados.get("queues", []):
            nome = builtins.str(fila.get("name") or "").strip()
            if nome and "folha" in nome.lower() and nome not in filas:
                filas.append(nome)

        if filas:
            return filas
    except Exception:
        pass

    return [POWERBI_FILA_FALLBACK]


def aplicar_filtro_fila_powerbi(query):
    fontes = query.setdefault("From", [])
    if not any(
        fonte.get("Entity") == "Filas" and fonte.get("Name") == "f"
        for fonte in fontes
    ):
        fontes.append({"Name": "f", "Entity": "Filas", "Type": 0})

    valores_fila = [
        [{"Literal": {"Value": f"'{fila}'"}}]
        for fila in carregar_filas_folha_bi()
    ]

    query.setdefault("Where", []).append(
        {
            "Condition": {
                "In": {
                    "Expressions": [
                        {
                            "Column": {
                                "Expression": {"SourceRef": {"Source": "f"}},
                                "Property": "Fila",
                            }
                        }
                    ],
                    "Values": valores_fila,
                }
            }
        }
    )


def aplicar_filtro_data_powerbi(query, data_referencia):
    fonte_data = next(
        (
            fonte
            for fonte in query.setdefault("From", [])
            if builtins.str(fonte.get("Entity", "")).startswith("LocalDateTable_")
        ),
        None,
    )
    if fonte_data is None:
        fonte_data = {
            "Name": "l",
            "Entity": "LocalDateTable_5b42e521-bbb7-4d39-8e7e-735e21b48675",
            "Type": 0,
        }
        query["From"].append(fonte_data)

    nome_fonte = fonte_data["Name"]
    query.setdefault("Where", []).insert(
        0,
        {
            "Condition": {
                "In": {
                    "Expressions": [
                        {
                            "Column": {
                                "Expression": {"SourceRef": {"Source": nome_fonte}},
                                "Property": "Ano",
                            }
                        },
                        {
                            "Column": {
                                "Expression": {"SourceRef": {"Source": nome_fonte}},
                                "Property": "Trimestre",
                            }
                        },
                        {
                            "Column": {
                                "Expression": {"SourceRef": {"Source": nome_fonte}},
                                "Property": "Mês",
                            }
                        },
                        {
                            "Column": {
                                "Expression": {"SourceRef": {"Source": nome_fonte}},
                                "Property": "Dia",
                            }
                        },
                    ],
                    "Values": [
                        [
                            {"Literal": {"Value": f"{data_referencia.year}L"}},
                            {
                                "Literal": {
                                    "Value": f"'{trimestre_powerbi(data_referencia)}'"
                                }
                            },
                            {
                                "Literal": {
                                    "Value": f"'{MESES_POWERBI[data_referencia.month]}'"
                                }
                            },
                            {"Literal": {"Value": f"{data_referencia.day}L"}},
                        ]
                    ],
                }
            }
        },
    )


def consulta_visual_powerbi(visual):
    if visual.get("query"):
        return json.loads(visual["query"])

    configuracao = json.loads(visual["config"])
    visual_config = configuracao["singleVisual"]
    query = json.loads(
        json.dumps(
            visual_config["prototypeQuery"],
            ensure_ascii=False,
        )
    )
    tipo_visual = visual_config.get("visualType")
    total_selecoes = len(query.get("Select", []))

    if tipo_visual == "tableEx":
        binding = {
            "Primary": {
                "Groupings": [
                    {
                        "Projections": list(range(total_selecoes)),
                        "Subtotal": 1,
                    }
                ]
            },
            "DataReduction": {
                "DataVolume": 3,
                "Primary": {"Window": {"Count": 500}},
            },
            "Version": 1,
        }
    else:
        binding = {
            "Primary": {
                "Groupings": [
                    {
                        "Projections": list(range(total_selecoes)),
                    }
                ]
            },
            "DataReduction": {
                "DataVolume": 3,
                "Primary": {"Top": {}},
            },
            "Version": 1,
        }

    return {
        "Commands": [
            {
                "SemanticQueryDataShapeCommand": {
                    "Query": query,
                    "Binding": binding,
                    "ExecutionMetricsKind": 1,
                }
            }
        ]
    }


def montar_consulta_powerbi(data_referencia):
    metadados = carregar_metadados_powerbi()
    secao_mapa = next(
        secao
        for secao in metadados["exploration"]["sections"]
        if secao.get("displayName") == "Mapa"
    )
    visual = secao_mapa["visualContainers"][0]
    consulta = consulta_visual_powerbi(visual)
    comando = consulta["Commands"][0]["SemanticQueryDataShapeCommand"]
    comando["Query"]["Where"] = []
    aplicar_filtro_data_powerbi(comando["Query"], data_referencia)
    aplicar_filtro_fila_powerbi(comando["Query"])

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


def montar_consulta_abandonadas_powerbi(data_referencia):
    metadados = carregar_metadados_powerbi()
    secao_abandonos = next(
        secao
        for secao in metadados["exploration"]["sections"]
        if secao.get("displayName") == "Abandonos"
    )
    visual = next(
        visual
        for visual in secao_abandonos["visualContainers"]
        if "Abandonadas.MAB" in visual.get("query", "")
        or "Abandonadas.MAB" in visual.get("config", "")
    )
    consulta = consulta_visual_powerbi(visual)
    comando = consulta["Commands"][0]["SemanticQueryDataShapeCommand"]
    query = comando["Query"]
    query["Where"] = []
    aplicar_filtro_data_powerbi(query, data_referencia)
    query["Select"] = [
        {
            "Aggregation": {
                "Expression": {
                    "Column": {
                        "Expression": {"SourceRef": {"Source": "a"}},
                        "Property": "Evento",
                    }
                },
                "Function": 5,
            },
            "Name": "CountNonNull(Abandonadas.Evento)",
            "NativeReferenceName": "Abandonadas",
        }
    ]
    aplicar_filtro_fila_powerbi(query)

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


def buscar_abandonadas_powerbi(data_referencia):
    sessao = criar_sessao_http()
    resposta = sessao.post(
        f"{POWERBI_API_BASE}/public/reports/querydata?synchronous=true",
        headers=headers_powerbi(),
        json=montar_consulta_abandonadas_powerbi(data_referencia),
        timeout=45,
    )
    resposta.raise_for_status()
    dados = resposta.json()
    data_shape = dados["results"][0]["result"]["data"]["dsr"]["DS"][0]["PH"][0]
    linhas = data_shape.get("DM0") or data_shape.get("DM1") or []
    if not linhas:
        return 0
    return int(round(float(linhas[0].get("M0") or 0)))


def atualizar_planilha_com_bi(data_referencia):
    registros_bi = buscar_produtividade_powerbi(data_referencia)
    abandonadas = buscar_abandonadas_powerbi(data_referencia)
    registros_por_agente = {
        normalizar_nome(registro["agente"]): registro
        for registro in registros_bi
    }

    wb = carregar_workbook_produtividade()
    ws = wb[ABA_PRODUTIVIDADE]
    colunas = cabecalhos(ws)
    col_data = coluna(colunas, "Data")
    col_tecnico = coluna(colunas, "Técnico")
    col_atendidas = coluna(colunas, "Atendidas")
    col_2min = coluna(colunas, " > 2min", "> 2min", ">2min")
    col_tma = coluna(colunas, "TMA")
    col_abandonadas = garantir_coluna(ws, COLUNA_ABANDONADAS)
    ws.column_dimensions[openpyxl.utils.get_column_letter(col_abandonadas)].hidden = True
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
        ws.cell(row=row, column=col_abandonadas).value = abandonadas
        chave_agente = aliases.get(normalizar_nome(tecnico), normalizar_nome(tecnico))
        registro = registros_por_agente.get(chave_agente)

        if not registro:
            sem_alias.append(tecnico)
            continue

        ws.cell(row=row, column=col_atendidas).value = registro["atendidas"]
        ws.cell(row=row, column=col_2min).value = registro["maior_2min"]
        ws.cell(row=row, column=col_tma).value = registro["tma"]
        atualizados.append(tecnico)

    zerar_colunas_sem_movimento_ws(ws)
    salvar_workbook_produtividade(wb)

    return {
        "data": data_referencia.strftime("%d/%m/%Y"),
        "fonte": len(registros_bi),
        "abandonadas": abandonadas,
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
        ultima_excecao = None
        for tentativa in range(1, SGD_RELATORIO_TENTATIVAS + 1):
            try:
                resposta = sessao.post(
                    SGD_RELATORIO_URL,
                    data=dados,
                    timeout=SGD_RELATORIO_TIMEOUT,
                )
                resposta.raise_for_status()
                break
            except requests.exceptions.ReadTimeout as exc:
                ultima_excecao = exc
                if tentativa == SGD_RELATORIO_TENTATIVAS:
                    periodo = (
                        f"{data_inicial.strftime('%d/%m/%Y')} a "
                        f"{data_final.strftime('%d/%m/%Y')}"
                    )
                    raise TimeoutError(
                        "O SGD demorou mais do que o esperado para gerar o "
                        f"relatÃ³rio do perÃ­odo {periodo}. Tente novamente em "
                        "alguns minutos."
                    ) from exc
                time.sleep(3 * tentativa)
        else:
            if ultima_excecao:
                raise ultima_excecao

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

    wb = carregar_workbook_produtividade()
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

    zerar_colunas_sem_movimento_ws(ws)
    salvar_workbook_produtividade(wb)

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

    wb = carregar_workbook_produtividade()
    ws = wb[ABA_PRODUTIVIDADE]
    colunas = cabecalhos(ws)
    col_data = coluna(colunas, "Data")
    col_tecnico = coluna(colunas, "Tecnico", "T?cnico", "T??cnico")
    col_ro = coluna(colunas, "RO")
    linhas_criadas = garantir_linhas_da_data(ws, colunas, data_referencia)
    tecnicos = tecnicos_da_planilha(ws, colunas)
    ws_aliases = garantir_aba_aliases(
        wb,
        tecnicos,
        dados_ro.get("tecnicos_origem", list(contagens.keys())),
        coluna_alias=6,
    )
    aliases = ler_aliases(ws_aliases, coluna_alias=6)

    atualizados = 0
    sem_alias = []

    for row in range(2, ws.max_row + 1):
        data_linha = ws.cell(row=row, column=col_data).value
        if not data_linha:
            continue
        data_linha = data_linha.date() if hasattr(data_linha, "date") else data_linha
        if data_linha != data_referencia:
            continue

        tecnico = ws.cell(row=row, column=col_tecnico).value
        chave_ro = aliases.get(normalizar_nome(tecnico), normalizar_nome(tecnico))
        valor_ro = contagens.get(chave_ro, 0)
        ws.cell(row=row, column=col_ro).value = valor_ro
        if not valor_ro and chave_ro not in contagens:
            sem_alias.append(tecnico)
        atualizados += 1

    zerar_colunas_sem_movimento_ws(ws)
    salvar_workbook_produtividade(wb)

    return {
        "data": data_referencia.strftime("%d/%m/%Y"),
        "arquivo": dados_ro["arquivo"],
        "fonte": sum(contagens.values()),
        "linhas_criadas": linhas_criadas,
        "atualizados": atualizados,
        "sem_alias": sorted(set(sem_alias)),
    }

def localizar_navegador_plug():
    candidatos = [
        os.getenv("PLUG_BROWSER_PATH", ""),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    for caminho in candidatos:
        if caminho and os.path.exists(caminho):
            return caminho
    return None


def periodo_chat_plug(data_referencia):
    inicio = datetime.combine(data_referencia, datetime.min.time())
    fim = datetime.combine(data_referencia, datetime.max.time()).replace(
        hour=23,
        minute=59,
        second=0,
        microsecond=0,
    )
    return (
        inicio.strftime("%d/%m/%Y %H:%M"),
        fim.strftime("%d/%m/%Y %H:%M"),
        inicio.strftime("%d/%m/%Y %H:%M")
        + " - "
        + fim.strftime("%d/%m/%Y %H:%M"),
    )


def abrir_sessao_plug():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "A biblioteca playwright não está instalada nesta máquina."
        ) from exc

    playwright = sync_playwright().start()
    contexto = None
    try:
        kwargs = {
            "user_data_dir": builtins.str(PLUG_PERFIL_DIR.resolve()),
            "headless": False,
        }
        executable_path = localizar_navegador_plug()
        if executable_path:
            kwargs["executable_path"] = executable_path
        contexto = playwright.chromium.launch_persistent_context(**kwargs)
        pagina = contexto.pages[0] if contexto.pages else contexto.new_page()
        pagina.goto(PLUG_CHATBOX_URL, wait_until="domcontentloaded", timeout=60000)

        limite = time.time() + 180
        while time.time() < limite:
            url_atual = pagina.url
            if "#/app/" in url_atual and "login" not in url_atual.lower():
                pagina.goto(
                    PLUG_CHATBOX_URL,
                    wait_until="domcontentloaded",
                    timeout=60000,
                )
                return playwright, contexto, pagina
            pagina.wait_for_timeout(1000)

        raise TimeoutError(
            "Faça login no PLUG na janela aberta e tente novamente."
        )
    except Exception:
        if contexto:
            contexto.close()
        playwright.stop()
        raise


def localizar_mes_no_calendario_plug(pagina, texto_mes_ano):
    locator = pagina.locator(".daterangepicker .month")
    textos = [texto.strip() for texto in locator.all_text_contents()]
    if texto_mes_ano in textos:
        return textos.index(texto_mes_ano)
    return -1


def ajustar_periodo_plug(pagina, data_referencia):
    _, _, periodo_texto = periodo_chat_plug(data_referencia)
    pagina.locator('input[name="datetimesCreated"]').click()
    pagina.wait_for_selector(".daterangepicker", timeout=15000)

    meses_pt = {
        1: "Janeiro",
        2: "Fevereiro",
        3: "Março",
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
    texto_mes_ano = f"{meses_pt[data_referencia.month]} {data_referencia.year}"

    for _ in range(24):
        indice_mes = localizar_mes_no_calendario_plug(pagina, texto_mes_ano)
        if indice_mes >= 0:
            break
        pagina.locator(".daterangepicker .prev.available").click()
        pagina.wait_for_timeout(200)
    else:
        raise RuntimeError("Não foi possível localizar o mês no calendário do PLUG.")

    calendario_alvo = ".drp-calendar.left" if indice_mes == 0 else ".drp-calendar.right"
    dia = builtins.str(data_referencia.day)
    celulas = pagina.locator(
        f"{calendario_alvo} td.available, {calendario_alvo} td.today.available"
    )

    encontrou = False
    for indice in range(celulas.count()):
        texto_celula = builtins.str(celulas.nth(indice).inner_text()).strip()
        if texto_celula == dia:
            celulas.nth(indice).click()
            pagina.wait_for_timeout(150)
            encontrou = True
            break
    if not encontrou:
        raise RuntimeError("Não foi possível selecionar o dia no calendário do PLUG.")

    pagina.locator(".drp-calendar.left .hourselect").select_option("0")
    pagina.locator(".drp-calendar.left .minuteselect").select_option("0")
    pagina.locator(".drp-calendar.right .hourselect").select_option("23")
    pagina.locator(".drp-calendar.right .minuteselect").select_option("59")
    pagina.locator(".daterangepicker .applyBtn").click()
    pagina.wait_for_timeout(1200)

    valor_atual = pagina.locator('input[name="datetimesCreated"]').input_value().strip()
    if valor_atual != periodo_texto:
        raise RuntimeError(
            f"Período do PLUG não foi aplicado corretamente: {valor_atual}"
        )


def selecionar_status_plug(pagina, status):
    pagina.locator("ng-select .ng-select-container").click()
    pagina.wait_for_selector('ng-dropdown-panel[role="listbox"]', timeout=10000)
    opcao = pagina.locator('ng-dropdown-panel [role="option"]', has_text=status)
    if opcao.count() != 1:
        raise RuntimeError(f"Status '{status}' não encontrado no PLUG.")
    opcao.click()
    pagina.wait_for_timeout(1200)


def obter_atendentes_plug(pagina):
    return pagina.evaluate(
        """() => Array.from(document.querySelectorAll('select'))[2]
            ? Array.from(document.querySelectorAll('select'))[2].options
                .filter((opt) => opt.value)
                .map((opt) => ({
                    value: opt.value,
                    text: opt.textContent.trim()
                }))
            : []"""
    )


def total_chat_plug_atendente(pagina, value_atendente):
    selects = pagina.locator("select")
    if selects.count() < 3:
        raise RuntimeError("Filtro de atendente não encontrado no PLUG.")
    selects.nth(2).select_option(value_atendente)
    pagina.wait_for_timeout(1200)
    total_texto = pagina.locator("text=Total:").first.inner_text(timeout=10000)
    match = re.search(r"Total:\s*(\d+)", total_texto)
    if not match:
        raise RuntimeError("Não foi possível ler o total de chats no PLUG.")
    return int(match.group(1))


def buscar_chat_plug(data_referencia):
    playwright, contexto, pagina = abrir_sessao_plug()
    try:
        pagina.goto(PLUG_CHATBOX_URL, wait_until="domcontentloaded", timeout=60000)
        pagina.wait_for_selector('button[title="Filtros"]', timeout=30000)
        pagina.locator('button[title="Filtros"]').click()
        pagina.wait_for_selector('input[name="datetimesCreated"]', timeout=15000)
        ajustar_periodo_plug(pagina, data_referencia)

        selects = pagina.locator("select")
        if selects.count() < 5:
            raise RuntimeError("Filtros principais do PLUG não foram carregados.")
        selects.nth(4).select_option(label=PLUG_GRUPO_FOLHA)
        pagina.wait_for_timeout(1200)
        selecionar_status_plug(pagina, PLUG_STATUS_FECHADO)

        atendentes = obter_atendentes_plug(pagina)
        contagens = {}
        for atendente in atendentes:
            contagens[normalizar_nome(atendente["text"])] = total_chat_plug_atendente(
                pagina,
                atendente["value"],
            )

        return {
            "contagens": contagens,
            "atendentes": [item["text"] for item in atendentes],
        }
    finally:
        contexto.close()
        playwright.stop()


def atualizar_planilha_com_chat(data_referencia):
    dados_chat = buscar_chat_plug(data_referencia)
    contagens = dados_chat["contagens"]

    wb = carregar_workbook_produtividade()
    ws = wb[ABA_PRODUTIVIDADE]
    colunas = cabecalhos(ws)
    col_data = coluna(colunas, "Data")
    col_tecnico = coluna(colunas, "Técnico")
    col_chat = coluna(colunas, "CHAT")
    linhas_criadas = garantir_linhas_da_data(ws, colunas, data_referencia)
    tecnicos = tecnicos_da_planilha(ws, colunas)
    ws_aliases = garantir_aba_aliases(
        wb,
        tecnicos,
        dados_chat["atendentes"],
        coluna_alias=5,
    )
    aliases = ler_aliases(ws_aliases, coluna_alias=5)

    atualizados = 0
    sem_alias = []

    for row in range(2, ws.max_row + 1):
        data_linha = ws.cell(row=row, column=col_data).value
        if not data_linha:
            continue
        data_linha = data_linha.date() if hasattr(data_linha, "date") else data_linha
        if data_linha != data_referencia:
            continue

        tecnico = ws.cell(row=row, column=col_tecnico).value
        chave_chat = aliases.get(normalizar_nome(tecnico), normalizar_nome(tecnico))
        if chave_chat not in contagens:
            ws.cell(row=row, column=col_chat).value = 0
            sem_alias.append(tecnico)
            continue

        ws.cell(row=row, column=col_chat).value = contagens[chave_chat]
        atualizados += 1

    zerar_colunas_sem_movimento_ws(ws)
    salvar_workbook_produtividade(wb)

    return {
        "data": data_referencia.strftime("%d/%m/%Y"),
        "fonte": len(contagens),
        "linhas_criadas": linhas_criadas,
        "atualizados": atualizados,
        "sem_alias": sorted(set(sem_alias)),
    }


def atualizar_planilha_com_pontoweb_mes(
    ano,
    mes,
    email,
    senha,
    banco_id="",
    banco_identificador="",
):
    login = login_pontoweb(email, senha)
    sessao = login["sessao"]
    access_token = login["access_token"]
    toolbar = buscar_toolbar_pontoweb(sessao, access_token)
    lista_bancos = localizar_lista_bancos(toolbar)
    banco = selecionar_banco_pontoweb(lista_bancos, banco_id, banco_identificador)
    identificador_banco = builtins.str(
        banco.get("Identificador", banco.get("identificador", ""))
    )
    funcionarios = buscar_funcionarios_pontoweb(sessao, access_token, identificador_banco)
    nomes_funcionarios = [
        funcionario.get("Nome", "")
        for funcionario in funcionarios
        if funcionario.get("Nome")
    ]

    wb = carregar_workbook_produtividade()
    ws = wb[ABA_PRODUTIVIDADE]
    colunas = cabecalhos(ws)
    col_data = coluna(colunas, "Data")
    col_tecnico = coluna(colunas, "Técnico")
    col_dias_meta = garantir_coluna(ws, COLUNA_DIAS_META)
    colunas = cabecalhos(ws)

    tecnicos_mes = sorted(
        {
            builtins.str(ws.cell(row=row, column=col_tecnico).value).strip()
            for row in range(2, ws.max_row + 1)
            if ws.cell(row=row, column=col_tecnico).value
            and (
                (
                    ws.cell(row=row, column=col_data).value.date()
                    if hasattr(ws.cell(row=row, column=col_data).value, "date")
                    else ws.cell(row=row, column=col_data).value
                ).year
                == ano
            )
            and (
                (
                    ws.cell(row=row, column=col_data).value.date()
                    if hasattr(ws.cell(row=row, column=col_data).value, "date")
                    else ws.cell(row=row, column=col_data).value
                ).month
                == mes
            )
        }
    )

    ws_aliases = garantir_aba_aliases(
        wb,
        tecnicos_mes,
        nomes_funcionarios,
        coluna_alias=4,
    )
    aliases_pontoweb = ler_aliases(ws_aliases, coluna_alias=4)

    dias_por_tecnico = {}
    erros = {}
    for tecnico in tecnicos_mes:
        try:
            nome_referencia = aliases_pontoweb.get(normalizar_nome(tecnico), tecnico)
            resumo = calcular_resumo_meta_pontoweb_contexto(
                sessao,
                access_token,
                identificador_banco,
                funcionarios,
                nome_referencia,
                ano,
                mes,
            )
            dias_por_tecnico[normalizar_nome(tecnico)] = resumo["dias_considerados"]
        except Exception as erro:
            erros[tecnico] = builtins.str(erro)

    atualizados = 0
    for row in range(2, ws.max_row + 1):
        data_linha = ws.cell(row=row, column=col_data).value
        if not data_linha:
            continue
        data_linha = data_linha.date() if hasattr(data_linha, "date") else data_linha
        if data_linha.year != ano or data_linha.month != mes:
            continue

        tecnico = ws.cell(row=row, column=col_tecnico).value
        chave = normalizar_nome(tecnico)
        if chave in dias_por_tecnico:
            ws.cell(row=row, column=col_dias_meta).value = dias_por_tecnico[chave]
            atualizados += 1

    salvar_workbook_produtividade(wb)
    return {
        "tecnicos_processados": len(dias_por_tecnico),
        "linhas_atualizadas": atualizados,
        "erros": erros,
    }


def zerar_colunas_sem_movimento_ws(ws):
    colunas = cabecalhos(ws)
    col_atendidas = coluna(colunas, "Atendidas")
    col_chat = coluna(colunas, "CHAT")
    col_ro = coluna(colunas, "RO")
    colunas_zeradas = [
        col_atendidas,
        coluna(colunas, " > 2min", "> 2min", ">2min"),
        coluna(colunas, "TMA"),
        col_ro,
        col_chat,
        coluna(colunas, "Realizado"),
        coluna(colunas, "Esperado"),
        coluna(colunas, "Desvio"),
    ]

    for row in range(2, ws.max_row + 1):
        atendidas = ws.cell(row=row, column=col_atendidas).value or 0
        chat = ws.cell(row=row, column=col_chat).value or 0
        ro = ws.cell(row=row, column=col_ro).value or 0
        try:
            atendidas = float(atendidas)
        except Exception:
            atendidas = 0
        try:
            chat = float(chat)
        except Exception:
            chat = 0
        try:
            ro = float(ro)
        except Exception:
            ro = 0

        if atendidas == 0 and chat == 0 and ro == 0:
            for indice_coluna in colunas_zeradas:
                ws.cell(row=row, column=indice_coluna).value = 0


def recalcular_colunas_derivadas(df):
    coluna_2min = " > 2min" if " > 2min" in df.columns else "> 2min"
    for coluna in [coluna_2min, "RO", "CHAT"]:
        if coluna in df.columns:
            df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0)

    if "Atendidas" in df.columns:
        df["Atendidas"] = pd.to_numeric(df["Atendidas"], errors="coerce").fillna(0)
    if COLUNA_ABANDONADAS not in df.columns:
        df[COLUNA_ABANDONADAS] = 0
    df[COLUNA_ABANDONADAS] = pd.to_numeric(
        df[COLUNA_ABANDONADAS],
        errors="coerce",
    ).fillna(0)
    if "TMA" in df.columns:
        df["TMA"] = pd.to_numeric(df["TMA"], errors="coerce").fillna(0)

    linhas_sem_movimento = (df["Atendidas"] == 0) & (df["CHAT"] == 0) & (df["RO"] == 0)
    colunas_para_zerar = [coluna_2min, "RO", "CHAT"]
    if "TMA" in df.columns:
        colunas_para_zerar.append("TMA")
    for nome_coluna in colunas_para_zerar:
        if nome_coluna in df.columns:
            df.loc[linhas_sem_movimento, nome_coluna] = 0

    df["Realizado"] = df[coluna_2min] + df["RO"] + df["CHAT"]
    meta_por_linha = df["Nível"].map(meta_esperada_nivel).fillna(0)
    tecnicos_normalizados = df["Técnico"].apply(normalizar_nome)
    atendidas_validas = df["Atendidas"].where(
        ~tecnicos_normalizados.isin(TECNICOS_DESCONSIDERADOS_ESPERADO),
        0,
    )
    total_meta_mes = total_meta_mensal_por_linha(df)
    total_atendidas_dia = df["Atendidas"].groupby(df["Data"]).transform("sum")
    total_abandonadas_dia = df[COLUNA_ABANDONADAS].groupby(df["Data"]).transform("max")
    total_base_fila_dia = total_atendidas_dia + total_abandonadas_dia
    proporcional = total_base_fila_dia * (
        meta_por_linha / total_meta_mes.replace(0, pd.NA)
    )
    proporcional = proporcional.fillna(0)
    df["Esperado"] = proporcional.clip(upper=meta_por_linha).map(
        arredondar_esperado
    )
    df.loc[linhas_sem_movimento, "Esperado"] = 0
    df["Desvio"] = df["Realizado"] - df["Esperado"]
    df["Classificação"] = df["Desvio"].apply(status_por_desvio)
    return df


def ultima_data_com_valor(dados, coluna, permitir_zero=False):
    if coluna not in dados.columns or dados.empty:
        return pd.NaT

    serie = pd.to_numeric(dados[coluna], errors="coerce")
    if permitir_zero:
        filtro = serie.notna()
    else:
        filtro = serie.fillna(0) > 0

    dados_validos = dados.loc[filtro]
    if dados_validos.empty:
        return pd.NaT
    return dados_validos["Data"].max()


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
        df_atual = ler_dataframe_produtividade()
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

df = ler_dataframe_produtividade()
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
env_local = carregar_env_local()
pontoweb_email = env_local.get("PONTOWEB_EMAIL", os.getenv("PONTOWEB_EMAIL", ""))
pontoweb_senha = env_local.get("PONTOWEB_SENHA", os.getenv("PONTOWEB_SENHA", ""))
pontoweb_banco_id = env_local.get(
    "PONTOWEB_BANCO_ID", os.getenv("PONTOWEB_BANCO_ID", "")
)
pontoweb_banco_identificador = env_local.get(
    "PONTOWEB_BANCO_IDENTIFICADOR",
    os.getenv("PONTOWEB_BANCO_IDENTIFICADOR", ""),
)

st.sidebar.title("🔐 Login")
for chave, valor_padrao in {
    "auth_ok": False,
    "auth_modo_gestao": False,
    "auth_usuario": "",
    "auth_tecnico": "",
}.items():
    if chave not in st.session_state:
        st.session_state[chave] = valor_padrao

if not st.session_state["auth_ok"]:
    with st.sidebar.form("form_login"):
        usuario_input = st.text_input("Usuário", key="login_usuario_input")
        senha_input = st.text_input("Senha", type="password", key="login_senha_input")
        entrar = st.form_submit_button("Entrar")

    if entrar:
        usuario_digitado = builtins.str(usuario_input).lower().strip()
        senha_digitada = builtins.str(senha_input).replace(".0", "").strip()
        login_usuario = usuarios[
            (usuarios["usuario"] == usuario_digitado)
            & (usuarios["senha"] == senha_digitada)
        ]
        usuario_apoio_valido = (
            usuario_digitado in USUARIOS_APOIO and not login_usuario.empty
        )

        if usuario_digitado == "gestao" and senha_digitada == "30071997":
            st.session_state["auth_ok"] = True
            st.session_state["auth_modo_gestao"] = True
            st.session_state["auth_usuario"] = "gestao"
            st.session_state["auth_tecnico"] = ""
            st.rerun()
        elif usuario_apoio_valido:
            st.session_state["auth_ok"] = True
            st.session_state["auth_modo_gestao"] = True
            st.session_state["auth_usuario"] = usuario_digitado
            st.session_state["auth_tecnico"] = ""
            st.rerun()
        elif not login_usuario.empty:
            st.session_state["auth_ok"] = True
            st.session_state["auth_modo_gestao"] = False
            st.session_state["auth_usuario"] = usuario_digitado
            st.session_state["auth_tecnico"] = builtins.str(
                login_usuario.iloc[0]["tecnico"]
            ).lower().strip()
            st.rerun()
        else:
            st.sidebar.error("Usuário ou senha inválidos.")

    st.warning("Digite usuário e senha.")
    st.stop()

modo_gestao = st.session_state["auth_modo_gestao"]
usuario_digitado = st.session_state["auth_usuario"]
tecnico = st.session_state["auth_tecnico"]

if st.sidebar.button("Sair", key="logout_painel"):
    for chave, valor_padrao in {
        "auth_ok": False,
        "auth_modo_gestao": False,
        "auth_usuario": "",
        "auth_tecnico": "",
        "login_usuario_input": "",
        "login_senha_input": "",
    }.items():
        st.session_state[chave] = valor_padrao
    st.rerun()

if modo_gestao:
    if usuario_digitado in USUARIOS_APOIO:
        st.sidebar.success(f"Bem-vinda {usuario_digitado.title()}")
    else:
        st.sidebar.success("Bem-vinda Gestão")

    data_referencia = data_dia_anterior()
    env_local = carregar_env_local()
    usuario_sgd_padrao = env_local.get("SGD_USUARIO", os.getenv("SGD_USUARIO", ""))
    senha_sgd_padrao = env_local.get("SGD_SENHA", os.getenv("SGD_SENHA", ""))

    st.sidebar.divider()
    st.sidebar.caption(f"Atualização para {data_referencia.strftime('%d/%m/%Y')}")
    chat_local_disponivel = servico_local_disponivel()

    if st.sidebar.button("Atualizar BI do dia anterior"):
        with st.spinner("Buscando dados no PowerBI e atualizando a planilha..."):
            try:
                origem_atualizacao = "painel"
                try:
                    resultado = atualizar_via_servico_local("bi", data_referencia)["bi"]
                    origem_atualizacao = "planilha local"
                except Exception:
                    resultado = atualizar_planilha_com_bi(data_referencia)
            except PermissionError:
                st.sidebar.error("Feche a produtividade.xlsx no Excel e tente novamente.")
            except Exception as erro:
                st.sidebar.error(f"Não foi possível atualizar o BI: {erro}")
            else:
                st.sidebar.success(
                    f"{resultado['atualizados']} técnicos atualizados em {resultado['data']}."
                )
                st.sidebar.caption(f"Origem da gravação: {origem_atualizacao}.")
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
                origem_atualizacao = "painel"
                try:
                    resultado = atualizar_via_servico_local("ro", data_referencia)["ro"]
                    origem_atualizacao = "planilha local"
                except Exception:
                    resultado = atualizar_planilha_com_ro(data_referencia)
            except PermissionError:
                st.sidebar.error("Feche a produtividade.xlsx no Excel e tente novamente.")
            except Exception as erro:
                st.sidebar.error(f"Não foi possível atualizar o RO: {erro}")
            else:
                st.sidebar.success(
                    f"{resultado['atualizados']} técnicos atualizados em {resultado['data']}."
                )
                st.sidebar.caption(f"Origem da gravação: {origem_atualizacao}.")
                st.sidebar.info(
                    f"Origem: {resultado['arquivo']}. Total de RO válidos: {resultado['fonte']}."
                )
                if resultado["linhas_criadas"]:
                    st.sidebar.info(
                        f"{resultado['linhas_criadas']} linhas foram criadas para essa data."
                    )
                st.cache_data.clear()
                st.rerun()

    if not chat_local_disponivel:
        st.sidebar.info(
            "O CHAT só pode ser atualizado no painel local "
            "(http://localhost:8501) com o serviço local ativo, "
            "porque ele depende do navegador desta máquina logado no PLUG."
        )

    if st.sidebar.button(
        "Atualizar CHAT do dia anterior",
        disabled=not chat_local_disponivel,
    ):
        with st.spinner("Buscando chats fechados no PLUG e atualizando a planilha..."):
            try:
                resultado = atualizar_via_servico_local("chat", data_referencia)["chat"]
                origem_atualizacao = "planilha local"
            except PermissionError:
                st.sidebar.error("Feche a produtividade.xlsx no Excel e tente novamente.")
            except Exception as erro:
                st.sidebar.error(f"Não foi possível atualizar o CHAT: {erro}")
            else:
                st.sidebar.success(
                    f"{resultado['atualizados']} técnicos atualizados em {resultado['data']}."
                )
                st.sidebar.caption(f"Origem da gravação: {origem_atualizacao}.")
                if resultado["linhas_criadas"]:
                    st.sidebar.info(
                        f"{resultado['linhas_criadas']} linhas foram criadas para essa data."
                    )
                if resultado["sem_alias"]:
                    st.sidebar.warning(
                        "Revise a coluna Agente Chat da aba Aliases para: "
                        + ", ".join(resultado["sem_alias"][:8])
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
                    origem_atualizacao = "painel"
                    try:
                        resultado = atualizar_via_servico_local("sgd", data_referencia)["sgd"]
                        origem_atualizacao = "planilha local"
                    except Exception:
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
                    st.sidebar.caption(f"Origem da gravação: {origem_atualizacao}.")
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
    st.sidebar.success(f"Bem-vindo(a), {tecnico.title()}")

if modo_gestao:
    st.sidebar.divider()
    with st.sidebar.expander("PontoWeb", expanded=False):
        pontoweb_email = st.text_input(
            "Usuário Ponto",
            value=pontoweb_email,
            key="pontoweb_email_input",
        )
        pontoweb_senha = st.text_input(
            "Senha Ponto",
            value=pontoweb_senha,
            type="password",
            key="pontoweb_senha_input",
        )
        pontoweb_banco_id = st.text_input(
            "Banco ID (opcional)",
            value=pontoweb_banco_id,
            key="pontoweb_banco_id_input",
        )
        pontoweb_banco_identificador = st.text_input(
            "Banco identificador (opcional)",
            value=pontoweb_banco_identificador,
            key="pontoweb_banco_identificador_input",
        )
        if st.button("Salvar credenciais do Ponto", key="salvar_pontoweb"):
            try:
                salvar_env_local(
                    {
                        "PONTOWEB_EMAIL": pontoweb_email,
                        "PONTOWEB_SENHA": pontoweb_senha,
                        "PONTOWEB_BANCO_ID": pontoweb_banco_id,
                        "PONTOWEB_BANCO_IDENTIFICADOR": pontoweb_banco_identificador,
                    }
                )
            except Exception as erro_salvar_pontoweb:
                st.sidebar.error(
                    f"Não foi possível salvar as credenciais do PontoWeb: {erro_salvar_pontoweb}"
                )
            else:
                st.sidebar.success("Credenciais do PontoWeb salvas.")
                st.cache_data.clear()
                st.rerun()

        if st.button("Atualizar Ponto do mês", key="atualizar_pontoweb_mes"):
            if not pontoweb_email or not pontoweb_senha:
                st.sidebar.error("Informe usuário e senha do PontoWeb.")
            else:
                with st.spinner("Buscando abatimentos no PontoWeb e atualizando a planilha..."):
                    try:
                        resultado = atualizar_planilha_com_pontoweb_mes(
                            data_referencia.year,
                            data_referencia.month,
                            pontoweb_email,
                            pontoweb_senha,
                            pontoweb_banco_id,
                            pontoweb_banco_identificador,
                        )
                    except PermissionError:
                        st.sidebar.error("Feche a produtividade.xlsx no Excel e tente novamente.")
                    except Exception as erro:
                        st.sidebar.error(f"Não foi possível atualizar o PontoWeb: {erro}")
                    else:
                        st.sidebar.success(
                            f"{resultado['tecnicos_processados']} técnicos processados. "
                            f"{resultado['linhas_atualizadas']} linhas atualizadas."
                        )
                        if resultado["erros"]:
                            st.sidebar.warning(
                                "Sem retorno do PontoWeb para: "
                                + ", ".join(list(resultado["erros"].keys())[:6])
                            )
                        st.cache_data.clear()
                        st.rerun()

    if not pontoweb_email or not pontoweb_senha:
        st.sidebar.caption(
            "O quadro Dias Úteis para Meta usa o Ponto quando o usuário e a senha do PontoWeb são informados."
        )

    tecnicos_disponiveis = sorted(df["Técnico"].unique())
    tecnico = st.session_state.get("tecnico_gestao", tecnicos_disponiveis[0])
    if tecnico not in tecnicos_disponiveis:
        tecnico = tecnicos_disponiveis[0]

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
ultima_data_produtividade = ultima_data_com_valor(dados_mes_atual, "Realizado")
if pd.isna(ultima_data_produtividade):
    ultima_data_produtividade = ultima_data_mes
dados_ultimo_dia = dados_mes_atual[dados_mes_atual["Data"] == ultima_data_produtividade]

ultima_data_satisfacao = ultima_data_com_valor(dados_mes_atual, "Satisfação")
if pd.isna(ultima_data_satisfacao):
    ultima_data_satisfacao = ultima_data_mes
dados_ultimo_dia_satisfacao = dados_mes_atual[
    dados_mes_atual["Data"] == ultima_data_satisfacao
]

ultima_data_votacao = ultima_data_com_valor(dados_mes_atual, "Votação")
if pd.isna(ultima_data_votacao):
    ultima_data_votacao = ultima_data_mes
dados_ultimo_dia_votacao = dados_mes_atual[
    dados_mes_atual["Data"] == ultima_data_votacao
]


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
    mostrar_lista_apoio_gestao()

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
    percentuais_absorcao_mes = percentuais_absorcao_por_mes(
        df,
        ano_atual,
        mes_atual,
    )
    col1, col2, col3, col4, col5 = st.columns(5)
    for (nivel, percentual), coluna in zip(
        percentuais_absorcao_mes.items(),
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
        base_ranking_geral[base_ranking_geral["Data"] == ultima_data_produtividade]
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

if modo_gestao:
    tecnico = st.selectbox(
        "Selecione o Técnico",
        tecnicos_disponiveis,
        index=tecnicos_disponiveis.index(tecnico),
        key="tecnico_gestao",
    )

st.divider()
st.subheader(f"📌 Resultados Individuais - {tecnico.title()}")

nivel_mode = dados_mes_atual["Nível"].dropna().mode()
nivel_tecnico = None if nivel_mode.empty else nivel_mode.iloc[0]

if modo_gestao:
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
else:
    col1, col2, col3, col4, col5, col6, col7, col8, col9 = st.columns(9)

with col1:
    st.metric("Atendidas BI Total", int(dados_mes_atual["Atendidas"].sum()))

with col2:
    st.metric("Realizado Total", int(dados_mes_atual["Realizado"].sum()))

with col3:
    st.metric("SSC Total", int(dados_mes_atual["SSC"].sum()))

with col4:
    st.metric("RO Total", int(dados_mes_atual["RO"].sum()))

with col5:
    coluna_votacao = nome_coluna_dataframe(
        dados_ultimo_dia_votacao,
        "Votação",
        "VotaÃ§Ã£o",
        "Vota??o",
    )
    votacao_ultimo_dia = round(dados_ultimo_dia_votacao[coluna_votacao].mean(), 2)
    st.metric("Votação Média", f"{votacao_ultimo_dia}%")

with col6:
    coluna_satisfacao = nome_coluna_dataframe(
        dados_ultimo_dia_satisfacao,
        "Satisfação",
        "SatisfaÃ§Ã£o",
        "Satisfa??o",
    )
    satisfacao_ultimo_dia = round(
        dados_ultimo_dia_satisfacao[coluna_satisfacao].mean(),
        2,
    )
    st.metric("Satisfação", f"{satisfacao_ultimo_dia}%")

with col7:
    coluna_classificacao = nome_coluna_dataframe(
        dados_mes_atual,
        "Classificação",
        "ClassificaÃ§Ã£o",
        "Classifica??o",
    )
    classificacao_mode = dados_mes_atual[coluna_classificacao].dropna().mode()
    classificacao = (
        "Sem classificação"
        if classificacao_mode.empty
        else classificacao_mode.iloc[0]
    )
    st.metric("Classificação", classificacao)

if not modo_gestao:
    percentual_absorcao = percentual_absorcao_tecnico(
        df,
        ano_atual,
        mes_atual,
        nivel_tecnico,
    )
    dias_meta_valor = float(dias_uteis_para_meta(ano_atual, mes_atual))
    fonte_dias_meta = "Calendário do mês"
    dias_segunda_sexta = sum(
        1
        for dia in range(
            1,
            calendar.monthrange(ano_atual, mes_atual)[1] + 1,
        )
        if date(ano_atual, mes_atual, dia).weekday() < 5
    )
    feriados_em_dias_uteis = int(dias_segunda_sexta - dias_meta_valor)
    detalhe_dias_meta = (
        f"{dias_segunda_sexta} dias de segunda a sexta"
        f" - {feriados_em_dias_uteis} feriado(s)"
        f" = {int(dias_meta_valor)} dias."
    )
    if COLUNA_DIAS_META in dados_tecnico.columns:
        dias_meta_salvos = (
            dados_tecnico[
                (dados_tecnico["Data"].dt.month == mes_atual)
                & (dados_tecnico["Data"].dt.year == ano_atual)
            ][COLUNA_DIAS_META]
            .dropna()
        )
        if not dias_meta_salvos.empty:
            dias_meta_valor = float(dias_meta_salvos.iloc[-1])
            fonte_dias_meta = "Planilha"
            detalhe_dias_meta = "Valor atualizado pela gestão a partir do PontoWeb."
    elif pontoweb_email and pontoweb_senha:
        try:
            resumo_meta_pontoweb = obter_resumo_meta_pontoweb(
                tecnico,
                ano_atual,
                mes_atual,
                pontoweb_email,
                pontoweb_senha,
                pontoweb_banco_id,
                pontoweb_banco_identificador,
            )
            dias_meta_valor = float(resumo_meta_pontoweb["dias_considerados"])
            fonte_dias_meta = "PontoWeb"
            detalhe_dias_meta = (
                f"Dias base: {resumo_meta_pontoweb['dias_base']} | "
                f"Abatimento: {resumo_meta_pontoweb['abatimento']}"
            )
        except Exception as erro_pontoweb:
            detalhe_dias_meta = f"PontoWeb indisponível: {erro_pontoweb}"

    if dias_meta_valor.is_integer():
        dias_meta_exibicao = builtins.str(int(dias_meta_valor))
    else:
        dias_meta_exibicao = f"{dias_meta_valor:.2f}".replace(".", ",")

    with col7:
        st.metric("Percentual de Absorção", f"{percentual_absorcao:.2f}%")
    with col8:
        st.metric(
            "Dias Úteis para Meta",
            dias_meta_exibicao,
        )
    st.caption(f"Fonte: {fonte_dias_meta}. {detalhe_dias_meta}")

    st.caption("Referência da classificação")
    col_leg_1, col_leg_2, col_leg_3, col_leg_4 = st.columns(4)
    for status, descricao, coluna in zip(
        ["CRÍTICO", "ATENÇÃO", "BOM", "EXCELENTE"],
        [
            "Desvio menor que -5",
            "Desvio entre -5 e menor que 0",
            "Desvio entre 0 e 5",
            "Desvio acima de 5",
        ],
        [col_leg_1, col_leg_2, col_leg_3, col_leg_4],
    ):
        with coluna:
            st.markdown(
                f"""
                <div style="
                    border:1px solid #E5E7EB;
                    border-left:6px solid {CORES_STATUS[status]};
                    border-radius:10px;
                    padding:10px 12px;
                    background-color:#FFFFFF;
                    min-height:74px;
                ">
                    <div style="font-weight:700;color:#111827;font-size:13px;">
                        {status}
                    </div>
                    <div style="color:#6B7280;font-size:12px;margin-top:4px;">
                        {descricao}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.divider()
    st.subheader("Lista de Apoio")
    col_atualizar_apoio_tecnico, col_status_apoio_tecnico = st.columns([1, 3])
    with col_atualizar_apoio_tecnico:
        if st.button("Atualizar minhas ajudas", key="atualizar_lista_apoio_tecnico"):
            st.rerun()
    with col_status_apoio_tecnico:
        st.caption("Use o botão para atualizar suas ajudas sem sair do painel.")

    with st.form("form_lista_apoio", clear_on_submit=True):
        topico_apoio = st.selectbox("Tópico do problema", TOPICOS_LISTA_APOIO)
        resumo_apoio = st.text_area(
            "Descreva a dúvida de forma resumida",
            height=110,
        )
        enviar_apoio = st.form_submit_button("Registrar dúvida")

    if enviar_apoio:
        if not resumo_apoio.strip():
            st.warning("Descreva a dúvida antes de registrar.")
        else:
            try:
                registrar_duvida_apoio(
                    tecnico,
                    topico_apoio,
                    resumo_apoio,
                )
            except PermissionError:
                st.error("Feche a lista_apoio.xlsx e tente registrar novamente.")
            except Exception as erro_lista_apoio:
                st.error(f"Não foi possível registrar a dúvida: {erro_lista_apoio}")
            else:
                st.success("Dúvida registrada para apoio.")

    try:
        lista_apoio_tecnico = ler_lista_apoio()
        lista_apoio_tecnico = lista_apoio_tecnico[
            lista_apoio_tecnico["Nome do Técnico"].map(normalizar_nome) == normalizar_nome(tecnico)
        ].copy()
    except Exception as erro_lista_apoio:
        st.caption(f"Não foi possível carregar suas ajudas registradas: {erro_lista_apoio}")
        lista_apoio_tecnico = pd.DataFrame(columns=CABECALHOS_LISTA_APOIO)

    if lista_apoio_tecnico.empty:
        st.caption("Você ainda não possui ajudas registradas.")
    else:
        lista_apoio_tecnico["Data"] = pd.to_datetime(
            lista_apoio_tecnico["Carimbo de data/hora"],
            dayfirst=True,
            errors="coerce",
        )
        lista_apoio_tecnico = lista_apoio_tecnico.sort_values(
            by="Data",
            ascending=False,
            na_position="last",
        )
        lista_apoio_tecnico["Carimbo de data/hora"] = lista_apoio_tecnico[
            "Carimbo de data/hora"
        ].fillna("")

        st.caption("Acompanhe abaixo a situação das ajudas que você registrou.")
        st.dataframe(
            lista_apoio_tecnico[
                [
                    "Carimbo de data/hora",
                    "Selecione o tópico",
                    "Descreva o problema/pedido em poucas palavras",
                    "Situação",
                    "Responsável/Apoio",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

if not modo_gestao:
    if nivel_tecnico:
        base_ranking_nivel = df[
            (df["Data"].dt.month == mes_atual)
            & (df["Data"].dt.year == ano_atual)
            & (df["Nível"] == nivel_tecnico)
        ]

        ranking_nivel_diario = (
            base_ranking_nivel[base_ranking_nivel["Data"] == ultima_data_produtividade]
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

data_inicial_padrao = dados_produtividade["Data"].min()
data_final_padrao = dados_produtividade["Data"].max()

col_periodo_1, col_periodo_2 = st.columns(2)
with col_periodo_1:
    periodo_inicial_texto = st.text_input(
        "Período inicial",
        value=data_inicial_padrao.strftime("%d/%m/%Y"),
    )
with col_periodo_2:
    periodo_final_texto = st.text_input(
        "Período final",
        value=data_final_padrao.strftime("%d/%m/%Y"),
    )

try:
    periodo_inicial = datetime.strptime(periodo_inicial_texto, "%d/%m/%Y").date()
    periodo_final = datetime.strptime(periodo_final_texto, "%d/%m/%Y").date()
except ValueError:
    st.error("Informe o período no formato dd/mm/aaaa.")
    st.stop()

if periodo_inicial > periodo_final:
    st.error("O período inicial não pode ser maior que o período final.")
    st.stop()

dados_produtividade = dados_produtividade[
    (dados_produtividade["Data"].dt.date >= periodo_inicial)
    & (dados_produtividade["Data"].dt.date <= periodo_final)
]
dados_produtividade = dados_produtividade[
    (dados_produtividade["Data"].dt.weekday < 5)
    & ~dados_produtividade["Data"].dt.date.map(eh_feriado_federal)
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
