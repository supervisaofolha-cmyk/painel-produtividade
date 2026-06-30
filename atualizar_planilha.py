import argparse
import pathlib
from datetime import datetime


APP_PATH = pathlib.Path(__file__).with_name("app.py")
CORTE_APP = "# STREAMLIT_APP_START"


def carregar_funcoes_app():
    codigo = APP_PATH.read_text(encoding="utf-8")
    try:
        corte = codigo.index(CORTE_APP)
    except ValueError as exc:
        raise RuntimeError(
            "Nao foi possivel localizar o marcador de inicio da interface no app.py."
        ) from exc
    namespace = {}
    exec(codigo[:corte], namespace)
    return namespace


def parse_data(valor):
    return datetime.strptime(valor, "%d/%m/%Y").date()


def main():
    parser = argparse.ArgumentParser(
        description="Atualiza produtividade.xlsx sem alterar o painel."
    )
    parser.add_argument(
        "--data",
        help="Data de referencia no formato DD/MM/AAAA. Se omitido, usa o ultimo dia util.",
    )
    parser.add_argument("--sem-bi", action="store_true", help="Nao atualiza BI.")
    parser.add_argument("--sem-ro", action="store_true", help="Nao atualiza RO.")
    parser.add_argument("--sem-sgd", action="store_true", help="Nao atualiza SGD.")
    args = parser.parse_args()

    ns = carregar_funcoes_app()
    data_referencia = (
        parse_data(args.data) if args.data else ns["data_dia_anterior"]()
    )

    print(f"Data de referencia: {data_referencia.strftime('%d/%m/%Y')}")

    if not args.sem_bi:
        resultado_bi = ns["atualizar_planilha_com_bi"](data_referencia)
        print(f"BI: {resultado_bi}")

    if not args.sem_ro:
        resultado_ro = ns["atualizar_planilha_com_ro"](data_referencia)
        print(f"RO: {resultado_ro}")

    if not args.sem_sgd:
        env_local = ns["carregar_env_local"]()
        usuario_sgd = env_local.get("SGD_USUARIO") or ns["os"].getenv("SGD_USUARIO")
        senha_sgd = env_local.get("SGD_SENHA") or ns["os"].getenv("SGD_SENHA")

        if not usuario_sgd or not senha_sgd:
            raise RuntimeError(
                "Credenciais do SGD ausentes. Defina SGD_USUARIO e SGD_SENHA no .env."
            )

        resultado_sgd = ns["atualizar_planilha_com_sgd"](
            data_referencia,
            usuario_sgd,
            senha_sgd,
        )
        print(f"SGD: {resultado_sgd}")


if __name__ == "__main__":
    main()
