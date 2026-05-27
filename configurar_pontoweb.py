import argparse
import getpass
import pathlib


ENV_PATH = pathlib.Path(__file__).with_name(".env")


def carregar_env():
    valores = {}
    if not ENV_PATH.exists():
        return valores

    for linha in ENV_PATH.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, valor = linha.split("=", 1)
        valores[chave.strip()] = valor.strip().strip('"').strip("'")
    return valores


def salvar_env(valores):
    linhas = [f"{chave}={valor}" for chave, valor in sorted(valores.items())]
    ENV_PATH.write_text("\n".join(linhas) + "\n", encoding="utf-8")


def valor_argumento_ou_prompt(valor_atual, prompt, secreto=False):
    if valor_atual:
        return valor_atual
    if secreto:
        return getpass.getpass(f"{prompt}: ").strip()
    return input(f"{prompt}: ").strip()


def main():
    parser = argparse.ArgumentParser(
        description="Configura rapidamente as credenciais do PontoWeb no .env."
    )
    parser.add_argument("--usuario", help="Usuário/login do PontoWeb.")
    parser.add_argument("--senha", help="Senha do PontoWeb.")
    parser.add_argument("--banco-id", help="Banco ID do PontoWeb.")
    parser.add_argument(
        "--banco-identificador",
        help="Identificador do banco/empresa no PontoWeb.",
    )
    args = parser.parse_args()

    valores = carregar_env()

    usuario = valor_argumento_ou_prompt(args.usuario, "Usuário PontoWeb")
    senha = valor_argumento_ou_prompt(args.senha, "Senha PontoWeb", secreto=True)
    banco_id = args.banco_id
    banco_identificador = args.banco_identificador

    if banco_id is None:
        banco_id = input("Banco ID (opcional): ").strip()
    if banco_identificador is None:
        banco_identificador = input("Banco identificador (opcional): ").strip()

    valores["PONTOWEB_EMAIL"] = usuario
    valores["PONTOWEB_SENHA"] = senha
    valores["PONTOWEB_BANCO_ID"] = banco_id
    valores["PONTOWEB_BANCO_IDENTIFICADOR"] = banco_identificador

    salvar_env(valores)
    print(f"Credenciais do PontoWeb salvas em {ENV_PATH}")


if __name__ == "__main__":
    main()
