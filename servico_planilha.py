import json
import pathlib
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer


HOST = "127.0.0.1"
PORT = 8765
APP_PATH = pathlib.Path(__file__).with_name("app.py")
CORTE_APP = "try:\n    locale.setlocale"


def carregar_funcoes_app():
    codigo = APP_PATH.read_text(encoding="utf-8")
    corte = codigo.index(CORTE_APP)
    namespace = {}
    exec(codigo[:corte], namespace)
    return namespace


NS = carregar_funcoes_app()


def parse_data(valor):
    if not valor:
        return NS["data_dia_anterior"]()
    return datetime.strptime(valor, "%d/%m/%Y").date()


def atualizar_fontes(fonte, data_referencia):
    resultados = {}

    if fonte in {"bi", "todas"}:
        resultados["bi"] = NS["atualizar_planilha_com_bi"](data_referencia)

    if fonte in {"ro", "todas"}:
        resultados["ro"] = NS["atualizar_planilha_com_ro"](data_referencia)

    if fonte in {"chat", "todas"}:
        resultados["chat"] = NS["atualizar_planilha_com_chat"](data_referencia)

    if fonte in {"sgd", "todas"}:
        env_local = NS["carregar_env_local"]()
        usuario_sgd = env_local.get("SGD_USUARIO") or NS["os"].getenv("SGD_USUARIO")
        senha_sgd = env_local.get("SGD_SENHA") or NS["os"].getenv("SGD_SENHA")
        if not usuario_sgd or not senha_sgd:
            raise RuntimeError(
                "Credenciais do SGD ausentes no .env para executar a atualização."
            )
        resultados["sgd"] = NS["atualizar_planilha_com_sgd"](
            data_referencia,
            usuario_sgd,
            senha_sgd,
        )

    return resultados


class Handler(BaseHTTPRequestHandler):
    server_version = "ServicoPlanilha/1.0"

    def _json(self, status, payload):
        corpo = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(corpo)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()
        self.wfile.write(corpo)

    def do_OPTIONS(self):
        self._json(200, {"ok": True})

    def do_GET(self):
        if self.path == "/health":
            return self._json(
                200,
                {
                    "ok": True,
                    "service": "planilha",
                    "host": HOST,
                    "port": PORT,
                },
            )
        self._json(404, {"ok": False, "erro": "Rota não encontrada."})

    def do_POST(self):
        if self.path != "/atualizar":
            return self._json(404, {"ok": False, "erro": "Rota não encontrada."})

        try:
            tamanho = int(self.headers.get("Content-Length", "0"))
            bruto = self.rfile.read(tamanho) if tamanho else b"{}"
            dados = json.loads(bruto.decode("utf-8"))
            fonte = str(dados.get("fonte", "todas")).strip().lower()
            if fonte not in {"bi", "sgd", "ro", "chat", "todas"}:
                raise ValueError("Fonte inválida. Use bi, sgd, ro, chat ou todas.")

            data_referencia = parse_data(dados.get("data"))
            resultados = atualizar_fontes(fonte, data_referencia)
            self._json(
                200,
                {
                    "ok": True,
                    "fonte": fonte,
                    "data": data_referencia.strftime("%d/%m/%Y"),
                    "resultados": resultados,
                },
            )
        except Exception as erro:
            self._json(500, {"ok": False, "erro": str(erro)})


def main():
    servidor = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Serviço local da planilha rodando em http://{HOST}:{PORT}")
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        servidor.server_close()


if __name__ == "__main__":
    main()
