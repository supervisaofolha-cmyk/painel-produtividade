import argparse
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


SRC = Path(r"C:\Users\esther.queiroz\Downloads\clientes_vs_abandonos 052026.xlsx")
OUT = Path(r"C:\Users\esther.queiroz\Downloads\clientes_vs_abandonos 052026_enriquecido.xlsx")
PROFILE = Path(r"C:\Users\esther.queiroz\AppData\Local\Temp\sgd_playwright_profile")
SEARCH_URL = "https://sgd.dominiosistemas.com.br/sgsc/faces/loc-cliente.html"
LOGIN_URL_PART = "sgd.dominiosistemas.com.br/login"
CDP_URL = "http://127.0.0.1:9222"

NEW_HEADERS = [
    "Nome_Licenciado",
    "Codigo_DECA",
    "SGD_Status",
    "SGD_Observacao",
    "Consultado_Em",
]


def digits(value):
    return re.sub(r"\D+", "", str(value or ""))


def ensure_workbook():
    if OUT.exists():
        wb = load_workbook(OUT)
    else:
        wb = load_workbook(SRC)

    ws = wb["Resultado da consulta"]
    headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
    for header in NEW_HEADERS:
        if header not in headers:
            ws.cell(1, ws.max_column + 1).value = header
            headers.append(header)

    wb.save(OUT)
    return wb, ws


def header_map(ws):
    return {ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)}


def normalize_text(text):
    return re.sub(r"\s+", " ", text or "").strip()


def extract_from_text(text, phone):
    clean = normalize_text(text)
    zero_occurrence = re.search(r"Encontrada\(s\)\s*0\s*ocorr", clean, re.IGNORECASE)
    if zero_occurrence:
        return {
            "status": "Nao encontrado",
            "nome": "",
            "deca": "",
            "obs": "Nenhum resultado na pesquisa por telefone",
        }

    if "Nenhum registro" in clean or "não encontrado" in clean.lower():
        return {
            "status": "Nao encontrado",
            "nome": "",
            "deca": "",
            "obs": "Nenhum resultado na pesquisa por telefone",
        }

    deca_match = re.search(r"(?:DECA|Deca|deca)\D{0,20}(\d{2,})", clean)
    code_match = re.search(r"(?:Código|Codigo|Cód\.?|Cod\.?)\D{0,20}(\d{2,})", clean)

    deca = deca_match.group(1) if deca_match else ""
    code = code_match.group(1) if code_match else ""

    # Prefer labels commonly used by the client details page.
    name = ""
    for pattern in [
        r"(?:Licenciado|Cliente|Nome|Razão Social|Razao Social)\s*:?\s*([A-Z0-9][^:\n\r]{3,120})",
        r"\b\d{2,}\s+([A-Z][A-Z0-9 .,&/-]{5,120})",
    ]:
        match = re.search(pattern, clean)
        if match:
            name = normalize_text(match.group(1))
            break

    if deca or code or name:
        return {
            "status": "Encontrado",
            "nome": name,
            "deca": deca or code,
            "obs": "",
        }

    if phone in clean:
        return {
            "status": "Encontrado",
            "nome": "",
            "deca": "",
            "obs": "Resultado encontrado, mas campos nao identificados automaticamente",
        }

    return {
        "status": "Nao encontrado",
        "nome": "",
        "deca": "",
        "obs": "Telefone nao apareceu no resultado da pesquisa",
    }


def extract_result_rows(page):
    return page.eval_on_selector_all(
        "table.tableSorter tbody tr",
        """rows => rows.map(row => {
            const cells = Array.from(row.querySelectorAll('td')).map(td =>
                td.innerText.replace(/\\s+/g, ' ').trim()
            );
            const details = row.querySelector('a[href*="d-cliente.html?clienteID="]');
            return {
                codigo: cells[0] || '',
                razao: cells[1] || '',
                representante: cells[2] || '',
                tipo: cells[6] || '',
                telefone: cells[9] || '',
                detailsHref: details ? details.getAttribute('href') : ''
            };
        }).filter(row => row.codigo || row.razao)""",
    )


def result_from_row(row):
    obs = []
    if row.get("representante"):
        obs.append(f"Representante: {row['representante']}")
    if row.get("telefone"):
        obs.append(f"Telefone suporte: {row['telefone']}")
    if row.get("tipo"):
        obs.append(f"Tipo: {row['tipo']}")
    return {
        "status": "Encontrado",
        "nome": row.get("razao", ""),
        "deca": row.get("codigo", ""),
        "obs": "; ".join(obs),
    }


def save_result(ws, cols, row, result):
    ws.cell(row, cols["Nome_Licenciado"]).value = result["nome"]
    ws.cell(row, cols["Codigo_DECA"]).value = result["deca"]
    ws.cell(row, cols["SGD_Status"]).value = result["status"]
    ws.cell(row, cols["SGD_Observacao"]).value = result["obs"]
    ws.cell(row, cols["Consultado_Em"]).value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def wait_for_login(page):
    page.goto(SEARCH_URL, wait_until="load")
    if LOGIN_URL_PART not in page.url:
        return

    print("Chrome aberto na tela de login do SGD. Faça login nessa janela.")
    print("Depois do login, o script continua automaticamente.")
    deadline = time.time() + 600
    while time.time() < deadline:
        if LOGIN_URL_PART not in page.url:
            page.goto(SEARCH_URL, wait_until="load")
            return
        time.sleep(2)
    raise RuntimeError("Login nao concluido em 10 minutos.")


def search_phone(page, phone, debug_dir, row):
    page.goto(SEARCH_URL, wait_until="load")
    page.locator('select[name="locForm:usuario"]').select_option("5")
    page.locator('input[name="locForm:palavraChave"]').fill(phone)

    page.locator('input[name="locForm:localizarBtn"]').click(no_wait_after=True)
    page.wait_for_load_state("domcontentloaded", timeout=30000)
    page.wait_for_timeout(1500)

    result_rows = extract_result_rows(page)
    if len(result_rows) == 1:
        return result_from_row(result_rows[0])
    if len(result_rows) > 1:
        phone_matches = [
            item for item in result_rows if digits(item.get("telefone")) == phone
        ]
        if len(phone_matches) == 1:
            return result_from_row(phone_matches[0])

        candidates = " | ".join(
            f"{item.get('codigo')} - {item.get('razao')}" for item in result_rows
        )
        debug_path = debug_dir / f"row_{row}_{phone}_multiplo.html"
        debug_path.write_text(page.content(), encoding="utf-8")
        return {
            "status": "Multiplo resultado",
            "nome": "",
            "deca": "",
            "obs": f"{len(result_rows)} resultados: {candidates}; HTML: {debug_path}",
        }

    body_text = page.locator("body").inner_text(timeout=10000)
    result = extract_from_text(body_text, phone)
    if result["status"] != "Encontrado" or not result["deca"] or not result["nome"]:
        debug_path = debug_dir / f"row_{row}_{phone}.html"
        debug_path.write_text(page.content(), encoding="utf-8")
        if result["obs"]:
            result["obs"] = f"{result['obs']}; HTML: {debug_path}"
        else:
            result["obs"] = f"HTML: {debug_path}"
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=10)
    parser.add_argument("--start-row", type=int, default=2)
    parser.add_argument("--save-every", type=int, default=10)
    parser.add_argument("--connect-cdp", action="store_true")
    args = parser.parse_args()

    wb, ws = ensure_workbook()
    cols = header_map(ws)
    debug_dir = OUT.with_suffix("")
    debug_dir.mkdir(exist_ok=True)

    rows = []
    for row in range(args.start_row, ws.max_row + 1):
        phone = digits(ws.cell(row, 1).value)
        if not phone:
            continue
        if ws.cell(row, cols["SGD_Status"]).value:
            continue
        rows.append((row, phone))
        if args.limit and len(rows) >= args.limit:
            break

    print(f"Linhas pendentes selecionadas: {len(rows)}")
    if not rows:
        return 0

    with sync_playwright() as p:
        browser = None
        if args.connect_cdp:
            browser = p.chromium.connect_over_cdp(CDP_URL)
            context = browser.contexts[0]
            page = context.pages[0] if context.pages else context.new_page()
        else:
            context = p.chromium.launch_persistent_context(
                str(PROFILE),
                channel="chrome",
                headless=False,
                viewport={"width": 1280, "height": 800},
            )
            page = context.pages[0] if context.pages else context.new_page()
        wait_for_login(page)

        done = 0
        for row, phone in rows:
            try:
                print(f"Consultando linha {row}: {phone}")
                result = search_phone(page, phone, debug_dir, row)
            except PlaywrightTimeoutError as exc:
                result = {
                    "status": "Erro",
                    "nome": "",
                    "deca": "",
                    "obs": f"Timeout na consulta: {exc}",
                }
            except Exception as exc:
                result = {
                    "status": "Erro",
                    "nome": "",
                    "deca": "",
                    "obs": f"Erro na consulta: {exc}",
                }

            save_result(ws, cols, row, result)
            done += 1
            print(f"  -> {result['status']} | {result['nome']} | {result['deca']}")
            if done % args.save_every == 0:
                wb.save(OUT)
                print(f"Progresso salvo em {OUT}")

        wb.save(OUT)
        if browser:
            browser.close()
        else:
            context.close()

    print(f"Finalizado. Arquivo: {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
