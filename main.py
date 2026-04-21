#!/usr/bin/env python3
"""
Gerador automático de relatórios semanais de marketing digital.
Uso: python main.py
"""
import json
import os
import sys

from data_processor import process_ecommerce, process_lead
from pptx_filler import fill_template


def ask(prompt: str, options=None) -> str:
    while True:
        resp = input(prompt).strip()
        if options is None:
            if resp:
                return resp
        else:
            if resp.isdigit() and int(resp) in options:
                return resp
        print("  Opção inválida, tente novamente.")


def main():
    config_path = os.path.join(os.path.dirname(__file__), "clients_config.json")
    with open(config_path, encoding="utf-8") as f:
        config = json.load(f)

    clients = list(config.keys())

    print("\nClientes disponíveis:")
    for i, c in enumerate(clients, 1):
        print(f"  {i}. {c}")

    idx = int(ask("Escolha o cliente: ", range(1, len(clients) + 1))) - 1
    client_key = clients[idx]
    client = config[client_key]

    # Tipo vem do config; só pergunta se não estiver definido
    tipo = client.get("tipo", "").lower()
    if tipo not in ("ecommerce", "lead"):
        print("\nTipo do cliente:")
        print("  1. Ecommerce")
        print("  2. Lead")
        tipo_idx = int(ask("Escolha o tipo: ", range(1, 3)))
        tipo = "ecommerce" if tipo_idx == 1 else "lead"
    else:
        print(f"\nTipo: {tipo}")

    periodo = ask("\nPeríodo analisado (ex: 14/04 a 20/04): ")
    periodo_comp = ask("Período comparado  (ex: 07/04 a 13/04): ")

    planilha = client.get("planilha", f"{client_key}.xlsx")
    xlsx_path = os.path.join(os.path.dirname(__file__), "clientes", planilha)
    if not os.path.exists(xlsx_path):
        sys.exit(f"Planilha não encontrada: {xlsx_path}")

    sheet_id = client.get("sheet_id")

    print("\nProcessando dados...")
    try:
        if tipo == "ecommerce":
            dados = process_ecommerce(xlsx_path, periodo, periodo_comp, client["metas"], sheet_id)
        else:
            dados = process_lead(xlsx_path, periodo, periodo_comp, client["metas"], sheet_id)
    except Exception as e:
        sys.exit(f"Erro ao processar planilha: {e}")

    dados["nome_cliente"] = client["nome"]
    dados["periodo_analisado"] = periodo
    dados["periodo_comparado"] = periodo_comp

    print("Gerando apresentação...")
    try:
        out = fill_template(tipo, dados, client)
    except FileNotFoundError as e:
        sys.exit(str(e))

    print(f"\nRelatório salvo em: {out}")


if __name__ == "__main__":
    main()
