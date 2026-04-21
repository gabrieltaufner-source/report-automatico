import io
import json
import os
import tempfile

from flask import Flask, render_template, request, send_file, jsonify

from data_processor import process_ecommerce, process_lead
from pptx_filler import fill_template_to_buffer

app = Flask(__name__)

BASE_DIR = os.path.dirname(__file__)
CONFIG_PATH = os.path.join(BASE_DIR, "clients_config.json")

with open(CONFIG_PATH, encoding="utf-8") as f:
    CONFIG = json.load(f)


@app.route("/")
def index():
    clients = [
        {"key": k, "nome": v["nome"], "tipo": v.get("tipo", "")}
        for k, v in CONFIG.items()
    ]
    return render_template("index.html", clients=clients)


@app.route("/gerar", methods=["POST"])
def gerar():
    client_key = request.form.get("cliente")
    periodo = request.form.get("periodo", "").strip()
    periodo_comp = request.form.get("periodo_comp", "").strip()

    if not client_key or not periodo or not periodo_comp:
        return jsonify({"erro": "Preencha todos os campos."}), 400

    if client_key not in CONFIG:
        return jsonify({"erro": "Cliente não encontrado."}), 400

    client = CONFIG[client_key]
    tipo = client.get("tipo", "ecommerce")
    sheet_id = client.get("sheet_id")
    planilha = client.get("planilha", f"{client_key}.xlsx")
    xlsx_path = os.path.join(BASE_DIR, "clientes", planilha)

    try:
        if tipo == "ecommerce":
            dados = process_ecommerce(xlsx_path, periodo, periodo_comp, client["metas"], sheet_id)
        else:
            dados = process_lead(xlsx_path, periodo, periodo_comp, client["metas"], sheet_id)
    except Exception as e:
        return jsonify({"erro": f"Erro ao processar dados: {e}"}), 500

    dados["nome_cliente"] = client["nome"]
    dados["periodo_analisado"] = periodo
    dados["periodo_comparado"] = periodo_comp

    try:
        buf = fill_template_to_buffer(tipo, dados)
    except Exception as e:
        return jsonify({"erro": f"Erro ao gerar apresentação: {e}"}), 500

    nome = client["nome"].replace(" ", "_")
    filename = f"{nome}_{periodo.replace('/', '-').replace(' ', '')}.pptx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
