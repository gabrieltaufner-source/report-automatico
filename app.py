import io
import json
import os

from flask import Flask, render_template, request, send_file, jsonify

from data_processor import process_ecommerce, process_lead
from pptx_filler import fill_template_to_buffer

app = Flask(__name__)

BASE_DIR = os.path.dirname(__file__)
CONFIG_PATH = os.path.join(BASE_DIR, "clients_config.json")

with open(CONFIG_PATH, encoding="utf-8") as f:
    CONFIG = json.load(f)


def _gerar_um(client_key: str, periodo: str, periodo_comp: str):
    """Gera o relatório de um cliente, salva no Drive e retorna (filename, BytesIO)."""
    from google_sheets import upload_to_drive

    client = CONFIG[client_key]
    tipo = client.get("tipo", "ecommerce")
    sheet_id = client.get("sheet_id")
    planilha = client.get("planilha", f"{client_key}.xlsx")
    xlsx_path = os.path.join(BASE_DIR, "clientes", planilha)

    if tipo == "ecommerce":
        dados = process_ecommerce(xlsx_path, periodo, periodo_comp, client["metas"], sheet_id)
    else:
        dados = process_lead(xlsx_path, periodo, periodo_comp, client["metas"], sheet_id)

    dados["nome_cliente"] = client["nome"]
    dados["periodo_analisado"] = periodo
    dados["periodo_comparado"] = periodo_comp

    buf = fill_template_to_buffer(tipo, dados)
    nome = client["nome"].replace(" ", "_")
    filename = f"{nome}_{periodo.replace('/', '-').replace(' ', '')}.pptx"

    # Upload para o Google Drive (não bloqueia o download em caso de falha)
    try:
        upload_to_drive(io.BytesIO(buf.getvalue()), filename)
        print(f"[Drive] ✅ Salvo: {filename}")
    except Exception as e:
        import traceback
        print(f"[Drive] ❌ Falha ao salvar {filename}: {e}")
        traceback.print_exc()

    buf.seek(0)
    return filename, buf


@app.route("/")
def index():
    clients = [
        {"key": k, "nome": v["nome"], "tipo": v.get("tipo", "")}
        for k, v in CONFIG.items()
    ]
    return render_template("index.html", clients=clients)


@app.route("/gerar", methods=["POST"])
def gerar():
    """Gera sempre um único relatório por requisição (o ZIP é montado no browser)."""
    client_key = request.form.get("cliente")
    periodo = request.form.get("periodo", "").strip()
    periodo_comp = request.form.get("periodo_comp", "").strip()

    if not client_key or not periodo or not periodo_comp:
        return jsonify({"erro": "Preencha todos os campos."}), 400

    if client_key not in CONFIG:
        return jsonify({"erro": "Cliente não encontrado."}), 400

    try:
        filename, buf = _gerar_um(client_key, periodo, periodo_comp)
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


@app.route("/test-drive")
def test_drive():
    import traceback
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
        from google_sheets import _get_credentials, DRIVE_OUTPUT_FOLDER

        creds = _get_credentials()
        service = build("drive", "v3", credentials=creds)

        # Lista os Shared Drives acessíveis pela service account (diagnóstico)
        drives = service.drives().list(fields="drives(id,name)").execute()

        # Tenta upload direto para a pasta configurada
        buf = io.BytesIO(b"teste de upload")
        buf.seek(0)
        media = MediaIoBaseUpload(buf, mimetype="text/plain", resumable=False)
        result = service.files().create(
            body={"name": "_teste_conexao.txt", "parents": [DRIVE_OUTPUT_FOLDER]},
            media_body=media,
            fields="id,parents",
            supportsAllDrives=True,
        ).execute()

        return jsonify({
            "status": "ok",
            "mensagem": "Upload funcionando!",
            "file_id": result.get("id"),
            "drives_visiveis": drives.get("drives", []),
            "folder_id_usado": DRIVE_OUTPUT_FOLDER,
        })
    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e), "detalhe": traceback.format_exc()}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
