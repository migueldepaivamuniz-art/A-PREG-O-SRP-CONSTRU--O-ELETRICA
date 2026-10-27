"""
================================================================================
SISTEMA DE PESQUISA DE PREÇOS DE MERCADO — PREGÃO SRP (LEI 14.133/2021)
SERVIDOR WEB LOCAL (FLASK) & REST API
================================================================================
Fornece interface web completa no navegador com o mesmo layout, recursos e regras
de negócio da aplicação desktop, permitindo execução local simples e compartilhada.
================================================================================
"""

import os
import sys
import json
import time
import glob
import threading
import webbrowser
from flask import Flask, render_template, jsonify, request, send_from_directory, send_file
from werkzeug.utils import secure_filename
from playwright.sync_api import sync_playwright

# Importa motor de negócio e persistência do app.py
from app import (
    load_sheet_items,
    cotar_item_core,
    save_and_export_results,
    create_export_zip_file,
    generate_template_excel,
    reset_all_data,
    read_sheet_file,
    parse_items_dataframe,
    save_active_sheet,
    get_active_sheet_info,
    reset_active_sheet_to_default,
    get_gemini_client,
    RESULTS_JSON,
    RESULTS_EXCEL,
    PRINTS_DIR,
    QUOTA_FILE
)

app = Flask(__name__, template_folder="templates")
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB max upload

# ESTADO GLOBAL DO ROBÔ DE PESQUISA
search_state = {
    "running": False,
    "progress": 0,
    "total": 0,
    "current_index": 0,
    "current_lote": None,
    "status_msg": "Pronto para iniciar pesquisa de mercado.",
    "stop_requested": False
}
search_lock = threading.Lock()

# ==============================================================================
# ROTAS DE PÁGINAS E ARQUIVOS ESTÁTICOS
# ==============================================================================
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/prints/<path:filename>")
def serve_print(filename):
    return send_from_directory(os.path.abspath(PRINTS_DIR), filename)

# ==============================================================================
# ROTAS DE DOWNLOAD DE ARQUIVOS GERADOS
# ==============================================================================
@app.route("/download/excel")
def download_excel():
    if os.path.exists(RESULTS_EXCEL):
        return send_file(os.path.abspath(RESULTS_EXCEL), as_attachment=True, download_name="cotacoes_pesquisa_mercado.xlsx")
    return jsonify({"error": "Planilha ainda não gerada. Execute cotações primeiro."}), 404

@app.route("/download/zip")
def download_zip():
    dest_zip = "Pacote_Pesquisa_Precos_Pregao.zip"
    create_export_zip_file(dest_zip)
    if os.path.exists(dest_zip):
        return send_file(os.path.abspath(dest_zip), as_attachment=True, download_name=dest_zip)
    return jsonify({"error": "Erro ao gerar arquivo compactado."}), 500

@app.route("/download/modelo")
def download_modelo():
    modelo_path = "Modelo_Itens_Cotacao.xlsx"
    generate_template_excel(modelo_path)
    return send_file(os.path.abspath(modelo_path), as_attachment=True, download_name=modelo_path)

# ==============================================================================
# API DE DADOS E KPIS
# ==============================================================================
@app.route("/api/data")
def api_get_data():
    # 1. Cotações existentes
    cotacoes = []
    if os.path.exists(RESULTS_JSON):
        try:
            with open(RESULTS_JSON, "r", encoding="utf-8") as f:
                cotacoes = json.load(f)
        except Exception:
            cotacoes = []

    # 2. Itens ativos da planilha
    itens_ativos = load_sheet_items()
    total_itens = len(itens_ativos)

    # 3. Comprovantes existentes
    prints = []
    if os.path.exists(PRINTS_DIR):
        prints = [os.path.basename(p) for p in glob.glob(os.path.join(PRINTS_DIR, "*.png"))]

    # 4. Cálculo de KPIs
    homologados_count = 0
    precos_validos = []
    for q in cotacoes:
        st = str(q.get("status", "")).lower()
        if "homologado" in st:
            homologados_count += 1
        vu = q.get("valor_unitario")
        if vu and isinstance(vu, (int, float)) and vu > 0:
            precos_validos.append(float(vu))

    pct = f"{(homologados_count / total_itens * 100):.1f}%" if total_itens > 0 else "0.0%"
    media_val = sum(precos_validos) / len(precos_validos) if precos_validos else 0.0
    media_str = f"R$ {media_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    # Cota Gemini
    rpd_used = 0
    if os.path.exists(QUOTA_FILE):
        try:
            with open(QUOTA_FILE, "r", encoding="utf-8") as f:
                qdata = json.load(f)
                rpd_used = qdata.get("rpd_used", 0)
        except Exception:
            pass

    return jsonify({
        "cotacoes": cotacoes,
        "itens_ativos": itens_ativos,
        "kpis": {
            "total": total_itens,
            "homologados": homologados_count,
            "homologados_pct": pct,
            "media_preco": media_str,
            "comprovantes": len(prints),
            "gemini_rpd": rpd_used
        },
        "active_sheet": get_active_sheet_info(),
        "prints": prints
    })

# ==============================================================================
# API DO MOTOR DE PESQUISA AUTOMÁTICA
# ==============================================================================
@app.route("/api/search/status")
def api_search_status():
    with search_lock:
        return jsonify(search_state)

@app.route("/api/search/start", methods=["POST"])
def api_search_start():
    with search_lock:
        if search_state["running"]:
            return jsonify({"success": False, "error": "Uma pesquisa já está em andamento."}), 400

    data = request.get_json() or {}
    mode = data.get("mode", "all")
    start = int(data.get("start", 1)) - 1
    end = int(data.get("end", 5))
    single_lote = int(data.get("lote", 1))

    items = load_sheet_items()
    if not items:
        return jsonify({"success": False, "error": "A planilha de itens está vazia."}), 400

    if mode == "all":
        target_items = items
    elif mode == "range":
        target_items = items[max(0, start):max(1, end)]
    elif mode == "single":
        target_items = [it for it in items if int(it.get("lote", -1)) == single_lote]
        if not target_items:
            return jsonify({"success": False, "error": f"Lote nº {single_lote} não encontrado na planilha."}), 404
    else:
        target_items = items

    # Dispara thread assíncrona
    th = threading.Thread(target=_run_search_thread, args=(target_items,), daemon=True)
    th.start()

    return jsonify({"success": True, "target_count": len(target_items)})

@app.route("/api/search/stop", methods=["POST"])
def api_search_stop():
    with search_lock:
        if search_state["running"]:
            search_state["stop_requested"] = True
            search_state["status_msg"] = "Interrupção solicitada pelo usuário..."
            return jsonify({"success": True, "message": "Parada solicitada."})
    return jsonify({"success": False, "message": "Nenhuma pesquisa ativa."})

def _run_search_thread(items_to_quote):
    with search_lock:
        search_state["running"] = True
        search_state["stop_requested"] = False
        search_state["total"] = len(items_to_quote)
        search_state["current_index"] = 0
        search_state["progress"] = 0
        search_state["status_msg"] = f"Iniciando pesquisa para {len(items_to_quote)} item(ns)..."

    try:
        gemini_client = get_gemini_client()
    except Exception as e_key:
        with search_lock:
            search_state["running"] = False
            search_state["status_msg"] = f"Erro de chave Gemini: {e_key}"
        return

    # Lê cotações preexistentes
    existing = []
    if os.path.exists(RESULTS_JSON):
        try:
            with open(RESULTS_JSON, "r", encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = []

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(
                headless=True,
                args=["--disable-blink-features=AutomationControlled"]
            )
        except Exception as e_launch:
            with search_lock:
                search_state["running"] = False
                search_state["status_msg"] = f"Erro ao iniciar Chromium: {e_launch}"
            return

        total = len(items_to_quote)
        for idx, item in enumerate(items_to_quote):
            with search_lock:
                if search_state["stop_requested"]:
                    search_state["status_msg"] = "Pesquisa cancelada pelo usuário."
                    break
                search_state["current_index"] = idx + 1
                search_state["current_lote"] = item.get("lote")
                search_state["progress"] = int((idx / total) * 100)
                search_state["status_msg"] = f"Cotando Lote {item.get('lote')} ({idx + 1}/{total}): {item.get('descricao_original')[:40]}..."

            def status_cb(info):
                with search_lock:
                    store = info.get("store", "")
                    step = info.get("step", "")
                    search_state["status_msg"] = f"Lote {info.get('current_lote')}: {store} — {step}"

            def check_stop():
                with search_lock:
                    return search_state["stop_requested"]

            res = cotar_item_core(browser, gemini_client, item, status_callback=status_cb, check_stop=check_stop)
            if res:
                existing = [e for e in existing if e.get("lote") != item.get("lote")]
                existing.append(res)
                save_and_export_results(existing)

            with search_lock:
                search_state["progress"] = int(((idx + 1) / total) * 100)

        browser.close()

    with search_lock:
        search_state["running"] = False
        search_state["progress"] = 100
        search_state["status_msg"] = "Pesquisa finalizada com sucesso."

# ==============================================================================
# API DE UPLOAD E GESTÃO DE PLANILHAS
# ==============================================================================
@app.route("/api/upload-sheet", methods=["POST"])
def api_upload_sheet():
    if 'sheet_file' not in request.files:
        return jsonify({"success": False, "error": "Nenhum arquivo enviado."}), 400

    file = request.files['sheet_file']
    if not file or file.filename == '':
        return jsonify({"success": False, "error": "Arquivo vazio."}), 400

    filename = secure_filename(file.filename)
    temp_path = os.path.join(os.path.abspath(PRINTS_DIR), f"temp_{filename}")
    os.makedirs(PRINTS_DIR, exist_ok=True)
    file.save(temp_path)

    try:
        df = read_sheet_file(temp_path)
        items, stats = parse_items_dataframe(df)
        if not items:
            return jsonify({"success": False, "error": "Nenhum item com descrição válida foi identificado na planilha."}), 400

        save_active_sheet(items, source_name=filename)
        return jsonify({
            "success": True,
            "total": len(items),
            "source": filename
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

@app.route("/api/restore-default-sheet", methods=["POST"])
def api_restore_default_sheet():
    reset_active_sheet_to_default()
    return jsonify({"success": True})

# ==============================================================================
# API DE RESET E BACKUP
# ==============================================================================
@app.route("/api/reset", methods=["POST"])
def api_reset():
    try:
        reset_all_data(backup=True)
        # Identifica a pasta de backup mais recente criada
        backups = sorted(glob.glob("backup_cotacoes_*"), reverse=True)
        backup_folder = backups[0] if backups else "Backup criado"
        return jsonify({"success": True, "backup_folder": backup_folder})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

# ==============================================================================
# INICIALIZAÇÃO DO SERVIDOR WEB
# ==============================================================================
def run_server(host="0.0.0.0", port=5000, open_browser=True):
    print("=" * 80)
    print("🏛️  SISTEMA DE PESQUISA DE PREÇOS — SERVIDOR WEB LOCAL (LEI 14.133/2021)")
    print("=" * 80)
    print(f"[*] Servidor Web Local ATIVO:")
    print(f"    👉 http://127.0.0.1:{port}/")
    print(f"    👉 http://localhost:{port}/")
    print("[*] Pressione Ctrl+C para encerrar o servidor a qualquer momento.")
    print("=" * 80)

    if open_browser:
        def _open():
            time.sleep(1.2)
            webbrowser.open(f"http://127.0.0.1:{port}/")
        threading.Thread(target=_open, daemon=True).start()

    try:
        app.run(host=host, port=port, debug=False)
    except Exception as e:
        print(f"Aviso no servidor web: {e}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Servidor Web Local do Sistema de Pesquisa de Preços")
    parser.add_argument("--host", default="0.0.0.0", help="Endereço de escuta (padrão: 0.0.0.0)")
    parser.add_argument("--port", type=int, default=5000, help="Porta do servidor (padrão: 5000)")
    parser.add_argument("--no-browser", action="store_true", help="Não abrir o navegador automaticamente")
    args = parser.parse_args()

    run_server(host=args.host, port=args.port, open_browser=not args.no_browser)
