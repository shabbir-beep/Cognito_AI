"""
Cognito AI — Desktop Entrypoint & Multithreaded API Server
Project Code: SIH26117 | MRPL
Cross-platform compatible: Windows & Linux
"""

import sys
import os
import json
import time
import atexit
import signal
import threading
import http.server
import socketserver
import urllib.parse
import webbrowser

# --- ENCODING FIX ---
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

from sovereign_engine import (
    profile_hardware,
    execute_agent_task,
    get_available_models_catalog,
    save_and_open_desktop_document,
    get_ollama_status,
    ensure_ollama_installed,
    start_ollama_server,
    stop_ollama_server,
    setup_ollama_background,
    list_local_models,
    stream_pull_model,
    cancel_model_pull,
    delete_local_model,
    load_all_sessions,
    save_session_to_disk,
    delete_session_from_disk,
    load_server_config,
    save_server_config,
    test_server_connection,
    query_model_chat,
    export_response_document,
    create_job_object,
    cleanup_orphan_ollama,
    unload_model,
    unload_all_models,
    switch_model,
    classify_prompt_and_select_model,
    recommend_model_by_preferences,
    load_all_projects,
    save_project_to_disk,
    delete_project_from_disk,
    is_writable,
    write_error_msg,
    APP_DIR,
)

def get_base_dir():
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

class ThreadedTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    daemon_threads = True
    allow_reuse_address = True

class CognitoHTTPHandler(http.server.SimpleHTTPRequestHandler):
    """Handles frontend UI serving and REST/SSE endpoints."""

    def log_message(self, format, *args):
        pass

    # ==============================================================
    # GET ENDPOINTS
    # ==============================================================
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)

        # 1. UI Root
        if parsed.path in ("/", "/index.html"):
            self._serve_ui()
            return

        # 1b. Static Icons
        if parsed.path in ("/icon.png", "/app_icon.png", "/favicon.ico", "/icon.ico"):
            icon_file = "icon.png" if "png" in parsed.path else "icon.ico"
            icon_path = os.path.join(get_base_dir(), icon_file)
            if not os.path.exists(icon_path):
                icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), icon_file)
            if os.path.exists(icon_path):
                mime = "image/png" if icon_file.endswith(".png") else "image/x-icon"
                self.send_response(200)
                self.send_header("Content-Type", mime)
                self.send_header("Cache-Control", "public, max-age=86400")
                self.end_headers()
                with open(icon_path, "rb") as f:
                    self.wfile.write(f.read())
                return

        # 2. App Status & Read-Only Check
        if parsed.path == "/api/app_status":
            self._json_response({
                "writable": is_writable,
                "error": write_error_msg,
                "app_dir": APP_DIR
            })
            return

        # 3. Hardware Profiling
        if parsed.path == "/api/hardware":
            self._json_response(profile_hardware())
            return

        # 3. Dynamic Models Catalog with Hardware Ratings
        if parsed.path == "/api/available_models":
            self._json_response({"catalog": get_available_models_catalog()})
            return

        # 4. Ollama Installation & Service Status
        if parsed.path == "/api/ollama_status":
            self._json_response(get_ollama_status())
            return

        # 5. Installed Local Models
        if parsed.path == "/api/local_models":
            self._json_response({"models": list_local_models()})
            return

        # 6. Persistent Chat Sessions
        if parsed.path == "/api/sessions":
            self._json_response({"sessions": load_all_sessions()})
            return

        # 6b. Project Workspaces (ChatGPT / Gemini Style)
        if parsed.path == "/api/projects":
            self._json_response({"projects": load_all_projects()})
            return

        # 7. Active Server Configuration
        if parsed.path == "/api/server_config":
            self._json_response(load_server_config())
            return

        # 8. Save & Open Document on Desktop
        if parsed.path == "/api/save_desktop":
            query = urllib.parse.parse_qs(parsed.query)
            fname = query.get("file", ["Hydrocracker_Inspection_Approval_Note.docx"])[0]
            desktop_path = save_and_open_desktop_document(fname)
            self._json_response({"status": "Success", "path": desktop_path})
            return

        # 9. Deliverables / Exports Download
        if parsed.path == "/api/download":
            query = urllib.parse.parse_qs(parsed.query)
            filepath = query.get("file", [""])[0]
            if not os.path.isabs(filepath):
                exp_path = os.path.join(APP_DIR, 'data', 'exports', filepath)
                deliv_path = os.path.join(os.getcwd(), "deliverables", filepath)
                filepath = exp_path if os.path.exists(exp_path) else deliv_path
            if os.path.exists(filepath):
                ext = os.path.splitext(filepath)[1].lower()
                mime = "application/pdf" if ext == ".pdf" else ("application/vnd.openxmlformats-officedocument.wordprocessingml.document" if ext == ".docx" else "text/plain")
                self.send_response(200)
                self.send_header("Content-Type", mime)
                self.send_header("Content-Disposition", f'attachment; filename="{os.path.basename(filepath)}"')
                self.end_headers()
                with open(filepath, "rb") as f:
                    self.wfile.write(f.read())
            else:
                self.send_response(404)
                self.end_headers()
            return

        # 10. Streaming Pull Model (GET request support)
        if parsed.path == "/api/pull_model":
            query = urllib.parse.parse_qs(parsed.query)
            model_name = query.get("model", [""])[0]
            self.send_response(200)
            self.send_header("Content-Type", "application/x-ndjson")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("X-Accel-Buffering", "no")
            self.end_headers()

            try:
                for line_bytes in stream_pull_model(model_name):
                    self.wfile.write(line_bytes)
                    self.wfile.flush()
            except Exception as e:
                err = json.dumps({"status": "error", "error": str(e)[:200]}).encode('utf-8')
                try:
                    self.wfile.write(err + b'\n')
                    self.wfile.flush()
                except Exception:
                    pass
            return

        super().do_GET()

    # ==============================================================
    # POST & DELETE ENDPOINTS
    # ==============================================================
    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        content_length = int(self.headers.get('Content-Length', 0))
        post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
        try:
            req = json.loads(post_data.decode('utf-8'))
        except Exception:
            req = {}

        # 1. Multi-Turn Conversational Chat Execution (Full Context)
        if parsed.path in ("/api/chat", "/api/execute"):
            messages = req.get("messages", [])
            if not messages and "prompt" in req:
                messages = [{"role": "user", "content": req["prompt"]}]

            model_name = req.get("model", "phi3.5:3.8b")
            file_name = req.get("file_name", None)
            server_cfg = req.get("server_config", None)

            res = execute_agent_task(
                messages=messages,
                model_name=model_name,
                file_path=file_name,
                server_config=server_cfg
            )
            self._json_response(res)
            return

        # 2. Save Session to Persistent Storage
        if parsed.path == "/api/sessions":
            saved = save_session_to_disk(req)
            self._json_response({"status": "saved", "session": saved})
            return

        # 3. Save Server Configuration
        if parsed.path == "/api/server_config":
            save_server_config(req)
            self._json_response({"status": "saved"})
            return

        # 4. Test Server Connectivity (Local, LAN, Remote)
        if parsed.path == "/api/test_server":
            res = test_server_connection(req)
            self._json_response(res)
            return

        # 5. Pull Model (POST Streaming NDJSON with Cancel Support)
        if parsed.path == "/api/pull_model":
            model_name = req.get("model", "")
            self.send_response(200)
            self.send_header("Content-Type", "application/x-ndjson")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("X-Accel-Buffering", "no")
            self.end_headers()

            try:
                for line_bytes in stream_pull_model(model_name):
                    self.wfile.write(line_bytes)
                    self.wfile.flush()
            except Exception as e:
                err = json.dumps({"status": "error", "error": str(e)[:200]}).encode('utf-8')
                try:
                    self.wfile.write(err + b'\n')
                    self.wfile.flush()
                except Exception:
                    pass
            return

        # 6. Cancel Model Pull
        if parsed.path == "/api/cancel_pull":
            model_name = req.get("model", "")
            ok = cancel_model_pull(model_name)
            self._json_response({"status": "cancelled" if ok else "not_found"})
            return

        # 7. Unload Model from VRAM
        if parsed.path == "/api/unload_model":
            model_name = req.get("model", "")
            ok = unload_model(model_name)
            self._json_response({"status": "unloaded" if ok else "error"})
            return

        # 8. Switch Model (Unloads previous model from memory)
        if parsed.path == "/api/switch_model":
            model_name = req.get("model", "")
            switch_model(model_name)
            self._json_response({"status": "switched", "model": model_name})
            return

        # 9. Delete Local Model
        if parsed.path == "/api/delete_model":
            model_name = req.get("model", "")
            ok = delete_local_model(model_name)
            self._json_response({"status": "deleted" if ok else "error"})
            return

        # 10. Export AI Response Document (PDF, DOCX, TXT)
        if parsed.path in ("/api/export", "/api/export_doc"):
            text = req.get("text", "")
            fmt = req.get("format", "pdf")
            title = req.get("title", "Cognito_AI_Report")
            filePath = export_response_document(text, format_type=fmt, title=title)

            if os.path.exists(filePath):
                filename = os.path.basename(filePath)
                self._json_response({
                    "status": "success",
                    "filename": filename,
                    "download_url": f"/api/download?file={filename}"
                })
            else:
                self._json_response({"status": "error", "error": "Export generation failed"}, status=500)
            return

        # 11. Classify Prompt & Select Best Model (Smart Model Router)
        if parsed.path == "/api/route_model":
            prompt = req.get("prompt", "")
            fname = req.get("file_name", "")
            installed = req.get("installed_models", None)
            if installed is None:
                local_m = list_local_models()
                installed = [m.get("name", "") for m in local_m]
            route_res = classify_prompt_and_select_model(prompt, fname, installed)
            self._json_response({
                "status": "success",
                "selected_model": route_res.get("model"),
                "intent": route_res.get("intent"),
                "category": route_res.get("category"),
                "label": route_res.get("label"),
                "reason": route_res.get("reason")
            })
            return

        # 11b. Dynamic Use-Case & Language Model Recommender
        if parsed.path == "/api/recommend_model":
            use_cases = req.get("use_cases", ["general"])
            lang = req.get("language", "English")
            rec_result = recommend_model_by_preferences(use_cases=use_cases, language=lang)
            self._json_response({
                "status": "success",
                "recommendation": rec_result
            })
            return

        # 12. Save Project Workspace
        if parsed.path == "/api/projects":
            saved = save_project_to_disk(req)
            self._json_response({"status": "saved", "project": saved})
            return

        self.send_response(404)
        self.end_headers()

    def do_DELETE(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/sessions":
            query = urllib.parse.parse_qs(parsed.query)
            sess_id = query.get("id", [""])[0]
            ok = delete_session_from_disk(sess_id)
            self._json_response({"status": "deleted" if ok else "error"})
            return

        if parsed.path == "/api/projects":
            query = urllib.parse.parse_qs(parsed.query)
            proj_id = query.get("id", [""])[0]
            ok = delete_project_from_disk(proj_id)
            self._json_response({"status": "deleted" if ok else "error"})
            return

        self.send_response(404)
        self.end_headers()

    # ==============================================================
    # HELPERS
    # ==============================================================
    def _serve_ui(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        ui_path = os.path.join(get_base_dir(), "app_ui.html")
        if os.path.exists(ui_path):
            with open(ui_path, "rb") as f:
                self.wfile.write(f.read())
        else:
            self.wfile.write(b"<h1>UI File app_ui.html Not Found</h1>")

    def _json_response(self, data, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))

# ==============================================================
# SERVER & APP LAUNCHER
# ==============================================================
def start_backend_server(port=8085):
    handler = CognitoHTTPHandler
    with ThreadedTCPServer(("127.0.0.1", port), handler) as httpd:
        httpd.serve_forever()

def launch_desktop():
    # 0. Initialize Windows Job Object for auto-cleanup on exit / crash
    create_job_object()

    # 0b. Clean up leftover/orphaned processes from prior crashed session
    cleanup_orphan_ollama()

    port = 8085

    # 1. Start HTTP Server in background thread
    server_thread = threading.Thread(target=start_backend_server, args=(port,), daemon=True)
    server_thread.start()
    time.sleep(0.4)

    # 2. Trigger Ollama background discovery / service start
    ollama_thread = threading.Thread(target=setup_ollama_background, daemon=True)
    ollama_thread.start()

    # 3. Clean exit teardown on normal termination and OS signals
    atexit.register(stop_ollama_server)

    def _signal_handler(sig, frame):
        stop_ollama_server()
        sys.exit(0)

    try:
        signal.signal(signal.SIGINT, _signal_handler)
        if hasattr(signal, 'SIGTERM'):
            signal.signal(signal.SIGTERM, _signal_handler)
        if hasattr(signal, 'SIGBREAK'):
            signal.signal(signal.SIGBREAK, _signal_handler)
    except Exception:
        pass

    app_url = f"http://127.0.0.1:{port}"
    print("=" * 65)
    print("  >> COGNITO AI WORKBENCH v2.0 ACTIVE")
    print(f"  URL: {app_url}")
    print("  Status: Air-Gapped Industrial AI Workstation")
    print("=" * 65)

    # 4. Open in PyWebView or default browser
    try:
        import webview
        window = webview.create_window(
            "Cognito AI",
            url=app_url,
            width=1480,
            height=920,
            resizable=True,
        )
        webview.start()
    except Exception:
        print(f"Opening browser at {app_url}...")
        webbrowser.open(app_url)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("Stopping Cognito AI.")

    # 5. Final exit teardown
    stop_ollama_server()

if __name__ == '__main__':
    launch_desktop()
