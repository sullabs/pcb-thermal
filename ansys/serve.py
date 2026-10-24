"""Local server for the page. Reads and writes files in <folder>; starts setup_study.py on request.

    python ansys/serve.py <folder> [port]       (use ansys\run.ps1 serve.py <folder>)

Open http://localhost:8765. Nothing leaves this PC.
"""
import json
import subprocess
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).parent
FOLDER = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else None
PORT = int(sys.argv[2]) if len(sys.argv) > 2 else 8765
job = {"running": False, "what": "", "log": ""}
lock = threading.Lock()


def jread(name):
    p = FOLDER / name
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else None


def run_script(script, what):
    with lock:
        if job["running"]:
            return False
        job.update(running=True, what=what, log="")

    def work():
        proc = subprocess.Popen([sys.executable, "-u", str(HERE / script), str(FOLDER)], stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
        for line in proc.stdout:
            if "INFO" not in line:
                job["log"] += line
        proc.wait()
        job["running"] = False

    threading.Thread(target=work, daemon=True).start()
    return True


class H(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json"):
        data = body if isinstance(body, bytes) else json.dumps(body).encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path == "/":
            self._send(200, (HERE / "page.html").read_bytes(), "text/html; charset=utf-8")
        elif self.path == "/api/state":
            self._send(200, {"folder": str(FOLDER), "board": jread("board.json"), "study": jread("study.json"),
                             "checks": jread("checks.json"), "job": job})
        else:
            self._send(404, {"error": "not found"})

    def do_PUT(self):
        if self.path == "/api/study":
            n = int(self.headers.get("Content-Length", 0))
            study = json.loads(self.rfile.read(n))
            (FOLDER / "study.json").write_text(json.dumps(study, indent=1), encoding="utf-8")
            self._send(200, {"ok": True})
        else:
            self._send(404, {"error": "not found"})

    def do_POST(self):
        script = {"/api/read": "read_board.py", "/api/run": "setup_study.py"}.get(self.path)
        if not script:
            return self._send(404, {"error": "not found"})
        ok = run_script(script, "Reading the board" if script == "read_board.py" else "Setting up")
        self._send(200 if ok else 409, {"ok": ok})

    def log_message(self, *a):
        pass


if __name__ == "__main__":
    if not FOLDER or not FOLDER.is_dir():
        raise SystemExit(__doc__)
    print(f"Page: http://localhost:{PORT}   folder: {FOLDER}", flush=True)
    ThreadingHTTPServer(("127.0.0.1", PORT), H).serve_forever()
