"""Local server for the page. Reads and writes files in <folder>; starts setup_study.py on request.

    .\ansys\run.ps1 serve.py [folder] [port]      (default folder: ansys\example)

Open http://localhost:8765. Nothing leaves this PC.
"""
import json
import subprocess
import webbrowser
import sys
import threading
import time
from urllib.parse import parse_qs
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).parent
FOLDER = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else HERE / "example"
PORT = int(sys.argv[2]) if len(sys.argv) > 2 else 8765
WHAT = {"build": "Building the project", "update": "Applying changes", "mesh": "Generating the mesh",
        "read": "Reading AEDT"}
job = {"running": False, "what": "", "log": "", "silent": False}
lock = threading.Lock()


def jread(name):
    p = FOLDER / name
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else None


def run_script(script, what, *args, silent=False):
    """One job at a time. A silent job (the page's background read of AEDT) never blocks a real one for long."""
    deadline = time.time() + 90
    while True:
        with lock:
            if not job["running"]:
                job.update(running=True, what=what, silent=silent)
                if not silent:
                    job["log"] = ""
                break
            if silent or not job["silent"] or time.time() > deadline:
                return False
        time.sleep(0.5)

    def work():
        proc = subprocess.Popen([sys.executable, "-u", str(HERE / script), str(FOLDER), *args], stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
        timer = threading.Timer(60 if silent else 900, proc.kill)    # a hung gRPC call must not block the page
        timer.start()
        for line in proc.stdout:
            if not silent and "INFO" not in line:
                job["log"] += line
        proc.wait()
        timer.cancel()
        if not silent and proc.returncode not in (0, 1):
            job["log"] += "[error] The script was stopped (exit code %s). AEDT may be busy or frozen.\n" % proc.returncode
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
                             "checks": jread("checks.json"), "state": jread("state.json"), "job": job})
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
        path, _, query = self.path.partition("?")
        q = parse_qs(query)
        mode = (q.get("mode") or ["build"])[0]
        if path == "/api/read":
            ok = run_script("read_board.py", "Reading the board")
        elif path == "/api/run" and mode in WHAT:
            ok = run_script("setup_study.py", WHAT[mode], mode, silent="silent" in q)
        else:
            return self._send(404, {"error": "not found"})
        self._send(200 if ok else 409, {"ok": ok})

    def log_message(self, *a):
        pass


if __name__ == "__main__":
    if not FOLDER or not FOLDER.is_dir():
        raise SystemExit(__doc__)
    print(f"Page: http://localhost:{PORT}   folder: {FOLDER}", flush=True)
    threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    ThreadingHTTPServer(("127.0.0.1", PORT), H).serve_forever()
