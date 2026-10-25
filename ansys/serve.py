"""Local server for the page. Reads and writes files in <folder>; starts setup_study.py on request.

    .\ansys\run.ps1 serve.py [folder] [port]      (default folder: ansys\example)
    ansys\start.bat [folder]                       (the same, and starts AEDT first)

Open http://localhost:8765. Nothing leaves this PC. The page can switch to <folder> or any of its sub-folders that
holds a .edb or .aedb export.
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
ROOT = FOLDER
PORT = int(sys.argv[2]) if len(sys.argv) > 2 else 8765
WHAT = {"build": "Building the project", "update": "Applying changes", "mesh": "Generating the mesh",
        "read": "Reading AEDT", "solve": "Solving"}
job = {"running": False, "what": "", "log": "", "silent": False}
lock = threading.Lock()


def jread(name):
    p = FOLDER / name
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else None


def boards():
    """board folders the page can switch to: ROOT and its sub-folders that hold an export or a board.json"""
    has = lambda d: (d / "board.json").exists() or any(x.is_dir() for p in ("*.edb", "*.aedb") for x in d.glob(p))  # noqa: E731
    return [str(d) for d in [ROOT, *sorted(x for x in ROOT.iterdir() if x.is_dir())] if has(d)]


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
        timer, killed = None, threading.Event()
        try:
            proc = subprocess.Popen([sys.executable, "-u", str(HERE / script), str(FOLDER), *args], stdout=subprocess.PIPE,
                                    stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
            limit = 60 if silent else 3600 if args[:1] == ("solve",) else 900
            timer = threading.Timer(limit, lambda: (killed.set(), proc.kill()))    # a hung gRPC call must not block the page
            timer.start()
            for line in proc.stdout:
                if not silent and "INFO" not in line:
                    job["log"] += line
            proc.wait()
            if silent:
                pass
            elif killed.is_set():    # on Windows a killed process exits with 1, like a normal "not ready"
                job["log"] += "[error] The script was stopped after %d min, the time limit. AEDT may be busy or frozen.\n" % (limit // 60)
            elif proc.returncode not in (0, 1):
                job["log"] += "[error] The script was stopped (exit code %s). AEDT may be busy or frozen.\n" % proc.returncode
        except Exception as ex:  # noqa: BLE001  the job must always end, or every later request gets 409
            if not silent:
                job["log"] += "[error] The script could not run: %s: %s\n" % (type(ex).__name__, ex)
        finally:
            if timer:
                timer.cancel()
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
            self._send(200, {"folder": str(FOLDER), "boards": boards(), "board": jread("board.json"), "study": jread("study.json"),
                             "checks": jread("checks.json"), "results": jread("results.json"), "state": jread("state.json"), "job": job})
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
        global FOLDER
        path, _, query = self.path.partition("?")
        q = parse_qs(query)
        mode = (q.get("mode") or ["build"])[0]
        if path == "/api/folder":
            want = (q.get("path") or [""])[0]
            if job["running"] or want not in boards():     # a background read still writes to the old folder
                return self._send(409, {"ok": False})
            FOLDER = Path(want)
            return self._send(200, {"ok": True})
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
    try:
        server = ThreadingHTTPServer(("127.0.0.1", PORT), H)
    except OSError:     # the page is running already (or the port is taken): show it
        print(f"Port {PORT} is in use, the page is probably running already. Opening it.", flush=True)
        webbrowser.open(f"http://localhost:{PORT}")
        sys.exit(0)
    print(f"Page: http://localhost:{PORT}   folder: {FOLDER}", flush=True)
    threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    server.serve_forever()
