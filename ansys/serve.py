"""Local server for the page. Reads and writes files in <folder>; starts setup_study.py on request.

    .\ansys\run.ps1 serve.py [folder] [port]      (default folder: the last one opened, else ansys\example)
    ansys\start.bat [folder]                       (the same, and starts AEDT first)

Open http://localhost:8765. Nothing leaves this PC. The page can open any folder on the PC (Windows folder dialog);
it lists the folders opened before and the sub-folders of the start folder that hold a .edb or .aedb export.
"""
import json
import os
import subprocess
import webbrowser
import sys
import threading
import time
from urllib.parse import parse_qs
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).parent
RECENT = Path(os.environ.get("APPDATA") or HERE) / "pcb-thermal" / "recent.json"     # folders opened before, newest first


def recent():
    try:
        return [d for d in json.loads(RECENT.read_text(encoding="utf-8")) if Path(d).is_dir()]
    except (OSError, ValueError):
        return []


def remember(folder):
    try:
        RECENT.parent.mkdir(parents=True, exist_ok=True)
        RECENT.write_text(json.dumps([str(folder), *[d for d in recent() if d != str(folder)]][:15], indent=1), encoding="utf-8")
    except OSError:
        pass


FOLDER = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 and sys.argv[1] else Path((recent() or [HERE / "example"])[0])
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
    """board folders the page lists: the open one, those opened before, ROOT's sub-folders with an export or a board.json"""
    has = lambda d: (d / "board.json").exists() or any(x.is_dir() for p in ("*.edb", "*.aedb") for x in d.glob(p))  # noqa: E731
    subs = [str(d) for d in [ROOT, *sorted(x for x in ROOT.iterdir() if x.is_dir())] if has(d)]
    return list(dict.fromkeys([str(FOLDER), *recent(), *subs]))


def pick_folder():
    """the Windows folder dialog, in its own process (Tk wants the main thread). Empty when cancelled"""
    code = ("import sys, tkinter as tk, tkinter.filedialog as fd; r = tk.Tk(); r.withdraw(); r.attributes('-topmost', True); "
            "print(fd.askdirectory(parent=r, initialdir=sys.argv[1], title='Board folder (with the .edb or .aedb export)'))")
    out = subprocess.run([sys.executable, "-c", code, str(FOLDER.parent)], capture_output=True, text=True, encoding="utf-8")
    return out.stdout.strip()


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

    def _foreign(self):
        """a web site open in the same browser must not drive the page: allow only requests from the page itself"""
        origin = self.headers.get("Origin")
        if origin and origin not in (f"http://localhost:{PORT}", f"http://127.0.0.1:{PORT}"):
            self._send(403, {"error": "forbidden"})
            return True
        return False

    def do_PUT(self):
        if self._foreign():
            return
        if self.path == "/api/study":
            n = int(self.headers.get("Content-Length", 0))
            study = json.loads(self.rfile.read(n))
            (FOLDER / "study.json").write_text(json.dumps(study, indent=1), encoding="utf-8")
            self._send(200, {"ok": True})
        else:
            self._send(404, {"error": "not found"})

    def do_POST(self):
        global FOLDER
        if self._foreign():
            return
        path, _, query = self.path.partition("?")
        q = parse_qs(query)
        mode = (q.get("mode") or ["build"])[0]
        if path in ("/api/folder", "/api/browse"):
            # A real job must finish in its folder first; the page's background read of AEDT may go on, its script
            # was given the old folder and writes only there
            busy = lambda: job["running"] and not job["silent"]  # noqa: E731
            if busy():
                return self._send(409, {"ok": False, "error": "A job is running. Switch when it has finished."})
            want = (q.get("path") or [""])[0] if path == "/api/folder" else pick_folder()
            if not want:
                return self._send(200, {"ok": False, "cancelled": True})
            want = Path(want).resolve()
            if not want.is_dir():
                return self._send(404, {"ok": False, "error": f"{want} is not a folder."})
            if busy():      # started while the dialog was open
                return self._send(409, {"ok": False, "error": "A job is running. Switch when it has finished."})
            FOLDER = want
            remember(FOLDER)
            return self._send(200, {"ok": True, "folder": str(FOLDER)})
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
    remember(FOLDER)
    print(f"Page: http://localhost:{PORT}   folder: {FOLDER}", flush=True)
    threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    server.serve_forever()
