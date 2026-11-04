"""Local server for the page. Reads and writes files in <folder>; starts setup_study.py on request.

    .\ansys\run.ps1 serve.py [folder] [port]      (default folder: the last one opened, else ansys\example)
    ansys\start.bat [folder]                       (the same, and starts AEDT first)

Open http://localhost:8765. Nothing leaves this PC. The page can open any folder on the PC (Windows folder dialog);
it lists the folders opened before and the sub-folders of the start folder that hold a .edb or .aedb export.
"""
import json
import os
import re
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
        "read": "Reading AEDT", "solve": "Solving", "meshcheck": "Checking the mesh (two solves)", "open": "Opening the project"}
LIMIT = {"solve": 12 * 3600, "meshcheck": 24 * 3600, "mesh": 4 * 3600}     # s; any other job 30 min, the silent read 60 s
job = {"running": False, "what": "", "log": "", "silent": False}
lock = threading.Lock()


def jread(name):
    p = FOLDER / name
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else None


RUN_ID = re.compile(r"^\d{8}-\d{6}$")     # runs/<id>.json, as setup_study.py names them
_runs = {}      # path -> (mtime, parsed run): the page asks every 5 s, a run file changes only with its note


def runs():
    """every kept solve, newest first, with its file name as id"""
    out = []
    for p in sorted((FOLDER / "runs").glob("*.json"), reverse=True):
        if RUN_ID.match(p.stem):
            try:
                t = p.stat().st_mtime
                if _runs.get(p, (None,))[0] != t:
                    _runs[p] = (t, {"id": p.stem, **json.loads(p.read_text(encoding="utf-8"))})
                out.append(_runs[p][1])
            except (OSError, ValueError):
                continue
    return out


def run_path(q):
    rid = (q.get("id") or [""])[0]
    return FOLDER / "runs" / f"{rid}.json" if RUN_ID.match(rid) else None


def project_file():
    """name of the built project in the folder, None if there is none: Build replaces it"""
    b = jread("board.json")
    name = b and f"{b['name']}_icepak.aedt"
    return name if name and (FOLDER / name).exists() else None


def boards():
    """board folders the page lists: the open one, those opened before, ROOT's sub-folders with an export or a board.json"""
    has = lambda d: (d / "board.json").exists() or any(x.is_dir() for p in ("*.edb", "*.aedb") for x in d.glob(p))  # noqa: E731
    subs = [str(d) for d in [ROOT, *sorted(x for x in ROOT.iterdir() if x.is_dir())] if has(d)]
    return list(dict.fromkeys([str(FOLDER), *recent(), *subs]))


def pick_folder():
    """the Windows folder dialog, in its own process (Tk wants the main thread). Empty when cancelled"""
    code = ("import sys, tkinter as tk, tkinter.filedialog as fd; r = tk.Tk(); r.withdraw(); r.attributes('-topmost', True); "
            "print(fd.askdirectory(parent=r, initialdir=sys.argv[1], title='Board folder (with the .edb or .aedb export)'))")
    # -X utf8: else the child prints to the pipe in the Windows code page, and a folder named Café or in Arabic breaks
    out = subprocess.run([sys.executable, "-X", "utf8", "-c", code, str(FOLDER.parent)], capture_output=True, text=True, encoding="utf-8")
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
            # a hung gRPC call must not block the page, but a solve or mesh of a big board may take hours
            limit = 60 if silent else LIMIT.get(args[0] if args else "", 1800)
            timer = threading.Timer(limit, lambda: (killed.set(), proc.kill()))
            timer.start()
            for line in proc.stdout:
                if not silent and "INFO" not in line:
                    job["log"] += line
            proc.wait()
            if silent:
                pass
            elif killed.is_set():    # on Windows a killed process exits with 1, like a normal "not ready"
                job["log"] += ("[error] The script was stopped after %d min, the time limit. AEDT may be busy or frozen; "
                               "a solve or mesh may still run on in AEDT.\n" % (limit // 60))
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
        path, _, query = self.path.partition("?")
        if path == "/":
            self._send(200, (HERE / "page.html").read_bytes(), "text/html; charset=utf-8")
        elif path == "/api/state":
            st = {"folder": str(FOLDER), "board": jread("board.json"), "study": jread("study.json"),
                  "checks": jread("checks.json"), "results": jread("results.json"), "meshcheck": jread("meshcheck.json"), "runs": runs(),
                  "built": jread("built.json"), "meshed": jread("meshed.json"), "applied": jread("applied.json"),
                  "projectFile": project_file(), "state": jread("state.json"), "job": job}
            if "boards" in parse_qs(query, keep_blank_values=True):      # the folder list: the page asks once
                st["boards"] = boards()
            self._send(200, st)
        else:
            self._send(404, {"error": "not found"})

    def _foreign(self):
        """a web site open in the same browser must not drive the page: allow only requests from the page itself"""
        origin = self.headers.get("Origin")
        if origin and origin not in (f"http://localhost:{PORT}", f"http://127.0.0.1:{PORT}"):
            self._send(403, {"error": "forbidden"})
            return True
        return False

    def _moved(self, q):
        """the page names the folder it shows: a write for another folder (another tab switched the server) is refused"""
        f = (q.get("folder") or [None])[0]
        if f is not None and f != str(FOLDER):
            self._send(409, {"ok": False, "moved": True, "error": "The server shows another folder now."})
            return True
        return False

    def do_PUT(self):
        if self._foreign():
            return
        path, _, query = self.path.partition("?")
        if self._moved(parse_qs(query)):
            return
        n = int(self.headers.get("Content-Length", 0))
        if path == "/api/study":
            study = json.loads(self.rfile.read(n))
            (FOLDER / "study.json").write_text(json.dumps(study, indent=1), encoding="utf-8")
            self._send(200, {"ok": True})
        elif path == "/api/runnote":       # the note typed next to a kept run
            p = run_path(parse_qs(query))
            if not p or not p.exists():
                return self._send(404, {"ok": False, "error": "No such run."})
            r = json.loads(p.read_text(encoding="utf-8"))
            r["note"] = str(json.loads(self.rfile.read(n)).get("note", ""))[:300]
            p.write_text(json.dumps(r, indent=1), encoding="utf-8")
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
        if self._moved(q):
            return
        if path == "/api/read":
            ok = run_script("read_board.py", "Reading the heights from the STEP", "--step") if "step" in q else                 run_script("read_board.py", "Reading the board")
        elif path == "/api/runremove":    # a kept run the user no longer wants in the table
            p = run_path(q)
            if p and p.exists():
                p.unlink()
            return self._send(200, {"ok": bool(p)})
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
