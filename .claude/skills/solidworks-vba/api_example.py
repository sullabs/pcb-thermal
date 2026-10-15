# Looks up the official VBA examples for a SolidWorks API member from the online 2023 help and prints the code.
# The site blocks plain scripts and builds pages with JavaScript, so this drives headless Chromium (Playwright).
# Results are cached in example_cache/, so a repeat lookup works offline.
#
#   python api_example.py IFace2.GetTessTriangles              print every VBA example linked from that member page
#   python api_example.py IFeatureManager.FeatureExtrusion2 --list     only the example titles
#   python api_example.py IBody2.GetTessellation --title Non-manifold   only examples whose title contains the text
#   python api_example.py ICWStudyManager.CreateNewStudy3 --ns cosworks   Simulation API (older page layout)
#
# Needs: pip install playwright, and any Chromium: Playwright's own, or the one under %LOCALAPPDATA%\ms-playwright.
# Do NOT crawl the whole site. The method reference is already local (api_doc.ps1, reference/*.md).
import glob
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "example_cache")
BASES = {
    "sldworks": ("https://help.solidworks.com/2023/english/api/sldworksapi/",
                 "SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.%s~%s.html"),
    "cosworks": ("https://help.solidworks.com/2023/english/api/swsimulationapi/",
                 "SolidWorks.Interop.cosworks~SolidWorks.Interop.cosworks.%s~%s.html"),
}


def find_browser():
    pats = [os.path.expandvars(r"%LOCALAPPDATA%\ms-playwright\chromium-*\chrome-win*\chrome.exe"),
            os.path.expandvars(r"%LOCALAPPDATA%\ms-playwright\chromium_headless_shell-*\chrome-headless-shell-win64\*.exe"),
            r"C:\Program Files\Google\Chrome\Application\chrome.exe"]
    for pat in pats:
        hits = sorted(glob.glob(pat))
        if hits:
            return hits[-1]
    return None


def slug(t):
    return re.sub(r"[^A-Za-z0-9]+", "_", t).strip("_")


def code_of(text):
    i = text.find("Example (")
    if i < 0:
        i = 0
    i = text.rfind("\n", 0, i) + 1
    j = text.find("Knowledge Base", i)
    body = text[i:j if j > 0 else None]
    return "\n".join(ln.rstrip() for ln in body.split("\n") if ln.strip())


def main():
    args = [a for a in sys.argv[1:]]
    if not args or "-h" in args:
        print(__doc__)
        return
    ns = "sldworks"
    title = ""
    only_list = False
    members = []
    k = 0
    while k < len(args):
        a = args[k]
        if a == "--ns":
            ns = args[k + 1]; k += 2; continue
        if a == "--title":
            title = args[k + 1].lower(); k += 2; continue
        if a == "--list":
            only_list = True; k += 1; continue
        members.append(a); k += 1

    os.makedirs(CACHE, exist_ok=True)
    idx_path = os.path.join(CACHE, "index.json")
    idx = json.load(open(idx_path)) if os.path.exists(idx_path) else {}
    base, pat = BASES[ns]
    todo = [m for m in members if (ns + ":" + m) not in idx]
    if todo:
        from playwright.sync_api import sync_playwright
        exe = find_browser()
        with sync_playwright() as p:
            b = p.chromium.launch(executable_path=exe) if exe else p.chromium.launch()
            pg = b.new_page()
            for m in todo:
                iface, mem = m.split(".", 1)
                try:
                    pg.goto(base + pat % (iface, mem), wait_until="domcontentloaded", timeout=60000)
                    pg.wait_for_timeout(5000)
                except Exception as e:
                    print("could not open the page for", m, ":", str(e)[:120]); continue
                links = pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim(), e.href])")
                if len(links) < 5:
                    print("page for", m, "did not load fully (", len(links), "links); not cached, try again"); continue
                ex = [[t, h] for t, h in links if re.search(r"\((VBA|VB)\)", t) or re.search(r"Example", t) and "VB" in h]
                ent = []
                for t, h in ex:
                    fn = slug(t) + ".txt"
                    fp = os.path.join(CACHE, fn)
                    if not os.path.exists(fp):
                        try:
                            pg.goto(h, wait_until="domcontentloaded", timeout=60000)
                            pg.wait_for_timeout(4000)
                            open(fp, "w", encoding="utf-8").write(h + "\n" + code_of(pg.inner_text("body")))
                        except Exception as e:
                            print("could not fetch", t, str(e)[:120]); continue
                    ent.append([t, fn])
                idx[ns + ":" + m] = ent
            b.close()
        json.dump(idx, open(idx_path, "w"), indent=1)

    for m in members:
        ent = idx.get(ns + ":" + m)
        if ent is None:
            continue
        print("=== %s: %d VBA example(s)" % (m, len(ent)))
        for t, fn in ent:
            if title and title not in t.lower():
                continue
            if only_list:
                print("  -", t)
                continue
            print("\n##### " + t)
            txt = open(os.path.join(CACHE, fn), encoding="utf-8").read()
            print(txt.replace(" ", " ").encode("ascii", "replace").decode("ascii"))


main()
