"""
check_generators.py - finds out why the Generators tile isn't showing.

Put it in the page folder (next to page_template.html and build_page.py) and run:
    python check_generators.py
It changes nothing - it only reads the files and reports.
"""
import glob
import os
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
ok_all = True


def check(label, ok, fix=""):
    global ok_all
    print(("  OK    " if ok else "  FIX   ") + label)
    if not ok:
        ok_all = False
        if fix:
            print("        -> " + fix)


def read(name):
    try:
        return open(name, encoding="utf-8").read()
    except OSError:
        return None


def when(name):
    return time.strftime("%d-%m-%Y %H:%M", time.localtime(os.path.getmtime(name))) if os.path.exists(name) else "-"


print(f"\nPage folder: {HERE}\n")
print("1) The files the patch needs")
for f in ("page_template.html", "bridge.py", "pylibs.zip", "build_page.py", "tool_generator_sync.py", "patch_generators.py"):
    check(f"{f} is in this folder", os.path.exists(f), f"put {f} in this folder")

print("\n2) Was the patch applied?  (python patch_generators.py)")
tpl, br = read("page_template.html") or "", read("bridge.py") or ""
check("page_template.html has the Generators tile + page", "function pageGen(" in tpl and 'id:"gens"' in tpl,
      "run: python patch_generators.py  (and read what it prints)")
check("bridge.py has the gs_* functions", "def gs_probe(" in br, "run: python patch_generators.py")
try:
    names = zipfile.ZipFile("pylibs.zip").namelist()
    check("pylibs.zip has msoffcrypto + olefile",
          any(n.startswith("msoffcrypto/") for n in names) and any(n.startswith("olefile") for n in names),
          "pip install msoffcrypto-tool, then run: python patch_generators.py")
except (OSError, zipfile.BadZipFile):
    check("pylibs.zip can be opened", False)

print("\n3) Was index.html rebuilt afterwards?  (python build_page.py)")
idx = read("index.html")
check("index.html exists in this folder", idx is not None, "run: python build_page.py")
if idx is not None:
    check("index.html contains the Generators tool", "function pageGen(" in idx and "tool_generator_sync.py" in idx,
          "run: python build_page.py  (it must list tool_generator_sync.py in what it prints)")
    sources = [f for f in ["page_template.html", "bridge.py", "pylibs.zip"] + glob.glob("tool_*.py") if os.path.exists(f)]
    newest = max(sources, key=os.path.getmtime) if sources else None
    if newest:
        check(f"index.html ({when('index.html')}) is newer than {newest} ({when(newest)})",
              os.path.getmtime("index.html") >= os.path.getmtime(newest), "run: python build_page.py")

print("\n4) Is the browser showing THIS index.html?")
print(f"   This folder's page:  {os.path.join(HERE, 'index.html')}")
print("   * Opened by double-clicking: check the address bar shows exactly that path.")
print("   * Opened through the server (python app.py, http://...:8080): the server shows the")
print("     index.html in ITS own folder - copy the new index.html there (or build there).")
print("   * Opened from GitHub Pages / another web address: upload / commit the new")
print("     index.html there; it can take a minute or two to update.")
print("   * Then reload without the cache: Ctrl + F5.")
print("   Quick test in the browser: press F12 > Console, type   typeof pageGen   and Enter.")
print("   'function' = the new page is loaded;  'undefined' = the browser has an old index.html.")

print("\n" + ("All files here are fine - the browser is showing a different / older index.html (see 4)."
              if ok_all else "Fix the items marked FIX above, top to bottom, then refresh with Ctrl + F5."))
