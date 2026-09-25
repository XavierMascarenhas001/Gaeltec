"""
build_page.py - builds index.html (one self-contained file) from:
  page_template.html  - the page (layout, drop zones, filters)
  worker.js           - runs Python in the browser (Pyodide)
  dashboards.css, dash_core.js, dash_pages.js - the three dashboards
  bridge.py, py_boot.py, tool_*.py, dash_*.py, engine.py - the Python that actually does the work
  pylibs.zip          - pure-Python libraries the browser doesn't have
                        (openpyxl, et_xmlfile, xlsxwriter, python-docx,
                        pyxlsb, rapidfuzz's pure-Python version)
After changing any of those files, run:   python build_page.py
"""
import base64, glob, json, os

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
PY_FILES = ["bridge.py", "py_boot.py", "engine.py"] + sorted(glob.glob("tool_*.py")) + sorted(glob.glob("dash_*.py"))
sources = {f: open(f, encoding="utf-8").read() for f in PY_FILES}
page = open("page_template.html", encoding="utf-8").read()
page = page.replace("/*__DASH_CSS__*/", open("dashboards.css", encoding="utf-8").read())
page = page.replace("/*__DASH_JS__*/", open("dash_core.js", encoding="utf-8").read() + "\n" + open("dash_pages.js", encoding="utf-8").read())
page = page.replace("/*__WORKER_JS__*/", open("worker.js", encoding="utf-8").read())
page = page.replace("/*__PY_SOURCES__*/", json.dumps(sources).replace("</", "<\\/"))
page = page.replace("/*__PYLIBS_B64__*/", base64.b64encode(open("pylibs.zip", "rb").read()).decode())
open("index.html", "w", encoding="utf-8").write(page)
print(f"index.html built ({os.path.getsize('index.html') // 1024} KB) with {', '.join(PY_FILES)}")
