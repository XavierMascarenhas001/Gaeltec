# Gaeltec Tools – runs in your browser, no server

`index.html` is the whole app. Open it in Chrome or Edge, pick a tool, and
drag your files onto it. Your Python runs inside the browser and results
download like normal files. Nothing is uploaded anywhere, and there's no
server to start.

| # | Tile | Python file |
|---|------|-------------|
| 1 | Target Price → Control File | `tool_tp_to_cf.py` |
| 2 | Aggregate Control Files | `tool_aggregate_cf.py` |
| 3 | Build Master Parquet | `tool_build_master.py` |
| 4 | Outputs (CV Excel Report) | `tool_materials_report.py` |
| 5 | Work Instructions & Map Check | `tool_work_instructions.py` |

## Using it

* **Open:** double-click `index.html`, or host it on GitHub Pages and share the
  link.
* **First load:** the page downloads Python (Pyodide 0.29.5 from
  `cdn.jsdelivr.net`), which takes about 30 seconds the first time. After that
  the browser keeps a copy. The pill at the top right says **Python ready**
  when it's done.
* **Adding files:** drag files onto a box, or click the box to choose them.
  The file-type filters are the same as the old dialogs.
* **Saving:** outputs download automatically. Each one also gets a
  **Download** button, plus **Save to a folder…**, which lets you pick any
  folder, including one on the network share.
* **Map check:** drag a **month** or **year** folder from
  `1 - Outages Programme` onto tab 3, or the whole programme folder. The
  scan uses the dates from tab 1, exactly as before. Map names in the report
  link to the real PDFs on `\\gaeltec-gl`.
* **Outputs logos:** drop `GaeltecImage.png` and `SPEN.png` once. The page
  remembers them.

## What's the same, what's different

* The tool files are the same code you run in Jupyter: same mappings, rules,
  filters and Excel/Word formatting. They're built into `index.html` together
  with the libraries they need.
* The browser can't use threads, so PDFs and folders are read one after
  another instead of 20 at a time. The results are the same, but big map
  scans take longer.
* rapidfuzz uses its own built-in pure-Python version. Its scores were checked
  identical to the normal rapidfuzz on 4,000 string pairs.
* In the map scan, an outage folder that is completely **empty** isn't listed
  as "had no Workpack zones folder", because browsers don't pass empty
  folders through. The results are otherwise unchanged.
* Very large jobs, such as aggregating 75+ Control Files, run slower than in
  Jupyter. They may hit the browser's memory limit; if that happens, do them
  in batches.

## Dashboards

The three Streamlit dashboards can't run inside a web page on their own.
Start them from Jupyter as usual (Network Job Tracker 8501, Master Control
8502, Materials Breakdown 8503) and the **Dashboards** section shows them
while they're running.

`network_job_tracker.py` here has one extra line that keeps the outage
programme's PID column as text. It stops the repeated "Could not convert
'Not delivered' … column PID" warnings in its log. Copy it over yours if you
want that.

## Changing a tool

Edit the `tool_*.py` file, then run `python build_page.py` to rebuild
`index.html`. The rest of the files are the parts `index.html` is built
from:

* `page_template.html` – the page layout
* `worker.js` – runs Python in the browser
* `bridge.py` – connects the page to the tools
* `py_boot.py` – swaps in the no-threads and pure-rapidfuzz versions
* `pylibs.zip` – openpyxl, xlsxwriter, python-docx, pyxlsb and rapidfuzz

## If it doesn't work

* **"Couldn't download Python from cdn.jsdelivr.net"**: the PC is offline or
  the site is blocked. Ask IT to allow `cdn.jsdelivr.net`.
* **A dashboard shows "Not running"**: start it from Jupyter. The page checks
  every few seconds and shows it once it's up.
