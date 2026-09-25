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

The three dashboards are now built into the page, so Streamlit is no longer
needed. Open one from the **Dashboards** section and drop your files on it.
The switcher at the top moves between them, and each keeps its filters and tab
while you're away.

| Dashboard | Files to drop | Python file |
|-----------|---------------|-------------|
| Network Job Tracker | Master file (.parquet or .csv); optionally High-level_planning_2026.xlsx, the Service Partner Workbank and the outage calendar .ics | `dash_tracker.py` |
| Master Control | Master file | `dash_master.py` |
| Materials Breakdown | materials_all.parquet + the Master file (maps optional, on the Maps tab) | `dash_materials.py` + `engine.py` |

* A Master file dropped once is shared by all three dashboards.
* All the filters sit in one bar above the charts, with the same options as
  the sidebar before (date field, quick range, District, Project, PID, and so
  on). Circuit and Pole still narrow to what the other filters allow.
* Every chart has a **Table** view, and every table has **CSV** download and
  search. **◐ Auto / ☀ Light / ☾ Dark** in the top bar sets the theme.
* Materials **Export** downloads the same `Materials_Breakdown_<date>.xlsx`
  workbook as before.
* The numbers were checked against the original Streamlit apps on the same
  data. Every KPI, card, total and row count matched.

`network_job_tracker.py` here has one extra line that keeps the outage
programme's PID column as text. It stops the repeated "Could not convert
'Not delivered' … column PID" warnings in its log. It's only needed if you
still run the Streamlit version.

## Changing a tool

Edit the `tool_*.py` / `dash_*.py` file, then run `python build_page.py` to rebuild
`index.html`. The rest of the files are the parts `index.html` is built
from:

* `page_template.html` – the page layout
* `dashboards.css`, `dash_core.js`, `dash_pages.js` – the dashboards (layout, filters, charts)
* `worker.js` – runs Python in the browser
* `bridge.py` – connects the page to the tools
* `py_boot.py` – swaps in the no-threads and pure-rapidfuzz versions
* `pylibs.zip` – openpyxl, xlsxwriter, python-docx, pyxlsb and rapidfuzz

## If it doesn't work

* **"Couldn't download Python from cdn.jsdelivr.net"**: the PC is offline or
  the site is blocked. Ask IT to allow `cdn.jsdelivr.net`.
* **A dashboard shows a red message after dropping a file**: it's the same
  error the Streamlit app would have given, for example a missing column. On
  the Network Job Tracker, open **⚙ Column mapping** and pick the right column.
