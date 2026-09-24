# Gaeltec Tools – web launcher

One web page with five tiles that run the Gaeltec Python tools. You pick files
from the **server's** view of the network share, so there's nothing to install
on each user's PC. Everyone just opens a link.

| # | Tile | Python file | Was |
|---|------|-------------|-----|
| 1 | Target Price → Control File | `tool_tp_to_cf.py` | TP → Template_CF script |
| 2 | Aggregate Control Files | `tool_aggregate_cf.py` | Block1 / PA CONTROL aggregator |
| 3 | Build Master Parquet | `tool_build_master.py` | both Build Master scripts (Man_day tick box) |
| 4 | Outputs (CV Excel Report) | `tool_materials_report.py` | "Excel Export Tool" window |
| 5 | Work Instructions & Map Check | `tool_work_instructions.py` | 3-tab "Pole / Work Instructions Reporting Tool" |

`app.py` is the web server and `index.html` is the page.

## Dashboards

The three Streamlit dashboards are started by the server and shown inside the
page, under **Dashboards** on the home screen:

| Tile | File | Port |
|------|------|------|
| Network Job Tracker | `network_job_tracker.py` | 8501 |
| Master Control | `master_control_dashboard.py` | 8502 |
| Materials Breakdown | `materials_breakdown_dashboard.py` (was `app.py`) + `engine.py` | 8503 |

* With `dashboard_autostart: true` in `config.json`, all three start with the
  server. Otherwise each one starts when you first open it. The first load can
  take up to 90 seconds on the network share.
* If one is already running, for example from Jupyter with
  `launch_dashboard(...)`, the page just shows it and doesn't start a second
  copy.
* Each dashboard writes its output to `<key>_dashboard.log` next to `app.py`.
  If a dashboard fails to start, the page shows the last lines of that log.
* **Restart** reloads a dashboard. **Open in new tab** opens it full screen.
* The dashboards run on their own ports, so the `allowed_clients` list does
  **not** protect them. Use Windows Firewall to limit ports 8501–8503 to the
  same PCs.
* The Materials Breakdown file was renamed from `app.py` so it doesn't clash
  with the web server. Its code is unchanged.
* One line was added to `network_job_tracker.py`: the outage programme's
  PID column is kept as text. That stops the repeated "Could not convert
  'Not delivered' … column PID" warnings in the log.

## What changed and what didn't

* The mapping dictionaries, rules, filters and Excel/Word formatting are
  copied **word for word** from the original scripts. Only the tkinter
  windows and dialogs were replaced.
* The file dialogs became a file browser in the page. It only shows the
  folders listed in `config.json`, and it keeps the same file-type filters
  (e.g. the Target Price picker still defaults to `*.xlsm`).
* The filters work the same way as before:
  * **Materials:** the "All / N selected" tick-list boxes, the custom-value
    **Add** box, All / None / OK / Cancel, "All (No Connections)", the fixed
    Const/Mat category, the multi date filter (plan1 / plan2 / done /
    datetouse / all / unplanned / undone, combined with OR), and the CV Groups
    box.
  * **Work Instructions:** search + click-to-toggle lists with counts,
    All (visible) / Clear, a YYYY-MM-DD date range that also drives the
    network scan, and a 500-row preview. The maps you pick are shared between
    tabs 2 and 3, and are only read once.
* The "Choose Line" pop-up (Target Price) became a review step: press
  **Check for similar lines**, then pick *Use line 1 / Use line 2 / Keep both*
  for each pair.
* **Browse** opens the file browser straight away. If the server can't be
  reached, it says so inside the browser.
* **Upload** (next to Browse) sends a file from your own PC to the server. It
  is saved in the **Uploaded files** folder, which you can also browse later.
* Saving asks for a folder and a file name. If the file already exists, you
  get an "already exists – replace it?" prompt. Every output also gets a
  **Download** button.
* The two Build Master scripts are now one. With **Man_day** ticked it works
  like the newer script; unticked, it works like the older one.
* A missing comma after `"Reinforcement_Lanark"` in the aggregator's
  `file_project_mapping` was fixed. Without it, that script can't run.

## Set up (once, on a Windows PC or server that can reach `\\gaeltec-gl`)

1. Install Python 3.10+ (or use Anaconda).
2. Copy this folder onto that machine, or `git clone` it from GitHub.
3. Edit **`config.json`**:
   * `allowed_clients` – the PCs that may use it. Add each PC's IP address
     (run `ipconfig` on the PC) or a whole subnet such as `"10.20.30.0/24"`.
     Anyone else gets "not on the allowed list".
   * `roots` – the only folders the file browser can open. `writable: true`
     lets people save outputs there.
   * `access_key` – optional shared password on top of the IP list.
4. Double-click **`start_server.bat`**. It installs the requirements and
   starts the server on port 8080. Leave the window open.
5. On the allowed PCs, open `http://<server-name>:8080`.
   You may need to allow port 8080 through Windows Firewall on the server.

The server account needs read and write access to the share. The page shows
"x/y folders reachable" at the top right.

## GitHub

Put this folder in a GitHub repo so everyone runs the same version. To
update, run `git pull` on the server and restart `start_server.bat`.

You can also host `index.html` on **GitHub Pages** and point it at your
server with `https://<you>.github.io/<repo>/?api=http://<server>:8080`. For
that, add the Pages address to `cors_origins` in `config.json`. **Be aware:**
browsers block an https page from calling a plain-http internal server
("mixed content"). So unless the server has HTTPS, opening
`http://<server>:8080` directly is the reliable option. That address serves
exactly the same page.

## Running a tool without the web page

Each `tool_*.py` file can still be used from Jupyter, for example:

```python
import tool_build_master
tool_build_master.run(r"...\CF_aggregated.parquet", r"...\Project Tracker.parquet",
                      r"...\miscelaneous.parquet", r"...\Master_24-09-2026.parquet")
```

## "Server not reachable" / can't browse

The page is only the front end. `app.py` must be running before anything
works.

1. On a PC that can open `\\gaeltec-gl\...`, double-click
   **`start_server.bat`**. Leave the black window open, because closing it
   stops the server.
   * If the window says **"Address already in use"**, something else is using
     port 8080. Change `"port"` in `config.json` (e.g. `8090`) and use that
     number below.
   * If it says **"No module named …"**, run
     `pip install -r requirements.txt` (in an Anaconda Prompt if you use
     Anaconda).
2. Open **http://localhost:8080** on that PC, or
   **http://&lt;that-PC-name&gt;:8080** from other PCs. Don't double-click
   `index.html`, and don't open the GitHub copy.
3. If the page was opened as a file or from GitHub, use the red box at the
   top of the page to type the server address and press **Connect**.
4. **"This computer is not allowed"** means you need to add that PC's IP to
   `allowed_clients` in `config.json` and restart `start_server.bat`.
