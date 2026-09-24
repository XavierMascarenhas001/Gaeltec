"""
Start the Gaeltec Tools server from JupyterLab (instead of start_server.bat).

In a notebook cell:

    %run "C:/path/to/gaeltec_web/start_from_jupyter.py"

It starts app.py as a separate process (never inside the Jupyter kernel),
waits for it to answer, and opens http://localhost:8080 in your browser.
Server output goes to server.log in this folder.
Run it again at any time: if the server is already up it just says so.
"""
import json
import os
import socket
import subprocess
import sys
import time
import webbrowser

HERE = os.path.dirname(os.path.abspath(__file__))


def _port(cfg_path=os.path.join(HERE, "config.json")):
    try:
        with open(cfg_path, encoding="utf-8") as f:
            return int(json.load(f).get("port", 8080))
    except Exception:
        return 8080


def _up(port):
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=0.5):
            return True
    except OSError:
        return False


def start_server(timeout=90):
    port = _port()
    url = f"http://localhost:{port}"
    if _up(port):
        print(f"Gaeltec Tools server is already running -> {url}")
        webbrowser.open(url)
        return
    log_path = os.path.join(HERE, "server.log")
    log = open(log_path, "w", encoding="utf-8")
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    env = dict(os.environ, GAELTEC_NO_BROWSER="1")  # we open the page ourselves below
    proc = subprocess.Popen([sys.executable, os.path.join(HERE, "app.py")], cwd=HERE,
                            stdout=log, stderr=subprocess.STDOUT, creationflags=flags, env=env)
    print(f"Starting Gaeltec Tools server (up to {timeout}s)...")
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _up(port):
            print(f"Running -> {url}   (log: {log_path})")
            webbrowser.open(url)
            return proc
        if proc.poll() is not None:
            break
        time.sleep(1)
    log.close()
    print("The server did not start. Last lines of server.log:\n")
    with open(log_path, encoding="utf-8", errors="replace") as f:
        print("".join(f.readlines()[-30:]))


if __name__ == "__main__":
    start_server()
