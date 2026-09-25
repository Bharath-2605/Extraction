#!/usr/bin/env python3
"""Start PDF URL Capture from Anaconda, Jupyter, or a terminal.

Uses the current Python. Does not create or require a virtualenv.
"""
from __future__ import annotations

import os
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from argparse import ArgumentParser
from pathlib import Path

ROOT = Path(__file__).resolve().parent

os.chdir(ROOT)
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _http_options() -> tuple[str, int, str, bool]:
    parser = ArgumentParser(add_help=False)
    parser.add_argument("--anaconda-project-host", action="append", default=[])
    parser.add_argument("--anaconda-project-port", type=int, default=None)
    parser.add_argument("--anaconda-project-iframe-hosts", action="append")
    parser.add_argument("--anaconda-project-no-browser", action="store_true")
    parser.add_argument("--anaconda-project-use-xheaders", action="store_true")
    parser.add_argument("--anaconda-project-url-prefix", default="")
    parser.add_argument("--anaconda-project-address", default="0.0.0.0")
    args, _unknown = parser.parse_known_args()
    address = args.anaconda_project_address or "0.0.0.0"
    port = args.anaconda_project_port
    if port is None:
        port = int(os.environ.get("PORT", "4747"))
    prefix = (args.anaconda_project_url_prefix or os.environ.get("ANACONDA_PROJECT_URL_PREFIX") or "").rstrip("/")
    no_browser = bool(args.anaconda_project_no_browser or prefix)
    return address, port, prefix, no_browser


def _public_url(port: int, prefix: str) -> str:
    return f"http://localhost:{port}{prefix}"


def _health_url(port: int, prefix: str) -> str:
    return f"http://127.0.0.1:{port}{prefix}/health"


def _ensure_packages() -> None:
    try:
        import fastapi  # noqa: F401
        import openpyxl  # noqa: F401
        import pandas  # noqa: F401
        import pymupdf  # noqa: F401
        import uvicorn  # noqa: F401
        import cv2  # noqa: F401
        return
    except ImportError:
        pass

    requirements = ROOT / "requirements.txt"
    print("Installing required packages into this Python (no virtualenv)...")
    print(f"Using: {sys.executable}")
    cmd = [sys.executable, "-m", "pip", "install", "-r", str(requirements)]
    try:
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(
            "Could not install packages with pip. "
            "On Anaconda Workbench, wait for the project environment to finish building, "
            "then run launch.py again."
        ) from exc


def _port_open(port: int, prefix: str) -> bool:
    try:
        with urllib.request.urlopen(_health_url(port, prefix), timeout=1.5) as response:
            return response.status == 200
    except (urllib.error.URLError, TimeoutError, OSError):
        return False


def _port_busy(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        return sock.connect_ex(("127.0.0.1", port)) == 0


def _in_notebook() -> bool:
    return "ipykernel" in sys.modules


def _build_app(prefix: str):
    from fastapi import FastAPI
    from fastapi.middleware.cors import CORSMiddleware

    import server as server_module

    app = server_module.app
    if not prefix:
        return app
    wrapper = FastAPI()
    wrapper.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_methods=["*"],
        allow_headers=["*"],
        expose_headers=["Content-Disposition"],
    )
    wrapper.mount(prefix, app)
    return wrapper


def start_app(open_browser: bool = True) -> str:
    """Start the Python API + web UI. Keep the notebook kernel running."""
    _ensure_packages()
    address, port, prefix, workbench_no_browser = _http_options()
    url = _public_url(port, prefix)
    should_open = open_browser and not workbench_no_browser

    if _port_open(port, prefix):
        if should_open:
            webbrowser.open(url)
        print(f"PDF URL Capture is already running at {url}")
        return url

    if _port_busy(port):
        raise RuntimeError(f"Port {port} is in use by another program. Stop it and try again.")

    import uvicorn

    serve_app = _build_app(prefix)
    config = uvicorn.Config(
        serve_app,
        host=address,
        port=port,
        log_level="info",
        proxy_headers=True,
        forwarded_allow_ips="*",
    )
    server = uvicorn.Server(config)

    if not _in_notebook():
        if should_open:
            threading.Thread(target=_open_when_ready, args=(port, prefix, url), daemon=True).start()
        print(f"PDF URL Capture is starting at {url}")
        print("Keep this window running while you use the webpage.")
        server.run()
        return url

    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()
    for _ in range(80):
        if _port_open(port, prefix):
            break
        time.sleep(0.1)
    else:
        raise RuntimeError("The app started but did not become ready. Check the server logs.")

    if should_open:
        webbrowser.open(url)
    print(f"PDF URL Capture is running at {url}")
    print("Keep this window/kernel running while you use the webpage.")
    return url


def _open_when_ready(port: int, prefix: str, url: str) -> None:
    for _ in range(80):
        if _port_open(port, prefix):
            webbrowser.open(url)
            return
        time.sleep(0.1)


if __name__ == "__main__":
    try:
        start_app()
    except KeyboardInterrupt:
        print("\nStopped.")
