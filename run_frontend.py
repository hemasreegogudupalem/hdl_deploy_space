"""Hdl Automation Frontend Launcher.

Starts the FastAPI web server and launches the web UI.
"""

import json
import socket
import sys
import threading
import time
from urllib.error import URLError
from urllib.request import urlopen
import webbrowser
import uvicorn


def _is_port_in_use(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as connection:
        return connection.connect_ex((host, port)) == 0


def _is_hdl_server_running(url: str) -> bool:
    try:
        with urlopen(f"{url}/api/status", timeout=2) as response:
            status = response.status
            payload = json.load(response)
    except (OSError, URLError, ValueError):
        return False

    return (
        status == 200
        and isinstance(payload, dict)
        and isinstance(payload.get("input_files_count"), int)
        and isinstance(payload.get("outputs_count"), int)
    )


def _open_browser_delayed(url: str, delay: float = 1.0):
    time.sleep(delay)
    try:
        webbrowser.open(url)
    except Exception:
        pass


def main():
    host = "127.0.0.1"
    port = 8000
    url = f"http://{host}:{port}"

    print("=" * 68)
    print("   HDL AUTOMATION (WEB UI)")
    print("=" * 68)
    print(f" -> Local Web App URL: {url}")
    print(" -> Existing single-template flow preserved; batch generation enabled.")
    print(" -> Press Ctrl+C in this console to stop the web server.")
    print("=" * 68)

    if _is_port_in_use(host, port):
        if _is_hdl_server_running(url):
            print(" -> HDL Automation is already running; reusing the existing server.")
            return
        raise SystemExit(
            f"Port {port} is already in use by another application. "
            "Stop that application before starting HDL Automation."
        )

    # Launch browser after uvicorn initializes socket listener
    threading.Thread(target=_open_browser_delayed, args=(url, 1.2), daemon=True).start()

    uvicorn.run("frontend.app:app", host=host, port=port, log_level="info")


if __name__ == "__main__":
    main()
