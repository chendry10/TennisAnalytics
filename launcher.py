from __future__ import annotations

import os
import socket
import sys
from pathlib import Path


def _find_open_port(start: int = 8501, attempts: int = 20) -> int:
    for port in range(start, start + attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return start


def main() -> int:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    os.chdir(base_dir)

    app_path = base_dir / "app.py"
    if not app_path.exists():
        app_path = Path(__file__).resolve().parent / "app.py"

    port = _find_open_port()

    os.environ["STREAMLIT_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_BROWSER_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_SERVER_ADDRESS"] = "127.0.0.1"
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.port",
        str(port),
        "--browser.serverPort",
        str(port),
        "--server.address",
        "127.0.0.1",
    ]

    try:
        from streamlit.web import cli as stcli
    except Exception:
        import streamlit.web.cli as stcli

    return stcli.main()


if __name__ == "__main__":
    raise SystemExit(main())
