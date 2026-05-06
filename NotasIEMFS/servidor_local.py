#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sirve los archivos estáticos y reenvía POST /notas-gas-proxy → Apps Script.
El navegador no puede seguir bien el 302 de Google con el cuerpo POST; curl -L sí.
Mantén DEFAULT_GAS_EXEC alineado con API_DIRECT en src/config.js.
"""

from __future__ import annotations

import os
import re
import subprocess
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).parent.resolve()
PROXY_TIMEOUT_SEC = int(os.environ.get("NOTAS_PROXY_TIMEOUT_SEC", "300"))

DEFAULT_GAS_EXEC = (
    "https://script.google.com/macros/s/AKfycbyLKQ4M-oXEFPkWnhoFUUtsbb2-aHtKCsVQCKCGeIITAecOHykZP63rC0s6j7Ws20tQNg/exec"
)


def _gas_exec_from_config_js() -> str | None:
    """
    Misma URL que `export const API` en src/config.js.
    Así, en localhost el proxy y la app no se desalinean (antes solo config.js cambiaba y el proxy seguía al GAS viejo).
    """
    cfg = ROOT / "src" / "config.js"
    if not cfg.is_file():
        return None
    try:
        text = cfg.read_text(encoding="utf-8")
    except OSError:
        return None
    m = re.search(r'export const API\s*=\s*(?:\n\s*)?"([^"]+)"', text)
    if not m:
        return None
    url = m.group(1).strip()
    return url or None


def _json_error(status_code: int, message: str) -> bytes:
    msg = str(message or "").replace("\\", "\\\\").replace('"', '\\"')
    return ('{"ok":false,"error":"' + msg + '"}').encode("utf-8")


GAS_EXEC = (
    (os.environ.get("NOTAS_GAS_EXEC") or "").strip()
    or _gas_exec_from_config_js()
    or DEFAULT_GAS_EXEC
)


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_POST(self):
        path_only = self.path.split("?", 1)[0]
        if path_only != "/notas-gas-proxy":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0
        body = self.rfile.read(length) if length else b""

        curl_exe = "curl.exe" if sys.platform == "win32" else "curl"
        try:
            proc = subprocess.run(
                [
                    curl_exe,
                    "-s",
                    "-S",
                    "-L",
                    "--connect-timeout",
                    "20",
                    "--max-time",
                    str(PROXY_TIMEOUT_SEC),
                    "-H",
                    "Content-Type: text/plain;charset=utf-8",
                    "--data-binary",
                    "@-",
                    GAS_EXEC,
                ],
                input=body,
                capture_output=True,
                timeout=PROXY_TIMEOUT_SEC + 10,
            )
        except FileNotFoundError:
            self.send_response(500)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(
                _json_error(500, "No se encontro curl. En Windows 10+ suele estar curl.exe en el sistema.")
            )
            return

        if proc.returncode != 0:
            err = (proc.stderr or b"").decode("utf-8", errors="replace")[:800]
            self.send_response(502)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(_json_error(502, "Error al llamar a Apps Script (curl): " + err))
            return

        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(proc.stdout)

    def log_message(self, format, *args):
        sys.stderr.write("%s - - [%s] %s\n" % (self.address_string(), self.log_date_time_string(), format % args))


def main() -> None:
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
    os.chdir(ROOT)
    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print("Carpeta:", ROOT)
    print("URL:    http://127.0.0.1:%s/" % port)
    print("Proxy GAS:", GAS_EXEC[:70] + ("..." if len(GAS_EXEC) > 70 else ""))
    print("Timeout proxy (s):", PROXY_TIMEOUT_SEC)
    httpd.serve_forever()


if __name__ == "__main__":
    main()
