"""Lokaler Testserver für die WebApp: python serve.py → http://localhost:5173/ (Demo: /?demo).

Eigenes Skript statt `python -m http.server`, weil Windows .js je nach Registry als text/plain meldet
und der Browser Module dann nicht lädt.
"""

import functools
import http.server
from pathlib import Path

PORT = 5173
Handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(Path(__file__).parent))
Handler.func.extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map,
                               ".js": "text/javascript", ".webmanifest": "application/manifest+json"}

if __name__ == "__main__":
    print(f"http://localhost:{PORT}/")
    http.server.ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
