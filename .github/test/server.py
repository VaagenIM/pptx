import os
import shutil
import subprocess
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


class CorsHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        super().end_headers()


if __name__ == "__main__":
    root = Path(__file__).resolve().parents[2]
    os.chdir(root)
    npm = shutil.which("npm.cmd" if sys.platform == "win32" else "npm")
    if not npm:
        raise SystemExit("npm is required to start the viewer server.")
    viewer_root = root / ".github" / "viewer"
    overrides_root = root / ".github" / "viewer-overrides"
    marker = "/* pptx repository local overrides */"
    styles = viewer_root / "src" / "styles.css"
    main = viewer_root / "src" / "main.js"
    css_override = (overrides_root / "styles-additions.css").read_text(encoding="utf-8")
    js_override = (overrides_root / "main-additions.js").read_text(encoding="utf-8")
    styles_source = styles.read_text(encoding="utf-8").split(marker, 1)[0].rstrip()
    main_source = main.read_text(encoding="utf-8").split(marker, 1)[0].rstrip()
    styles.write_text(f"{styles_source}\n{marker}\n{css_override}", encoding="utf-8")
    main.write_text(f"{main_source}\n{marker}\n{js_override}", encoding="utf-8")
    viewer = subprocess.Popen(
        [npm, "run", "dev", "--", "--host", "localhost"],
        cwd=viewer_root,
    )
    server = ThreadingHTTPServer(("localhost", 8000), CorsHandler)
    print("Files: http://localhost:8000")
    print("Viewer: http://localhost:5173/pptx-web/")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        viewer.terminate()
        viewer.wait()
