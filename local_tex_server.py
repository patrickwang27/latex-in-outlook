#!/usr/bin/env python3
import base64, io, os, shutil, subprocess, tempfile
from flask import Flask, request, jsonify, make_response

app = Flask(__name__)

# Configure binaries (adjust if needed)
PDFLATEX = '/home/yuhengw/texlive/2024/bin/x86_64-linux/pdflatex'
GHOSTSCRIPT='/usr/bin/gs'

if not PDFLATEX:
    raise RuntimeError("pdflatex not found on PATH")
if not GHOSTSCRIPT:
    raise RuntimeError("Ghostscript (gs) not found on PATH")

TEMPLATE = r"""
\documentclass[varwidth,border=2pt]{standalone}
\usepackage{amsmath,amssymb}
\usepackage{bm}
\usepackage{mathtools}
\begin{document}
%s
\end{document}
"""

def latex_to_png(eqn_tex: str, display: bool, dpi: int = 300) -> bytes:
    # Wrap as inline math or display math
    body = f"\\[{eqn_tex}\\]" if display else f"\\({eqn_tex}\\)"
    tex = TEMPLATE % body

    with tempfile.TemporaryDirectory() as td:
        tex_path = os.path.join(td, "eqn.tex")
        pdf_path = os.path.join(td, "eqn.pdf")
        png_path = os.path.join(td, "eqn.png")

        with open(tex_path, "w", encoding="utf-8") as f:
            f.write(tex)

        # Quiet pdflatex
        subprocess.run([PDFLATEX, "-interaction=nonstopmode", "-halt-on-error", tex_path],
                       cwd=td, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, check=True)

        # Ghostscript PDF -> PNG with transparency
        # -sDEVICE=pngalpha preserves alpha; -r sets DPI; -dTextAlphaBits/-dGraphicsAlphaBits for quality
        gs_cmd = [
            GHOSTSCRIPT,
            "-dSAFER", "-dBATCH", "-dNOPAUSE", "-dQUIET",
            "-sDEVICE=pngalpha",
            f"-r{dpi}",
            "-dTextAlphaBits=4", "-dGraphicsAlphaBits=4",
            f"-sOutputFile={png_path}",
            pdf_path
        ]
        subprocess.run(gs_cmd, cwd=td, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, check=True)

        with open(png_path, "rb") as f:
            return f.read()

@app.route("/render", methods=["POST"])
def render():
    try:
        data = request.get_json(force=True)
        tex = data.get("tex", "")
        display = bool(data.get("display", False))
        dpi = int(data.get("dpi", 300))
        if not tex.strip():
            return jsonify({"error": "empty tex"}), 400

        png_bytes = latex_to_png(tex, display, dpi)
        resp = make_response(png_bytes)
        resp.headers["Content-Type"] = "image/png"
        # CORS (not strictly needed for GM_xmlhttpRequest, but harmless)
        resp.headers["Access-Control-Allow-Origin"] = "*"
        return resp
    except subprocess.CalledProcessError as e:
        return jsonify({"error": "latex/gs failed", "log": e.stdout.decode("utf-8", "ignore")}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/health")
def health():
    return jsonify({"ok": True})

if __name__ == "__main__":
    # Listen only on localhost
    app.run(host="127.0.0.1", port=8765)
