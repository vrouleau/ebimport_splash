"""
ebimport_splash web UI.

Small Flask app that accepts an xlsx upload, runs one of the CLI
loaders (dry-run / MDB / Lenex) in a subprocess, parses the output
for the Summary + Issues sections, and offers the generated .mdb/.lxf
plus issues report as a ZIP download.

Stateless: each request gets its own temp dir, cleaned up when the
download is streamed or after `STAGING_TTL_SECS` of inactivity.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path

from flask import (
    Flask, jsonify, render_template, request, send_file, abort
)

# --------------------------------------------------------------------------- #
# Paths & config
# --------------------------------------------------------------------------- #
APP_DIR     = Path(__file__).parent.resolve()
REPO_ROOT   = APP_DIR.parent
MDB_LOADER  = REPO_ROOT / "load_to_mdb.py"
LNX_LOADER  = REPO_ROOT / "load_to_lenex.py"
EMPTY_MDB   = APP_DIR / "templates_mdb" / "empty_splash_meet.mdb"

STAGING_DIR = Path(os.environ.get("STAGING_DIR", "/tmp/ebimport_staging"))
STAGING_DIR.mkdir(parents=True, exist_ok=True)
STAGING_TTL_SECS = 30 * 60                # 30 minutes

# Upload size limits (bytes)
MAX_XLSX_BYTES = 50 * 1024 * 1024
MAX_MDB_BYTES  = 200 * 1024 * 1024

# --------------------------------------------------------------------------- #
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_XLSX_BYTES + MAX_MDB_BYTES + 1 * 1024 * 1024


# --------------------------------------------------------------------------- #
# Staging lifecycle
# --------------------------------------------------------------------------- #
@dataclass
class Staging:
    id: str
    dir: Path
    result_zip: Path
    created_at: float


_stagings: dict[str, Staging] = {}
_stagings_lock = threading.Lock()


def _new_staging() -> Staging:
    sid = uuid.uuid4().hex
    d = STAGING_DIR / sid
    d.mkdir(parents=True, exist_ok=True)
    s = Staging(id=sid, dir=d, result_zip=d / "result.zip",
                created_at=time.time())
    with _stagings_lock:
        _stagings[sid] = s
    return s


def _drop_staging(sid: str) -> None:
    with _stagings_lock:
        s = _stagings.pop(sid, None)
    if s is None:
        return
    try:
        shutil.rmtree(s.dir, ignore_errors=True)
    except Exception:
        pass


def _gc_stagings() -> None:
    """Remove staging dirs older than the TTL."""
    now = time.time()
    expired: list[str] = []
    with _stagings_lock:
        for sid, s in _stagings.items():
            if now - s.created_at > STAGING_TTL_SECS:
                expired.append(sid)
    for sid in expired:
        _drop_staging(sid)


# --------------------------------------------------------------------------- #
# Output parsing — extract Summary + Issues from loader stdout
# --------------------------------------------------------------------------- #
_BAR = re.compile(r"^=+\s*$")
_CAT = re.compile(r"^\s*\[(WARNING|NOTE)\]\s+([A-Za-z_]+)\s*:\s*(\d+)\s*$")
_BULLET = re.compile(r"^\s*-\s+(.*?)(?:\s+\(row\s+(\d+)\))?\s*$")


def parse_loader_output(text: str) -> dict:
    """Return {summary: [str], issues: {category: {...}}}.

    summary is the bulleted "+N new X" lines inside the 'Summary of
    changes' or 'Summary' block.  issues groups bullets under each
    `[WARNING] category: N` or `[NOTE] category: N` header."""
    lines = text.splitlines()

    summary: list[str] = []
    issues: dict[str, dict] = {}

    # Walk through and locate section headers.
    i = 0
    while i < len(lines):
        line = lines[i]
        # Section header is " Summary of changes" or " Summary" or " Issues..."
        stripped = line.strip()
        if stripped in ("Summary of changes", "Summary"):
            # Eat the following === line
            i += 1
            while i < len(lines) and _BAR.match(lines[i]):
                i += 1
            while i < len(lines) and not _BAR.match(lines[i]):
                ln = lines[i].rstrip()
                if ln.strip():
                    summary.append(ln.strip())
                i += 1
            i += 1
            continue
        if "Issues found" in stripped:
            # Walk the issues block
            i += 1
            while i < len(lines) and _BAR.match(lines[i]):
                i += 1
            cur_cat: dict | None = None
            while i < len(lines) and not _BAR.match(lines[i]):
                ln = lines[i].rstrip()
                mcat = _CAT.match(ln)
                if mcat:
                    severity, cat, n = mcat.group(1), mcat.group(2), int(mcat.group(3))
                    cur_cat = {"severity": severity, "category": cat,
                               "count": n, "items": []}
                    issues[cat] = cur_cat
                    i += 1
                    continue
                mit = _BULLET.match(ln)
                if mit and cur_cat is not None:
                    item = {"message": mit.group(1).rstrip()}
                    if mit.group(2):
                        item["row"] = int(mit.group(2))
                    cur_cat["items"].append(item)
                i += 1
            i += 1
            continue
        i += 1

    return {"summary": summary, "issues": issues}


# --------------------------------------------------------------------------- #
# Loader invocation
# --------------------------------------------------------------------------- #
def run_loader(mode: str,
               xlsx_path: Path,
               staging: Staging,
               user_mdb: Path | None = None) -> dict:
    """Run the appropriate loader and return parsed output plus the
    generated result file (or None if dry-run).

    mode: 'dry-run' | 'mdb' | 'lenex'
    """
    env = os.environ.copy()
    env.setdefault("UCANACCESS_DIR", str(REPO_ROOT / "vendor" / "ucanaccess"))

    result_file: Path | None = None

    if mode in ("dry-run", "mdb"):
        # Copy the MDB template (or user-supplied) into the staging dir
        src_mdb = user_mdb if user_mdb else EMPTY_MDB
        out_mdb = staging.dir / "meet.mdb"
        shutil.copy(src_mdb, out_mdb)
        cmd = [sys.executable, str(MDB_LOADER),
               "--xlsx", str(xlsx_path),
               "--mdb", str(out_mdb)]
        if mode == "dry-run":
            cmd.append("--dry-run")
        else:
            result_file = out_mdb
    elif mode == "lenex":
        out_lxf = staging.dir / "meet.lxf"
        cmd = [sys.executable, str(LNX_LOADER),
               "--xlsx", str(xlsx_path),
               "--out", str(out_lxf),
               "--zip"]
        result_file = out_lxf
    else:
        raise ValueError(f"unknown mode: {mode!r}")

    completed = subprocess.run(
        cmd, capture_output=True, text=True, env=env,
        cwd=str(staging.dir), timeout=600)

    combined = (completed.stdout or "") + "\n" + (completed.stderr or "")
    parsed = parse_loader_output(combined)
    parsed["returncode"] = completed.returncode
    parsed["raw_output"] = combined

    # Bundle the result file + an issues.txt into a zip, so the user
    # gets a self-describing archive.
    issues_txt = _render_issues_text(parsed, xlsx_path.name)
    zip_path = staging.result_zip
    with zipfile.ZipFile(zip_path, "w",
                         compression=zipfile.ZIP_DEFLATED) as z:
        if result_file is not None and result_file.exists():
            z.write(result_file, arcname=result_file.name)
        z.writestr("issues.txt", issues_txt)

    parsed["download_name"] = _download_name(mode, xlsx_path.name)
    parsed["download_id"]   = staging.id
    return parsed


def _render_issues_text(parsed: dict, xlsx_name: str) -> str:
    """Plain-text issues report for inclusion in the zip."""
    from datetime import datetime
    lines = [f"Rapport de qualité des données — {xlsx_name}",
             f"Généré : {datetime.now():%Y-%m-%d %H:%M:%S}",
             ""]
    summary = parsed.get("summary", [])
    if summary:
        lines.append("== Sommaire ==")
        lines.extend(summary)
        lines.append("")
    issues = parsed.get("issues", {})
    if not issues:
        lines.append("Aucun problème détecté.")
    else:
        lines.append("== Problèmes détectés ==")
        for cat, data in issues.items():
            lines.append(f"[{data['severity']}] {cat}: {data['count']}")
            for it in data["items"]:
                row = f" (ligne {it['row']})" if it.get("row") else ""
                lines.append(f"    - {it['message']}{row}")
    return "\n".join(lines) + "\n"


def _download_name(mode: str, xlsx_name: str) -> str:
    base = Path(xlsx_name).stem or "meet"
    suffix = {"dry-run": "dry-run", "mdb": "mdb", "lenex": "lenex"}[mode]
    return f"{base}-{suffix}.zip"


# --------------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------------- #
@app.route("/")
def index():
    _gc_stagings()
    return render_template("index.html")


@app.route("/api/run", methods=["POST"])
def api_run():
    _gc_stagings()
    mode = request.form.get("mode", "dry-run")
    if mode not in ("dry-run", "mdb", "lenex"):
        return jsonify({"error": f"mode invalide: {mode!r}"}), 400

    xlsx = request.files.get("xlsx")
    if xlsx is None or not xlsx.filename:
        return jsonify({"error": "Aucun fichier xlsx reçu."}), 400

    mdb_upload = request.files.get("mdb")
    staging = _new_staging()
    try:
        xlsx_path = staging.dir / "input.xlsx"
        xlsx.save(xlsx_path)
        if xlsx_path.stat().st_size > MAX_XLSX_BYTES:
            return jsonify({"error": "Fichier xlsx trop volumineux."}), 413

        user_mdb: Path | None = None
        if mdb_upload and mdb_upload.filename:
            mdb_path = staging.dir / "input.mdb"
            mdb_upload.save(mdb_path)
            if mdb_path.stat().st_size > MAX_MDB_BYTES:
                return jsonify({"error": "Fichier mdb trop volumineux."}), 413
            user_mdb = mdb_path

        parsed = run_loader(mode, xlsx_path, staging, user_mdb=user_mdb)
        return jsonify(parsed)
    except subprocess.TimeoutExpired:
        _drop_staging(staging.id)
        return jsonify({"error": "Dépassement du délai (10 min)."}), 504
    except Exception as e:
        _drop_staging(staging.id)
        app.logger.exception("loader failure")
        return jsonify({"error": f"Erreur interne: {e}"}), 500


@app.route("/api/download/<sid>")
def api_download(sid: str):
    with _stagings_lock:
        s = _stagings.get(sid)
    if s is None or not s.result_zip.exists():
        abort(404)
    name = request.args.get("name") or "result.zip"
    # Flask 2+: send_file supports file path
    resp = send_file(s.result_zip, mimetype="application/zip",
                     as_attachment=True, download_name=name)
    # Schedule cleanup after the response is sent
    @resp.call_on_close
    def _cleanup():
        _drop_staging(sid)
    return resp


@app.route("/healthz")
def health():
    return {"ok": True}


# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
