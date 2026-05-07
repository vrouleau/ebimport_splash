"""
ebimport_splash web UI.

Small Flask app that accepts an xlsx upload, runs the MDB loader in
a subprocess (dry-run or write mode), parses the output for the
Summary + Issues sections, and offers the resulting .mdb plus issues
report as a ZIP download.

Stateless: each request gets its own temp dir, cleaned up when the
download is streamed or after `STAGING_TTL_SECS` of inactivity.
"""
from __future__ import annotations

import os
import json
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
COPY_SCRIPT = REPO_ROOT / "copy_prelim_to_masters_final.py"
AUDIT_SCRIPT = REPO_ROOT / "audit_pdf.py"
DEFAULT_MDB = REPO_ROOT / "template.mdb"

BUILD_TIMESTAMP = (REPO_ROOT / "BUILD_TIMESTAMP").read_text().strip() or "dev"

STAGING_DIR = Path(os.environ.get("STAGING_DIR", "/tmp/ebimport_staging"))
STAGING_DIR.mkdir(parents=True, exist_ok=True)
STAGING_TTL_SECS = 10 * 60                # 10 minutes

# Upload size limits (bytes)
MAX_XLSX_BYTES = 50 * 1024 * 1024
MAX_MDB_BYTES  = 200 * 1024 * 1024

# --------------------------------------------------------------------------- #
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_XLSX_BYTES + MAX_MDB_BYTES + 1 * 1024 * 1024


@app.after_request
def _security_headers(resp):
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["Permissions-Policy"] = "geolocation=(), camera=(), microphone=()"
    return resp


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
               user_mdb: Path) -> dict:
    """Run the appropriate loader and return parsed output.

    mode: 'dry-run' | 'mdb' | 'lenex'
    user_mdb must always be provided — it supplies the meet's event
    structure, which the loader treats as authoritative.
    """
    env = os.environ.copy()
    env.setdefault("UCANACCESS_DIR", str(REPO_ROOT / "vendor" / "ucanaccess"))

    result_file: Path | None = None

    if mode in ("dry-run", "mdb"):
        out_mdb = staging.dir / "meet.mdb"
        shutil.copy(user_mdb, out_mdb)
        cmd = [sys.executable, str(MDB_LOADER),
               "--xlsx", str(xlsx_path),
               "--mdb", str(out_mdb)]
        if mode == "dry-run":
            cmd.append("--dry-run")
        else:
            result_file = out_mdb
    elif mode == "lenex":
        out_lxf = staging.dir / "meet.lxf"
        mdb_copy = staging.dir / "template.mdb"
        shutil.copy(user_mdb, mdb_copy)
        cmd = [sys.executable, str(LNX_LOADER),
               "--xlsx", str(xlsx_path),
               "--mdb", str(mdb_copy),
               "--out", str(out_lxf)]
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

    # If the loader aborted with a FATAL (exit 2), parse the list of
    # fatal errors out of the combined output so the UI can highlight
    # them prominently.  No result file is produced.
    parsed["fatal"] = _parse_fatals(combined) if completed.returncode == 2 else []

    # Bundle the result file + an issues.txt into a zip
    issues_txt = _render_issues_text(parsed, xlsx_path.name)
    zip_path = staging.result_zip
    with zipfile.ZipFile(zip_path, "w",
                         compression=zipfile.ZIP_DEFLATED) as z:
        if result_file is not None and result_file.exists():
            z.write(result_file, arcname=result_file.name)
        # Include the template mdb in lenex mode (as meet.mdb for consistency)
        if mode == "lenex":
            mdb_in_staging = staging.dir / "template.mdb"
            if mdb_in_staging.exists():
                z.write(mdb_in_staging, arcname="meet.mdb")
        # Include masters_transfer.vbs + .bat for MDB and Lenex modes
        if mode in ("mdb", "lenex"):
            vbs_path = REPO_ROOT / "masters_transfer.vbs"
            bat_path = REPO_ROOT / "masters_transfer.bat"
            if vbs_path.exists():
                z.write(vbs_path, arcname="masters_transfer.vbs")
            if bat_path.exists():
                z.write(bat_path, arcname="masters_transfer.bat")
        z.writestr("issues.txt", issues_txt)

    parsed["download_name"] = _download_name(mode, xlsx_path.name)
    parsed["download_id"]   = staging.id

    # Remove intermediate files — only keep the result zip
    for f in staging.dir.iterdir():
        if f != staging.result_zip:
            f.unlink(missing_ok=True)

    return parsed


def _parse_fatals(text: str) -> list[str]:
    """Pull the bullet list out of a FATAL: block in loader stdout."""
    out: list[str] = []
    in_fatal = False
    for line in text.splitlines():
        if "FATAL" in line and "template/xlsx" in line:
            in_fatal = True
            continue
        if in_fatal:
            if _BAR.match(line):
                # consecutive === lines are fine; end when we hit a non-bar,
                # non-bullet line
                continue
            m = re.match(r"^\s*-\s+(.*?)$", line)
            if m:
                out.append(m.group(1).rstrip())
            elif line.strip().startswith("fatal error"):
                break
    return out

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
    return render_template("index.html", build_timestamp=BUILD_TIMESTAMP)


@app.route("/api/run", methods=["POST"])
def api_run():
    _gc_stagings()
    mode = request.form.get("mode", "dry-run")
    if mode not in ("dry-run", "mdb", "lenex"):
        return jsonify({"error": f"mode invalide: {mode!r}"}), 400

    xlsx = request.files.get("xlsx")
    if xlsx is None or not xlsx.filename:
        return jsonify({"error": "Aucun fichier xlsx reçu."}), 400

    # MDB template — use uploaded file or fall back to bundled default.
    mdb_upload = request.files.get("mdb")
    user_mdb_path: Path | None = None

    staging = _new_staging()
    try:
        xlsx_path = staging.dir / "input.xlsx"
        xlsx.save(xlsx_path)
        if xlsx_path.stat().st_size > MAX_XLSX_BYTES:
            return jsonify({"error": "Fichier xlsx trop volumineux."}), 413

        if mdb_upload and mdb_upload.filename:
            mdb_path = staging.dir / "input.mdb"
            mdb_upload.save(mdb_path)
            if mdb_path.stat().st_size > MAX_MDB_BYTES:
                return jsonify({"error": "Fichier mdb trop volumineux."}), 413
            user_mdb_path = mdb_path
        else:
            user_mdb_path = DEFAULT_MDB

        parsed = run_loader(mode, xlsx_path, staging, user_mdb=user_mdb_path)
        # Remove uploaded files immediately — only keep the result zip
        xlsx_path.unlink(missing_ok=True)
        if user_mdb_path and user_mdb_path != DEFAULT_MDB:
            user_mdb_path.unlink(missing_ok=True)
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


@app.route("/api/copy-masters", methods=["POST"])
def api_copy_masters():
    """Run copy_prelim_to_masters_final.py on an uploaded .mdb."""
    _gc_stagings()
    mdb_upload = request.files.get("mdb")
    if mdb_upload is None or not mdb_upload.filename:
        return jsonify({"error": "Un fichier .mdb est requis."}), 400

    dry_run = request.form.get("dry_run", "false").lower() in ("1", "true", "yes")

    staging = _new_staging()
    try:
        mdb_path = staging.dir / "meet.mdb"
        mdb_upload.save(mdb_path)
        if mdb_path.stat().st_size > MAX_MDB_BYTES:
            return jsonify({"error": "Fichier mdb trop volumineux."}), 413

        env = os.environ.copy()
        env.setdefault("UCANACCESS_DIR", str(REPO_ROOT / "vendor" / "ucanaccess"))
        cmd = [sys.executable, str(COPY_SCRIPT), "--mdb", str(mdb_path)]
        if dry_run:
            cmd.append("--dry-run")

        completed = subprocess.run(
            cmd, capture_output=True, text=True, env=env, timeout=300)
        output = (completed.stdout or "") + "\n" + (completed.stderr or "")

        result = {
            "returncode": completed.returncode,
            "output": output.strip(),
            "dry_run": dry_run,
        }

        # Offer the modified mdb as download (unless dry-run)
        if not dry_run and completed.returncode == 0:
            zip_path = staging.result_zip
            with zipfile.ZipFile(zip_path, "w",
                                 compression=zipfile.ZIP_DEFLATED) as z:
                z.write(mdb_path, arcname="meet.mdb")
            result["download_id"] = staging.id

        return jsonify(result)
    except subprocess.TimeoutExpired:
        _drop_staging(staging.id)
        return jsonify({"error": "Dépassement du délai (5 min)."}), 504
    except Exception as e:
        _drop_staging(staging.id)
        app.logger.exception("copy-masters failure")
        return jsonify({"error": f"Erreur interne: {e}"}), 500


@app.route("/api/audit", methods=["POST"])
def api_audit():
    """Run PDF audit: compare SPLASH heat-sheet PDF against xlsx."""
    _gc_stagings()
    pdf = request.files.get("pdf")
    xlsx = request.files.get("xlsx")
    if not pdf or not pdf.filename:
        return jsonify({"error": "Un fichier PDF est requis."}), 400
    if not xlsx or not xlsx.filename:
        return jsonify({"error": "Un fichier xlsx est requis."}), 400

    staging = _new_staging()
    try:
        pdf_path = staging.dir / "heats.pdf"
        xlsx_path = staging.dir / "input.xlsx"
        pdf.save(pdf_path)
        xlsx.save(xlsx_path)

        completed = subprocess.run(
            [sys.executable, str(AUDIT_SCRIPT),
             "--pdf", str(pdf_path), "--xlsx", str(xlsx_path), "--json"],
            capture_output=True, text=True, timeout=120)

        # Clean up immediately
        pdf_path.unlink(missing_ok=True)
        xlsx_path.unlink(missing_ok=True)
        _drop_staging(staging.id)

        if completed.returncode != 0:
            return jsonify({"error": completed.stderr or "Audit failed"}), 500

        return jsonify(json.loads(completed.stdout))
    except subprocess.TimeoutExpired:
        _drop_staging(staging.id)
        return jsonify({"error": "Dépassement du délai."}), 504
    except Exception as e:
        _drop_staging(staging.id)
        app.logger.exception("audit failure")
        return jsonify({"error": f"Erreur interne: {e}"}), 500


@app.route("/healthz")
def health():
    return {"ok": True}


# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
