#!/usr/bin/env bash
# Common bootstrap sourced by run_mdb.sh / run_lenex.sh.
# - Verifies Python 3.10+ and Java 8+ are available.
# - Creates a local .venv in the package root on first run, installing
#   bundled wheels with --no-index so no network access is needed.
# - Exports UCANACCESS_DIR pointing at the vendored jars.

set -euo pipefail

# Resolve package root (directory containing this file's parent 'scripts/').
_bootstrap_dir=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
PKG_ROOT=$(cd "$_bootstrap_dir/.." && pwd)

VENV_DIR="$PKG_ROOT/.venv"
WHEELS_DIR="$PKG_ROOT/vendor/wheels"
export UCANACCESS_DIR="$PKG_ROOT/vendor/ucanaccess"

# ---- Python check -----------------------------------------------------------
PYTHON_BIN="${PYTHON:-python3}"
if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
    echo "ERROR: Python 3 not found on PATH (looked for '$PYTHON_BIN')." >&2
    echo "       Install Python 3.10 or newer, then re-run." >&2
    exit 1
fi
py_ver=$("$PYTHON_BIN" -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
py_major=${py_ver%%.*}
py_minor=${py_ver##*.}
if [ "$py_major" -lt 3 ] || { [ "$py_major" -eq 3 ] && [ "$py_minor" -lt 10 ]; }; then
    echo "ERROR: Python $py_ver found; need Python 3.10 or newer." >&2
    exit 1
fi

# ---- Java check -------------------------------------------------------------
# Only enforced by run_mdb.sh (Lenex doesn't need Java) — sourced callers
# opt in via NEED_JAVA=1.
if [ "${NEED_JAVA:-0}" = "1" ]; then
    if ! command -v java >/dev/null 2>&1; then
        echo "ERROR: Java not found on PATH." >&2
        echo "       The MDB loader needs Java 8+ (for UCanAccess)." >&2
        exit 1
    fi
fi

# ---- Venv bootstrap (first run only) ---------------------------------------
if [ ! -d "$VENV_DIR" ]; then
    echo ">> first run: creating local venv at $VENV_DIR"
    "$PYTHON_BIN" -m venv "$VENV_DIR"
    # shellcheck disable=SC1091
    source "$VENV_DIR/bin/activate"
    python -m pip install --quiet --upgrade pip
    echo ">> installing bundled wheels from $WHEELS_DIR"
    python -m pip install --quiet --no-index --find-links "$WHEELS_DIR" \
        openpyxl jaydebeapi JPype1
    deactivate
fi

# Make the venv active for the caller's exec.
# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"
