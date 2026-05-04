#!/usr/bin/env bash
# Entry point for the MDB loader.
# Usage:
#   ./run_mdb.sh --xlsx <file.xlsx> --mdb <file.mdb> [--dry-run] [--wipe]
set -euo pipefail

NEED_JAVA=1
# shellcheck disable=SC1091
source "$(dirname "${BASH_SOURCE[0]}")/_bootstrap.sh"

exec python "$PKG_ROOT/load_to_mdb.py" "$@"
