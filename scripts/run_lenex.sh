#!/usr/bin/env bash
# Entry point for the Lenex generator (no Java required).
# Usage:
#   ./run_lenex.sh --xlsx <file.xlsx> --out <file.lef|file.lxf> [--zip]
set -euo pipefail

# shellcheck disable=SC1091
source "$(dirname "${BASH_SOURCE[0]}")/_bootstrap.sh"

exec python "$PKG_ROOT/load_to_lenex.py" "$@"
