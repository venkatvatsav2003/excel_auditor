#!/usr/bin/env bash
set -euo pipefail

VERSION="2.0.0"

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; CYAN='\033[0;36m'; NC='\033[0m'

usage() {
    cat <<EOF
Data Auditor v$VERSION — CSV/Excel Data Quality Analysis

Usage: $0 [files...] [options]
   or: $0 --recursive
   or: find . -name "*.csv" | $0

Options:
  -c, --config FILE     Config file (default: config/profiles.yml)
  -o, --output DIR      Report output directory (default: reports/)
  -f, --format FORMAT    Output format: json, html, both (default: both)
  -r, --recursive       Scan directory recursively
  -p, --pattern PATTERN  File pattern for recursive (default: *.csv)
  --json                Print JSON to stdout
  -h, --help            Show this help

Examples:
  $0 data.csv
  $0 data1.csv data2.csv --json
  $0 -r -p "*.csv"
  ls *.csv | $0
EOF
    exit 0
}

log_info()  { echo -e "${CYAN}[*]${NC} $1"; }
log_ok()    { echo -e "${GREEN}[+]${NC} $1"; }

check_deps() {
    python3 -c "import yaml" 2>/dev/null || {
        log_info "Installing yaml..."
        pip3 install pyyaml -q
    }
}

ARGS=()
while [[ $# -gt 0 ]]; do
    case "$1" in
        -h|--help) usage ;;
        -c|--config) ARGS+=("-c" "$2"); shift 2 ;;
        -o|--output) ARGS+=("-o" "$2"); shift 2 ;;
        -f|--format) ARGS+=("-f" "$2"); shift 2 ;;
        -r|--recursive) ARGS+=("-r"); shift ;;
        -p|--pattern) ARGS+=("-p" "$2"); shift 2 ;;
        --json) ARGS+=("--json"); shift ;;
        *) ARGS+=("$1"); shift ;;
    esac
done

check_deps

mkdir -p reports 2>/dev/null || true
log_info "Data Auditor v$VERSION"
python3 auditor.py "${ARGS[@]}"
