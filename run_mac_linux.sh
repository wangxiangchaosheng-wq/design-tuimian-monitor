#!/usr/bin/env bash
set -euo pipefail
[ -d .venv ] || python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python crawler.py
printf '完成：请打开 docs/index.html\n'
