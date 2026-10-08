#!/usr/bin/env bash
# Start the app on http://127.0.0.1:8000. Use HOST=0.0.0.0 to open it to the local network.
cd "$(dirname "$0")"
[ -d .venv ] || { python3 -m venv .venv && .venv/bin/pip install -r requirements.txt; }
exec .venv/bin/uvicorn app.main:app --host "${HOST:-127.0.0.1}" --port "${PORT:-8000}"
