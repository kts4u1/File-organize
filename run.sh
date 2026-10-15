#!/usr/bin/env sh
cd "$(dirname "$0")"
[ -f .venv/bin/activate ] && . .venv/bin/activate
exec python3 -m streamlit run app.py
