#!/bin/bash
set -e

# Start unoserver as a background daemon (LibreOffice stays resident)
unoserver --daemon --port 2003 &
UNOPID=$!

# Give it a few seconds to initialize
sleep 5

exec gunicorn --bind "0.0.0.0:${PORT:-8000}" --timeout 300 --workers 2 --threads 4 app:app
