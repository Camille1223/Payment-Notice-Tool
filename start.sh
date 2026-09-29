#!/bin/bash
set -e

# Start unoserver as a background daemon (LibreOffice stays resident)
unoserver --daemon --port 2003 &

# Give it a few seconds to initialize
sleep 5

# Single worker + multiple threads: avoids cross-process job-state races
# while still handling concurrent requests.
exec gunicorn --bind "0.0.0.0:${PORT:-8000}" --timeout 300 --workers 1 --threads 8 app:app
