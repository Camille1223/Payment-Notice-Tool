#!/bin/bash
set -e

# Single worker + multiple threads: handles concurrent requests without
# cross-process job-state races (job state is file-based in JOB_DIR).
exec gunicorn --bind "0.0.0.0:${PORT:-8000}" --timeout 300 --workers 1 --threads 8 app:app
