#!/usr/bin/env bash
# Render build script

set -o errexit

pip install --upgrade pip
pip install -r requirements.txt
pip install gunicorn eventlet

# Create required directories
mkdir -p uploads extracted generated templates
