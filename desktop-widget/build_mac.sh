#!/bin/bash
# Run this ON a Mac (PyInstaller doesn't cross-compile).
# Produces dist/StatusUpdate.app — zip it up for teammates along with a
# config.json (copied from config.example.json, filled in) placed next to
# the .app before they double-click it, or inside the .app's Resources —
# simplest is just next to it in the same folder.
set -e
pip3 install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name StatusUpdate status_widget.py

echo
echo "Build done. dist/StatusUpdate.app"
echo "Since this isn't notarized, teammates will need to right-click > Open"
echo "the first time (Gatekeeper will otherwise block an unsigned app)."
