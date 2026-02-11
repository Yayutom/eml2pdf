#!/bin/bash
cd "$(dirname "$0")"
source venv/bin/activate
python3 eml2pdf.py
