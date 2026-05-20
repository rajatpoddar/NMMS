#!/bin/bash
# ============================================
# NMMS Tracker - macOS Installer Builder
# ============================================
# Builds a standalone macOS .app of the
# desktop GUI app using PyInstaller.
# Run this script locally on your Mac:
#   chmod +x build-mac.sh && ./build-mac.sh
# ============================================

set -e

echo ""
echo "============================================"
echo "  NMMS Tracker - macOS Build Script"
echo "============================================"
echo ""

# ---- Check Python ----
if ! command -v python3 &>/dev/null; then
    echo "❌ python3 not found. Please install Python 3.8+ from https://www.python.org/"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1)
echo "✅ Detected: $PYTHON_VERSION"

# ---- Check Chrome (required by Selenium) ----
if [ -d "/Applications/Google Chrome.app" ] || command -v google-chrome &>/dev/null; then
    echo "✅ Google Chrome found."
else
    echo "⚠️  Google Chrome not detected in /Applications."
    echo "   Selenium requires Chrome. Install from https://www.google.com/chrome/"
fi

# ---- Create virtual environment (optional but recommended) ----
VENV_DIR="build-venv"
if [ ! -d "$VENV_DIR" ]; then
    echo "📦 Creating virtual environment..."
    python3 -m venv "$VENV_DIR"
fi
source "$VENV_DIR/bin/activate"

# ---- Install dependencies ----
echo "📥 Installing client dependencies..."
pip install --upgrade pip --quiet
pip install --no-cache-dir -r requirements-client.txt --quiet
pip install pyinstaller>=6.10 --quiet
echo "✅ Dependencies installed."

# ---- Verify imports ----
echo "🔍 Verifying imports..."
python3 -c "
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
print('Checking imports...')
import customtkinter; print('[OK] customtkinter')
import requests; print('[OK] requests')
from bs4 import BeautifulSoup; print('[OK] bs4')
from selenium import webdriver; print('[OK] selenium')
from selenium.webdriver.chrome.service import Service; print('[OK] selenium.webdriver.chrome.service')
from selenium.webdriver.chrome.options import Options; print('[OK] selenium.webdriver.chrome.options')
from selenium.webdriver.common.by import By; print('[OK] selenium.webdriver.common.by')
from selenium.webdriver.support.ui import WebDriverWait, Select; print('[OK] selenium.webdriver.support.ui')
from selenium.webdriver.support import expected_conditions as EC; print('[OK] selenium.webdriver.support.expected_conditions')
from webdriver_manager.chrome import ChromeDriverManager; print('[OK] webdriver_manager.chrome')
import pandas as pd; print('[OK] pandas')
from openpyxl import Workbook; print('[OK] openpyxl')
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side; print('[OK] openpyxl.styles')
from openpyxl.drawing.image import Image as OpenpyxlImage; print('[OK] openpyxl.drawing.image')
from io import BytesIO; print('[OK] io.BytesIO')
from PIL import Image as PILImage; print('[OK] PIL.Image')
import getmac; print('[OK] getmac')
print('All imports successful!')
"
echo ""

# ---- Build with PyInstaller ----
echo "🔨 Building macOS app with PyInstaller..."
echo "    This may take a few minutes..."
echo ""

pyinstaller --onedir \
    --windowed \
    --name "NMMS_Tracker" \
    --hidden-import customtkinter \
    --hidden-import PIL._tkinter_finder \
    --hidden-import selenium \
    --hidden-import webdriver_manager \
    --hidden-import webdriver_manager.chrome \
    --hidden-import openpyxl \
    --hidden-import pandas \
    --hidden-import bs4 \
    --hidden-import getmac \
    --hidden-import jinja2 \
    app.py

# To add a custom app icon, uncomment and adjust:
# pyinstaller --onedir --windowed --name "NMMS_Tracker" --icon "app.icns" [...] app.py

echo ""
echo "============================================"
echo "  ✅ Build Complete!"
echo "============================================"

# ---- Show output ----
if [ -d "dist/NMMS_Tracker.app" ]; then
    SIZE_MB=$(du -sh "dist/NMMS_Tracker.app" | cut -f1)
    echo "📁 Output: dist/NMMS_Tracker.app"
    echo "📏 Size:   $SIZE_MB"
    echo ""
    echo "🚀 To run:"
    echo "   open dist/NMMS_Tracker.app"
    echo ""
    echo "💡 Tip: Copy 'dist/NMMS_Tracker.app' to your Applications folder"
    echo "   or run it from terminal: open dist/NMMS_Tracker.app"
elif [ -f "dist/NMMS_Tracker" ]; then
    SIZE_MB=$(du -h "dist/NMMS_Tracker" | cut -f1)
    echo "📁 Output: dist/NMMS_Tracker"
    echo "📏 Size:   $SIZE_MB"
    echo ""
    echo "🚀 To run from terminal:"
    echo "   ./dist/NMMS_Tracker"
else
    echo "⚠️  Could not find build output in dist/"
    ls -la dist/
fi

# ---- Cleanup ----
echo ""
echo "🧹 Cleaning up build artifacts..."
rm -rf build/ *.spec
echo "✅ Done."
