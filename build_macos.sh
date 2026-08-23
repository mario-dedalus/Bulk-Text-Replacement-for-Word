#!/usr/bin/env bash
set -e

# =============================================================
#  build_macos.sh — One-click macOS build and packaging script
#  Outputs: dist/WordTextReplacer.app        (macOS App Bundle)
#           dist/WordTextReplacer_macos.zip  (Release zip archive)
#           dist/WordTextReplacer_mac.dmg    (Optional DMG installer)
#
#  Usage:
#    ./build_macos.sh
#    ./build_macos.sh --version 2.0.0
# =============================================================

APP_NAME="WordTextReplacer"
VERSION=""

while [[ $# -gt 0 ]]; do
    case "$1" in
        --version)
            VERSION="$2"
            shift 2
            ;;
        *)
            shift
            ;;
    esac
done

echo "============================================================"
echo "  Building $APP_NAME for macOS"
if [ -n "$VERSION" ]; then
    echo "  Version: $VERSION"
fi
echo "============================================================"

# ------ 1. Check Python environment ------
echo "[1/5] Checking virtual environment..."
PYTHON_BIN=".venv/bin/python"
if [ ! -f "$PYTHON_BIN" ]; then
    echo "      Creating virtual environment (.venv)..."
    python3 -m venv .venv
    PYTHON_BIN=".venv/bin/python"
fi

# ------ 2. Install dependencies ------
echo "[2/5] Installing dependencies..."
"$PYTHON_BIN" -m pip install --quiet -r requirements.txt

# ------ 3. Run unit tests ------
echo "[3/5] Running tests..."
"$PYTHON_BIN" -m unittest discover -s tests -v

# ------ 4. Build .app with PyInstaller ------
echo "[4/5] Building macOS .app bundle with PyInstaller..."
mkdir -p build/pyi_config
export PYINSTALLER_CONFIG_DIR="$PWD/build/pyi_config"
.venv/bin/pyinstaller --clean --noconfirm WordTextReplacer_mac.spec

# ------ 5. Package into ZIP and DMG ------
echo "[5/5] Creating release distribution packages..."

if [ -n "$VERSION" ]; then
    ZIP_NAME="${APP_NAME}_v${VERSION}_macos.zip"
    DMG_NAME="${APP_NAME}_v${VERSION}_mac.dmg"
else
    ZIP_NAME="${APP_NAME}_macos.zip"
    DMG_NAME="${APP_NAME}_mac.dmg"
fi

# Create ZIP with ditto (preserves macOS app metadata and permissions)
rm -f "dist/$ZIP_NAME"
ditto -c -k --keepParent "dist/$APP_NAME.app" "dist/$ZIP_NAME"
echo "      Created ZIP archive: dist/$ZIP_NAME"

# Attempt DMG creation with hdiutil
DMG_DIR="build/dmg_pack"
rm -rf "$DMG_DIR" "dist/$DMG_NAME"
mkdir -p "$DMG_DIR"
cp -R "dist/$APP_NAME.app" "$DMG_DIR/"
ln -s /Applications "$DMG_DIR/Applications" 2>/dev/null || true

if hdiutil create -volname "$APP_NAME" -srcfolder "$DMG_DIR" -ov -format UDZO "dist/$DMG_NAME" 2>/dev/null; then
    echo "      Created DMG disk image: dist/$DMG_NAME"
fi
rm -rf "$DMG_DIR"

echo ""
echo "============================================================"
echo "  Build completed successfully!"
echo ""
echo "  App Bundle     : dist/$APP_NAME.app"
echo "  Release ZIP    : dist/$ZIP_NAME"
if [ -f "dist/$DMG_NAME" ]; then
    echo "  DMG Installer  : dist/$DMG_NAME"
fi
echo ""
echo "  To run: open dist/$APP_NAME.app"
echo "============================================================"
