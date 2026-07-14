#!/bin/bash
# Build CubeSat Builder.app for macOS
# Source code edits take effect immediately - no rebuild needed after code changes

APP_NAME="CubeSat Builder"
APP_DIR="/Applications/${APP_NAME}.app"
SRC_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "Building ${APP_NAME}.app..."

# Remove old app if exists
rm -rf "${APP_DIR}"

# Create app bundle structure
mkdir -p "${APP_DIR}/Contents/MacOS"
mkdir -p "${APP_DIR}/Contents/Resources"

# Create Info.plist
cat > "${APP_DIR}/Contents/Info.plist" << 'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>CubeSat Builder</string>
    <key>CFBundleDisplayName</key>
    <string>CubeSat Builder</string>
    <key>CFBundleIdentifier</key>
    <string>com.cubesat.builder</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>CFBundleExecutable</key>
    <string>launch</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
PLIST

# Create launcher script that activates venv and runs main.py
cat > "${APP_DIR}/Contents/MacOS/launch" << LAUNCHER
#!/bin/bash
cd "${SRC_DIR}"
source .venv/bin/activate
python main.py "\$@"
LAUNCHER

chmod +x "${APP_DIR}/Contents/MacOS/launch"

echo ""
echo "Done! ${APP_NAME}.app installed to /Applications/"
echo ""
echo "IMPORTANT: Source code lives at ${SRC_DIR}"
echo "Edit main.py or CubeSat-Builder.html and changes take effect immediately."
echo "Only re-run this script if you move the source folder."
