#!/usr/bin/env bash
# Cognito AI — Linux Launcher
# Mangalore Refinery and Petrochemicals Limited (MRPL) | SIH26117

set -e

echo "====================================================="
echo "⚡ Starting Cognito AI (Linux Air-Gapped)"
echo "====================================================="

# Check Python 3
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is required. Please install python3."
    exit 1
fi

# Check / install dependencies if needed
python3 -c "import docx, openpyxl, webview" 2>/dev/null || {
    echo "📦 Installing required lightweight Python libraries..."
    python3 -m pip install --quiet python-docx openpyxl pywebview requests || true
}

# Run the platform
python3 main.py
