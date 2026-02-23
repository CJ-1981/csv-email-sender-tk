#!/bin/bash
# CSV Email Sender - Launch Script (macOS/Linux)

# Get script directory
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if venv exists
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
    echo "Installing dependencies..."
    ./venv/bin/pip install -r requirements.txt
fi

# Check if tkinter is available
if ! ./venv/bin/python -c "import tkinter" 2>/dev/null; then
    echo "ERROR: Tkinter is not available in this Python installation."
    echo ""
    echo "To fix this on macOS with Homebrew:"
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "  brew install python-tk@3.13"
        echo "  brew unlink python@3.13"
        echo "  brew link python-tk@3.13 --force"
    else
        echo "  sudo apt-get install python3-tk"
    fi
    echo ""
    echo "Then delete the venv folder and run this script again."
    exit 1
fi

# Activate venv and run
echo "Starting CSV Email Sender..."
./venv/bin/python main.py
