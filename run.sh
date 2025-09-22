#!/bin/bash

# Klarity Template Comparer - Quick Start Script

echo "🚀 Starting Klarity Template Comparer..."
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed. Please install Python 3.7 or higher."
    exit 1
fi

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "❌ pip is not installed. Please install pip."
    exit 1
fi

# Install dependencies if requirements.txt exists
if [ -f "requirements.txt" ]; then
    echo "📦 Installing dependencies..."
    pip3 install -r requirements.txt
    if [ $? -eq 0 ]; then
        echo "✅ Dependencies installed successfully"
    else
        echo "❌ Failed to install dependencies"
        exit 1
    fi
else
    echo "⚠️ requirements.txt not found"
fi

echo ""
echo "🌟 Launching Klarity Template Comparer..."
echo "📖 The app will open in your default browser"
echo "🔄 Press Ctrl+C to stop the server"
echo ""

# Run the Streamlit app
streamlit run app.py