#!/bin/bash

# Jarvis Voice Assistant Runner
# This script activates the virtual environment and runs the voice assistant

echo "🤖 Starting Jarvis Voice Assistant..."
echo "📁 Activating virtual environment..."

# Activate virtual environment
source jarvis_env/bin/activate

echo "🚀 Launching Jarvis..."
echo "💡 Note: You may need to configure your Wolfram Alpha App ID in Jarvis/config.py"

# Run the main application
python3 main.py

echo "👋 Jarvis has been stopped."