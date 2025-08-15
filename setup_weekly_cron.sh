#!/bin/bash

# Setup weekly cron job for Kobo to SharePoint transfer
# This script will run the transfer every Sunday at 2:00 AM

echo "Setting up weekly cron job for Kobo to SharePoint transfer..."

# Get the current directory
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_SCRIPT="$SCRIPT_DIR/new.py"

# Check if the Python script exists
if [ ! -f "$PYTHON_SCRIPT" ]; then
    echo "❌ Error: new.py not found in $SCRIPT_DIR"
    exit 1
fi

# Check if .env file exists
if [ ! -f "$SCRIPT_DIR/.env" ]; then
    echo "❌ Error: .env file not found in $SCRIPT_DIR"
    echo "Please ensure your .env file is configured with all required credentials."
    exit 1
fi

# Create the cron job entry (every Sunday at 2:00 AM)
CRON_JOB="0 2 * * 0 cd $SCRIPT_DIR && /usr/bin/python3 $PYTHON_SCRIPT >> $SCRIPT_DIR/transfer.log 2>&1"

# Check if cron job already exists
if crontab -l 2>/dev/null | grep -q "$PYTHON_SCRIPT"; then
    echo "⚠️ Cron job already exists. Removing old entry..."
    crontab -l 2>/dev/null | grep -v "$PYTHON_SCRIPT" | crontab -
fi

# Add the new cron job
(crontab -l 2>/dev/null; echo "$CRON_JOB") | crontab -

echo "✅ Weekly cron job set up successfully!"
echo "📅 The script will run every Sunday at 2:00 AM"
echo "📁 Logs will be saved to: $SCRIPT_DIR/transfer.log"
echo ""
echo "To view the cron job: crontab -l"
echo "To remove the cron job: crontab -r"
echo "To edit the cron job: crontab -e"
