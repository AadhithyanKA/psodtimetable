#!/bin/bash
# Change to the directory of this script
cd "$(dirname "$0")"

# Print a nice startup message
echo "==========================================="
echo "  Starting PSOD Timetable Dev Server..."
echo "  The site will open automatically in your browser."
echo "  Close this terminal window to stop the server."
echo "==========================================="

# Run the development server and open the browser
npm run dev -- --open
