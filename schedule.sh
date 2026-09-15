#!/usr/bin/env bash
# Installs or removes the weekly launchd schedule for the timesheet script.
#
# Usage:
#   ./schedule.sh install    # runs cal2cat.py every Friday at 16:00
#   ./schedule.sh uninstall  # removes the schedule
#   ./schedule.sh status     # shows whether it's currently loaded

set -euo pipefail

PLIST_NAME="com.sap.timesheet.plist"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC_PLIST="$SCRIPT_DIR/$PLIST_NAME"
DEST_PLIST="$HOME/Library/LaunchAgents/$PLIST_NAME"
LABEL="com.sap.timesheet"

case "${1:-}" in
    install)
        mkdir -p "$HOME/Library/LaunchAgents"
        sed "s#__SCRIPT_DIR__#$SCRIPT_DIR#g" "$SRC_PLIST" > "$DEST_PLIST"
        launchctl unload "$DEST_PLIST" 2>/dev/null || true
        launchctl load "$DEST_PLIST"
        echo "Installed. Will run every Friday at 16:00 (edit $DEST_PLIST and re-run 'install' to change)."
        ;;
    uninstall)
        launchctl unload "$DEST_PLIST" 2>/dev/null || true
        rm -f "$DEST_PLIST"
        echo "Removed the schedule."
        ;;
    status)
        launchctl list | grep "$LABEL" || echo "Not currently scheduled."
        ;;
    *)
        echo "Usage: $0 {install|uninstall|status}"
        exit 1
        ;;
esac
