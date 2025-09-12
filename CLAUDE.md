# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script project managed with clasp (Command Line Apps Script Projects). The project uses the V8 runtime and is configured for the America/Chicago timezone.

## Architecture

- **Code.js**: Main entry point containing Google Apps Script functions
- **appsscript.json**: Project manifest defining runtime version, timezone, and dependencies
- **.clasp.json**: clasp configuration file linking to Google Apps Script project (ID: 1sHGYREBdPkMk1y_sOHR57ahHaPR8XiiJzS8zU9CawIBHKuyZ3lw6Ji9E)

## Development Commands

### Push code to Google Apps Script
```bash
clasp push
```

### Pull code from Google Apps Script
```bash
clasp pull
```

### Open the script in the web editor
```bash
clasp open
```

### Deploy the script
```bash
clasp deploy
```

### View logs
```bash
clasp logs
```

## File Structure

- JavaScript/Google Apps Script files use `.js` or `.gs` extensions
- HTML files use `.html` extension
- Configuration files use `.json` extension
- The `rootDir` is set to the current directory (no subdirectory structure)

## Runtime Configuration

- Runtime: V8
- Exception logging: STACKDRIVER
- Timezone: America/Chicago
- No external dependencies currently configured