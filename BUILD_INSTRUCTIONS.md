# Gmail Mailer Pro - Build Instructions (v2)

## Problem
When running as a standalone `.exe`, Playwright's chromium browser was not found because:
1. `sys.executable` in a PyInstaller bundle points to the `.exe` itself, NOT Python
2. So `subprocess.run([sys.executable, "-m", "playwright", "install", ...])` silently fails
3. Browsers were being cached in `%TEMP%` which gets cleaned up by Windows

## Solution (v2)

### Key Changes

| What | Old (Broken) | New (Fixed) |
|------|-------------|-------------|
| **Browser Install** | `sys.executable -m playwright install` | Uses `playwright._impl._driver.compute_driver_executable()` to call the Playwright CLI directly |
| **Browser Path** | `%TEMP%\playwright_browsers` (gets cleaned) | `%LOCALAPPDATA%\GmailMailerPro\playwright_browsers` (persistent) |
| **Driver Binary** | Not bundled | Bundled via `pw_driver_binaries` in spec |
| **First Run** | Silent failure | Shows info dialog + installs Chromium |
| **Pre-launch Check** | None | `_ensure_browser()` check before every conversion |
| **Console** | `False` | `True` for debugging (change to `False` when confirmed) |

### How It Works Now

1. **On startup**: App checks if Chromium is already installed at the persistent path
2. **If not found (first run)**: Shows a dialog telling user "downloading browser components", then installs using Playwright's own driver executable (works inside .exe!)
3. **Before each conversion**: `_ensure_browser()` double-checks and auto-installs if needed
4. **Browsers persist** in `%LOCALAPPDATA%\GmailMailerPro\playwright_browsers\` — won't get deleted by Windows cleanup

## Build Steps

### Prerequisites
```bash
# Activate your virtual environment
cd /path/to/Gmail-Mailer
source venv/bin/activate   # Mac/Linux
# OR
.\venv\Scripts\activate     # Windows
```

### Step 1: Install Playwright Browsers (on build machine)
```bash
python -m playwright install chromium
```

### Step 2: Build
```bash
pyinstaller build_app.spec
```

### Step 3: Test
- The `.exe` will be in `dist/GmailMailerPro.exe`
- **First run** will show a setup dialog and download Chromium (~150MB)
- Subsequent runs will start instantly
- Console window will show debug output (change `console=True` to `False` in spec when confirmed working)

## Troubleshooting

### If Chromium download fails inside .exe:
1. **Check internet connection** - the .exe needs internet on first run to download Chromium
2. **Check antivirus** - some AV software blocks downloading executables
3. **Manual install** - Open PowerShell and run:
   ```powershell
   $env:PLAYWRIGHT_BROWSERS_PATH = "$env:LOCALAPPDATA\GmailMailerPro\playwright_browsers"
   npx playwright install chromium
   ```

### If you see "compute_driver_executable" errors:
- The Playwright driver binary wasn't bundled correctly
- Verify the `pw_driver_binaries` section in `build_app.spec` found files
- Check that `playwright/driver/` directory exists in your venv's site-packages

### Console window shows up:
- This is intentional for debugging. Once everything works:
- Edit `build_app.spec` → change `console=True` to `console=False` → rebuild

## Files Modified
- `main.py` - Fixed browser install logic + persistent path + pre-launch checks
- `build_app.spec` - Added Playwright driver binaries + driver hidden import + console=True
