# Gmail Mailer Pro - Build Instructions (UPDATED)

## Problem Fixed
Playwright browsers (chromium) were not being found in the standalone .exe, causing PDF/Image attachment conversion to fail.

## Solution Implemented

### 1. **Updated PyInstaller Spec** (`build_app.spec`)
- Added `collect_submodules('playwright')` to include all Playwright modules
- Added hidden imports for Playwright implementations
- Included PIL and PPTX data files for image handling

### 2. **Updated main.py**
- Added `setup_playwright_env()` function to set `PLAYWRIGHT_BROWSERS_PATH` environment variable
- This tells Playwright where to cache/find browsers in temp directory
- Updated `install_browsers()` to work in frozen (bundled) applications
- Changed main block to always install browsers on startup

## How to Build

### Prerequisites
```bash
# Make sure you're in the project directory and venv is activated
cd /Users/akshitgupta/Desktop/Gmail-Mailer
source venv/bin/activate
```

### Step 1: Install Playwright Browsers (Important!)
```bash
python -m playwright install chromium
```

### Step 2: Build with PyInstaller
```bash
pyinstaller build_app.spec
```

### Step 3: Test the .exe
The .exe will be in `dist/GmailMailerPro.exe`

## How It Works on Windows

1. **First Run**: App checks if browsers are installed. If not, installs them to:
   - `C:\Users\{username}\AppData\Local\Temp\playwright_browsers\`

2. **Browser Path**: The app sets environment variable:
   ```python
   PLAYWRIGHT_BROWSERS_PATH = C:\Users\{username}\AppData\Local\Temp\playwright_browsers\
   ```

3. **Attachment Handling**: Temporary attachment files go to:
   - `C:\Users\{username}\AppData\Local\Temp\gmail_mailer_attachments\`

## Troubleshooting

### If still getting browser errors:
1. **Manual Installation** (on Windows machine):
   ```bash
   # After extracting the .exe, open PowerShell and run:
   python -m playwright install chromium --with-deps
   ```

2. **Check if Playwright is installed**:
   - Look for: `C:\Users\{username}\AppData\Local\Temp\playwright_browsers\chromium-1217`
   - Should contain `chrome-headless-shell.exe`

3. **Enable Console for Debugging**:
   - Edit `build_app.spec`, change `console=False` to `console=True`
   - Rebuild to see detailed error messages

## Key Changes Summary

| File | Change | Reason |
|------|--------|--------|
| `main.py` | Added `setup_playwright_env()` | Set correct browser path for .exe |
| `main.py` | Updated `install_browsers()` | Make it work in frozen apps |
| `build_app.spec` | Added `collect_submodules('playwright')` | Bundle all Playwright code |
| `build_app.spec` | Added PIL and PPTX data | Include all image conversion libs |

## Expected Behavior After Fix

✅ Standalone .exe launches normally  
✅ Attachments (PDF, PNG, PPTX, etc.) convert successfully  
✅ Emails send with attachments attached  
✅ No "Executable doesn't exist" errors  

## Files Modified
- `/Users/akshitgupta/Desktop/Gmail-Mailer/main.py` ✓
- `/Users/akshitgupta/Desktop/Gmail-Mailer/build_app.spec` ✓
