# Bug Report: Rox-Shooter v1.1

## Critical Bugs

### 1. **Missing Function Definitions** ⛔ CRITICAL
**Location:** Lines 1518-1519, 1935, 1943
**Severity:** CRASH
**Description:** Four functions are called but never defined:
- `generate_scrambled_html_invoice()`
- `generate_bidi_email_body()`
- `generate_vectorized_pdf()`
- `generate_html_mosaic()`

These are invoked in `preview_all_task()` and `generate_vec_preview_all()` / `generate_mosaic_preview_all()` methods.

**Impact:** Application will crash with `NameError` when:
- User clicks "Preview All" button
- Preview vector or mosaic formats

**Fix:** Either implement these functions or remove references to them.

---

## High-Priority Bugs

### 2. **Missing Module Import: `math`**
**Location:** Line 1231 (used in `arrange_windows_split()`)
**Severity:** CRASH
**Description:** `math.ceil()` and `math.sqrt()` are used but `math` is not imported at the top.

```python
# Line 1231
cols = math.ceil(math.sqrt(n))  # NameError: name 'math' is not defined
```

**Fix:** Add `import math` at the top of the file.

---

### 3. **Missing Module Import: `pandas`**
**Location:** Line 1404 (in `handle_load_file()`)
**Severity:** CRASH
**Description:** `pd.read_csv()` and `pd.read_excel()` are called but pandas is not imported.

```python
# Line 1404
df = pd.read_csv(file_path)  # NameError: name 'pd' is not defined
```

**Fix:** Add `import pandas as pd` at the top (or import locally in function).

---

### 4. **Potential Race Condition: Contexts Dict Access**
**Location:** Multiple locations (e.g., line 1629 in `blasting_engine_task()`)
**Severity:** HIGH - Data corruption/crashes under load
**Description:** `self.contexts` dictionary is accessed from multiple threads (UI thread + async background loop) without locks.

```python
# Thread-unsafe
active_ids = list(self.contexts.keys())  # Can change while iterating
window_id = active_ids[i % len(active_ids)]
page = self.contexts[window_id]["page"]  # window_id may have been removed
```

**Impact:** RuntimeError when window is closed during blasting.

**Fix:** Use `threading.Lock()` to protect dictionary access.

---

### 5. **Async/Await Context Mismatch in `apply_dynamic_personalization()`**
**Location:** Line 1609
**Severity:** HIGH - Function is synchronous but called in async context
**Description:** The nested function `apply_dynamic_personalization()` is defined as synchronous but should be async or properly handled.

```python
# Line 1629
parsed_html = apply_dynamic_personalization(parsed_html)  # Fine, but could cause issues
```

**Impact:** While it works, it can block the event loop.

**Fix:** Either move it outside the async function or mark it explicitly for sync context.

---

## Medium-Priority Bugs

### 6. **File Not Found Error in `obfuscate_and_attach()`**
**Location:** Line 1623
**Severity:** MEDIUM - Silent failure
**Description:** The function modifies a file by adding padding, but doesn't validate that write operations complete.

```python
with open(file_path, 'rb') as f:
    data = f.read()
# ... mutate_data ...
with open(file_path, 'wb') as f:
    f.write(mutated_data)  # Could fail silently
```

**Fix:** Add error handling and validation.

---

### 7. **Unhandled Exception: `asyncio.to_thread()` Not Available in Python < 3.9**
**Location:** Lines 1518-1519
**Severity:** MEDIUM - Version incompatibility
**Description:** `asyncio.to_thread()` was added in Python 3.9. Older versions will crash.

```python
("Scrambled HTML", asyncio.to_thread(generate_scrambled_html_invoice, ...))
```

**Fix:** Check Python version or use `loop.run_in_executor()` instead.

---

### 8. **Missing Error Handling: Gmail Automation Selectors**
**Location:** Lines 1737-1813 (`automate_gmail_send()`)
**Severity:** MEDIUM - Fragile UI automation
**Description:** Multiple selectors assume specific Gmail DOM structure that may change:

```python
to_input = page.locator('input[aria-label="To recipients"], ...').last  # May not exist
subject_input = page.locator('input[name="subjectbox"], ...').last  # May fail
```

**Impact:** Emails won't send if Gmail updates their UI.

**Fix:** Add comprehensive fallbacks and validation.

---

### 9. **Platform Detection Bug in `preview_attachment_task()`**
**Location:** Lines 1506-1514
**Severity:** MEDIUM - Platform-specific crashes
**Description:** Missing `subprocess` usage without proper error handling on unsupported platforms.

```python
if sys.platform == "darwin":
    subprocess.run(["open", path])
elif sys.platform == "win32":
    os.startfile(path)  # Not defined on macOS/Linux!
else:
    subprocess.run(["xdg-open", path])
```

**Fix:** Wrap in try-except or validate platform first.

---

## Low-Priority Bugs

### 10. **Potential Resource Leak: Playwright Contexts**
**Location:** `launch_browser_task()` and `close_window_task()`
**Severity:** LOW - Memory leak over time
**Description:** If `close_window_task()` fails, context objects are never cleaned up.

```python
async def close_window_task(self, window_id):
    # If exception occurs before cleanup, pw/context aren't properly closed
    await data["context"].close()  # Could raise exception
    if window_id in self.contexts:
        self.contexts.pop(window_id)  # May not execute
```

**Fix:** Use try-finally or context managers.

---

### 11. **Deprecated Method: `os.startfile()` on Windows**
**Location:** Multiple locations
**Severity:** LOW - Future compatibility
**Description:** `os.startfile()` is Windows-specific and may be deprecated.

**Fix:** Use `subprocess.Popen()` with cross-platform support instead.

---

### 12. **Integer Division Without Type Check**
**Location:** Line 1667
**Severity:** LOW - Edge case
**Description:** Delay value is cast to float but not validated for negative values.

```python
delay_sec = float(self.entry_delay.get() or 2)  # Could be negative
```

**Fix:** Add validation: `delay_sec = max(0, float(...))`

---

### 13. **Hardcoded Supabase Credentials in Source Code** 🔓 SECURITY
**Location:** Lines 38-39
**Severity:** HIGH - Security vulnerability
**Description:** API keys are hardcoded publicly in source code.

```python
SUPABASE_URL = "https://syzmaecfeiltzrtmlgoq.supabase.co"
SUPABASE_KEY = "sb_publishable_IPDIsxft6C9RRy4s9EPgOQ_rXVaC8N-"
```

**Fix:** Move to `.env` file and load with `python-dotenv`.

---

### 14. **Missing Validation: Email List Empty Check**
**Location:** Line 1609
**Severity:** LOW - User experience
**Description:** Blasting engine doesn't validate email count after filtering.

```python
emails = [e.strip() for e in raw_emails if e.strip()]
if not emails:
    self.log("Error: No recipient emails found.")
    # Good, but no early return to `self.is_blasting = False`
```

**Fix:** Ensure `self.is_blasting` flag is reset.

---

## Summary

| Severity | Count | Issues |
|----------|-------|--------|
| CRITICAL | 1 | Missing function definitions |
| HIGH | 4 | Missing imports, race conditions, async issues, hardcoded credentials |
| MEDIUM | 4 | File handling, Python version incompatibility, UI automation fragility, platform detection |
| LOW | 4 | Resource leaks, deprecated methods, type validation, user experience |

---

## Recommended Fixes (Priority Order)

1. **Implement or remove:** `generate_scrambled_html_invoice()`, `generate_bidi_email_body()`, `generate_vectorized_pdf()`, `generate_html_mosaic()`
2. **Add imports:** `import math` and `import pandas as pd`
3. **Move credentials:** Use `.env` file for Supabase keys
4. **Add locks:** Protect `self.contexts` dictionary with `threading.Lock()`
5. **Add Python version check:** For `asyncio.to_thread()` compatibility
6. **Improve error handling:** Gmail automation selectors with comprehensive fallbacks

