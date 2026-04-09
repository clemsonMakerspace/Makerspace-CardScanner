# DPI Scaling, Scraping & Training Data — Technical Notes

This document records the root causes and fixes for three interrelated issues:
DPI/fullscreen scaling, web scraping reliability, and training data persistence.

---

## 1. DPI Scaling (New User Popup)

### Problem
The new user popup rendered at 1707×960 instead of the actual 2560×1440 display.
The returning user popup worked fine.

### Root Cause
The new user flow creates **two** GUI windows in sequence:

1. `prompt_for_username()` creates a `customtkinter.CTk()` window
2. After the username is entered, that window is destroyed and a fresh `tkinter.Tk()` is created for the welcome popup

`customtkinter.CTk()` internally calls `SetProcessDpiAwareness()` (Windows API).
This changes the DPI context **for the entire process**. After that CTk window is
destroyed, the subsequent `tk.Tk()` window's `winfo_screenwidth()` returns the
**unscaled** dimensions (1707×960 at 150% scaling), and the popup renders at the
wrong size.

For returning users, no CTk window is ever created — only `tk.Tk()` — so the DPI
context stays consistent and fullscreen works correctly.

### Fix
Call `SetProcessDpiAwareness(2)` **once at process startup**, before any GUI
toolkit imports. This locks in the DPI context so CTk can't change it mid-process:

```python
import sys
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    pass

import customtkinter as ctk  # Now CTk can't change DPI awareness
import tkinter as tk          # tk.Tk() will report true resolution
```

Additionally, `prompt_for_username()`'s CTk window must use
`root.attributes('-fullscreen', True)` instead of `root.geometry(f"{w}x{h}+0+0")`,
because CTk's `geometry()` method applies its own DPI scaling on top of the
OS-level awareness — resulting in a window request of 3840×2160 on a 2560×1440
screen.

### Key Lesson
`SetProcessDpiAwareness` can only be called **once** per process. If it's called
at startup, subsequent calls (by CTk or anything else) are silently ignored.
Never call it mid-process after windows have already been created.

---

## 2. Web Scraping (New Users)

### Problem
Scraping (Selenium/Chrome → my.clemson.edu) appeared to start but never completed.
The process would exit after the popup closed (~2.5s), killing the scrape.

### Root Cause
`scrape_user_thread()` ran with `daemon=True`. Daemon threads are killed
immediately when the main thread (and all non-daemon threads) exit. The flow was:

1. Popup displays for 2.5s
2. `root.destroy()` → `mainloop()` returns
3. Process exits → daemon scrape thread killed mid-HTTP-request

### Fix
Changed the scrape thread to **non-daemon** and added `thread.join(timeout=30)`
after `mainloop()` returns:

```python
def scrape_user_thread(username, callback):
    thread = threading.Thread(target=run)  # NOT daemon
    thread.start()
    return thread

# After mainloop exits:
if scrape_thread and scrape_thread.is_alive():
    scrape_thread.join(timeout=30)
```

The 30s timeout prevents infinite hangs if the scrape gets stuck.

---

## 3. Training Data Not Persisting

### Problem
API calls appeared in the log (3 of 10 courses) but training data was never saved
to the database. On next scan, there was no cached data to display.

### Root Cause
Same daemon thread issue, but in `update_training_data_async()`:

**Returning user flow:**
- `update_training_data_async(username)` spawned a **daemon thread**
- Popup showed for 2.5–4s → mainloop returned → process exited
- Daemon killed after ~3 API calls (each ~700ms) — data never saved

**New user flow:**
- Scrape thread (now non-daemon) calls `on_scrape_complete`
- `on_scrape_complete` called `update_training_data_async` — which spawned another
  **daemon thread** for training
- Scrape thread completed → `join()` returned → process exited
- Nested daemon training thread killed

### Fix

**Returning user:** Changed `update_training_data_async` thread to non-daemon,
return the thread handle, and `join(timeout=30)` after mainloop:

```python
def update_training_data_async(username):
    thread = threading.Thread(target=background_update)  # NOT daemon
    thread.start()
    return thread

# After mainloop:
training_thread = update_training_data_async(username)
root.mainloop()
if training_thread and training_thread.is_alive():
    training_thread.join(timeout=30)
```

**New user:** Eliminated the nested thread entirely. Since `on_scrape_complete`
already runs in a background thread (the scrape thread), the training fetch runs
**synchronously** inside it:

```python
def on_scrape_complete(first_name, last_name, major):
    add_user_to_sheet(...)
    # Fetch training synchronously (already in background thread)
    if BRIDGE_API_AVAILABLE:
        status = get_all_training_status(username)
        if status:
            save_training_data_to_excel(username, status)
```

This way, `scrape_thread.join()` waits for both the scrape AND the training save.

---

## 4. "Last Updated" Timestamp

### Problem
The popup showed the current time instead of when the training data was actually
fetched from the API.

### Root Cause
Two timestamps exist:
- `fetch_time` — stored inside the training JSON, set when the API call starts
- `training_last_updated` — the DB column, set to `datetime.now()` at write time

The popup displayed `last_db_update` (DB column) first, falling back to
`fetch_time`. The DB write can happen 7–10 seconds after the fetch starts (10
sequential API calls), so the timestamp was always wrong.

### Fix
Reversed the priority:

```python
last_updated = training_status.get("fetch_time") or training_status.get("last_db_update")
```

---

## Summary of Thread Lifecycle Rules

| Thread | daemon? | Wait strategy |
|--------|---------|---------------|
| Scrape (new user) | No | `join(timeout=30)` after mainloop |
| Training update (returning user) | No | `join(timeout=30)` after mainloop |
| Training fetch (new user) | N/A | Runs synchronously inside scrape thread |
| `fetch_training_status_async` | No | Used elsewhere, not critical path |

**Rule:** Any thread that writes to the database or performs I/O that must complete
must be **non-daemon** with an explicit `join()` before the process can exit.
