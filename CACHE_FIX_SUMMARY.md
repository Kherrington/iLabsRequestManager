# Cache Duplication Fix — Summary

## What Was Wrong

The iLab Request Manager was creating multiple cache files because:

1. **Relative cache path**: The cache file was referenced as `"ilab_requests_cache.csv"` (relative path)
2. **Working directory dependent**: Each time the app was run, the cache was created wherever the current working directory happened to be
3. **Multiple instances**: You and your colleague running the app from different locations resulted in separate cache files instead of one shared source of truth

**Example of the problem:**
```
C:\Users\NIC-ADMIN4\ilab_requests_cache.csv           ← Created when you run from here
C:\Users\NIC-ADMIN4\Workspace\ilab_requests_cache.csv ← Created when you run from here
C:\Users\NIC-ADMIN4\OneDrive - UCSF\...\ilab_requests_cache.csv ← Your colleague's copy
```

Each file had different data, and syncing was impossible.

---

## What Was Fixed

### 1. **Absolute Default Path** (`config.py`)
```python
# OLD (relative path)
DATA_FILE = "ilab_requests_cache.csv"

# NEW (absolute path in app directory)
from pathlib import Path
DATA_FILE = str(Path(__file__).parent / "ilab_requests_cache.csv")
```

Now the default location is always the app's directory, regardless of where the app is launched from.

### 2. **Path Resolution Logic** (`app.py`)
Added fallback logic that ensures relative paths are resolved to the app directory:
```python
# If user sets a relative path, resolve it to the app directory
# This provides backward compatibility while preventing duplicates
_data_path_obj = _Path(_data_path)
if not _data_path_obj.is_absolute():
    _data_path = str(_Path(__file__).parent / _data_path)
```

### 3. **Documentation** (`prefs.py`)
Updated the preferences comment to explain how to set a shared OneDrive path:
```python
"data_file": "",   # absolute path to ilab_requests_cache.csv
                   # To sync across multiple users, set to shared OneDrive/network path
```

---

## New Tools

### `consolidate_cache.py` — Merge Existing Duplicates

**Interactive utility** that:
- 🔍 Finds all `ilab_requests_cache.csv` files on your system
- 📊 Shows contents of each (request counts, file size, modification dates)
- 💾 Merges them into one (keeping the newest version of each request)
- 🗑️ Backs up or deletes the old files

**Usage:**
```bash
python consolidate_cache.py
```

Then choose which file to keep as the consolidated cache.

### `SHARED_CACHE_SETUP.md` — Configuration Guide

Step-by-step instructions for:
- Consolidating existing cache files
- Configuring a shared OneDrive cache path
- Troubleshooting sync issues
- Understanding what data is shared vs. local

---

## How to Set Up

### Immediate Steps

1. **Consolidate existing caches** (one-time):
   ```bash
   cd C:\Users\NIC-ADMIN4\Workspace\iLabsRequestManager
   python consolidate_cache.py
   ```
   
   Choose option 1 or 2 depending on where you want the master cache.

2. **Update both users' preferences** to point to the same cache file:
   
   **For you:**
   ```json
   {
     "data_file": "C:\\Users\\NIC-ADMIN4\\OneDrive - UCSF\\Documents - CALM\\ilab_requests_cache.csv"
   }
   ```
   
   **For your colleague:**
   ```json
   {
     "data_file": "C:\\Users\\<COLLEAGUE>\\OneDrive - UCSF\\Documents - CALM\\ilab_requests_cache.csv"
   }
   ```
   
   Both paths point to the same OneDrive file, so changes sync automatically.

3. **Restart the app** to load the new cache path.

### Going Forward

- No more duplicate cache files ✓
- Both users work from the same data ✓
- Changes sync automatically via OneDrive ✓
- Local notes, assignments, and labels are preserved ✓

---

## Technical Details

### What Changed
- `config.py`: Default cache path is now absolute
- `app.py`: Path resolution logic for backward compatibility
- `prefs.py`: Updated documentation

### What Stayed the Same
- All data structure (CSV format, columns) is unchanged
- API sync logic works the same
- Preferences system works the same
- Local-only fields still preserved

### Backward Compatibility
- Existing `prefs.json` files continue to work
- If `data_file` is empty, uses the new absolute default path
- Relative paths in `data_file` are auto-resolved to app directory

---

## Questions?

- **How do I know if the shared cache is working?** 
  - Both users will have the same records in their request list
  - When one user edits a note, the other sees it after clicking **Sync All**

- **What if OneDrive isn't synced?**
  - Check OneDrive sync status in your system tray
  - Make sure the OneDrive path exists on both machines
  - Try right-click → Sync on the cache file

- **Can we use a different shared path?**
  - Yes! Any network drive or shared folder works
  - Just set `data_file` to the absolute path
  - Both users must have read/write access

---

## Files Modified
- ✏️ `config.py` — DEFAULT path now absolute
- ✏️ `app.py` — Path resolution logic added
- ✏️ `prefs.py` — Documentation improved
- ✨ `consolidate_cache.py` — NEW utility
- ✨ `SHARED_CACHE_SETUP.md` — NEW guide
