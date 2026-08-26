# iLab Request Manager — Shared Cache Setup

## Problem: Duplicate Cache Files

When you and your colleague run the iLab Request Manager from different locations (or with different working directories), the app creates separate cache files instead of sharing one unified cache. This causes:

- Duplicate request records
- Sync conflicts
- Loss of locally-added information (assigned staff, labels, notes)
- Difficulty keeping multiple app instances in sync

## Root Cause

The cache file path used to be relative (`ilab_requests_cache.csv`), so it was created in whatever directory the app was run from. We've fixed this, but existing caches may still be scattered.

## Solution: Unified Cache

There are two parts to the solution:

### Part 1: Consolidate Existing Cache Files

If you have multiple cache files, consolidate them into one:

```bash
cd C:\Users\NIC-ADMIN4\Workspace\iLabsRequestManager
python consolidate_cache.py
```

This script will:
1. 🔍 Find all `ilab_requests_cache.csv` files on your system
2. 📊 Show you what's in each one
3. 💾 Merge them into a single file (keeping the newest version of each record)
4. 🗑️ Optionally back up or delete the old duplicates

### Part 2: Configure Shared Cache Path

Both you and your colleague should point to the same cache file in your preferences.

**Option A: Shared OneDrive Path (Recommended)**

This is the easiest method — the cache auto-syncs via OneDrive.

1. Run the app
2. Click **⚙ Preferences**
3. In the **Data Cache** section, set the path to:
   ```
   C:\Users\NIC-ADMIN4\OneDrive - UCSF\Documents - CALM\ilab_requests_cache.csv
   ```
4. Click **Save Preferences** and restart the app

Your colleague should use:
   ```
   C:\Users\<COLLEAGUE-USERNAME>\OneDrive - UCSF\Documents - CALM\ilab_requests_cache.csv
   ```

Both paths point to the same OneDrive file, so changes sync automatically.

**Option B: Shared Network Drive**

If you have network access:

1. Place the cache file on a shared network path both users can access
2. Set both users' `data_file` preference to that path
3. The app will reload the file on startup to pick up changes from the other user

**Option C: Manual via prefs.json**

If the Preferences UI doesn't have a Data Cache field yet, edit `prefs.json` directly:

```json
{
  "data_file": "C:\\Users\\NIC-ADMIN4\\OneDrive - UCSF\\Documents - CALM\\ilab_requests_cache.csv",
  ...
}
```

## How It Works

1. **Absolute paths prevent duplication**: The cache file is now stored in the app directory by default, not the current working directory.

2. **Preferences allow overrides**: You can point to any shared path via `prefs.json`.

3. **OneDrive auto-sync**: Since your cache is in OneDrive, both machines automatically get the latest version.

4. **Reload on startup**: The app re-reads the CSV at startup to pick up changes made by the other user.

## After Setup

- Both of you edit the same cache file
- Changes are automatically reflected across machines (via OneDrive sync)
- All local notes, assignments, and labels are preserved and shared
- No more duplicate records

## Troubleshooting

### "Permission Denied" Error

If you get a permission error, the other user might have the file locked. Try:
- Wait 30 seconds for OneDrive to release the file
- Restart the app
- Check that both users have write permission to the OneDrive path

### Changes Aren't Syncing

If changes from one user don't appear on the other's machine:

1. **Check OneDrive sync status**: Look at the OneDrive icon in your system tray
2. **Manual reload**: Click **Sync All** in the app to force a re-read of the cache file
3. **Verify the path**: Both users should be pointing to the exact same file path

### Old Duplicates Are Still There

After consolidation, you can manually clean up old cache files:

```bash
# List all cache files (safe)
python -c "
from consolidate_cache import find_cache_files
for f in find_cache_files():
    print(f)
"

# Delete a specific old file manually
del C:\path\to\old\ilab_requests_cache.csv
```

## What's Saved Locally (Not Synced)

The following are never shared and stay local to each machine:

- User preferences (xlsx paths, dark mode, etc.) — stored in `prefs.json`
- Preferences are machine-specific because each user may have different file paths

The following **are** shared in the cache file:

- All iLab request data (read from the API)
- Assigned staff
- Labels and tags
- Local notes
- Training information (date, microscope, etc.)
- Workflow progress (emails sent, etc.)

---

**Questions?** Check `config.py` and `data_store.py` in the app directory for implementation details.
