# What Changed in This Update

## 1. Plain-Text Passwords
Passwords are now stored as plain readable text in the database (no hashing).
- Existing admin account password remains: `Flask@Mitel#Walkathon26!`
- Any legacy hashed passwords are automatically converted on first startup.
- New registrations store passwords as-is.

## 2. Hourly Auto-Backup to Network Path
Every hour the application automatically exports a full Excel backup to your configured path.

**How to set the backup folder:**
```
# Windows (set environment variable before running):
set BACKUP_NETWORK_PATH=\\SERVER\Share\backups

# Linux:
export BACKUP_NETWORK_PATH=/mnt/share/backups
```

If not set, backups go to a `backups/` subfolder inside the project directory.

The last 48 backups (48 hours) are kept automatically. Older ones are deleted.

You can also change the interval (default 3600 seconds = 1 hour):
```
set BACKUP_INTERVAL_SECONDS=1800   # every 30 minutes
```

## 3. Improved Excel Export (Admin Page → Backup & Restore tab)
- All optional fields (weight, calories, team, notes, country) show **-** instead of blank cells.
- Download produces a proper `.xlsx` file with colour-coded headers.
- Contains 3 human-readable sheets (Users, Teams, Steps) + 3 raw restore sheets.

## 4. Upload Excel Backup / Restore Database (NEW Admin feature)
In the Admin page → **Backup & Restore** tab:
1. Click **Upload Excel Backup** or drag-and-drop your `.xlsx` file.
2. Click **Restore Database from File**.
3. The system reads all sheets, inserts any missing records, and reports how many were imported.
4. Existing records are never overwritten — only missing ones are added.

This works seamlessly even after a complete database loss.

## 5. New `requirements.txt` entry
```
openpyxl>=3.1.0
```
Install with:
```
pip install -r requirements.txt
```
