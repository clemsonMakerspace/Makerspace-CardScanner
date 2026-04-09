"""
EXAMPLE CONFIGURATION - Copy these settings to config.py to enable sync
"""

# ============================================================================
# EXAMPLE 1: Watt Building Scanner with Dropbox Sync
# ============================================================================

LOCATION = "Watt"
SYNC_ENABLED = True
SYNC_FOLDER = "C:\\Users\\MakerspaceUser\\Dropbox\\MakerspaceSync"
SCANNER_ID = None  # Will use "Watt" as ID
SYNC_TIMES = ["06:00", "18:00"]  # 6 AM and 6 PM
SYNC_ON_STARTUP = True
SYNC_MAX_AGE_DAYS = 7
CONFLICT_STRATEGY = "merge_all"

# ============================================================================
# EXAMPLE 2: Cooper Library Scanner with OneDrive Sync
# ============================================================================

LOCATION = "Cooper"
SYNC_ENABLED = True
SYNC_FOLDER = "C:\\Users\\LibraryPC\\OneDrive\\MakerspaceSync"
SCANNER_ID = None  # Will use "Cooper" as ID
SYNC_TIMES = ["06:00", "18:00"]  # 6 AM and 6 PM
SYNC_ON_STARTUP = True
SYNC_MAX_AGE_DAYS = 7
CONFLICT_STRATEGY = "merge_all"

# ============================================================================
# EXAMPLE 3: Multiple Scanners at Same Location
# ============================================================================

LOCATION = "Watt"
SYNC_ENABLED = True
SYNC_FOLDER = "C:\\Users\\Scanner1\\Dropbox\\MakerspaceSync"
SCANNER_ID = "Watt-Entrance"  # Unique ID for this specific scanner
SYNC_TIMES = ["06:00", "18:00"]
SYNC_ON_STARTUP = True
SYNC_MAX_AGE_DAYS = 7
CONFLICT_STRATEGY = "merge_all"

# Another scanner at same location:
# SCANNER_ID = "Watt-Exit"

# ============================================================================
# EXAMPLE 4: Sync Disabled (Default Operation)
# ============================================================================

LOCATION = "Watt"
SYNC_ENABLED = False  # No sync - operates independently
SYNC_FOLDER = ""
SCANNER_ID = None
SYNC_TIMES = []
SYNC_ON_STARTUP = False
SYNC_MAX_AGE_DAYS = 7
CONFLICT_STRATEGY = "merge_all"

# ============================================================================
# EXAMPLE 5: Frequent Sync Schedule (Multiple Times per Day)
# ============================================================================

LOCATION = "Watt"
SYNC_ENABLED = True
SYNC_FOLDER = "C:\\Users\\MakerspaceUser\\Dropbox\\MakerspaceSync"
SCANNER_ID = None
SYNC_TIMES = ["08:00", "12:00", "16:00", "20:00"]  # 4 times daily
SYNC_ON_STARTUP = True
SYNC_MAX_AGE_DAYS = 7
CONFLICT_STRATEGY = "merge_all"

# ============================================================================
# NOTES
# ============================================================================

# Cloud Folder Paths:
# - Dropbox:      "C:\\Users\\YourName\\Dropbox\\MakerspaceSync"
# - OneDrive:     "C:\\Users\\YourName\\OneDrive\\MakerspaceSync"
# - Google Drive: "G:\\My Drive\\MakerspaceSync"

# SCANNER_ID:
# - Leave as None to automatically use LOCATION as the identifier
# - Set custom ID if you have multiple scanners at same location

# SYNC_TIMES:
# - Use 24-hour format: "06:00" = 6 AM, "18:00" = 6 PM
# - Can have multiple sync times per day
# - Leave empty [] for manual-only sync

# CONFLICT_STRATEGY:
# - "merge_all": Keep all scans, only add new users (RECOMMENDED)
# - "latest_wins": Update user data if remote has newer timestamp
