"""
Cloud-Based Database Synchronization System
Merges multiple SQLite databases from different locations via cloud storage
"""

import sqlite3
import shutil
import os
import time
import json
from datetime import datetime, timedelta
from pathlib import Path
import threading
import schedule

try:
    from config import (
        SYNC_ENABLED, SYNC_FOLDER, SCANNER_ID, LOCATION, 
        SYNC_TIMES, SYNC_ON_STARTUP, SYNC_MAX_AGE_DAYS,
        CONFLICT_STRATEGY, DB_FILE
    )
except ImportError:
    # Fallback defaults if config.py doesn't exist
    SYNC_ENABLED = False
    SYNC_FOLDER = ""
    SCANNER_ID = None
    LOCATION = "Watt"
    SYNC_TIMES = ["18:00"]
    SYNC_ON_STARTUP = True
    SYNC_MAX_AGE_DAYS = 7
    CONFLICT_STRATEGY = "merge_all"
    DB_FILE = "hardware_users.db"

# Sync state tracking
SYNC_LOG_FILE = ".sync_log.json"
_sync_in_progress = False
_sync_lock = threading.Lock()


def get_scanner_id():
    """Get unique identifier for this scanner"""
    return SCANNER_ID if SCANNER_ID else LOCATION


def get_sync_metadata():
    """Load sync metadata from log file"""
    if os.path.exists(SYNC_LOG_FILE):
        try:
            with open(SYNC_LOG_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {
        "last_sync": None,
        "sync_count": 0,
        "last_error": None
    }


def save_sync_metadata(metadata):
    """Save sync metadata to log file"""
    try:
        with open(SYNC_LOG_FILE, 'w') as f:
            json.dump(metadata, f, indent=2)
    except Exception as e:
        print(f"Warning: Could not save sync metadata: {e}")


def is_sync_enabled():
    """Check if sync is enabled and properly configured"""
    if not SYNC_ENABLED:
        return False
    
    if not SYNC_FOLDER or not os.path.exists(SYNC_FOLDER):
        print(f"⚠ Sync enabled but cloud folder not accessible: {SYNC_FOLDER}")
        return False
    
    return True


def upload_local_database():
    """
    Upload local database to cloud sync folder
    Filename includes scanner ID and timestamp
    """
    try:
        if not os.path.exists(DB_FILE):
            print(f"⚠ Local database not found: {DB_FILE}")
            return False
        
        # Ensure sync folder exists
        os.makedirs(SYNC_FOLDER, exist_ok=True)
        
        # Generate filename with scanner ID and timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        scanner_id = get_scanner_id()
        remote_filename = f"db_{scanner_id}_{timestamp}.db"
        remote_path = os.path.join(SYNC_FOLDER, remote_filename)
        
        # Copy database to cloud folder
        shutil.copy2(DB_FILE, remote_path)
        print(f"✓ Uploaded database to cloud: {remote_filename}")
        
        # Also create a "latest" copy for quick reference
        latest_filename = f"db_{scanner_id}_LATEST.db"
        latest_path = os.path.join(SYNC_FOLDER, latest_filename)
        shutil.copy2(DB_FILE, latest_path)
        
        return True
        
    except Exception as e:
        print(f"✗ Failed to upload database: {e}")
        return False


def get_remote_databases():
    """
    Get list of all remote databases in cloud folder
    Returns list of (filepath, scanner_id, timestamp) tuples
    """
    databases = []
    
    try:
        if not os.path.exists(SYNC_FOLDER):
            return databases
        
        cutoff_date = datetime.now() - timedelta(days=SYNC_MAX_AGE_DAYS)
        scanner_id = get_scanner_id()
        
        for filename in os.listdir(SYNC_FOLDER):
            if not filename.endswith('.db') or '_LATEST' in filename:
                continue
            
            # Parse filename: db_{scanner_id}_{timestamp}.db
            parts = filename.replace('.db', '').split('_')
            if len(parts) < 3 or parts[0] != 'db':
                continue
            
            file_scanner_id = parts[1]
            file_timestamp_str = parts[2]
            
            # Skip our own databases
            if file_scanner_id == scanner_id:
                continue
            
            # Parse timestamp
            try:
                file_timestamp = datetime.strptime(file_timestamp_str, "%Y%m%d")
                if file_timestamp < cutoff_date:
                    continue  # Too old
            except:
                continue
            
            filepath = os.path.join(SYNC_FOLDER, filename)
            databases.append((filepath, file_scanner_id, file_timestamp))
        
        # Sort by timestamp (newest first)
        databases.sort(key=lambda x: x[2], reverse=True)
        
    except Exception as e:
        print(f"⚠ Error scanning remote databases: {e}")
    
    return databases


def merge_databases(remote_dbs):
    """
    Merge remote databases into local database
    Implements conflict resolution based on CONFLICT_STRATEGY
    """
    if not remote_dbs:
        print("ℹ No remote databases to merge")
        return True
    
    try:
        # Connect to local database
        local_conn = sqlite3.connect(DB_FILE)
        local_cursor = local_conn.cursor()
        
        merged_users = 0
        merged_scans = 0
        
        for remote_path, remote_scanner_id, remote_timestamp in remote_dbs:
            print(f"  Merging from {remote_scanner_id} ({remote_timestamp.strftime('%Y-%m-%d')})...")
            
            try:
                # Attach remote database
                local_cursor.execute(f"ATTACH DATABASE '{remote_path}' AS remote")
                
                # Merge users table
                if CONFLICT_STRATEGY == "latest_wins":
                    # Replace user if remote has newer update timestamp
                    merge_users_query = """
                        INSERT OR REPLACE INTO users 
                        SELECT * FROM remote.users 
                        WHERE username NOT IN (SELECT username FROM users)
                           OR updated_at > (SELECT updated_at FROM users WHERE users.username = remote.users.username)
                    """
                else:  # merge_all
                    # Keep existing, only add new users
                    merge_users_query = """
                        INSERT OR IGNORE INTO users 
                        SELECT * FROM remote.users
                    """
                
                result = local_cursor.execute(merge_users_query)
                merged_users += result.rowcount
                
                # Merge scans table (always merge all unique scans)
                # Avoid duplicates by checking hardware_id + timestamp combination
                merge_scans_query = """
                    INSERT OR IGNORE INTO scans (hardware_id, username, timestamp, location)
                    SELECT hardware_id, username, timestamp, location FROM remote.scans
                    WHERE NOT EXISTS (
                        SELECT 1 FROM scans 
                        WHERE scans.hardware_id = remote.scans.hardware_id 
                        AND scans.timestamp = remote.scans.timestamp
                    )
                """
                
                result = local_cursor.execute(merge_scans_query)
                merged_scans += result.rowcount
                
                # Detach remote database
                local_cursor.execute("DETACH DATABASE remote")
                
            except Exception as e:
                print(f"  ⚠ Error merging {remote_scanner_id}: {e}")
                try:
                    local_cursor.execute("DETACH DATABASE remote")
                except:
                    pass
        
        local_conn.commit()
        local_conn.close()
        
        print(f"✓ Merge complete: {merged_users} users, {merged_scans} scans")
        return True
        
    except Exception as e:
        print(f"✗ Merge failed: {e}")
        return False


def cleanup_old_remote_databases():
    """
    Remove old database files from cloud folder to save space
    Keeps only the latest from each scanner
    """
    try:
        if not os.path.exists(SYNC_FOLDER):
            return
        
        # Group databases by scanner_id
        scanner_dbs = {}
        
        for filename in os.listdir(SYNC_FOLDER):
            if not filename.endswith('.db') or '_LATEST' in filename:
                continue
            
            parts = filename.replace('.db', '').split('_')
            if len(parts) < 3 or parts[0] != 'db':
                continue
            
            scanner_id = parts[1]
            timestamp_str = parts[2]
            
            try:
                timestamp = datetime.strptime(timestamp_str, "%Y%m%d")
                filepath = os.path.join(SYNC_FOLDER, filename)
                
                if scanner_id not in scanner_dbs:
                    scanner_dbs[scanner_id] = []
                scanner_dbs[scanner_id].append((filepath, timestamp))
            except:
                continue
        
        # For each scanner, keep only the 2 most recent
        deleted_count = 0
        for scanner_id, db_list in scanner_dbs.items():
            db_list.sort(key=lambda x: x[1], reverse=True)
            
            # Delete all but the 2 most recent
            for filepath, timestamp in db_list[2:]:
                try:
                    os.remove(filepath)
                    deleted_count += 1
                except Exception as e:
                    print(f"  Warning: Could not delete {filepath}: {e}")
        
        if deleted_count > 0:
            print(f"  Cleaned up {deleted_count} old database files")
            
    except Exception as e:
        print(f"⚠ Cleanup error: {e}")


def perform_sync():
    """
    Main sync operation:
    1. Upload local database to cloud
    2. Download and merge remote databases
    3. Cleanup old files
    """
    global _sync_in_progress
    
    if _sync_in_progress:
        print("⚠ Sync already in progress, skipping...")
        return False
    
    if not is_sync_enabled():
        return False
    
    print(f"\n{'='*60}")
    print(f"🔄 Starting database sync - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}")
    
    with _sync_lock:
        _sync_in_progress = True
        metadata = get_sync_metadata()
        
        try:
            # Step 1: Upload our database
            print("📤 Step 1: Uploading local database...")
            if not upload_local_database():
                raise Exception("Upload failed")
            
            # Step 2: Get remote databases
            print("📥 Step 2: Scanning for remote databases...")
            remote_dbs = get_remote_databases()
            print(f"  Found {len(remote_dbs)} remote databases to merge")
            
            # Step 3: Merge remote databases
            if remote_dbs:
                print("🔀 Step 3: Merging remote databases...")
                if not merge_databases(remote_dbs):
                    raise Exception("Merge failed")
            else:
                print("ℹ Step 3: No remote databases to merge")
            
            # Step 4: Cleanup
            print("🧹 Step 4: Cleaning up old files...")
            cleanup_old_remote_databases()
            
            # Update metadata
            metadata["last_sync"] = datetime.now().isoformat()
            metadata["sync_count"] = metadata.get("sync_count", 0) + 1
            metadata["last_error"] = None
            save_sync_metadata(metadata)
            
            print(f"{'='*60}")
            print(f"✅ Sync completed successfully!")
            print(f"   Total syncs: {metadata['sync_count']}")
            print(f"{'='*60}\n")
            
            return True
            
        except Exception as e:
            metadata["last_error"] = str(e)
            save_sync_metadata(metadata)
            
            print(f"{'='*60}")
            print(f"❌ Sync failed: {e}")
            print(f"{'='*60}\n")
            
            return False
            
        finally:
            _sync_in_progress = False


def schedule_sync():
    """
    Schedule automatic syncs at configured times
    Runs in background thread
    """
    if not is_sync_enabled():
        return
    
    for sync_time in SYNC_TIMES:
        schedule.every().day.at(sync_time).do(perform_sync)
        print(f"✓ Scheduled sync for {sync_time} daily")
    
    def run_scheduler():
        while True:
            schedule.run_pending()
            time.sleep(60)  # Check every minute
    
    scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
    scheduler_thread.start()
    print("✓ Sync scheduler started")


def init_sync():
    """
    Initialize sync system
    Call this on program startup
    """
    if not is_sync_enabled():
        print("ℹ Database sync is disabled")
        return
    
    print(f"✓ Database sync enabled - Scanner ID: {get_scanner_id()}")
    print(f"  Cloud folder: {SYNC_FOLDER}")
    print(f"  Sync times: {', '.join(SYNC_TIMES)}")
    
    # Perform startup sync if enabled
    if SYNC_ON_STARTUP:
        print("  Performing startup sync...")
        threading.Thread(target=perform_sync, daemon=True).start()
    
    # Schedule periodic syncs
    schedule_sync()


if __name__ == "__main__":
    # Test/manual sync
    print("Manual Database Sync")
    print("=" * 60)
    
    if not is_sync_enabled():
        print("⚠ Sync is disabled in config.py")
        print("  Set SYNC_ENABLED = True and configure SYNC_FOLDER")
    else:
        perform_sync()
