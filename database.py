"""
SQLite Database Module - Robust backup for Excel data
Maintains an identical copy of all Excel data in SQLite format
Automatically syncs on every write operation
"""

import sqlite3
import json
import os
import threading
from datetime import datetime
from contextlib import contextmanager

# Database configuration
DB_FILE = "hardware_users.db"
_db_lock = threading.Lock()

def init_database():
    """
    Initialize the SQLite database with tables matching Excel structure
    Creates tables if they don't exist
    """
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # Users table (matches Users sheet)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY,
                hardware_id INTEGER UNIQUE,
                login_count INTEGER DEFAULT 0,
                first_name TEXT,
                last_name TEXT,
                major TEXT,
                training_data TEXT,
                training_last_updated TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Scans table (matches Scans sheet)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS scans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                hardware_id INTEGER,
                username TEXT,
                timestamp TIMESTAMP,
                location TEXT DEFAULT 'Watt',
                FOREIGN KEY (username) REFERENCES users(username)
            )
        ''')
        
        # Create indexes for performance
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_users_hardware_id 
            ON users(hardware_id)
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_scans_username 
            ON scans(username)
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_scans_timestamp 
            ON scans(timestamp)
        ''')
        
        conn.commit()
        print("✓ Database initialized successfully")

@contextmanager
def get_db_connection():
    """
    Context manager for database connections with thread safety
    Automatically commits and closes connection
    """
    conn = None
    try:
        _db_lock.acquire()
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row  # Enable column access by name
        yield conn
        conn.commit()
    except Exception as e:
        if conn:
            conn.rollback()
        print(f"Database error: {e}")
        raise
    finally:
        if conn:
            conn.close()
        if _db_lock.locked():
            _db_lock.release()

def add_or_update_user(username, hardware_id, first_name=None, last_name=None, major=None):
    """
    Add a new user or update existing user information
    Returns True if successful, False otherwise
    """
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            
            # Check if user exists
            cursor.execute('SELECT username FROM users WHERE username = ?', (username,))
            exists = cursor.fetchone()
            
            if exists:
                # Update existing user
                cursor.execute('''
                    UPDATE users 
                    SET hardware_id = ?,
                        first_name = COALESCE(?, first_name),
                        last_name = COALESCE(?, last_name),
                        major = COALESCE(?, major),
                        updated_at = CURRENT_TIMESTAMP
                    WHERE username = ?
                ''', (hardware_id, first_name, last_name, major, username))
                print(f"Updated user: {username}")
            else:
                # Insert new user
                cursor.execute('''
                    INSERT INTO users (username, hardware_id, first_name, last_name, major)
                    VALUES (?, ?, ?, ?, ?)
                ''', (username, hardware_id, first_name, last_name, major))
                print(f"Added new user to database: {username}")
            
            return True
    except sqlite3.IntegrityError as e:
        print(f"Database integrity error for user {username}: {e}")
        return False
    except Exception as e:
        print(f"Error adding/updating user {username}: {e}")
        return False

def add_scan(hardware_id, username, timestamp=None, location='Watt'):
    """
    Record a new scan in the database
    Returns True if successful, False otherwise
    """
    if timestamp is None:
        timestamp = datetime.now().strftime('%m/%d/%Y %H:%M:%S')
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO scans (hardware_id, username, timestamp, location)
                VALUES (?, ?, ?, ?)
            ''', (hardware_id, username, timestamp, location))
            
            # Increment login count
            cursor.execute('''
                UPDATE users 
                SET login_count = login_count + 1,
                    updated_at = CURRENT_TIMESTAMP
                WHERE username = ?
            ''', (username,))
            
            return True
    except Exception as e:
        print(f"Error adding scan for {username}: {e}")
        return False

def update_training_data(username, training_status):
    """
    Update training data for a user
    Stores as JSON in the database
    """
    if not training_status:
        return False
    
    try:
        training_json = json.dumps(training_status, separators=(',', ':'))
        timestamp = datetime.now().strftime('%m/%d/%Y %H:%M:%S')
        
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE users 
                SET training_data = ?,
                    training_last_updated = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE username = ?
            ''', (training_json, timestamp, username))
            
            if cursor.rowcount > 0:
                print(f"Training data updated in database for {username}")
                return True
            else:
                print(f"Warning: User {username} not found for training update")
                return False
    except Exception as e:
        print(f"Error updating training data for {username}: {e}")
        return False

def get_user_by_hardware_id(hardware_id):
    """
    Get user information by hardware ID
    Returns dict with user data or None if not found
    """
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT username, hardware_id, login_count, first_name, 
                       last_name, major, training_data, training_last_updated
                FROM users 
                WHERE hardware_id = ?
            ''', (hardware_id,))
            
            row = cursor.fetchone()
            if row:
                return dict(row)
            return None
    except Exception as e:
        print(f"Error getting user by hardware ID {hardware_id}: {e}")
        return None

def get_user_by_username(username):
    """
    Get user information by username
    Returns dict with user data or None if not found
    """
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT username, hardware_id, login_count, first_name, 
                       last_name, major, training_data, training_last_updated
                FROM users 
                WHERE username = ?
            ''', (username,))
            
            row = cursor.fetchone()
            if row:
                return dict(row)
            return None
    except Exception as e:
        print(f"Error getting user {username}: {e}")
        return None

def get_training_data(username):
    """
    Get training data for a user
    Returns parsed JSON dict or None
    """
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT training_data, training_last_updated
                FROM users 
                WHERE username = ?
            ''', (username,))
            
            row = cursor.fetchone()
            if row and row['training_data']:
                training_data = json.loads(row['training_data'])
                training_data['last_db_update'] = row['training_last_updated']
                return training_data
            return None
    except Exception as e:
        print(f"Error getting training data for {username}: {e}")
        return None

def get_recent_scans(limit=100):
    """
    Get recent scans from the database
    Returns list of scan records
    """
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT s.id, s.hardware_id, s.username, s.timestamp, s.location,
                       u.first_name, u.last_name
                FROM scans s
                LEFT JOIN users u ON s.username = u.username
                ORDER BY s.timestamp DESC
                LIMIT ?
            ''', (limit,))
            
            return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"Error getting recent scans: {e}")
        return []

def get_user_scan_count(username):
    """
    Get total number of scans for a user
    """
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT COUNT(*) as count
                FROM scans 
                WHERE username = ?
            ''', (username,))
            
            row = cursor.fetchone()
            return row['count'] if row else 0
    except Exception as e:
        print(f"Error getting scan count for {username}: {e}")
        return 0

def get_database_stats():
    """
    Get statistics about the database
    Returns dict with counts and info
    """
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            
            # Total users
            cursor.execute('SELECT COUNT(*) as count FROM users')
            total_users = cursor.fetchone()['count']
            
            # Total scans
            cursor.execute('SELECT COUNT(*) as count FROM scans')
            total_scans = cursor.fetchone()['count']
            
            # Users with training data
            cursor.execute('SELECT COUNT(*) as count FROM users WHERE training_data IS NOT NULL')
            users_with_training = cursor.fetchone()['count']
            
            # Database file size
            db_size = os.path.getsize(DB_FILE) if os.path.exists(DB_FILE) else 0
            
            return {
                'total_users': total_users,
                'total_scans': total_scans,
                'users_with_training': users_with_training,
                'database_size_bytes': db_size,
                'database_size_mb': round(db_size / 1024 / 1024, 2)
            }
    except Exception as e:
        print(f"Error getting database stats: {e}")
        return {}

def backup_database(backup_dir='backups'):
    """
    Create a backup copy of the database
    """
    import shutil
    
    try:
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        
        timestamp = datetime.now().strftime("%m_%d_%Y___%H_%M_%S")
        backup_file = os.path.join(backup_dir, f"hardware_users_db_Watt_{timestamp}.db")
        
        shutil.copy2(DB_FILE, backup_file)
        print(f"✓ Database backup created: {backup_file}")
        return backup_file
    except Exception as e:
        print(f"Error backing up database: {e}")
        return None

def verify_database_integrity():
    """
    Run SQLite integrity check
    Returns True if database is OK, False if corrupted
    """
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('PRAGMA integrity_check')
            result = cursor.fetchone()
            
            if result and result[0] == 'ok':
                print("✓ Database integrity check passed")
                return True
            else:
                print(f"⚠ Database integrity check failed: {result}")
                return False
    except Exception as e:
        print(f"Error checking database integrity: {e}")
        return False

# Initialize database on module import
if __name__ != "__main__":
    if not os.path.exists(DB_FILE):
        print("Creating new database...")
        init_database()
    else:
        # Verify existing database
        verify_database_integrity()

# Test/demo code
if __name__ == "__main__":
    print("=" * 70)
    print("SQLite Database Module Test")
    print("=" * 70)
    
    # Initialize
    init_database()
    
    # Test adding user
    print("\n1. Adding test user...")
    add_or_update_user("testuser", 999999, "Test", "User", "Computer Science")
    
    # Test adding scan
    print("\n2. Adding test scan...")
    add_scan(999999, "testuser")
    
    # Test getting user
    print("\n3. Getting user by hardware ID...")
    user = get_user_by_hardware_id(999999)
    print(f"   Found: {user}")
    
    # Test stats
    print("\n4. Database statistics...")
    stats = get_database_stats()
    for key, value in stats.items():
        print(f"   {key}: {value}")
    
    print("\n" + "=" * 70)
    print("Test complete!")
