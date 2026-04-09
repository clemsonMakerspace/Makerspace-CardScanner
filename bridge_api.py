"""
Bridge LMS API Integration Module
Handles fetching training/course completion status from Bridge LMS
"""

import requests
import time
from datetime import datetime
import os
import sys

# Try to import config, provide helpful error if not found
API_CONFIGURED = False
try:
    from config import BRIDGE_API_URL, BRIDGE_AUTH_TOKEN, TRAINING_CACHE_DURATION
    if BRIDGE_API_URL and BRIDGE_AUTH_TOKEN:
        API_CONFIGURED = True
    else:
        print("Warning: Bridge API credentials are empty in config.py")
except ImportError:
    print("Warning: config.py not found. Training status will not be displayed.")
    print("To enable training status, copy config.example.py to config.py and add your API credentials.")
    BRIDGE_API_URL = None
    BRIDGE_AUTH_TOKEN = None
    TRAINING_CACHE_DURATION = 60

# Course definitions - organized by category
# Test Course IDs
TEST_COURSES = {
    487: {"name": "Makerspace Test Course", "category": "test", "required": False}
}

# Production Course IDs
PRODUCTION_COURSES = {
    # Required Trainings
    5424: {"name": "Makerspace Waiver", "category": "required", "required": True, "order": 1},
    5422: {"name": "Safety Quiz", "category": "required", "required": True, "order": 2},
    
    # Priority Equipment (3D Printing)
    5473: {"name": "3D Printing", "category": "priority", "required": False, "order": 3},
    
    # Optional Equipment Trainings
    5462: {"name": "Formlabs (SLA)", "category": "optional", "required": False, "order": 4},
    5472: {"name": "3D Scanner", "category": "optional", "required": False, "order": 5},
    5455: {"name": "Epilog Laser", "category": "optional", "required": False, "order": 6},
    5456: {"name": "Sticker Printer", "category": "optional", "required": False, "order": 7},
    5457: {"name": "Vinyl Cutter", "category": "optional", "required": False, "order": 8},
    5463: {"name": "Othermill (CNC)", "category": "optional", "required": False, "order": 9},
    5461: {"name": "Fabric Printer", "category": "optional", "required": False, "order": 10},
}

# Determine which courses to use based on API URL
def get_courses():
    """Return appropriate course list based on API URL"""
    if BRIDGE_API_URL and "clemsontest" in BRIDGE_API_URL:
        return TEST_COURSES
    return PRODUCTION_COURSES

# Simple in-memory cache: {username: {"timestamp": time, "data": training_status}}
_training_cache = {}

def _is_cache_valid(username):
    """Check if cached data for user is still valid"""
    if username not in _training_cache:
        return False
    cache_entry = _training_cache[username]
    elapsed = time.time() - cache_entry["timestamp"]
    return elapsed < TRAINING_CACHE_DURATION

def _get_cached_data(username):
    """Get cached training data for user"""
    if _is_cache_valid(username):
        return _training_cache[username]["data"]
    return None

def _cache_data(username, data):
    """Cache training data for user"""
    _training_cache[username] = {
        "timestamp": time.time(),
        "data": data
    }

def check_course_completion(username, course_id):
    """
    Check if a user has completed a specific course
    
    Args:
        username: Clemson username (without @clemson.edu)
        course_id: Bridge course ID
    
    Returns:
        dict with keys: completed (bool), completed_at (str or None), error (str or None)
    """
    if not BRIDGE_API_URL or not BRIDGE_AUTH_TOKEN:
        return {"completed": False, "completed_at": None, "error": "API not configured"}
    
    email = f"{username}@clemson.edu"
    url = f"{BRIDGE_API_URL}author/course_templates/{course_id}/enrollments"
    
    params = {
        "search": email,
        "sort": "-due_date"
    }
    
    headers = {
        "Authorization": BRIDGE_AUTH_TOKEN,
        "Content-Type": "application/json"
    }
    
    try:
        api_start = time.perf_counter()  # API DEBUG
        response = requests.get(url, params=params, headers=headers, timeout=10)
        api_time = (time.perf_counter() - api_start) * 1000  # API DEBUG
        print(f"[API DEBUG] Course {course_id} check: {api_time:.0f}ms")  # API DEBUG
        response.raise_for_status()
        
        data = response.json()
        enrollments = data.get("enrollments", [])
        
        if not enrollments:
            return {"completed": False, "completed_at": None, "error": None}
        
        # Check the most recent enrollment
        enrollment = enrollments[0]
        completed_at = enrollment.get("completed_at")
        
        return {
            "completed": completed_at is not None,
            "completed_at": completed_at,
            "error": None
        }
        
    except requests.exceptions.Timeout:
        return {"completed": False, "completed_at": None, "error": "Request timeout"}
    except requests.exceptions.RequestException as e:
        return {"completed": False, "completed_at": None, "error": str(e)}
    except Exception as e:
        return {"completed": False, "completed_at": None, "error": str(e)}

def get_all_training_status(username):
    """
    Get completion status for all courses for a user
    
    Args:
        username: Clemson username (without @clemson.edu)
    
    Returns:
        dict with course info and completion status, organized by category
        Returns None if API is not configured or a critical error occurs
    """
    # Return None immediately if API not configured
    if not API_CONFIGURED:
        print("Bridge API not configured, skipping training status fetch")
        return None
    
    # Check cache first
    cached = _get_cached_data(username)
    if cached is not None:
        print(f"[API DEBUG] Using cached training data for {username} (no API calls made)")
        return cached
    
    api_total_start = time.perf_counter()  # API DEBUG: Track total API time
    print(f"Fetching training data for {username} from Bridge API...")
    
    try:
        courses = get_courses()
        results = {
            "required": [],
            "priority": [],
            "optional": [],
            "username": username,
            "fetch_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "has_errors": False
        }
        
        error_count = 0
        for course_id, course_info in courses.items():
            status = check_course_completion(username, course_id)
            
            entry = {
                "course_id": course_id,
                "name": course_info["name"],
                "category": course_info["category"],
                "required": course_info.get("required", False),
                "order": course_info.get("order", 99),
                "completed": status["completed"],
                "completed_at": status["completed_at"],
                "error": status["error"]
            }
            
            if status["error"]:
                results["has_errors"] = True
                error_count += 1
            
            # Add to appropriate category
            category = course_info["category"]
            if category == "required":
                results["required"].append(entry)
            elif category == "priority":
                results["priority"].append(entry)
            elif category == "test":
                results["required"].append(entry)  # Put test courses in required for visibility
            else:
                results["optional"].append(entry)
        
        # If ALL requests failed, return None to hide training display
        if error_count == len(courses):
            print(f"All API requests failed for {username}, hiding training display")
            return None
        
        # Sort each category by order
        results["required"].sort(key=lambda x: x["order"])
        results["priority"].sort(key=lambda x: x["order"])
        results["optional"].sort(key=lambda x: x["order"])
        
        # Calculate summary stats
        all_courses = results["required"] + results["priority"] + results["optional"]
        results["total_courses"] = len(all_courses)
        results["completed_courses"] = sum(1 for c in all_courses if c["completed"])
        results["required_complete"] = all(c["completed"] for c in results["required"])
        
        # Cache the results
        _cache_data(username, results)
        
        api_total_time = (time.perf_counter() - api_total_start) * 1000  # API DEBUG
        print(f"[API DEBUG] TOTAL API TIME for {username}: {api_total_time:.0f}ms ({len(courses)} courses)")  # API DEBUG
        
        return results
        
    except Exception as e:
        api_total_time = (time.perf_counter() - api_total_start) * 1000  # API DEBUG
        print(f"[API DEBUG] TOTAL API TIME (with error): {api_total_time:.0f}ms")  # API DEBUG
        print(f"Critical error in get_all_training_status: {e}")
        return None

def clear_cache(username=None):
    """Clear cache for specific user or all users"""
    global _training_cache
    if username:
        _training_cache.pop(username, None)
    else:
        _training_cache = {}

# Test function
if __name__ == "__main__":
    # Test with sample usernames
    test_users = ["mlalpho", "freuden"]  # Complete and incomplete test users
    
    for user in test_users:
        print(f"\n{'='*50}")
        print(f"Testing user: {user}")
        print('='*50)
        
        status = get_all_training_status(user)
        
        print(f"\nRequired Trainings:")
        for course in status["required"]:
            icon = "✓" if course["completed"] else "✗"
            print(f"  [{icon}] {course['name']}")
        
        print(f"\nPriority Equipment:")
        for course in status["priority"]:
            icon = "✓" if course["completed"] else "✗"
            print(f"  [{icon}] {course['name']}")
        
        print(f"\nOptional Equipment:")
        for course in status["optional"]:
            icon = "✓" if course["completed"] else "✗"
            print(f"  [{icon}] {course['name']}")
        
        print(f"\nSummary: {status['completed_courses']}/{status['total_courses']} complete")
        print(f"Required complete: {status['required_complete']}")
