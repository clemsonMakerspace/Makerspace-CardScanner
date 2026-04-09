# Excel Editing Guide - Manual Database Management

## 📊 **Overview**

The system now supports **bidirectional synchronization** between Excel and the SQLite database. This means you can:

✅ **Edit Excel manually** to update user information  
✅ **Add users via Excel** instead of card scanning  
✅ **Modify training data** by editing JSON in Excel  
✅ **Changes automatically sync** to database on next program start  

---

## 🔄 **How Bidirectional Sync Works**

### **On Startup**
```
Program Starts
    ↓
Check file modification times
    ↓
If Excel is newer → Excel updates Database
If Database is newer → Database updates Excel
    ↓
Both sync in both directions for consistency
    ↓
Program runs with synced data
```

### **On Shutdown**
```
Program Closes
    ↓
Database → Excel sync runs
    ↓
Excel file updated with latest scans
    ↓
Ready for manual editing
```

---

## ✏️ **Editing User Data in Excel**

### **Step 1: Close the Scanner Program**
- Make sure **MakerspaceSignInTablet.py** or **CardReaderMakerspace.py** is NOT running
- This prevents file conflicts

### **Step 2: Open `hardware_users.xlsx`**
Navigate to the **Users** sheet (bottom tab)

### **Step 3: Edit Data**

**Column Structure:**
```
A: Username       (e.g., "jlbrk")
B: Hardware ID    (e.g., 111111)
C: Login Count    (auto-updated by system)
D: First Name     (e.g., "John")
E: Last Name      (e.g., "Smith")
F: Major          (e.g., "Computer Science")
G: Training Data  (JSON - see below)
H: Training Updated (timestamp)
```

**Example Edits:**

**Add a New User:**
```
Username: jsmith
Hardware ID: 123456
First Name: Jane
Last Name: Smith
Major: Mechanical Engineering
```

**Update Existing User:**
```
Change major from "Computer Science" → "Electrical Engineering"
Just edit cell in column F
```

**Fix Typo in Name:**
```
Change "Jhon" → "John" in column D
```

### **Step 4: Save and Close Excel**
- Click **File → Save**
- Close Excel completely

### **Step 5: Restart Scanner Program**
- Run **MakerspaceSignInTablet.py** or **CardReaderMakerspace.py**
- Watch console for sync confirmation:

```
🔄 ============================================================
   SMART BIDIRECTIONAL SYNC - Excel ↔ Database
🔄 ============================================================
📊 Excel modified:    2026-01-15 14:30:25
💾 Database modified: 2026-01-15 09:15:10
→ Excel is newer - User likely made manual edits
  Priority: Excel → Database
✓ Synced 15 users from Excel to Database
✓ Synced 0 new scans from Excel to Database
✅ Bidirectional sync complete!
```

---

## 🎓 **Editing Training Data**

Training data is stored in **Column G** as JSON. You can manually edit it if needed.

### **Format:**
```json
{
  "required": [
    {"name": "Safety Training", "completed": true},
    {"name": "Tool Basics", "completed": false}
  ],
  "priority": [
    {"name": "3D Printing", "completed": true}
  ],
  "optional": [
    {"name": "Laser Cutter", "completed": false},
    {"name": "CNC Mill", "completed": true}
  ],
  "total_courses": 5,
  "completed_courses": 3,
  "required_complete": false
}
```

### **To Edit:**

1. **Open Cell G for the user**
2. **Copy JSON to text editor** (Notepad++ or VS Code)
3. **Edit values:**
   - Change `"completed": false` → `"completed": true`
   - Add new courses to arrays
   - Update counts
4. **Copy back to Excel cell G**
5. **Update Cell H** with current timestamp: `01/15/2026 14:30:00`
6. **Save and restart program**

---

## ⚠️ **Important Rules**

### **DO:**
✅ Edit Excel when program is **closed**  
✅ Use consistent formatting (dates, usernames, etc.)  
✅ Keep Column A (Username) unique  
✅ Keep Column B (Hardware ID) unique  
✅ Save Excel before reopening program  

### **DON'T:**
❌ Edit Excel while program is **running** (file lock conflicts)  
❌ Delete column headers (row 1)  
❌ Use duplicate usernames  
❌ Use duplicate hardware IDs  
❌ Leave required fields blank (Username, Hardware ID)  
❌ Manually edit **Scans** sheet (let program manage it)  

---

## 📋 **Common Editing Tasks**

### **Task 1: Add Multiple New Users**
```
1. Close scanner program
2. Open hardware_users.xlsx → Users sheet
3. Add rows at bottom:
   Row 10: jsmith | 123456 | 0 | Jane | Smith | Mech Eng
   Row 11: bwilson | 789012 | 0 | Bob | Wilson | Civil Eng
4. Save and close Excel
5. Restart scanner
6. Console shows: "✓ Synced 2 users from Excel to Database"
```

### **Task 2: Fix User Information**
```
1. Close scanner
2. Find user row in Excel (search for username)
3. Edit columns D, E, F as needed
4. Save Excel
5. Restart scanner
6. Changes sync to database automatically
```

### **Task 3: Bulk Update Training Status**
```
1. Close scanner
2. Export Column G to text file
3. Use find/replace: "completed": false → "completed": true
4. Import back to Column G
5. Update Column H timestamps
6. Save and restart
```

### **Task 4: Remove Invalid User**
```
1. Close scanner
2. Delete entire row in Users sheet
3. Save Excel
4. Restart scanner
5. User removed from database on sync
```

---

## 🔍 **Verification**

### **Check Sync Happened:**
Look for console output on startup:
```
🚀 STARTUP SYNC
🔄 ============================================================
   SMART BIDIRECTIONAL SYNC - Excel ↔ Database
✓ Synced 15 users from Excel to Database
✅ Bidirectional sync complete!
```

### **Check Sync Direction:**
- **Excel → Database:** "Excel is newer - User likely made manual edits"
- **Database → Excel:** "Database is newer - Program made updates"

### **Verify User Added:**
Scan their card or type username - should work immediately

---

## 🛠️ **Troubleshooting**

### **Problem: Changes Not Appearing**

**Solution:**
1. Check you saved Excel before closing
2. Verify sync ran on startup (check console)
3. Confirm no errors in sync output
4. Try manual sync: `python excel_db_sync.py`

### **Problem: "File is locked" Error**

**Solution:**
- Close scanner program completely
- Close Excel completely
- Check Task Manager for lingering processes
- Reopen Excel and try again

### **Problem: Duplicate Hardware IDs**

**Solution:**
- Database will reject duplicate hardware IDs
- Check Excel for duplicate values in Column B
- Remove or modify duplicates
- Restart program

### **Problem: Invalid JSON in Training Data**

**Solution:**
- Use online JSON validator: jsonlint.com
- Copy Column G value and validate
- Fix syntax errors (missing commas, quotes)
- Paste corrected JSON back to Excel

---

## 📊 **Sync Timing**

### **When Syncs Happen:**

| Event | Sync Direction | Purpose |
|-------|----------------|---------|
| **Program Startup** | Excel ↔ Database (smart) | Import manual edits |
| **Program Shutdown** | Database → Excel | Export latest scans |
| **Every Scan** | Write to Database only | Fast performance |
| **Manual** | Both directions | Testing/troubleshooting |

### **File Modification Times:**
- System compares file timestamps
- If difference < 5 seconds: "Already in sync"
- If Excel newer: Excel edits imported
- If Database newer: Database exported

---

## 🎯 **Best Practices**

1. **Make Edits During Off-Hours**
   - Edit Excel when no one is scanning cards
   - Prevents confusion with live data

2. **Backup Before Major Changes**
   - Copy `hardware_users.xlsx` before bulk edits
   - Database backups in `backups/` folder

3. **Test Changes on One User First**
   - Edit one user, save, restart
   - Verify sync works before bulk edits

4. **Use Excel Features**
   - Sort by username for easy finding
   - Filter to find incomplete training
   - Conditional formatting for quick visual checks

5. **Document Changes**
   - Keep notes of what you changed
   - Helps troubleshoot if issues arise

---

## 🧪 **Testing Manual Sync**

You can test sync without restarting the full program:

```powershell
python excel_db_sync.py
```

**Output:**
```
Manual Bidirectional Sync Test
🔄 ============================================================
   SMART BIDIRECTIONAL SYNC - Excel ↔ Database
🔄 ============================================================
📊 Excel modified:    2026-01-15 14:30:25
💾 Database modified: 2026-01-15 14:25:10
→ Excel is newer - User likely made manual edits
  Priority: Excel → Database
✓ Synced 15 users from Excel to Database
✓ Synced 2 new scans from Excel to Database
✅ Bidirectional sync complete!
```

---

## 🔐 **Data Safety**

### **Protection Mechanisms:**

1. **Backups:** System creates backups every 10 hours
2. **Database Integrity:** Foreign key constraints prevent orphaned data
3. **Duplicate Prevention:** Unique constraints on username and hardware_id
4. **Rollback Support:** Excel preserved, can always re-import
5. **Atomic Operations:** Database transactions ensure data consistency

### **Recovery:**

**If you break something:**
1. Restore Excel from `backups/` folder
2. Delete `hardware_users.db`
3. Restart program (will rebuild database from Excel)

---

## 📖 **Examples**

### **Example 1: Add 5 New Students**

**Excel Edit:**
```
Row 15: student1 | 111111 | 0 | Alice | Anderson | Biology
Row 16: student2 | 222222 | 0 | Bob | Brown | Chemistry  
Row 17: student3 | 333333 | 0 | Carol | Carter | Physics
Row 18: student4 | 444444 | 0 | David | Davis | Math
Row 19: student5 | 555555 | 0 | Eve | Evans | Engineering
```

**Save → Restart → Console:**
```
✓ Synced 5 users from Excel to Database
✅ Bidirectional sync complete!
```

### **Example 2: Update Major for 10 Students**

**Excel Edit:**
```
Find all rows with Major = "Undeclared"
Change to specific majors
Save
```

**Restart → Automatic Sync**

### **Example 3: Mark Training Complete**

**Excel Edit (Column G):**
```
Find user: jsmith
Current: {"completed": false}
Change to: {"completed": true}
Update Column H: 01/15/2026 15:00:00
Save
```

**Next scan shows updated training status**

---

## ✅ **Summary**

**You Can Now:**
- ✅ Edit Excel to update database
- ✅ Add users without scanning cards
- ✅ Bulk update user information
- ✅ Modify training status manually
- ✅ Fix data errors easily

**System Automatically:**
- ✅ Syncs Excel ↔ Database on startup
- ✅ Updates Excel on shutdown
- ✅ Preserves manual edits
- ✅ Prevents duplicates
- ✅ Maintains data integrity

**Remember:**
- ⚠️ Always close program before editing Excel
- ⚠️ Always save Excel before restarting program
- ⚠️ Check console for sync confirmation

---

**Happy editing! Your manual changes will now automatically sync to the database.** 🎉
