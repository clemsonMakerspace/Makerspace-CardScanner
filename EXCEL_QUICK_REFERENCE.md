# Quick Reference: Excel Database Editing

## 🎯 **Quick Steps to Edit User Data**

### **1. Close Scanner Program**
```
Stop MakerspaceSignInTablet.py or CardReaderMakerspace.py
```

### **2. Open Excel**
```
Open: hardware_users.xlsx
Sheet: Users (bottom tab)
```

### **3. Make Changes**
```
Edit any user data in columns A-F:
- Username (A)
- Hardware ID (B)
- First Name (D)
- Last Name (E)
- Major (F)
```

### **4. Save & Close**
```
File → Save → Close Excel
```

### **5. Restart Scanner**
```
Run program again
Check console for: "✓ Synced X users from Excel to Database"
```

---

## 📊 **Column Reference**

| Column | Field | Example | Notes |
|--------|-------|---------|-------|
| **A** | Username | `jsmith` | Must be unique |
| **B** | Hardware ID | `123456` | Must be unique, 6 digits |
| **C** | Login Count | `15` | Auto-updated by system |
| **D** | First Name | `Jane` | Editable |
| **E** | Last Name | `Smith` | Editable |
| **F** | Major | `Computer Science` | Editable |
| **G** | Training Data | `{JSON}` | Advanced - JSON format |
| **H** | Training Updated | `01/15/2026 14:30` | Auto-updated |

---

## ⚡ **Common Tasks**

### **Add New User**
```
Add row with: Username | Hardware_ID | 0 | First | Last | Major
Example: jdoe | 789012 | 0 | John | Doe | Engineering
```

### **Update User Info**
```
Find user row → Edit columns D, E, F → Save
```

### **Fix Typo**
```
Double-click cell → Edit → Enter → Save
```

### **Remove User**
```
Right-click row number → Delete → Save
```

---

## ✅ **Verification Checklist**

- [ ] Scanner program is closed
- [ ] Excel file saved after edits
- [ ] Excel file closed completely
- [ ] Restarted scanner program
- [ ] Console shows sync message
- [ ] Tested with card scan (works!)

---

## ⚠️ **Rules**

✅ **DO:** Edit when program is closed  
✅ **DO:** Save before closing Excel  
✅ **DO:** Keep usernames unique  
✅ **DO:** Keep hardware IDs unique  

❌ **DON'T:** Edit while program running  
❌ **DON'T:** Delete column headers  
❌ **DON'T:** Create duplicate usernames  
❌ **DON'T:** Create duplicate hardware IDs  

---

## 🔧 **Troubleshooting**

**Changes not appearing?**
→ Check console for sync confirmation
→ Verify you saved Excel
→ Try `python excel_db_sync.py`

**File locked error?**
→ Close scanner program
→ Close Excel completely
→ Check Task Manager for processes

**Duplicate error?**
→ Check for duplicate usernames or hardware IDs
→ Fix in Excel and restart

---

## 🔄 **When Syncs Happen**

| Event | Direction | What Happens |
|-------|-----------|--------------|
| **Startup** | Smart | Newer file updates older |
| **Shutdown** | Database → Excel | Export latest scans |
| **Each Scan** | Write to DB | Fast performance |
| **Manual** | Both ways | Test: `python excel_db_sync.py` |

---

## 📞 **Need Help?**

See full guide: **EXCEL_EDITING_GUIDE.md**

---

**Remember: Close program → Edit Excel → Save → Restart program** ✨
