# Location-Specific Background Images - Implementation Summary

## Date: January 14, 2026

## Overview
Added support for location-specific background images. When the `Location` variable is set to "Cooper", the application will display `BackgroundAdobe.png` instead of the default backgrounds.

## Changes Made

### 1. CardReaderMakerspace.py (Training Popup)
- **Line ~290-298**: Modified background image loading logic
- **Before**: Always loaded `backgroundLarge.png`
- **After**: Checks `Location` variable:
  - If `Location == "Cooper"` → uses `BackgroundAdobe.png`
  - Otherwise → uses `BackgroundWatt.png`

### 2. MakerspaceSignInTablet.py (Sign-In Screen)
- **Two locations updated**:

  **A. Popup Window (~Line 66-75)**:
  - Before: Always loaded `backgroundLarge.png`
  - After: Location-based selection (same as above)
  
  **B. Main Tablet Screen (~Line 122-128)**:
  - Before: Always loaded `BackgroundTablet.png`
  - After: Checks `Location` variable:
    - If `Location == "Cooper"` → uses `BackgroundAdobe.png`
    - Otherwise → uses `BackgroundTablet.png`

### 3. setup.py (Build Configuration)
- **Line ~28**: Added `BackgroundAdobe.png` to the include_files list
- Ensures the new background is bundled when creating portable builds

## How to Use

### To switch to Cooper location:
1. Open `CardReaderMakerspace.py`
2. Change line 49 from:
   ```python
   Location = "Watt"
   ```
   to:
   ```python
   Location = "Cooper"
   ```

3. Open `MakerspaceSignInTablet.py`
4. Change line 26 similarly

### Adding More Locations:
To add backgrounds for other locations, update the conditional logic:
```python
if Location == "Cooper":
    background_filename = "BackgroundAdobe.png"
elif Location == "NewLocation":
    background_filename = "BackgroundNewLocation.png"
else:
    background_filename = "BackgroundWatt.png"
```

## Files Affected
- ✓ CardReaderMakerspace.py
- ✓ MakerspaceSignInTablet.py
- ✓ setup.py

## Testing
- ✓ Background selection logic verified with test script
- ✓ All required background files present:
  - BackgroundTablet.png (default for tablet)
  - BackgroundAdobe.png (Cooper location)
  - BackgroundWatt.png (Watt location / default for popups)

## Notes
- The `Location` variable must be set in both files to ensure consistency
- Background images are automatically resized to fit the screen
- No changes needed to database or Excel structure
- Backward compatible - works with existing setups
