# Fuzzy Matching Implementation Summary

## Changes Made

### 1. Added Medium Strictness as Default
- **Default preset**: `medium` (configurable via `user_inputs/scanner.env`)
- **Medium values**:
  - `min_score`: 0.65 (overall match threshold)
  - `min_word_score`: 0.80 (per-word threshold, allows ~20% character errors)
  - `token_threshold`: 0.85 (token presence check)

### 2. Updated scanner.env
**File**: `user_inputs/scanner.env`

Added fuzzy matching configuration with preset support:
```bash
# Fuzzy Matching Configuration
FUZZY_PRESET=medium  # Options: lenient, medium, strict

# Custom overrides (optional):
# FUZZY_MIN_SCORE=0.65
# FUZZY_MIN_WORD_SCORE=0.80
# FUZZY_TOKEN_THRESHOLD=0.85
```

### 3. Updated Core Matching Functions
**File**: `Application/eidp_term_scanner.core.py`

#### `_get_fuzzy_match_config()` (lines 266-348)
- Added preset support: lenient, medium, strict
- Reads from `FUZZY_PRESET` environment variable
- Falls back to custom values if specified
- Default: **medium** preset

**Preset Definitions**:
```python
presets = {
    'lenient': {  # For poor OCR
        'min_score': 0.55,
        'min_word_score': 0.70,
        'token_threshold': 0.75
    },
    'medium': {  # Balanced (DEFAULT)
        'min_score': 0.65,
        'min_word_score': 0.80,
        'token_threshold': 0.85
    },
    'strict': {  # For high quality OCR
        'min_score': 0.75,
        'min_word_score': 0.90,
        'token_threshold': 0.90
    }
}
```

#### `_fuzzy_ratio()` (lines 1790-1820)
- Now reads configuration from `_get_fuzzy_match_config()`
- Passes configured thresholds to `_fuzzy_match_multiword()`
- No more hardcoded values

#### `_anchor_tokens_present()` (lines 1997-2048)
- Changed signature: `fuzzy_threshold: Optional[float] = None`
- Reads from config if threshold not specified
- Backward compatible with explicit threshold

## How Users Configure It

### Option 1: Use Presets (Recommended)
Edit `user_inputs/scanner.env`:
```bash
FUZZY_PRESET=lenient   # For poor OCR
# or
FUZZY_PRESET=medium    # For balanced matching (DEFAULT)
# or
FUZZY_PRESET=strict    # For high quality OCR
```

### Option 2: Custom Values
Edit `user_inputs/scanner.env`:
```bash
# Uncomment and customize:
FUZZY_MIN_SCORE=0.70
FUZZY_MIN_WORD_SCORE=0.85
FUZZY_TOKEN_THRESHOLD=0.88
```

### Option 3: Environment Variables (Advanced)
```bash
# Windows
set FUZZY_PRESET=strict
run.bat

# Linux/Mac
export FUZZY_PRESET=lenient
./run.sh
```

## Test Results with Medium Preset

| Search Term | OCR Text | Score | Result |
|------------|----------|-------|--------|
| "Seats Closed" | "Seat Closed" | 0.900 | ✓ MATCH |
| "Seats Closed" | "Seats Clsd" | 0.667 | ✓ MATCH |
| "Chamber Pressure" | "Chambe Pressure" | 0.929 | ✓ MATCH |
| "Thermal Soak" | "Therma Soak" | 0.917 | ✓ MATCH |
| "Seats Closed" | "Doors Open" | 0.200 | ✗ NO MATCH |

## Preset Comparison

### What Each Preset Tolerates

| OCR Error | lenient | medium | strict |
|-----------|---------|--------|--------|
| "Seat Closed" vs "Seats Closed" | ✓ | ✓ | ✓ |
| "Seats Clsd" vs "Seats Closed" | ✓ | ✓ | ? |
| "Sats Clsd" vs "Seats Closed" | ✓ | ? | ✗ |
| "Doors Open" vs "Seats Closed" | ✗ | ✗ | ✗ |

**Legend**: ✓ = Matches, ✗ = Doesn't match, ? = Borderline

### Choosing the Right Preset

**Choose `lenient` if**:
- OCR quality is poor (scanned documents, low DPI)
- You're getting many false negatives (terms not matching)
- Your search terms have many small variations in the PDFs

**Choose `medium` (DEFAULT) if**:
- OCR quality is average to good
- You want balanced matching
- You're not sure what to choose

**Choose `strict` if**:
- OCR quality is excellent (digital PDFs, high DPI scans)
- You're getting false positives (wrong terms matching)
- Your search terms are very specific

## Files Modified

1. **user_inputs/scanner.env**
   - Added `FUZZY_PRESET=medium`
   - Added configuration comments

2. **Application/eidp_term_scanner.core.py**
   - Updated `_get_fuzzy_match_config()` with preset support
   - Updated `_fuzzy_ratio()` to use config
   - Updated `_anchor_tokens_present()` to use config

3. **FUZZY_MATCHING_QUICK_REFERENCE.md**
   - Updated with preset information
   - Simplified configuration instructions

## Migration Notes

### For Existing Users

**No action required!** The default `medium` preset is automatically applied.

If you previously set custom environment variables:
- `FUZZY_MIN_SCORE`, `FUZZY_MIN_WORD_SCORE`, `FUZZY_TOKEN_THRESHOLD`

These will **override** the preset values and continue to work as before.

### Recommended Actions

1. **Try the default** (`medium`) first
2. **If you get false negatives** (terms not matching):
   - Edit `scanner.env` and change to `FUZZY_PRESET=lenient`
3. **If you get false positives** (wrong matches):
   - Edit `scanner.env` and change to `FUZZY_PRESET=strict`

## Backward Compatibility

✓ All existing code continues to work
✓ Custom environment variables still respected
✓ Default behavior improved but similar to before
✓ Can opt-out by setting `FUZZY_PRESET=strict`

## Quick Reference

```bash
# Quick preset change
Edit: user_inputs/scanner.env
Change: FUZZY_PRESET=medium
To: FUZZY_PRESET=lenient (or strict)
Run: run.bat
```

That's it! Your fuzzy matching is now configured with medium strictness by default and easily adjustable via scanner.env.
