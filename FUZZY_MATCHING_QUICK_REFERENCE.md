# Fuzzy Matching Quick Reference

## Problem Solved

**Before**: Search terms like "Seats Closed" failed to match OCR text with small errors:
- "Seat Closed" (missing 's') → ✗ NO MATCH
- "Seats Clsd" (OCR error) → ✗ NO MATCH
- "Seats  Closed" (extra space) → ✗ NO MATCH

**After**: Same searches now match correctly:
- "Seat Closed" → ✓ MATCH (score: 0.900)
- "Seats Clsd" → ✓ MATCH (score: 0.667)
- "Seats  Closed" → ✓ MATCH (score: 1.000)

## What Changed

1. **Levenshtein Distance** - Better character-level similarity matching
2. **Word-by-Word Matching** - Multi-word terms matched intelligently
3. **Fuzzy Token Presence** - Allows ~15% character errors per word (medium preset)
4. **Configurable via scanner.env** - Easy preset selection or custom tuning

## How to Use

### Default Behavior (Medium Strictness)

Just run your scanner as normal. The improved matching uses **medium** strictness by default:

```bash
run.bat
```

Configuration is in **user_inputs/scanner.env** with `FUZZY_PRESET=medium` (already set).

### Quick Preset Changes

Edit `user_inputs/scanner.env` and change the `FUZZY_PRESET` line:

**For Poor OCR Quality** (more lenient):
```
FUZZY_PRESET=lenient
```

**For High OCR Quality** (stricter):
```
FUZZY_PRESET=strict
```

**For Balanced (default)**:
```
FUZZY_PRESET=medium
```

Then run normally:
```bash
run.bat
```

### Custom Fine-Tuning

For advanced users, uncomment and modify individual parameters in `scanner.env`:

```bash
# Example: Custom strictness between lenient and medium
FUZZY_MIN_SCORE=0.60
FUZZY_MIN_WORD_SCORE=0.75
FUZZY_TOKEN_THRESHOLD=0.80
```

## Configuration Presets

| Preset | Use Case | min_score | min_word_score | token_threshold |
|--------|----------|-----------|----------------|-----------------|
| **lenient** | Poor OCR quality | 0.55 | 0.70 | 0.75 |
| **medium** | Balanced (default) | 0.65 | 0.80 | 0.85 |
| **strict** | High OCR quality | 0.75 | 0.90 | 0.90 |

**Lower values** = More lenient (fewer false negatives, more false positives)
**Higher values** = Stricter (fewer false positives, more false negatives)

## Testing

Verify the improvements work:

```bash
python test_fuzzy_matching.py
```

Expected: ~90% of tests should PASS

## Examples

### Multi-Word Terms

| Your Search Term | Will Match OCR Text |
|------------------|---------------------|
| "Seats Closed" | "Seat Closed", "Seats Clsd", "Seals Closed" |
| "Thermal Soak High" | "Thermal Soak High", "Therma Soak High" |
| "Chamber Pressure" | "Chambe Pressure", "Chamber Presure" |
| "Random Vib Z" | "Randon Vib Z", "Random Vib (Z)" |

### Single-Word Terms

| Your Search Term | Will Match OCR Text |
|------------------|---------------------|
| "Program" | "Program", "Progran", "Pragram" |
| "Serial" | "Serial", "Seria1", "Serlal" |
| "TC-02" | "TC-02", "TC-O2", "TC 02" |

## Troubleshooting

### Terms Still Not Matching

1. Check the `debug_info` in scan results JSON
2. Look at the `smart_snap_context` field to see what text was matched
3. Lower thresholds: `set FUZZY_MIN_WORD_SCORE=0.7`
4. Run tests to verify: `python test_fuzzy_matching.py`

### Too Many Wrong Matches

1. Increase thresholds: `set FUZZY_MIN_SCORE=0.7`
2. Check if search terms are too generic (e.g., "Test" matches "Best", "Rest", etc.)
3. Use more specific multi-word terms (e.g., "Test Results" instead of "Test")

## Key Files

- **eidp_term_scanner.core.py** - Main fuzzy matching implementation
- **test_fuzzy_matching.py** - Test suite
- **FUZZY_MATCHING_IMPROVEMENTS.md** - Detailed documentation

## Quick Diagnostic

Run this to see if fuzzy matching is working:

```python
from Application.eidp_term_scanner.core import _fuzzy_ratio

# Should score ~0.9 (high similarity)
score = _fuzzy_ratio("Seats Closed", "Seat Closed")
print(f"Score: {score}")  # Expected: ~0.9

# Should score <0.3 (low similarity)
score = _fuzzy_ratio("Seats Closed", "Doors Open")
print(f"Score: {score}")  # Expected: ~0.2
```

## Support

- See detailed documentation: `FUZZY_MATCHING_IMPROVEMENTS.md`
- Run tests: `python test_fuzzy_matching.py`
- Check scan results JSON for `debug_info.smart_snap_context`
