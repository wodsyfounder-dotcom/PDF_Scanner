```powershell
python debug\ocr_debug_runner.py path\to\file.pdf --pages 1


PS C:\Users\zachs\Documents\DevProjects\PDF_Scanner> $env:PYTHONPATH="$PWD\EIDAT_App_Files;$env:PYTHONPATH"
>> $env:DEBUG_MODE="1"
>> python -m ui_next.qt_main



C:\Users\zachs\Documents\DevProjects\PDF_Scanner\ocr_runs_stress.py

eidp_term_scanner.core.py --reset-state --reset-confirm RESET [--reset-include-debug]
```

For **numeric** values without Smart Position:

 Builds candidates from right-of-label tokens and scores each:

 **Scoring Components (max ~4.7 points):**

| Component                  | Max Points | Description                                                                          |
| -------------------------- | ---------- | ------------------------------------------------------------------------------------ |
| **Secondary Header** | 2.0        | X-distance from secondary_term column header (exponential decay:`1.0/(1+dist/20)`) |
| **Value Header**     | 2.0        | Fallback when no secondary term: distance from "Value" header                        |
| **Range Validation** | 2.0        | Within range=2.0; 20-50% outside=-10% penalty; boundary match=1.0                    |
| **Format Match**     | 0.2        | Matches user-specified format pattern                                                |
| **Units Hint**       | 0.4        | Units match configured hint                                                          |
| **Label Proximity**  | 0.1        | X-distance from label end (`1.0/(1+dx/10)`)                                        |

**Nullifier Rule** ([lines 3384-3388](vscode-webview://0qetrc4n0jh91juptl6gu53q2rk10djvnfktbk0m54e3sbjgj4fd/Application/eidp_term_scanner.core.py#L3384-L3388)):

* Candidates **>50% outside** configured range are completely excluded
* If all candidates nullified, extraction fails with diagnostic

**Tie-Breaking:**

* If multiple candidates score within 0.1 points → marks as conflict
* Prioritizes candidates with matching units hints
* Returns highest scorer

---

### Alt Search Direction ([lines 2585-2673](vscode-webview://0qetrc4n0jh91juptl6gu53q2rk10djvnfktbk0m54e3sbjgj4fd/Application/eidp_term_scanner.core.py#L2585-L2673))

When `alt_search="below"` or `"above"` for **numeric** Smart Snap:

1. Finds anchor row by Y-position
2. Scans rows vertically in specified direction
3. Searches **entire row** for numeric values
4. Returns **first value** within configured range
5. Stops after first match

Use case: Values in separate row from labels (e.g., vertically stacked tables)

---

### Smart Type Detection ([lines 2526-2537](vscode-webview://0qetrc4n0jh91juptl6gu53q2rk10djvnfktbk0m54e3sbjgj4fd/Application/eidp_term_scanner.core.py#L2526-L2537))

If `smart_snap_type` not specified, auto-detects by priority:

1. **date** → if DATE_REGEX matches
2. **time** → if TIME_REGEX matches
3. **number** → if NUMBER_REGEX matches
4. **title** → default fallback

Each type uses different extraction patterns:

* **number** : Extracts first numeric match, validates range, extracts units
* **date/time** : Regex pattern matching
* **title/text** : Returns all text after label (or uses format pattern if specified)

---

### Key Innovations

1. **Header-aware extraction** : Uses column header positions for robust tabular data extraction
2. **Label boundary extension** ([line 2999](vscode-webview://0qetrc4n0jh91juptl6gu53q2rk10djvnfktbk0m54e3sbjgj4fd/Application/eidp_term_scanner.core.py#L2999)): Handles split labels like "Serial / Component"
3. **Units neighbor detection** ([lines 3068-3092](vscode-webview://0qetrc4n0jh91juptl6gu53q2rk10djvnfktbk0m54e3sbjgj4fd/Application/eidp_term_scanner.core.py#L3068-L3092)): Looks ahead 3 tokens to find units in adjacent columns
4. **Proportional scoring** : Secondary header uses distance-based decay for robust column alignment
5. **Range-based nullification** : Prevents obviously wrong values from being selected
