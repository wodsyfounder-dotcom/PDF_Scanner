#!/usr/bin/env python3
"""
EIDP Term Scanner (Application deliverable)
This file is a self-contained copy of the main scanner script so that
the /Application folder can be transferred as the deliverable.
"""

# NOTE: For brevity in this patch, this wrapper delegates to the root script if present
# and otherwise informs the user. If you are running from the deliverable that only
# contains /Application, place eidp_term_scanner.py here (already done during packaging).

import runpy
from pathlib import Path
import sys

HERE = Path(__file__).resolve().parent
LOCAL = HERE / "eidp_term_scanner.core.py"
ROOT = HERE.parents[0]
LEGACY = ROOT / "eidp_term_scanner.py"

if LOCAL.exists():
    runpy.run_path(str(LOCAL), run_name="__main__")
elif LEGACY.exists():
    sys.path.insert(0, str(ROOT))
    runpy.run_path(str(LEGACY), run_name="__main__")
else:
    raise SystemExit("No scanner core found in /Application. Expected 'eidp_term_scanner.core.py'.")
