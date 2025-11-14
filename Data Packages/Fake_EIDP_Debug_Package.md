# Synthetic End Item Data Package (Debug Reference)

This document intentionally blends nominal, off-nominal, and incomplete data so the scanner can be exercised against a broad set of scenarios (duplicate terms, missing units, multi-value tables, etc.). Only synthetic values are used.

## Document Profile

| Field | Value |
| --- | --- |
| Program | Hyperion Dragonfly Propulsion Demo |
| Vehicle | SV3 (Block 2) |
| Serial / Component | SN42-AX |
| Build Revision | Rev E (04) |
| Date Compiled | 2025-01-17 |
| Compiler | QA Test Harness |
| Source Bundle | /Data Packages/Hyperion_SV3_SN42_RevE |

---

## 1. Functional Acceptance Snapshot

| Term | Description | Requirement | Measured Value | Units | Page | Data Quality | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| ignition_delay | Time between arm and ignition | ≤ 0.85 | 0.62 | s | 5 | PASS | sample nominal case |
| isp | Steady-state specific impulse | 43 ± 3 | 39.4 | sec | 6 | LOW | flagged low but inside tolerance window |
| chamber_pressure | Mean P<sub>cc</sub> | 185 ± 10 | 201 | psia | 6 | WARN | value above requirement, watch for drift |
| burn_duration | Hot-fire duration | 42–48 | 0 (not testable) | s | 6 | NOT_RUN | use to test textual statuses |
| resistance | Injector harness resistance | 4.8–5.3 | 6.1 | Ω | 8 | FAIL | triggers out-of-range numeric failure |
| valve_alignment | Valve alignment photo | Visual match | Image #12 | n/a | 9 | PASS | non-numeric text extraction |
| leak_check | Helium leak rate | ≤ 1.0e-5 | **Not Found** | atm·cc/s | 10 | MISSING | ensures missing term handling |

---

## 2. Performance KPI Table

| KPI | Min | Target | Max | Value | Source | Confidence |
| --- | --- | --- | --- | --- | --- | --- |
| thrust_nominal | 4.8 | 5.1 | 5.5 | 5.32 | Table 3 | 0.97 |
| thrust_peak | – | 6.2 | 6.8 | 6.95 | Table 3 | 0.91 |
| power_draw_idle | 75 | 80 | 95 | 104 | Table 4 | 0.77 |
| power_draw_peak | – | 140 | 155 | 131 | Table 4 | 0.85 |
| thermal_margin | 8 | 12 | – | 6 | Fig. 5 | 0.63 |

> Use this table to confirm multi-column parsing and comparisons against asymmetric spec bands (no min or max defined).

---

## 3. Environmental Endurance Matrix

| Test | Temp (°C) | Duration (hr) | Requirement Met? | Deviation | Log Ref |
| --- | --- | --- | --- | --- | --- |
| Thermal Soak Low | -45 | 2.0 | YES | ±0.5 | A-112 |
| Thermal Soak High | +95 | 1.5 | NO | Chamber tripped at +88°C | A-118 |
| Thermal Ramp Cycling | -35↔+85 | 6 cycles | YES | – | A-126 |
| Random Vib (X) | 14.1 grms | 7 min | YES | – | V-031 |
| Random Vib (Z) | 12.0 grms | 7 min | CONDITIONAL | Sensor SAT | V-035 |
| Sine Sweep | 5–500 Hz | 0.5 g | **Not Recorded** | – | V-041 |

Mix of numeric, range (arrow), textual, and missing entries stresses parser assumptions.

---

## 4. Calibration & Trim Table

| Channel | Pre Trim | Post Trim | Allowed Delta | Status | Comment |
| --- | --- | --- | --- | --- | --- |
| TC-01 | 1.0034 | 0.9991 | ±0.0100 | PASS | Temperatures within tolerance |
| TC-02 | 0.9989 | 1.0142 | ±0.0100 | FAIL | Out-of-family high |
| PT-03 | 1.0000 | 1.0000 | ±0.0020 | PASS | Perfect match |
| PT-07 | 0.9977 | 0.9974 | ±0.0020 | INVESTIGATE | Delta close to limit |
| LVDT-2 | 1.0040 | – | ±0.0050 | INCOMPLETE | Missing post-trim data |

---

## 5. Anomaly & Waiver Log

| ID | Severity | Term | Description | Containment | Status |
| --- | --- | --- | --- | --- | --- |
| DR-2025-014 | Major | chamber_pressure | Persistent +8% over spec during burn #2 | Added relief orifice, retest pending | OPEN |
| NCR-25-008 | Minor | resistance | Harness swapped due to 6.1 Ω reading | Re-test verified 5.05 Ω | CLOSED |
| WA-25-002 | Info | leak_check | Leak test deferred pending new fixture | Approved by Chief QA | APPROVED |

---

## 6. Attachment Index (Synthetic)

| Attachment | File Name | Pages | Notes |
| --- | --- | --- | --- |
| Acceptance Test Report | ATR_Hyperion_SV3_SN42.pdf | 124 | Contains Sections 3 & 4 tables |
| Thermal Test Log | THERM_SV3_SN42_revC.pdf | 56 | Page numbers offset by +2 in footer |
| Vib Test Photos | VIB_IMG_SN42.zip | 24 images | Use to test non-PDF assets |
| Calibration Data | CAL_SN42_revE.xlsx | 4 sheets | Multi-tab workbook |

---

### Usage Hints

1. Place this Markdown file alongside PDFs to validate mixed content detection.
2. Convert to PDF (e.g., via pandoc) if you want to test page-scoped extraction.
3. Adjust term lists so at least one expected value lives in every section above.
