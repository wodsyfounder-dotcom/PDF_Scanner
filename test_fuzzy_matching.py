#!/usr/bin/env python3
"""
Test script for improved fuzzy matching functionality.

This tests the Levenshtein-based fuzzy matching for multi-word OCR search terms.
"""

import sys
import os

# Add the Application directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'Application'))

# Import the core module
import importlib.util
spec = importlib.util.spec_from_file_location(
    "eidp_term_scanner_core",
    os.path.join(os.path.dirname(__file__), 'Application', 'eidp_term_scanner.core.py')
)
core = importlib.util.module_from_spec(spec)
spec.loader.exec_module(core)

# Import the fuzzy matching functions
_levenshtein_distance = core._levenshtein_distance
_levenshtein_ratio = core._levenshtein_ratio
_fuzzy_match_multiword = core._fuzzy_match_multiword
_fuzzy_ratio = core._fuzzy_ratio
_anchor_tokens_present = core._anchor_tokens_present


def test_levenshtein_distance():
    """Test basic Levenshtein distance calculations."""
    print("=" * 60)
    print("Testing Levenshtein Distance")
    print("=" * 60)

    test_cases = [
        ("seats", "seats", 0),  # Identical
        ("seats", "seat", 1),   # One deletion
        ("seats", "seals", 2),  # Two substitutions
        ("seats", "seatss", 1), # One insertion
        ("Seats Closed", "Seat Closed", 1),  # Case + deletion
        ("chamber_pressure", "chamberpressure", 1),  # Underscore removed
    ]

    for s1, s2, expected in test_cases:
        distance = _levenshtein_distance(s1.lower(), s2.lower())
        status = "PASS" if distance == expected else "FAIL"
        print(f"[{status}] '{s1}' vs '{s2}': distance={distance} (expected={expected})")
    print()


def test_levenshtein_ratio():
    """Test Levenshtein similarity ratio."""
    print("=" * 60)
    print("Testing Levenshtein Ratio (Similarity)")
    print("=" * 60)

    test_cases = [
        ("seats", "seats", 1.0),     # Perfect match
        ("seats", "seat", 0.8),      # 1 char diff in 5 chars = 0.8
        ("seats", "seals", 0.6),     # 2 char diff in 5 chars = 0.6
        ("seats closed", "seat closed", 0.923),  # ~92% similar
    ]

    for s1, s2, expected in test_cases:
        ratio = _levenshtein_ratio(s1.lower(), s2.lower())
        status = "PASS" if abs(ratio - expected) < 0.05 else "FAIL"
        print(f"[{status}] '{s1}' vs '{s2}': ratio={ratio:.3f} (expected ~{expected:.3f})")
    print()


def test_fuzzy_match_multiword():
    """Test multi-word fuzzy matching for OCR scenarios."""
    print("=" * 60)
    print("Testing Multi-Word Fuzzy Matching")
    print("=" * 60)

    # Test cases: (search_term, ocr_text, should_match)
    test_cases = [
        # Perfect matches
        ("Seats Closed", "Seats Closed", True),
        ("Thermal Soak High", "Thermal Soak High", True),

        # OCR errors: missing letters
        ("Seats Closed", "Seat Closed", True),   # Missing 's'
        ("Seats Closed", "Seats Clsd", True),    # Missing 'ose'
        ("Chamber Pressure", "Chambe Pressure", True),  # Missing 'r'

        # OCR errors: extra spaces
        ("Seats Closed", "Seats  Closed", True),  # Extra space
        ("SeatsClose", "Seats Closed", True),     # Space added by OCR

        # OCR errors: substitutions
        ("Seats Closed", "Seals Closed", True),  # 't' -> 'l'
        ("Random Vib", "Randon Vib", True),      # 'm' -> 'n'

        # Should NOT match
        ("Seats Closed", "Doors Open", False),
        ("Thermal Soak", "Chamber Pressure", False),

        # Partial word matches (context sensitive)
        ("TC-02", "TC-02 Post Trim Data", True),  # Term is substring
        ("Harness Resistance", "resistance harness 4.8-5.3 6.1", True),  # Words reordered
    ]

    for search_term, ocr_text, should_match in test_cases:
        score, debug = _fuzzy_match_multiword(
            search_term,
            ocr_text,
            min_word_score=0.75,
            min_overall_score=0.6,
            require_all_words=True
        )

        matched = score >= 0.6
        status = "PASS" if matched == should_match else "FAIL"
        result_str = "MATCH" if matched else "NO MATCH"

        print(f"[{status}] '{search_term}' in '{ocr_text}'")
        print(f"   Score: {score:.3f} -> {result_str} (expected: {'MATCH' if should_match else 'NO MATCH'})")
        if debug.get('word_matches'):
            for wm in debug['word_matches']:
                print(f"   - '{wm['search_word']}' matched '{wm['matched_word']}' (score: {wm['score']:.3f})")
        print()


def test_fuzzy_ratio():
    """Test the updated _fuzzy_ratio function."""
    print("=" * 60)
    print("Testing Updated _fuzzy_ratio()")
    print("=" * 60)

    test_cases = [
        # Single word
        ("Program", "Program Hyperion", 0.85),

        # Multi-word
        ("Seats Closed", "Seat Closed", 0.85),
        ("Thermal Soak High", "Thermal Soak High +95", 0.75),
        ("Source Bundle", "IData Packages/Hyperion_SV3 Source Bundle", 0.70),
    ]

    for search_term, target_text, min_expected in test_cases:
        ratio = _fuzzy_ratio(search_term, target_text)
        status = "PASS" if ratio >= min_expected else "FAIL"
        print(f"[{status}] '{search_term}' vs '{target_text}'")
        print(f"   Score: {ratio:.3f} (expected >= {min_expected:.3f})")
        print()


def test_anchor_tokens_present():
    """Test the improved _anchor_tokens_present function."""
    print("=" * 60)
    print("Testing Fuzzy Token Presence Check")
    print("=" * 60)

    # Test cases: (anchor, text, should_be_present)
    test_cases = [
        # Exact matches
        ("Seats Closed", "Seats Closed", True),
        ("TC-02", "TC-02 Post Trim", True),

        # Fuzzy matches
        ("Seats Closed", "Seat Closed", True),  # Missing 's'
        ("Chamber Pressure", "Chambe Pressure Mean", True),  # Missing 'r'
        ("Thermal Soak", "Therma Soak High", True),  # Missing 'l'

        # Should NOT match
        ("Seats Closed", "Doors", False),
        ("Chamber Pressure", "Thermal Margin", False),

        # Partial matches (at least half the tokens)
        ("Thermal Soak High", "Thermal Soak", True),  # 2 of 3 words
        ("Random Vib Z", "Random Vib", True),  # 2 of 3 words
    ]

    for anchor, text, should_be_present in test_cases:
        present = _anchor_tokens_present(anchor, text, fuzzy_threshold=0.8)
        status = "PASS" if present == should_be_present else "FAIL"
        result_str = "PRESENT" if present else "NOT PRESENT"

        print(f"[{status}] '{anchor}' in '{text}'")
        print(f"   Result: {result_str} (expected: {'PRESENT' if should_be_present else 'NOT PRESENT'})")
        print()


def main():
    """Run all tests."""
    print("\n" + "=" * 60)
    print("FUZZY MATCHING TEST SUITE")
    print("=" * 60 + "\n")

    test_levenshtein_distance()
    test_levenshtein_ratio()
    test_fuzzy_match_multiword()
    test_fuzzy_ratio()
    test_anchor_tokens_present()

    print("=" * 60)
    print("All tests completed!")
    print("=" * 60)


if __name__ == '__main__':
    main()
