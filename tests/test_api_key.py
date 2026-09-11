import os
import sys

# Add the project root to the path. Computed from this file so the test
# keeps working if the project is moved or cloned somewhere else.
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.logic.api_key import check_key, EMPTY, LEGACY, MALFORMED, OK

# A new-format auth key (AQ.Ab...) and an old standard key (AIza..., 39
# chars). Both are fake, only the shape matters.
#
# Assembled at runtime instead of written out as one literal: a string
# shaped like a real key trips GitHub's secret scanner and blocks the push,
# which is exactly the behavior you want from it. Do not "fix" this by
# inlining the strings.
NEW_KEY = "AQ." + "Ab" + "n" * 45
OLD_KEY = "AIza" + "Sy" + "x" * 33


def check(label, key, expected):
    result = check_key(key)
    status = "PASS" if result == expected else "FAIL"
    shown = (key[:12] + "...") if key and len(key) > 12 else repr(key)
    print(f"{label}: {shown} | Result: {result} | Expected: {expected} | {status}")
    return result == expected


if __name__ == "__main__":
    results = [
        # The whole point of the fix: the 2026 key format must be accepted.
        check("New auth key", NEW_KEY, OK),
        # Old keys still run, with a warning, so nobody is locked out mid-migration.
        check("Old standard key", OLD_KEY, LEGACY),
        # Length is no longer a rule: a key longer or shorter than 39 is fine.
        check("Unknown future format", "ZZ.SomeFutureFormat_2030-abcdefghijkl", OK),
        # Paste accidents.
        check("Empty", "", EMPTY),
        check("Only spaces", "   ", EMPTY),
        check("None", None, EMPTY),
        check("Too short", NEW_KEY[:8], MALFORMED),
        check("Pasted a sentence", f"my api key is {NEW_KEY}", MALFORMED),
        check("Contains a line break", NEW_KEY[:20] + "\n" + NEW_KEY[20:], MALFORMED),
        # Surrounding whitespace is trimmed, not rejected.
        check("Padded with spaces", f"  {NEW_KEY}  ", OK),
    ]

    passed = sum(results)
    print(f"\n{passed}/{len(results)} passed")
    sys.exit(0 if passed == len(results) else 1)
