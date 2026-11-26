"""Dataset quality gate shared by the Hungary jobs.

OK passes. STATIC_REFERENCE and SOURCE_PARTIAL (upstream publishes nothing for that
day or direction) are warnings, never counted OK. Everything else (PARTIAL, FAILED,
SKIPPED, unknown) is a problem and turns the job red.
"""
import json
import sys

WARN_STATES = ("STATIC_REFERENCE", "SOURCE_PARTIAL")


def issues(status):
    details = status.get("dataset_details", {})
    if not details:
        return ["Dataset state missing"]
    out = []
    for name, item in details.items():
        state = item.get("state", "")
        if state == "OK":
            continue
        if state in WARN_STATES:
            print(f"::warning::{name}: {state}: {item.get('detail', '')}")
            continue
        out.append(f"{name}: {state}: {item.get('detail', '')}")
    return out


if __name__ == "__main__":
    with open(sys.argv[1], encoding="utf-8") as f:
        found = issues(json.load(f))
    for p in found:
        print(f"::error::{p}")
    sys.exit(1 if found else 0)
