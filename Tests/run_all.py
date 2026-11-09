"""
MCP Suite Unified Test Runner
==============================
Executes test suites across all 5 MCP servers:
  1. Blender MCP
  2. Excel MCP
  3. FreeCAD MCP
  4. PowerPoint MCP
  5. Word MCP

Usage:
  python Tests/run_all.py             # Runs all 5 test suites
  python Tests/run_all.py blender     # Runs only Blender tests
  python Tests/run_all.py excel       # Runs only Excel tests
  python Tests/run_all.py freecad     # Runs only FreeCAD tests
  python Tests/run_all.py powerpoint  # Runs only PowerPoint tests
  python Tests/run_all.py word        # Runs only Word tests
"""

import sys
import os
import subprocess
import time

TEST_SUITES = {
    "blender": {
        "name": "Blender MCP",
        "file": "test_blender.py",
        "description": "FastMCP tools, argument schemas, mock bridge, and syntax checks",
    },
    "excel": {
        "name": "Excel MCP",
        "file": "test_excel.py",
        "description": "Workbook, styling, formula, multi/single chart, analysis, and COM/openpyxl engine",
    },
    "freecad": {
        "name": "FreeCAD MCP",
        "file": "test_freecad.py",
        "description": "All 29 tools, headless CAD pipeline, RPC commands, HTTP bridge security",
    },
    "powerpoint": {
        "name": "PowerPoint MCP",
        "file": "test_powerpoint.py",
        "description": "All 26 tools, presentations, cards, tables, charts, timelines, themes, templates",
    },
    "word": {
        "name": "Word MCP",
        "file": "test_word.py",
        "description": "All 22 tools, styles, headings, tables, callouts, TOC, comments, COM/docx engine",
    },
}


def run_suite(key: str, info: dict) -> bool:
    test_file = os.path.join(os.path.dirname(__file__), info["file"])
    print("\n" + "=" * 65)
    print(f"  RUNNING: {info['name']} ({info['file']})")
    print(f"  Scope:   {info['description']}")
    print("=" * 65 + "\n")

    start_time = time.time()
    result = subprocess.run([sys.executable, test_file])
    duration = time.time() - start_time

    success = result.returncode == 0
    status_str = "PASSED" if success else "FAILED"
    print(f"\n>> {info['name']}: {status_str} in {duration:.2f}s (Exit code: {result.returncode})\n")
    return success


def main():
    args = [a.lower().strip() for a in sys.argv[1:] if not a.startswith("-")]

    targets = args if args else list(TEST_SUITES.keys())
    for t in targets:
        if t not in TEST_SUITES:
            print(f"Unknown target: '{t}'. Available targets: {', '.join(TEST_SUITES.keys())}")
            sys.exit(1)

    print("\n" + "#" * 65)
    print("   MONOREPO MCP TEST SUITE - UNIFIED RUNNER")
    print(f"   Executing {len(targets)} test suite(s): {', '.join(targets)}")
    print("#" * 65)

    results = {}
    total_start = time.time()

    for key in targets:
        results[key] = run_suite(key, TEST_SUITES[key])

    total_duration = time.time() - total_start

    print("\n" + "#" * 65)
    print("   FINAL TEST RESULTS SUMMARY")
    print("#" * 65)
    all_passed = True
    for key, passed in results.items():
        name = TEST_SUITES[key]["name"]
        status = "PASSED" if passed else "FAILED"
        symbol = "[OK]" if passed else "[FAIL]"
        if not passed:
            all_passed = False
        print(f"  {symbol} {name:<20}: {status}")

    print("-" * 65)
    print(f"  Total Duration: {total_duration:.2f}s")
    if all_passed:
        print("  Status: ALL TEST SUITES PASSED SUCCESSFULLY!")
    else:
        print("  Status: SOME TEST SUITES FAILED.")
    print("#" * 65 + "\n")

    sys.exit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
