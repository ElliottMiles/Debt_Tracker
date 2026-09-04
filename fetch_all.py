"""
Runs every data-fetch script in this project, one after another, and prints
a combined summary. Does NOT run build_report.py -- rebuilding the HTML
report from whatever's now on disk is still a separate, deliberate step
(`python build_report.py`), same as always.

Each fetch script runs as its own subprocess, not an in-process import:
  - fetch_yield_curve.py reads an optional start-year argument straight from
    sys.argv[1]. Importing it and calling its main() directly would have it
    read *this* script's arguments instead of its own -- a real bug, not a
    style nitpick. Subprocesses each get a clean, correct argv.
  - A crash or hang in one script can't take down the others, and each one's
    output stays clearly attributed to it instead of interleaving.

A failed script doesn't stop the rest from running -- every source is
independent, so one broken fetch (a changed API, a network hiccup) shouldn't
block updating everything else. The summary at the end reports success/
failure per script, and the exit code is nonzero if anything failed, so this
is still safe to use from cron/CI even though it presses on past failures.

Usage:
    python fetch_all.py
"""

import subprocess
import sys
import time
from pathlib import Path

SCRIPTS = [
    "fetch_auctions.py",
    "fetch_yield_curve.py",
]

DIR = Path(__file__).resolve().parent


def run_script(name):
    # flush=True: without it, these prints sit in this process's own stdout
    # buffer until it exits, while the subprocess below writes straight to
    # the shared stdout and flushes at its own exit -- so without flushing
    # here, every script's output would appear to happen before any of this
    # script's own headers, out of chronological order. Only shows up when
    # stdout isn't a live terminal (piped, redirected to a log, cron), which
    # is exactly when getting the order right matters most.
    print(f"\n{'=' * 60}", flush=True)
    print(f"Running {name}", flush=True)
    print("=" * 60, flush=True)
    t0 = time.time()
    result = subprocess.run([sys.executable, str(DIR / name)], cwd=DIR)
    elapsed = time.time() - t0
    ok = result.returncode == 0
    status = "OK" if ok else f"FAILED (exit code {result.returncode})"
    print(f"-- {name}: {status} in {elapsed:.1f}s", flush=True)
    return ok


def main():
    results = {name: run_script(name) for name in SCRIPTS}

    print(f"\n{'=' * 60}")
    print("Summary")
    print("=" * 60)
    for name, ok in results.items():
        print(f"  {'ok  ' if ok else 'FAIL'}  {name}")

    if all(results.values()):
        print("\nAll data sources updated. Run `python build_report.py` to refresh the report.")
    else:
        failed = [name for name, ok in results.items() if not ok]
        print(f"\n{len(failed)} script(s) failed: {', '.join(failed)}. See output above for details.")
        sys.exit(1)


if __name__ == "__main__":
    main()
