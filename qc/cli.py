"""Command-line audit runner.

    python -m qc.cli deck.pptx --profile prezlab_en [--modules font,color_palette]
                               [--json manifest.json] [--verbose]

This is the developer/test harness for Flow B (audit only). The web API in
the PRD wraps this same engine.
"""

import argparse
import sys
from pathlib import Path

from .engine import run_audit
from .records import MODULES


def main(argv=None):
    ap = argparse.ArgumentParser(description="Audit a .pptx against a formatting profile.")
    ap.add_argument("deck", help="Path to the .pptx to audit")
    ap.add_argument("--profile", default="prezlab_en",
                    help="Profile name in qc/profiles/ or a path to a profile JSON")
    ap.add_argument("--modules", default=None,
                    help=f"Comma-separated subset of: {','.join(MODULES)}")
    ap.add_argument("--json", dest="json_out", default=None,
                    help="Write the full manifest JSON here")
    ap.add_argument("--verbose", action="store_true", help="Print every record")
    args = ap.parse_args(argv)

    modules = args.modules.split(",") if args.modules else None
    result = run_audit(args.deck, args.profile, modules)

    print(f"\nAudit: {Path(args.deck).name}  |  profile: {result.profile_id} "
          f"v{result.profile_version}  |  {result.slides} slides")
    print(f"Findings: {result.summary['total']}  "
          f"(errors: {result.summary['by_severity'].get('error', 0)}, "
          f"warnings: {result.summary['by_severity'].get('warning', 0)}, "
          f"info: {result.summary['by_severity'].get('info', 0)}, "
          f"arabic-flagged: {result.summary['arabic_flagged']})")

    if result.summary["by_issue_type"]:
        print("\nBy issue type:")
        width = max(len(k) for k in result.summary["by_issue_type"])
        for issue, n in sorted(result.summary["by_issue_type"].items()):
            print(f"  {issue.ljust(width)}  {n}")

    if args.verbose:
        print("\nRecords:")
        for r in result.records:
            ar = " [AR]" if r.arabic_flag else ""
            print(f"  slide {r.slide_index + 1:>3}  {r.severity.upper():7} "
                  f"{r.issue_type}{ar}: {r.message}")

    if args.json_out:
        result.save_manifest(args.json_out)
        print(f"\nManifest written: {args.json_out}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
