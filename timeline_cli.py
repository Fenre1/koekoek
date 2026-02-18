from __future__ import annotations

import argparse
from pathlib import Path

from timeline_combined_viewer import generate_combined_timeline
from timeline_horizontal import generate_horizontal_timeline
from timeline_vertical_filterable import generate_vertical_timeline


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="timeline",
        description="Generate timeline HTML from an Excel file.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    horizontal = subparsers.add_parser("horizontal", help="Generate horizontal timeline HTML.")
    horizontal.add_argument("-i", "--input", required=True, type=Path, help="Path to input Excel file.")
    horizontal.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("timeline_horizontal.html"),
        help="Path to output HTML file.",
    )

    vertical = subparsers.add_parser("vertical", help="Generate vertical timeline HTML.")
    vertical.add_argument("-i", "--input", required=True, type=Path, help="Path to input Excel file.")
    vertical.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("timeline_vertical_filterable.html"),
        help="Path to output HTML file.",
    )

    combined = subparsers.add_parser("combined", help="Generate combined timeline viewer HTML.")
    combined.add_argument("-i", "--input", required=True, type=Path, help="Path to input Excel file.")
    combined.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("timeline_combined.html"),
        help="Path to output HTML file.",
    )
    return parser


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    if args.command == "horizontal":
        generate_horizontal_timeline(args.input, args.output)
        return 0
    if args.command == "vertical":
        generate_vertical_timeline(args.input, args.output)
        return 0
    if args.command == "combined":
        generate_combined_timeline(args.input, args.output)
        return 0

    parser.error(f"Unknown command: {args.command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
