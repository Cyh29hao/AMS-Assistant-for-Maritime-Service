from __future__ import annotations

import argparse
import getpass
import json
import sys
from pathlib import Path

THIS_DIR = Path(__file__).resolve().parent
ROOT = THIS_DIR.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from desktop_app.private_pack import build_private_pack


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build an encrypted AMS private business pack.")
    parser.add_argument("--source", required=True, help="Source directory containing private-pack.json and overlay files.")
    parser.add_argument("--output", required=True, help="Output .amspack path.")
    parser.add_argument("--password", help="Encryption password. If omitted, prompt securely.")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    password = args.password or getpass.getpass("Private pack password: ")
    if not password:
        raise SystemExit("Password cannot be empty.")
    summary = build_private_pack(Path(args.source), Path(args.output), password)
    print(
        json.dumps(
            {
                "output_path": str(Path(args.output).expanduser().resolve()),
                "summary": summary.to_dict(),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
