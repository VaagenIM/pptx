"""Generate the stable PowerPoint UUID-to-path map for the public branch."""

from __future__ import annotations

import argparse
import json
import uuid
from pathlib import Path, PurePosixPath

UUID_NAMESPACE = uuid.UUID("7d3f6b1e-4f1a-5d0c-9b6a-2a4d8c0e1f73")


def presentation_id(path: str) -> str:
    """Return the deterministic ID for a repository-relative presentation path."""
    return str(uuid.uuid5(UUID_NAMESPACE, path))


def build_map(source: Path) -> dict[str, str]:
    paths = sorted(
        PurePosixPath(file.relative_to(source).as_posix()).as_posix()
        for file in source.rglob("*.pptx")
    )
    return {presentation_id(path): f"powerpoints/{path}" for path in paths}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(build_map(args.source), indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
