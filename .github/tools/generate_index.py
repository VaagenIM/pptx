"""Generate the static directory index used by the Pages site."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def build_index(source: Path) -> list[dict[str, int | str]]:
    entries: list[dict[str, int | str]] = []
    directories = {source}
    for file in sorted(source.rglob("*")):
        relative = file.relative_to(source).as_posix()
        path = f"powerpoints/{relative}"
        if file.is_dir():
            directories.add(file)
            continue
        if file.suffix.lower() != ".pptx":
            continue
        entries.append({"path": path, "type": "blob", "size": file.stat().st_size})
        parent = file.parent
        while parent != source:
            directories.add(parent)
            parent = parent.parent

    for directory in sorted(directories - {source}):
        entries.append({
            "path": f"powerpoints/{directory.relative_to(source).as_posix()}",
            "type": "tree",
            "size": 0,
        })
    return sorted(entries, key=lambda entry: (str(entry["path"]), str(entry["type"])))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(build_index(args.source), indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
