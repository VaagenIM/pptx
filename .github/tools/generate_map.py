"""Generate the stable PowerPoint UUID-to-path map for the data branch."""

from __future__ import annotations

import argparse
import json
import subprocess
import uuid
from pathlib import Path, PurePosixPath

UUID_NAMESPACE = uuid.UUID("7d3f6b1e-4f1a-5d0c-9b6a-2a4d8c0e1f73")


def presentation_id(path: str) -> str:
    """Return the deterministic ID for a repository-relative presentation path."""
    return str(uuid.uuid5(UUID_NAMESPACE, path))


def build_map(source: Path) -> dict[str, str]:
    return build_map_with_previous(source, {}, Path.cwd())


def build_map_with_previous(source: Path, previous: dict[str, str], git_root: Path) -> dict[str, str]:
    paths = sorted(
        PurePosixPath(file.relative_to(source).as_posix()).as_posix()
        for file in source.rglob("*.pptx")
    )
    rename_forward = read_renames(git_root)
    previous_by_path = {
        path.removeprefix("powerpoints/"): identifier
        for identifier, path in previous.items()
        if isinstance(identifier, str) and isinstance(path, str) and path.startswith("powerpoints/")
    }
    result = {}
    for path in paths:
        old_path = next(
            (candidate for candidate in previous_by_path if follow_renames(candidate, rename_forward) == path),
            path,
        )
        identifier = previous_by_path.get(old_path, presentation_id(path))
        result[identifier] = f"powerpoints/{path}"
    return result


def read_renames(git_root: Path) -> dict[str, str]:
    output = subprocess.run(
        ["git", "-C", str(git_root), "log", "--format=", "--name-status", "--find-renames", "--", "powerpoints"],
        check=True,
        capture_output=True,
        text=True,
    ).stdout
    renames = {}
    for line in output.splitlines():
        parts = line.split("\t")
        if len(parts) == 3 and parts[0].startswith("R"):
            renames[parts[1].removeprefix("powerpoints/")] = parts[2].removeprefix("powerpoints/")
    return renames


def follow_renames(path: str, rename_forward: dict[str, str]) -> str:
    seen = set()
    while path in rename_forward and path not in seen:
        seen.add(path)
        path = rename_forward[path]
    return path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--previous-map", type=Path)
    args = parser.parse_args()

    previous = json.loads(args.previous_map.read_text(encoding="utf-8")) if args.previous_map and args.previous_map.exists() else {}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(build_map_with_previous(args.source, previous, Path.cwd()), indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
