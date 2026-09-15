"""Add optimized copies of newly added PPTX files to a separate branch."""

from __future__ import annotations

import argparse
import io
import zipfile
from pathlib import Path

from PIL import Image

MAX_IMAGE_PIXELS = 1920 * 1080
IMAGE_FORMATS = {".png": "PNG", ".jpg": "JPEG", ".jpeg": "JPEG"}


def optimize_image(data: bytes, suffix: str) -> bytes:
    image_format = IMAGE_FORMATS.get(suffix.lower())
    if not image_format:
        return data
    try:
        with Image.open(io.BytesIO(data)) as image:
            image.load()
            scale = min(1, (MAX_IMAGE_PIXELS / (image.width * image.height)) ** 0.5)
            size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
            if size != image.size:
                image = image.resize(size, Image.Resampling.LANCZOS)
            output = io.BytesIO()
            save_options = {"optimize": True}
            if image_format == "JPEG":
                if image.mode not in ("RGB", "L"):
                    image = image.convert("RGB")
                save_options["quality"] = 88
            image.save(output, format=image_format, **save_options)
            return output.getvalue()
    except (OSError, ValueError):
        return data


def optimize_pptx(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as input_zip, zipfile.ZipFile(
        destination, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
    ) as output_zip:
        for item in input_zip.infolist():
            data = input_zip.read(item.filename)
            if item.filename.startswith("ppt/media/"):
                data = optimize_image(data, Path(item.filename).suffix)
            output_zip.writestr(item, data)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--destination", type=Path, required=True)
    args = parser.parse_args()
    for source in args.source.rglob("*.pptx"):
        destination = args.destination / source.relative_to(args.source)
        if not destination.exists():
            optimize_pptx(source, destination)
            print(f"Optimized {source} -> {destination}")


if __name__ == "__main__":
    main()
