"""Synchronize optimized PPTX files to a separate branch."""

from __future__ import annotations

import argparse
import io
import posixpath
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from PIL import Image

MAX_IMAGE_DIMENSION = 1600
MAX_IMAGE_PIXELS = 1600 * 900
REFERENCE_SLIDE_WIDTH = 1920
IMAGE_FORMATS = {".png": "PNG", ".jpg": "JPEG", ".jpeg": "JPEG"}
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def optimize_image(data: bytes, suffix: str, target_size: tuple[int, int] | None = None) -> bytes:
    image_format = IMAGE_FORMATS.get(suffix.lower())
    if not image_format:
        return data
    try:
        with Image.open(io.BytesIO(data)) as image:
            image.load()
            scale = min(
                1,
                MAX_IMAGE_DIMENSION / max(image.width, image.height),
                (MAX_IMAGE_PIXELS / (image.width * image.height)) ** 0.5,
            )
            if target_size:
                target_width, target_height = target_size
                scale = min(
                    scale,
                    max(target_width / image.width, target_height / image.height),
                )
            size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
            if size != image.size:
                image = image.resize(size, Image.Resampling.LANCZOS)
            output = io.BytesIO()
            save_options = {"optimize": True}
            if image_format == "JPEG":
                if image.mode not in ("RGB", "L"):
                    image = image.convert("RGB")
                save_options["quality"] = 75
            image.save(output, format=image_format, **save_options)
            optimized = output.getvalue()
            return optimized if len(optimized) < len(data) else data
    except (OSError, ValueError):
        return data


def media_target_sizes(input_zip: zipfile.ZipFile) -> dict[str, tuple[int, int]]:
    try:
        presentation = ElementTree.fromstring(input_zip.read("ppt/presentation.xml"))
        slide_size = presentation.find("p:sldSz", NS)
        if slide_size is None:
            return {}
        slide_width = int(slide_size.attrib["cx"])
        slide_height = int(slide_size.attrib["cy"])
        reference_slide_height = REFERENCE_SLIDE_WIDTH * slide_height / slide_width
        targets = {}
        for slide_name in input_zip.namelist():
            if not slide_name.startswith("ppt/slides/slide") or not slide_name.endswith(".xml"):
                continue
            relationships_name = posixpath.join(
                posixpath.dirname(slide_name),
                "_rels",
                f"{posixpath.basename(slide_name)}.rels",
            )
            if relationships_name not in input_zip.namelist():
                continue
            relationships = {
                relationship.attrib["Id"]: relationship.attrib["Target"]
                for relationship in ElementTree.fromstring(input_zip.read(relationships_name))
                if relationship.tag.rsplit("}", 1)[-1] == "Relationship"
            }
            slide = ElementTree.fromstring(input_zip.read(slide_name))
            for picture in slide.findall(".//p:pic", NS):
                blip = picture.find("p:blipFill/a:blip", NS)
                extent = picture.find("p:spPr/a:xfrm/a:ext", NS)
                if blip is None:
                    continue
                relationship_id = blip.attrib.get(f"{{{NS['r']}}}embed")
                target = relationships.get(relationship_id)
                if not target or extent is None:
                    continue
                media_name = posixpath.normpath(posixpath.join(posixpath.dirname(slide_name), target))
                width = round(int(extent.attrib["cx"]) / slide_width * REFERENCE_SLIDE_WIDTH)
                height = round(int(extent.attrib["cy"]) / slide_height * reference_slide_height)
                current = targets.get(media_name, (0, 0))
                targets[media_name] = (max(current[0], width), max(current[1], height))
        return targets
    except (AttributeError, KeyError, TypeError, ValueError, ElementTree.ParseError):
        return {}


def optimize_pptx(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as input_zip, zipfile.ZipFile(
        destination, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
    ) as output_zip:
        target_sizes = media_target_sizes(input_zip)
        for item in input_zip.infolist():
            data = input_zip.read(item.filename)
            if item.filename.startswith("ppt/media/"):
                data = optimize_image(data, Path(item.filename).suffix, target_sizes.get(item.filename))
            output_zip.writestr(item, data)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--destination", type=Path, required=True)
    parser.add_argument("--force-compress", action="store_true")
    args = parser.parse_args()
    source_paths = {
        source.relative_to(args.source)
        for source in args.source.rglob("*.pptx")
    }
    for destination in args.destination.rglob("*.pptx"):
        if destination.relative_to(args.destination) not in source_paths:
            destination.unlink()
    for source in args.source.rglob("*.pptx"):
        destination = args.destination / source.relative_to(args.source)
        if args.force_compress or not destination.exists():
            optimize_pptx(source, destination)
            print(f"Optimized {source} -> {destination}")


if __name__ == "__main__":
    main()
