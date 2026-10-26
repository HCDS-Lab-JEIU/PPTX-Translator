"""PPTX에서 위치를 포함한 텍스트와 원본 이미지를 추출하는 CLI."""

import argparse
from collections import defaultdict
import hashlib
import json
import math
from pathlib import Path
import re
import sys
from zipfile import BadZipFile

from pptx_text import frame_text, equation_map, math_text_nodes, shape_elements
from pptx_styles import styled_frame

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.exc import PackageNotFoundError
except ImportError as exc:
    raise SystemExit(
        "python-pptx 설치가 필요합니다: python -m pip install -r requirements.txt"
    ) from exc

EMU_PER_CM = 360000
IDENTITY = (1, 0, 0, 1, 0, 0)
NOISE_TYPES = {
    PP_PLACEHOLDER.HEADER, PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE,
}
TITLE_TYPES = {
    PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE,
    PP_PLACEHOLDER.VERTICAL_TITLE,
}
PAGE_NUMBER = re.compile(
    r"(?:[-–—]\s*)?(?:page\s*|p\.?\s*)?\d{1,4}"
    r"(?:\s*(?:/|of)\s*\d{1,4})?(?:\s*쪽|\s*페이지)?(?:\s*[-–—])?",
    re.IGNORECASE,
)


def compose(a, b):
    """Affine a(b(point)); nested groups retain slide coordinates."""
    aa, ab, ac, ad, ae, af = a
    ba, bb, bc, bd, be, bf = b
    return (aa * ba + ac * bb, ab * ba + ad * bb,
            aa * bc + ac * bd, ab * bc + ad * bd,
            aa * be + ac * bf + ae, ab * be + ad * bf + af)


def orientation(shape):
    """Shape rotation/flip about its centre in parent coordinates."""
    angle = math.radians(shape.rotation)
    c, s = math.cos(angle), math.sin(angle)
    transforms = shape.element.xpath("./p:spPr/a:xfrm | ./p:grpSpPr/a:xfrm | ./p:xfrm")
    xfrm = transforms[0] if transforms else None
    fx = -1 if xfrm is not None and xfrm.get("flipH") in ("1", "true") else 1
    fy = -1 if xfrm is not None and xfrm.get("flipV") in ("1", "true") else 1
    a, b, c, d = c * fx, s * fx, -s * fy, c * fy
    cx = shape.left + shape.width / 2
    cy = shape.top + shape.height / 2
    return (a, b, c, d, cx - a * cx - c * cy, cy - b * cx - d * cy)


def walk_shapes(shapes, parent=IDENTITY, *, unsupported=None):
    for element in shape_elements(shapes._spTree, unsupported):
        shape = shapes._shape_factory(element)
        matrix = compose(parent, orientation(shape))
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            xfrm = shape.element.grpSpPr.xfrm
            sx = shape.width / xfrm.chExt.cx if xfrm.chExt.cx else 1
            sy = shape.height / xfrm.chExt.cy if xfrm.chExt.cy else 1
            child = (sx, 0, 0, sy,
                     shape.left - xfrm.chOff.x * sx,
                     shape.top - xfrm.chOff.y * sy)
            yield from walk_shapes(shape.shapes, compose(matrix, child), unsupported=unsupported)
        else:
            yield shape, matrix


def bbox(matrix, x, y, width, height):
    a, b, c, d, e, f = matrix
    points = [(a * px + c * py + e, b * px + d * py + f)
              for px, py in ((x, y), (x + width, y),
                             (x, y + height), (x + width, y + height))]
    xs, ys = zip(*points)
    return {"x": round(min(xs) / EMU_PER_CM, 4),
            "y": round(min(ys) / EMU_PER_CM, 4),
            "width": round((max(xs) - min(xs)) / EMU_PER_CM, 4),
            "height": round((max(ys) - min(ys)) / EMU_PER_CM, 4)}


def clean_text(text):
    return "\n".join(line.rstrip() for line in text.replace("\v", "\n").splitlines()).strip()


def file_sha256(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def text_records(shape, matrix, page):
    placeholder = shape.placeholder_format.type if shape.is_placeholder else None
    if placeholder in NOISE_TYPES:
        return
    base = {"page": page, "shape_id": shape.shape_id,
            "is_title": placeholder in TITLE_TYPES}
    if shape.has_table:
        table = shape.table
        y = shape.top
        for row_index, row in enumerate(table.rows):
            x = shape.left
            for col_index, cell in enumerate(row.cells):
                text = clean_text(frame_text(cell.text_frame))
                if text and not cell.is_spanned:
                    width = sum(table.columns[i].width for i in
                                range(col_index, col_index + cell.span_width))
                    height = sum(table.rows[i].height for i in
                                 range(row_index, row_index + cell.span_height))
                    record = {**base, "kind": "table_cell", "text": text,
                           "row": row_index + 1, "column": col_index + 1,
                           "row_span": cell.span_height, "column_span": cell.span_width,
                           "position_cm": bbox(matrix, x, y, width, height)}
                    yield from frame_records(record, cell.text_frame)
                x += table.columns[col_index].width
            y += row.height
    elif shape.has_text_frame:
        text = clean_text(frame_text(shape.text_frame))
        if text:
            record = {**base, "kind": "text", "text": text,
                      "position_cm": bbox(matrix, shape.left, shape.top, shape.width, shape.height)}
            yield from frame_records(record, shape.text_frame)


def frame_records(record, frame):
    styled, _ = styled_frame(frame)
    if styled is not None:
        record["styled_text"] = styled.replace("\v", "\n").strip("\n")
    equations = [{"token": token, "text": "".join(node.text or "" for node in math_text_nodes_in_root(root))}
                 for root, token in equation_map(frame).items()]
    if equations:
        record["equations"] = equations
    yield record
    for index, node in enumerate(math_text_nodes(frame), 1):
        if re.search(r"[\u1100-\u11ff\u3130-\u318f\uac00-\ud7a3]", node.text or ""):
            yield {**record, "kind": "math_text", "math_text_index": index, "text": node.text,
                   "is_title": False, "styled_text": None}


def math_text_nodes_in_root(root):
    return root.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}t")


def margin_zone(record, slide_height):
    pos = record["position_cm"]
    # Only boxes entirely inside the top/bottom margin qualify.
    if pos["y"] + pos["height"] <= slide_height * 0.10:
        return "top"
    if pos["y"] >= slide_height * 0.90:
        return "bottom"
    return None


def filter_texts(slides, width, height):
    """Remove explicit placeholders, margin page numbers, repeated margin labels."""
    occurrences = defaultdict(set)

    def key(record):
        zone = margin_zone(record, height)
        if (not zone or record["is_title"] or record["kind"] != "text"
                or len(record["text"]) > 200):
            return None
        pos = record["position_cm"]
        return (zone, " ".join(record["text"].split()).casefold(),
                round(pos["x"] / width * 50), round(pos["y"] / height * 50))

    for slide in slides:
        for record in slide["texts"]:
            marker = key(record)
            if marker is not None:
                occurrences[marker].add(slide["page"])
    threshold = max(3, math.ceil(len(slides) * 0.60))
    removed = 0
    for slide in slides:
        kept = []
        for record in slide["texts"]:
            candidate = record["kind"] == "text" and not record["is_title"]
            page_number = (candidate and margin_zone(record, height)
                           and PAGE_NUMBER.fullmatch(" ".join(record["text"].split())))
            marker = key(record)
            repeated = marker is not None and len(occurrences[marker]) >= threshold
            if page_number or repeated:
                removed += 1
            else:
                kept.append(record)
        slide["texts"] = sorted(kept, key=lambda r: (
            r["position_cm"]["y"], r["position_cm"]["x"], r["shape_id"]))
    return removed


def extract_pptx(source, output_dir):
    source, output_dir = Path(source), Path(output_dir)
    if source.suffix.lower() != ".pptx":
        raise ValueError(".pptx 파일만 지원합니다. .ppt는 .pptx로 변환하세요.")
    if not source.is_file():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {source}")
    presentation = Presentation(str(source))
    width = presentation.slide_width / EMU_PER_CM
    height = presentation.slide_height / EMU_PER_CM
    if width <= 0 or height <= 0:
        raise ValueError("슬라이드 크기가 올바르지 않습니다.")
    resources = output_dir / "resources"
    resources.mkdir(parents=True, exist_ok=True)
    slides = []
    warnings = []
    omitted_placeholders = 0
    preserved_unsupported = []
    for page, slide in enumerate(presentation.slides, 1):
        result = {"page": page, "texts": [], "images": []}
        unsupported = []
        for z_order, (shape, matrix) in enumerate(walk_shapes(slide.shapes, unsupported=unsupported)):
            if shape.is_placeholder and shape.placeholder_format.type in NOISE_TYPES:
                omitted_placeholders += 1
                continue
            result["texts"].extend({**record, "z_order": z_order}
                                   for record in text_records(shape, matrix, page))
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    image = shape.image
                except (ValueError, KeyError) as exc:
                    warnings.append(f"페이지 {page}, 도형 {shape.shape_id}: 이미지 읽기 실패 ({exc})")
                    continue
                digest = hashlib.sha256(image.blob).hexdigest()
                relative = Path("resources") / f"image_{digest}.{image.ext}"
                (output_dir / relative).write_bytes(image.blob)
                result["images"].append({
                    "page": page, "shape_id": shape.shape_id,
                    "z_order": z_order,
                    "file": relative.as_posix(),
                    "position_cm": bbox(matrix, shape.left, shape.top, shape.width, shape.height),
                    "rotation_degrees": shape.rotation,
                    "crop": {"left": shape.crop_left, "top": shape.crop_top,
                             "right": shape.crop_right, "bottom": shape.crop_bottom},
                })
            if shape.has_chart:
                warnings.append(f"페이지 {page}, 도형 {shape.shape_id}: 차트 렌더링은 지원하지 않습니다.")
        if unsupported:
            kinds = sorted({element.tag.rsplit('}', 1)[-1] for element in unsupported})
            warnings.append(f"페이지 {page}: 편집 미지원 삽입 개체 {len(unsupported)}개는 원본 그대로 유지합니다. ({', '.join(kinds)})")
            preserved_unsupported.extend({'page':page, 'type':element.tag.rsplit('}', 1)[-1]}
                                         for element in unsupported)
        slides.append(result)
    removed = filter_texts(slides, width, height)
    document = {
        "source": source.name, "source_sha256": file_sha256(source), "text_schema_version": 3,
        "coordinate_unit": "cm",
        "coordinate_origin": "top-left", "slide_size_cm": {"width": width, "height": height},
        "removed_text_count": omitted_placeholders + removed,
        "slides": slides, "warnings": warnings, "preserved_unsupported_shapes": preserved_unsupported,
    }
    return document


def format_text(document):
    """One record per line; JSON string escaping keeps multiline text unambiguous."""
    lines = []
    for slide in document["slides"]:
        for record in slide["texts"]:
            pos = record["position_cm"]
            location = (f"x={pos['x']:.4f},y={pos['y']:.4f},"
                        f"w={pos['width']:.4f},h={pos['height']:.4f} cm; shape={record['shape_id']}")
            if "row" in record:
                location += f"; table={record['shape_id']},row={record['row']},col={record['column']}"
            if "math_text_index" in record:
                location += f"; math_text={record['math_text_index']}"
            content = ({"text": record["text"], "styled_text": record["styled_text"]}
                       if record.get("styled_text") else record["text"])
            lines.append(f"[page={slide['page']}][{location}]"
                         f"[{json.dumps(content, ensure_ascii=False)}]")
    return "\n".join(lines) + ("\n" if lines else "")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file", type=Path, help="입력 .pptx 경로")
    parser.add_argument("-o", "--output", type=Path, default=Path("output"), help="저장 폴더 (기본: output)")
    args = parser.parse_args(argv)
    # PowerShell redirection must preserve Korean and mathematical symbols.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    try:
        document = extract_pptx(args.file, args.output)
        text = format_text(document)
        (args.output / "texts.txt").write_text(text, encoding="utf-8")
        (args.output / "content.json").write_text(
            json.dumps(document, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        sys.stdout.write(text)
        image_count = sum(len(slide["images"]) for slide in document["slides"])
        print(f"저장 완료: {args.output.resolve()} "
              f"(페이지 {len(document['slides'])}, 이미지 배치 {image_count}, "
              f"제외 텍스트 {document['removed_text_count']})", file=sys.stderr)
        for warning in document["warnings"]:
            print(f"참고: {warning}", file=sys.stderr)
    except (OSError, ValueError, KeyError, BadZipFile, PackageNotFoundError, SyntaxError) as exc:
        print(f"추출 실패: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
