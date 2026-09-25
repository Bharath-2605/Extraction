#!/usr/bin/env python3
"""Show how one PDF page stores its pictures and boxes, and what the scan reads.

    python3 debug_page.py "09-02-26 JCP_7467.pdf" 1

Writes page_render.png and paybox_*.png next to this script so the crops can be checked.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf as fitz

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from pdf_url_extractor import (  # noqa: E402
    _bordered_boxes_in_render,
    _box_from_border_lines,
    _is_pay_box,
    _qr_page_zoom,
    extract_from_page,
    extract_qr_images,
    render_region,
)
from vision_scan import decode_qr_hits, image_to_png, recognize_text  # noqa: E402


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    pdf = Path(sys.argv[1]).expanduser()
    page_number = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    document = fitz.open(pdf)
    page = document.load_page(page_number - 1)
    rect = page.rect
    print(f"Page {page_number}: {rect.width:.0f} x {rect.height:.0f} pt, text chars={len(page.get_text().strip())}")

    print("\nPictures listed by the PDF:")
    for info in page.get_image_info(xrefs=True) or []:
        box = fitz.Rect(info["bbox"])
        print(
            f"  xref={info.get('xref')} box=({box.x0:.0f},{box.y0:.0f},{box.x1:.0f},{box.y1:.0f}) "
            f"size={box.width:.0f}x{box.height:.0f}pt pixels={info.get('width')}x{info.get('height')} "
            f"pay_box={_is_pay_box(box, rect)}"
        )

    print("\nPictures the scan keeps:")
    for item in extract_qr_images(page, document):
        box = item.get("bbox")
        shape = item["array"].shape
        label = "none" if box is None else f"({box.x0:.0f},{box.y0:.0f},{box.x1:.0f},{box.y1:.0f})"
        print(f"  box={label} pixels={shape[1]}x{shape[0]}")

    kinds: dict[str, int] = {}
    for entry in page.get_bboxlog() or []:
        kinds[str(entry[0])] = kinds.get(str(entry[0]), 0) + 1
    print("\nDrawing parts:", kinds)

    full = render_region(page, rect, _qr_page_zoom(page))
    (ROOT / "page_render.png").write_bytes(image_to_png(full, max_side=2400))
    lines = _box_from_border_lines(page, rect)
    red = _bordered_boxes_in_render(full, rect)
    print("\nBox from border lines:", lines)
    print("Red boxes in the render:", red)

    for index, box in enumerate(([lines] if lines else []) + red):
        crop = render_region(page, box, 900 / max(min(box.width, box.height), 1.0))
        (ROOT / f"paybox_{index}.png").write_bytes(image_to_png(crop, max_side=1400))
        hits = decode_qr_hits(crop, True)
        print(f"\nBox {index} {box}")
        print("  QR:", [hit.get("payload") for hit in hits])
        print("  Text:", recognize_text(crop, True).replace("\n", " | "))

    rows, _text = extract_from_page(page, document)
    print("\nRows the scan returns:")
    for row in rows:
        print(f"  {row.get('source')}: {row.get('url')}")
    document.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
