#!/usr/bin/env python3
"""Local image OCR and QR decoding. Tesseract is never used.

RapidOCR runs PP-OCR models through ONNX Runtime. Apple Vision is not used.
QR codes are read with OpenCV.
"""
from __future__ import annotations

import logging
import os
import sys
import threading
from functools import lru_cache

import cv2
import numpy as np

os.environ.setdefault("OMP_NUM_THREADS", "1")
os.environ.setdefault("ORT_NUM_THREADS", "1")

logger = logging.getLogger(__name__)

_THREAD = threading.local()
_INIT_LOCK = threading.Lock()
_BACKEND_NAME: str | None = None


def _to_rgb(image: np.ndarray) -> np.ndarray:
    if image is None or image.size == 0:
        return np.zeros((8, 8, 3), dtype=np.uint8)
    if image.ndim == 2:
        return cv2.cvtColor(image, cv2.COLOR_GRAY2RGB)
    if image.shape[2] == 4:
        return cv2.cvtColor(image, cv2.COLOR_RGBA2RGB)
    return image


def _to_gray(image: np.ndarray) -> np.ndarray:
    if image.ndim == 2:
        return image
    if image.shape[2] == 4:
        return cv2.cvtColor(image, cv2.COLOR_RGBA2GRAY)
    return cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)


def _ensure_min_side(image: np.ndarray, min_side: int) -> np.ndarray:
    height, width = image.shape[:2]
    shortest = min(height, width)
    if shortest >= min_side or shortest < 2:
        return image
    scale = min_side / float(shortest)
    return cv2.resize(image, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)


def _fit_image(image: np.ndarray, max_side: int) -> np.ndarray:
    height, width = image.shape[:2]
    longest = max(height, width)
    if longest <= max_side:
        return image
    scale = max_side / float(longest)
    return cv2.resize(image, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)


def _apple_vision_available() -> bool:
    return False


def _make_rapidocr():
    try:
        from rapidocr import RapidOCR
        return RapidOCR()
    except Exception:
        from rapidocr_onnxruntime import RapidOCR
        return RapidOCR()


def _rapidocr_texts(engine, image: np.ndarray, faint: bool = False) -> list[str]:
    # The line reader turns detection off, and RapidOCR keeps that setting, so turn it back on.
    # Faint mode lowers the 0.5 cut-offs so pale, tiny print in a pay box is not thrown away.
    options = {"use_det": True, "use_cls": True, "use_rec": True}
    options.update({"text_score": 0.3, "box_thresh": 0.3} if faint else {"text_score": 0.5, "box_thresh": 0.5})
    try:
        output = engine(_to_rgb(image), **options)
    except TypeError:
        output = engine(_to_rgb(image))
    texts: list[str] = []
    if output is None:
        return texts
    txts = getattr(output, "txts", None)
    if txts:
        return [str(item) for item in txts if str(item).strip()]
    result = output[0] if isinstance(output, tuple) else output
    if not result:
        return texts
    for item in result:
        if isinstance(item, (list, tuple)) and len(item) > 1:
            texts.append(str(item[1]))
        elif isinstance(item, str):
            texts.append(item)
    return [text for text in texts if text.strip()]


def _apple_texts(image: np.ndarray, accurate: bool = False) -> list[str]:
    from ocrmac import ocrmac
    from PIL import Image

    rgb = _to_rgb(image)
    pil = Image.fromarray(rgb)
    recognizer = ocrmac.OCR(
        pil,
        recognition_level="accurate" if accurate else "fast",
        language_preference=["en-US"],
    )
    records = recognizer.recognize() or []
    texts: list[str] = []
    for record in records:
        if isinstance(record, dict):
            text = record.get("text") or record.get("content") or ""
        elif isinstance(record, (list, tuple)) and record:
            text = record[0]
        else:
            text = record
        if text:
            texts.append(str(text))
    return [text for text in texts if text.strip()]


def _probe_image() -> np.ndarray:
    from PIL import Image, ImageDraw, ImageFont

    image = Image.new("RGB", (720, 160), "white")
    draw = ImageDraw.Draw(image)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 36)
    except Exception:
        font = ImageFont.load_default()
    draw.text((24, 52), "https://ocr-probe.example.com", fill="black", font=font)
    return np.array(image)


def _select_backend() -> str:
    global _BACKEND_NAME
    with _INIT_LOCK:
        if _BACKEND_NAME:
            return _BACKEND_NAME
        # Apple Vision is not used. Every machine, including macOS, uses RapidOCR.
        _BACKEND_NAME = "rapidocr"
        logger.info("OCR backend: RapidOCR PP-OCR (ONNX)")
        return _BACKEND_NAME


_RAPID_LOCK = threading.Lock()
_RAPID_ENGINE = None


def _thread_rapidocr():
    """One OCR engine for every page. Loading it again per thread is the slow part."""
    global _RAPID_ENGINE
    engine = _RAPID_ENGINE
    if engine is not None:
        return engine
    with _RAPID_LOCK:
        if _RAPID_ENGINE is None:
            _RAPID_ENGINE = _make_rapidocr()
        return _RAPID_ENGINE


def ocr_backend() -> str:
    return _select_backend()


def recognize_faint_text(image: np.ndarray) -> str:
    """Full OCR with lower cut-offs, for the small area beside a code only."""
    if image is None or image.size == 0 or min(image.shape[:2]) < 8:
        return ""
    image = _fit_image(_ensure_min_side(image, 300), 1200)
    try:
        engine = _thread_rapidocr()
        with _RAPID_LOCK:
            texts = _rapidocr_texts(engine, image, faint=True)
    except Exception as exc:
        logger.warning("RapidOCR failed: %s", exc)
        texts = []
    return "\n".join(texts).strip()


def recognize_text(image: np.ndarray, accurate: bool | None = None) -> str:
    """Read visible text from screenshots, logos, and other raster content."""
    if image is None or image.size == 0:
        return ""
    if min(image.shape[:2]) < 8:
        return ""
    image = _ensure_min_side(image, 300 if accurate else 220)
    image = _fit_image(image, 1200 if accurate else 900)
    if accurate is None:
        accurate = min(image.shape[:2]) < 420
    _select_backend()
    try:
        engine = _thread_rapidocr()
        with _RAPID_LOCK:
            texts = _rapidocr_texts(engine, image)
    except Exception as exc:
        logger.warning("RapidOCR failed: %s", exc)
        texts = []
    return "\n".join(texts).strip()


def recognize_lines(image: np.ndarray) -> str:
    """Read a small box of print one line at a time, without the text finder.

    The finder misses pale, tiny print in a pay box. Each line is cut out by
    its ink and read directly, which is also far faster than a full OCR pass.
    """
    if image is None or image.size == 0 or image.ndim != 3 or min(image.shape[:2]) < 8:
        return ""
    rgb = _to_rgb(image)
    red = rgb[:, :, 0].astype(np.int16)
    green = rgb[:, :, 1].astype(np.int16)
    blue = rgb[:, :, 2].astype(np.int16)
    # Print is anything clearly darker than the paper, so pale grey links still count.
    # A red border would make every row look like print.
    darkest = rgb.min(axis=2).astype(np.int16)
    paper = int(np.percentile(darkest, 90))
    ink = (darkest < paper - 18) & ~((red - green > 45) & (red - blue > 45))
    rows = ink.sum(axis=1) > max(2, int(0.01 * rgb.shape[1]))
    bands: list[tuple[int, int]] = []
    start = None
    for index, on in enumerate(list(rows) + [False]):
        if on and start is None:
            start = index
        elif not on and start is not None:
            if index - start >= 5:
                bands.append((start, index))
            start = None
    if not bands:
        return ""
    engine = _thread_rapidocr()
    texts: list[str] = []
    for y0, y1 in bands[:6]:
        columns = np.where(ink[y0:y1].any(axis=0))[0]
        if columns.size == 0:
            continue
        pad = max(4, (y1 - y0) // 3)
        line = rgb[max(0, y0 - pad) : y1 + pad, max(0, int(columns[0]) - pad) : int(columns[-1]) + pad]
        gray = _to_gray(line).astype(np.float32)
        low, high = float(gray.min()), float(max(gray.max(), gray.min() + 1))
        stretched = ((gray - low) * (255.0 / (high - low))).clip(0, 255).astype(np.uint8)
        try:
            with _RAPID_LOCK:
                output = engine(cv2.cvtColor(stretched, cv2.COLOR_GRAY2RGB), use_det=False, use_cls=False, use_rec=True)
        except Exception as exc:
            logger.debug("Line OCR failed: %s", exc)
            continue
        for text in getattr(output, "txts", None) or ():
            if str(text).strip():
                texts.append(str(text).strip())
    return "\n".join(texts)


def _thread_qr_detector():
    detector = getattr(_THREAD, "qr", None)
    if detector is None:
        detector = cv2.QRCodeDetector()
        try:
            detector.setEpsX(0.2)
            detector.setEpsY(0.2)
        except Exception:
            pass
        _THREAD.qr = detector
    return detector


def _decode_with_detector(image: np.ndarray) -> list[str]:
    payloads: list[str] = []
    gray = image if image.ndim == 2 else _to_gray(image)
    for factory in (_thread_qr_detector, _thread_aruco_detector):
        detector = factory()
        if detector is None:
            continue
        try:
            ok, decoded, _points, _ = detector.detectAndDecodeMulti(gray)
            if ok and decoded is not None:
                payloads.extend(str(item) for item in decoded if item)
        except Exception:
            pass
        if payloads:
            return payloads
        try:
            data, _points, _straight = detector.detectAndDecode(gray)
            if data:
                payloads.append(str(data))
        except Exception:
            try:
                data, _points, _ = detector.detectAndDecode(gray)
                if data:
                    payloads.append(str(data))
            except Exception:
                pass
        if payloads:
            return payloads
    return payloads


def _thread_aruco_detector():
    detector = getattr(_THREAD, "qr_aruco", None)
    if detector is False:
        return None
    if detector is None:
        try:
            detector = cv2.QRCodeDetectorAruco()
        except Exception:
            _THREAD.qr_aruco = False
            return None
        _THREAD.qr_aruco = detector
    return detector


_VISION_GATE = threading.Semaphore(4)


def _observation_payload(observation) -> str:
    value = getattr(observation, "payloadStringValue", None)
    if callable(value):
        try:
            value = value()
        except Exception:
            value = None
    return str(value or "").strip()


def _candidate_text(candidate) -> str:
    value = getattr(candidate, "string", None)
    if callable(value):
        try:
            value = value()
        except Exception:
            value = None
    return str(value or "").strip()


def _ns_image_data(image: np.ndarray) -> object | None:
    try:
        from Foundation import NSData
        from PIL import Image
        import io
    except Exception:
        return None
    rgb = _to_rgb(image)
    pil = Image.fromarray(rgb)
    buf = io.BytesIO()
    pil.save(buf, format="PNG")
    data = buf.getvalue()
    return NSData.dataWithBytes_length_(data, len(data))


def _recognize_with_apple_vision(image: np.ndarray, accurate: bool = True) -> list[str]:
    """Apple Vision OCR is disabled on every platform."""
    del image, accurate
    return []


def _decode_with_apple_vision(image: np.ndarray) -> list[str]:
    """Apple Vision QR decode is disabled on every platform."""
    del image
    return []


def _qr_variants(image: np.ndarray, extra: bool = False) -> list[np.ndarray]:
    gray = _to_gray(image)
    height, width = gray.shape[:2]
    longest = max(height, width)
    if longest < 280:
        gray = cv2.resize(gray, None, fx=2.4, fy=2.4, interpolation=cv2.INTER_CUBIC)
    elif longest > 1600:
        gray = _fit_image(gray, 1600)
    variants = [gray, cv2.bitwise_not(gray)]
    if extra:
        try:
            variants.append(
                cv2.adaptiveThreshold(
                    gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 5
                )
            )
        except Exception:
            pass
        _, otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        variants.append(otsu)
    return variants


def _order_points(pts: np.ndarray) -> np.ndarray:
    pts = np.array(pts, dtype=np.float32).reshape(-1, 2)
    total = pts.sum(axis=1)
    diff = np.diff(pts, axis=1).reshape(-1)
    return np.array(
        [pts[np.argmin(total)], pts[np.argmin(diff)], pts[np.argmax(total)], pts[np.argmax(diff)]],
        dtype=np.float32,
    )


def _crop_quad(image: np.ndarray, quad) -> list[np.ndarray]:
    pts = np.array(quad, dtype=np.float32).reshape(-1, 2)
    if pts.shape[0] < 4:
        return []
    height, width = image.shape[:2]
    pad = int(max(8, 0.18 * max(pts[:, 0].max() - pts[:, 0].min(), pts[:, 1].max() - pts[:, 1].min())))
    x0 = max(int(pts[:, 0].min()) - pad, 0)
    y0 = max(int(pts[:, 1].min()) - pad, 0)
    x1 = min(int(pts[:, 0].max()) + pad, width)
    y1 = min(int(pts[:, 1].max()) + pad, height)
    if x1 - x0 < 12 or y1 - y0 < 12:
        return []
    crops = [image[y0:y1, x0:x1]]
    try:
        side = max(int(max(x1 - x0, y1 - y0)), 280)
        src = _order_points(pts)
        dst = np.array([[0, 0], [side - 1, 0], [side - 1, side - 1], [0, side - 1]], dtype=np.float32)
        matrix = cv2.getPerspectiveTransform(src, dst)
        warped = cv2.warpPerspective(_to_rgb(image), matrix, (side, side))
        crops.append(warped)
    except Exception:
        pass
    return [crop for crop in crops if crop is not None and min(crop.shape[:2]) >= 12]


def _detected_qr_crops(image: np.ndarray) -> list[np.ndarray]:
    gray = _to_gray(image)
    crops: list[np.ndarray] = []
    for factory in (_thread_qr_detector, _thread_aruco_detector):
        detector = factory()
        if detector is None:
            continue
        points = None
        try:
            ok, _decoded, pts, _straight = detector.detectAndDecodeMulti(gray)
            if ok:
                points = pts
        except Exception:
            try:
                found, pts = detector.detect(gray)
                if found:
                    points = pts
            except Exception:
                points = None
        if points is None:
            continue
        for quad in np.array(points):
            crops.extend(_crop_quad(image, quad))
        if crops:
            break
    return crops


def _square_region_crops(image: np.ndarray, max_crops: int = 24) -> list[np.ndarray]:
    gray = _to_gray(image)
    height, width = gray.shape[:2]
    if min(height, width) < 40:
        return []
    try:
        binary = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 5
        )
    except Exception:
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    if float(np.mean(binary)) > 127:
        binary = cv2.bitwise_not(binary)
    contours, _ = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    boxes: list[tuple[int, int, int, int]] = []
    for contour in contours:
        x, y, cw, ch = cv2.boundingRect(contour)
        if cw < 22 or ch < 22 or cw > width * 0.9 or ch > height * 0.9:
            continue
        aspect = cw / float(ch)
        if aspect < 0.55 or aspect > 1.85:
            continue
        area = cv2.contourArea(contour)
        if area < 0.22 * cw * ch:
            continue
        boxes.append((x, y, cw, ch))
    boxes.sort(key=lambda item: item[2] * item[3], reverse=True)
    crops: list[np.ndarray] = []
    kept: list[tuple[int, int, int, int]] = []
    for x, y, cw, ch in boxes:
        if any(abs((x - ox) * (y - oy)) < 4 and abs(cw - ow) < 8 and abs(ch - oh) < 8 for ox, oy, ow, oh in kept):
            continue
        pad = int(max(cw, ch) * 0.2)
        x0 = max(0, x - pad)
        y0 = max(0, y - pad)
        x1 = min(width, x + cw + pad)
        y1 = min(height, y + ch + pad)
        crop = image[y0:y1, x0:x1]
        if min(crop.shape[:2]) < 16:
            continue
        crops.append(crop)
        kept.append((x, y, cw, ch))
        if len(crops) >= max_crops:
            break
    return crops


def _qr_tiles(gray: np.ndarray, rows: int = 3, cols: int = 3, overlap: float = 0.22) -> list[np.ndarray]:
    height, width = gray.shape[:2]
    if height < 40 or width < 40:
        return []
    tile_h = max(40, int(height / rows * (1 + overlap)))
    tile_w = max(40, int(width / cols * (1 + overlap)))
    step_y = max(1, (height - tile_h) // max(rows - 1, 1)) if rows > 1 else 0
    step_x = max(1, (width - tile_w) // max(cols - 1, 1)) if cols > 1 else 0
    tiles: list[np.ndarray] = []
    for row in range(rows):
        y0 = min(row * step_y, max(0, height - tile_h))
        y1 = min(height, y0 + tile_h)
        for col in range(cols):
            x0 = min(col * step_x, max(0, width - tile_w))
            x1 = min(width, x0 + tile_w)
            tiles.append(gray[y0:y1, x0:x1])
    return tiles


def image_to_png(image: np.ndarray, max_side: int = 200) -> bytes:
    """Encode a QR crop for the Excel sheet."""
    if image is None or getattr(image, "size", 0) == 0:
        return b""
    import io
    from PIL import Image

    rgb = _fit_image(_to_rgb(image), max_side)
    buffer = io.BytesIO()
    Image.fromarray(rgb).save(buffer, format="PNG")
    return buffer.getvalue()


def _clip_box(box, width: int, height: int, pad_frac: float = 0.14):
    if box is None:
        return None
    x0, y0, x1, y1 = [int(round(float(value))) for value in box]
    pad = int(max(8, max(1, x1 - x0, y1 - y0) * pad_frac))
    x0 = max(0, x0 - pad)
    y0 = max(0, y0 - pad)
    x1 = min(width, x1 + pad)
    y1 = min(height, y1 + pad)
    if x1 - x0 < 8 or y1 - y0 < 8:
        return None
    return x0, y0, x1, y1


def _crop_box(image: np.ndarray, box):
    if image is None or box is None:
        return None
    height, width = image.shape[:2]
    clipped = _clip_box(box, width, height)
    if clipped is None:
        return None
    x0, y0, x1, y1 = clipped
    crop = image[y0:y1, x0:x1]
    if crop.size == 0 or min(crop.shape[:2]) < 8:
        return None
    return crop


def _vision_box_pixels(rect, width: int, height: int):
    try:
        x = float(rect.origin.x)
        y = float(rect.origin.y)
        w = float(rect.size.width)
        h = float(rect.size.height)
    except Exception:
        return None
    if w <= 0 or h <= 0:
        return None
    x0 = int(round(x * width))
    x1 = int(round((x + w) * width))
    y0 = int(round((1.0 - (y + h)) * height))
    y1 = int(round((1.0 - y) * height))
    return x0, y0, x1, y1


def _quad_box(quad, width: int, height: int):
    try:
        pts = np.array(quad, dtype=np.float32).reshape(-1, 2)
    except Exception:
        return None
    if pts.shape[0] < 4:
        return None
    x0 = int(np.floor(float(pts[:, 0].min())))
    y0 = int(np.floor(float(pts[:, 1].min())))
    x1 = int(np.ceil(float(pts[:, 0].max())))
    y1 = int(np.ceil(float(pts[:, 1].max())))
    if x1 - x0 < 8 or y1 - y0 < 8:
        return None
    x0 = max(0, min(width - 1, x0))
    y0 = max(0, min(height - 1, y0))
    x1 = max(x0 + 1, min(width, x1))
    y1 = max(y0 + 1, min(height, y1))
    return x0, y0, x1, y1


def _make_qr_hit(payload: str, image: np.ndarray, box) -> dict | None:
    cleaned = str(payload or "").strip()
    if not cleaned:
        return None
    crop = _crop_box(image, box)
    if crop is None:
        crop = image
    return {"payload": cleaned, "png": image_to_png(crop), "box": box}


def _vision_qr_hits(image: np.ndarray) -> list[dict]:
    """Apple Vision is disabled. Callers that still import this name get nothing."""
    del image
    return []


def _order_quad(points: np.ndarray) -> np.ndarray:
    quad = np.array(points, dtype=np.float32).reshape(4, 2)
    sums = quad.sum(axis=1)
    diffs = np.diff(quad, axis=1).reshape(-1)
    return np.float32(
        [
            quad[int(np.argmin(sums))],
            quad[int(np.argmin(diffs))],
            quad[int(np.argmax(sums))],
            quad[int(np.argmax(diffs))],
        ]
    )


def _prepared_qr_text(gray: np.ndarray, points) -> str:
    """Straighten a located code and give it a white border so OpenCV can read it."""
    if points is None or gray is None or gray.size == 0:
        return ""
    try:
        quad = _order_quad(points)
    except Exception:
        return ""
    side = 700
    destination = np.float32([[0, 0], [side - 1, 0], [side - 1, side - 1], [0, side - 1]])
    try:
        warp = cv2.warpPerspective(
            gray,
            cv2.getPerspectiveTransform(quad, destination),
            (side, side),
            flags=cv2.INTER_CUBIC,
            borderValue=255,
        )
    except Exception:
        return ""
    quiet = cv2.copyMakeBorder(warp, 40, 40, 40, 40, cv2.BORDER_CONSTANT, value=255)
    try:
        data, _points, _straight = cv2.QRCodeDetector().detectAndDecode(quiet)
    except Exception:
        return ""
    return str(data or "").strip()


def _qr_points_present(image: np.ndarray) -> bool:
    """Cheap check so an empty piece is not redrawn at full size."""
    if image is None or getattr(image, "size", 0) == 0:
        return False
    gray = image if image.ndim == 2 else _to_gray(image)
    height, width = gray.shape[:2]
    views = [gray]
    if height > 240 and width > 240:
        views.append(gray[int(height * 0.15):, int(width * 0.25):])
    for view in views:
        try:
            found, points = cv2.QRCodeDetector().detect(view)
        except Exception:
            continue
        if found and points is not None:
            return True
    return False


def _opencv_prepared_hits(image: np.ndarray) -> list[dict]:
    """Read a code OpenCV can see but not decode until it is straightened."""
    if image is None or getattr(image, "size", 0) == 0:
        return []
    gray = image if image.ndim == 2 else _to_gray(image)
    rgb = _to_rgb(image)
    height, width = gray.shape[:2]
    views = [(gray, 0, 0)]
    if height > 240 and width > 240:
        views.append((gray[int(height * 0.15):, int(width * 0.25):], int(width * 0.25), int(height * 0.15)))
    hits = []
    for view, origin_x, origin_y in views:
        detector = cv2.QRCodeDetector()
        try:
            found, points = detector.detect(view)
        except Exception:
            continue
        if not found or points is None:
            continue
        payload = _prepared_qr_text(view, points)
        if not payload:
            continue
        box = _quad_box(points, int(view.shape[1]), int(view.shape[0]))
        if box is not None:
            x0, y0, x1, y1 = box
            box = (x0 + origin_x, y0 + origin_y, x1 + origin_x, y1 + origin_y)
        hit = _make_qr_hit(payload, rgb, box)
        if hit:
            hits.append(hit)
            return hits
    return hits


def _detector_qr_hits(image: np.ndarray) -> list[dict]:
    if image is None or image.size == 0:
        return []
    gray = image if image.ndim == 2 else _to_gray(image)
    rgb = _to_rgb(image)
    height, width = gray.shape[:2]
    hits: list[dict] = []
    for factory in (_thread_qr_detector, _thread_aruco_detector):
        detector = factory()
        if detector is None:
            continue
        decoded = None
        points = None
        try:
            ok, decoded, points, _straight = detector.detectAndDecodeMulti(gray)
            if not ok:
                decoded, points = None, None
        except Exception:
            decoded, points = None, None
        if decoded is not None and points is not None:
            for text, quad in zip(decoded, np.array(points)):
                hit = _make_qr_hit(str(text or ""), rgb, _quad_box(quad, width, height))
                if hit:
                    hits.append(hit)
        if hits:
            return hits
        data = ""
        pts = None
        try:
            data, pts, _straight = detector.detectAndDecode(gray)
        except Exception:
            try:
                data, pts, _straight = detector.detectAndDecode(gray)
            except Exception:
                data, pts = "", None
        if data:
            hit = _make_qr_hit(str(data), rgb, _quad_box(pts, width, height) if pts is not None else None)
            if hit:
                hits.append(hit)
        if hits:
            return hits
    return hits


_FINDER = np.array(
    [
        [1, 1, 1, 1, 1, 1, 1],
        [1, 0, 0, 0, 0, 0, 1],
        [1, 0, 1, 1, 1, 0, 1],
        [1, 0, 1, 1, 1, 0, 1],
        [1, 0, 1, 1, 1, 0, 1],
        [1, 0, 0, 0, 0, 0, 1],
        [1, 1, 1, 1, 1, 1, 1],
    ],
    dtype=np.uint8,
)


def _finder_score(grid: np.ndarray) -> float:
    n = grid.shape[0]
    corners = (grid[:7, :7], grid[:7, n - 7 :], grid[n - 7 :, :7])
    return float(np.mean([np.mean(corner == _FINDER) for corner in corners]))


def _qr_quads(gray: np.ndarray) -> list[np.ndarray]:
    """Where a code is, even when OpenCV cannot decode it."""
    quads: list[np.ndarray] = []
    for factory in (_thread_qr_detector, _thread_aruco_detector):
        detector = factory()
        if detector is None:
            continue
        try:
            found, points = detector.detect(gray)
        except Exception:
            continue
        if found and points is not None:
            quads.append(np.array(points, dtype=np.float32).reshape(-1, 2)[:4])
    return quads


def _redrawn_qr_hits(gray: np.ndarray, rgb: np.ndarray) -> list[dict]:
    """A blurry small code is sampled module by module and redrawn sharp for OpenCV.

    Low-resolution pay-box codes are too soft for OpenCV to decode directly.
    The corners are still found, so the grid is read cell by cell and a clean
    black-and-white copy is decoded instead.
    """
    height, width = gray.shape[:2]
    work = gray
    scale = 1.0
    if min(height, width) < 360:
        scale = 360.0 / float(min(height, width))
        work = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    quads = _qr_quads(work)
    if not quads:
        return []
    side = 580
    destination = np.float32([[0, 0], [side - 1, 0], [side - 1, side - 1], [0, side - 1]])
    for quad in quads:
        try:
            ordered = _order_quad(quad)
            warp = cv2.warpPerspective(
                work, cv2.getPerspectiveTransform(ordered, destination), (side, side),
                flags=cv2.INTER_CUBIC, borderValue=255,
            )
        except Exception:
            continue
        best = None
        for version in range(1, 8):
            modules = 17 + 4 * version
            cell = side / float(modules)
            centers = (np.arange(modules) + 0.5) * cell
            radius = max(1, int(cell * 0.22))
            samples = np.zeros((modules, modules), dtype=np.float32)
            for row, cy in enumerate(centers):
                y0, y1 = max(0, int(cy) - radius), min(side, int(cy) + radius + 1)
                for col, cx in enumerate(centers):
                    x0, x1 = max(0, int(cx) - radius), min(side, int(cx) + radius + 1)
                    samples[row, col] = float(warp[y0:y1, x0:x1].mean())
            threshold = (float(samples.min()) + float(samples.max())) / 2.0
            grid = (samples < threshold).astype(np.uint8)
            score = _finder_score(grid)
            if best is None or score > best[0]:
                best = (score, grid)
        if best is None or best[0] < 0.8:
            continue
        grid = best[1]
        clean = np.where(grid == 1, 0, 255).astype(np.uint8)
        clean = cv2.resize(clean, None, fx=12, fy=12, interpolation=cv2.INTER_NEAREST)
        clean = cv2.copyMakeBorder(clean, 48, 48, 48, 48, cv2.BORDER_CONSTANT, value=255)
        try:
            data, _points, _straight = cv2.QRCodeDetector().detectAndDecode(clean)
        except Exception:
            data = ""
        if not data:
            continue
        box = _quad_box(quad / scale, width, height)
        hit = _make_qr_hit(str(data), rgb, box)
        if hit:
            return [hit]
    return []


def _shift_box(box, ox: float, oy: float, scale: float):
    if box is None:
        return None
    x0, y0, x1, y1 = box
    if scale and scale != 1:
        x0, y0, x1, y1 = x0 / scale, y0 / scale, x1 / scale, y1 / scale
    return (x0 + ox, y0 + oy, x1 + ox, y1 + oy)


def _boxes_same(left, right) -> bool:
    if left is None or right is None:
        return False
    ax0, ay0, ax1, ay1 = left
    bx0, by0, bx1, by1 = right
    aw = max(1.0, float(ax1 - ax0))
    ah = max(1.0, float(ay1 - ay0))
    bw = max(1.0, float(bx1 - bx0))
    bh = max(1.0, float(by1 - by0))
    acx, acy = (ax0 + ax1) / 2, (ay0 + ay1) / 2
    bcx, bcy = (bx0 + bx1) / 2, (by0 + by1) / 2
    dist = ((acx - bcx) ** 2 + (acy - bcy) ** 2) ** 0.5
    if dist <= 0.45 * max(aw, ah, bw, bh):
        return True
    inter_w = max(0.0, min(ax1, bx1) - max(ax0, bx0))
    inter_h = max(0.0, min(ay1, by1) - max(ay0, by0))
    inter = inter_w * inter_h
    return inter / min(aw * ah, bw * bh) >= 0.4


def _prefer_hit(new_hit: dict, old_hit: dict) -> bool:
    new_box = new_hit.get("box")
    old_box = old_hit.get("box")

    def area(box) -> float:
        if not box:
            return 1e12
        return max(1.0, (box[2] - box[0]) * (box[3] - box[1]))

    if new_box and old_box and area(new_box) < area(old_box) * 0.85 and new_hit.get("png"):
        return True
    return len(new_hit.get("png") or b"") > len(old_hit.get("png") or b"")


def finder_boxes(image: np.ndarray) -> list[tuple[int, int, int, int]]:
    """Pixel boxes around QR finder patterns, including codes that are small on the page.

    Apple's barcode reader skips a code when it is only a small part of a large
    picture. The finder pattern is the nested square (black/white/black/white/black
    in the ratio 1:1:3:1:1), so a small code can still be located and cropped.
    """
    if image is None or getattr(image, "size", 0) == 0 or min(image.shape[:2]) < 48:
        return []
    gray = _to_gray(image)
    binaries = []
    try:
        adaptive = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 5
        )
        binaries.append(adaptive < 128)
    except Exception:
        pass
    try:
        _thr, otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        binaries.append(otsu < 128)
    except Exception:
        pass
    if not binaries:
        return []
    height, width = gray.shape[:2]
    boxes: list[tuple[int, int, int, int]] = []
    for black in binaries:
        boxes.extend(_boxes_from_finders(black, width, height))
        if boxes:
            break
    return _unique_boxes(boxes)[:6]


def _ratio_centers(black: np.ndarray, axis: int) -> list[tuple[float, float, float]]:
    """Centers of 1:1:3:1:1 runs. axis 0 scans rows, axis 1 scans columns."""
    lines = black if axis == 0 else black.T
    hits: list[tuple[float, float, float]] = []
    for index in range(0, lines.shape[0], 2):
        row = lines[index]
        if int(row.sum()) < 8:
            continue
        starts, lengths, colors = _run_lengths(row)
        count = len(lengths)
        if count < 5:
            continue
        for pos in range(count - 4):
            if not (
                colors[pos]
                and not colors[pos + 1]
                and colors[pos + 2]
                and not colors[pos + 3]
                and colors[pos + 4]
            ):
                continue
            left, gap_a, mid, gap_b, right = (float(lengths[pos + offset]) for offset in range(5))
            if mid < 3:
                continue
            unit = mid / 3.0
            if unit < 1.25 or unit > 48:
                continue
            low, high = 0.45 * unit, 1.7 * unit
            if not (low <= left <= high and low <= gap_a <= high and low <= gap_b <= high and low <= right <= high):
                continue
            center = float(starts[pos]) + left + gap_a + (mid / 2.0)
            if axis == 0:
                hits.append((center, float(index), unit))
            else:
                hits.append((float(index), center, unit))
    return hits


def _run_lengths(row: np.ndarray):
    changes = np.flatnonzero(row[1:] != row[:-1]) + 1
    starts = np.r_[0, changes]
    ends = np.r_[changes, row.size]
    return starts, ends - starts, row[starts]


def _cluster_finders(hits: list[tuple[float, float, float]]) -> list[list[float]]:
    clusters: list[list[float]] = []
    for x, y, unit in hits:
        matched = None
        for cluster in clusters:
            dist = ((x - cluster[0]) ** 2 + (y - cluster[1]) ** 2) ** 0.5
            if dist <= max(6.0, 3.4 * max(unit, cluster[2])):
                matched = cluster
                break
        if matched is None:
            clusters.append([x, y, unit, 1.0])
            continue
        seen = matched[3]
        matched[0] = (matched[0] * seen + x) / (seen + 1.0)
        matched[1] = (matched[1] * seen + y) / (seen + 1.0)
        matched[2] = (matched[2] * seen + unit) / (seen + 1.0)
        matched[3] = seen + 1.0
    return [item for item in clusters if item[3] >= 3]


def _line_has_finder(line: np.ndarray, index: int, unit: float) -> bool:
    if line.size < 8:
        return False
    index = int(max(0, min(line.size - 1, round(index))))
    starts, lengths, colors = _run_lengths(line)
    count = len(lengths)
    if count < 5:
        return False
    low, high = 0.4 * unit, 1.8 * unit
    for pos in range(count - 4):
        if not (
            colors[pos]
            and not colors[pos + 1]
            and colors[pos + 2]
            and not colors[pos + 3]
            and colors[pos + 4]
        ):
            continue
        run_start = int(starts[pos])
        run_end = run_start + int(sum(int(lengths[pos + offset]) for offset in range(5)))
        if index < run_start or index >= run_end:
            continue
        left, gap_a, mid, gap_b, right = (float(lengths[pos + offset]) for offset in range(5))
        if mid < 3:
            continue
        if low <= left <= high and low <= gap_a <= high and low <= gap_b <= high and low <= right <= high:
            if 0.55 * unit <= mid / 3.0 <= 1.8 * unit:
                return True
    return False


def _confirmed_finder(black: np.ndarray, finder: list[float]) -> bool:
    height, width = black.shape[:2]
    x = int(round(finder[0]))
    y = int(round(finder[1]))
    if x < 0 or y < 0 or x >= width or y >= height or not bool(black[y, x]):
        return False
    return _line_has_finder(black[y], x, finder[2]) and _line_has_finder(black[:, x], y, finder[2])


def _group_finders(finders: list[list[float]]) -> list[list[list[float]]]:
    groups: list[list[list[float]]] = []
    for finder in finders:
        unit = max(finder[2], 1.0)
        placed = False
        for group in groups:
            for other in group:
                dist = ((finder[0] - other[0]) ** 2 + (finder[1] - other[1]) ** 2) ** 0.5
                limit = 72.0 * max(unit, other[2])
                near = 8.0 * max(unit, other[2])
                if near < dist <= limit:
                    group.append(finder)
                    placed = True
                    break
            if placed:
                break
        if not placed:
            groups.append([finder])
    return groups


def _box_for_finders(group: list[list[float]], width: int, height: int):
    xs = [item[0] for item in group]
    ys = [item[1] for item in group]
    unit = float(np.median([item[2] for item in group]))
    # One confirmed finder still gets a crop large enough for a short link QR.
    # Three finders describe the real corners, so the pad only adds the quiet zone.
    pad = (8.0 if len(group) >= 2 else 36.0) * unit
    x0 = max(0, int(np.floor(min(xs) - pad)))
    y0 = max(0, int(np.floor(min(ys) - pad)))
    x1 = min(width, int(np.ceil(max(xs) + pad)))
    y1 = min(height, int(np.ceil(max(ys) + pad)))
    if x1 - x0 < 16 or y1 - y0 < 16:
        return None
    if len(group) == 1 and (x1 - x0) * (y1 - y0) > 0.28 * width * height:
        return None
    return x0, y0, x1, y1


def _boxes_from_finders(black: np.ndarray, width: int, height: int) -> list[tuple[int, int, int, int]]:
    # Rows are enough: a finder looks the same at every right angle, and each
    # hit is still confirmed down the column. Skipping the column sweep keeps
    # this off the slow path.
    hits = _ratio_centers(black, 0)
    finders = [item for item in _cluster_finders(hits) if _confirmed_finder(black, item)]
    boxes = []
    for group in _group_finders(finders):
        box = _box_for_finders(group, width, height)
        if box is not None:
            boxes.append(box)
    boxes.sort(key=lambda item: (item[2] - item[0]) * (item[3] - item[1]))
    return boxes


def _unique_boxes(boxes: list[tuple[int, int, int, int]]) -> list[tuple[int, int, int, int]]:
    kept: list[tuple[int, int, int, int]] = []
    for box in boxes:
        if any(_boxes_same(box, previous) for previous in kept):
            continue
        kept.append(box)
    return kept


def _tiles_with_origin(image: np.ndarray, rows: int = 2, cols: int = 2, overlap: float = 0.32):
    height, width = image.shape[:2]
    if height < 80 or width < 80:
        return []
    tile_h = max(64, int(height / rows * (1 + overlap)))
    tile_w = max(64, int(width / cols * (1 + overlap)))
    tile_h = min(tile_h, height)
    tile_w = min(tile_w, width)
    step_y = max(1, (height - tile_h) // max(rows - 1, 1)) if rows > 1 else 0
    step_x = max(1, (width - tile_w) // max(cols - 1, 1)) if cols > 1 else 0
    tiles = []
    for row in range(rows):
        y0 = min(row * step_y, max(0, height - tile_h))
        y1 = min(height, y0 + tile_h)
        for col in range(cols):
            x0 = min(col * step_x, max(0, width - tile_w))
            x1 = min(width, x0 + tile_w)
            tiles.append((image[y0:y1, x0:x1], (x0, y0)))
    return tiles


def _hits_from_image(image: np.ndarray) -> list[dict]:
    if image is None or getattr(image, "size", 0) == 0 or min(image.shape[:2]) < 8:
        return []
    boosted = _ensure_min_side(image, 240)
    scale = float(boosted.shape[1]) / float(max(image.shape[1], 1))
    hits: list[dict] = []
    for hit in _detector_qr_hits(boosted):
        hit["box"] = _shift_box(hit.get("box"), 0, 0, scale)
        hits.append(hit)
    return hits


def decode_qr_hits(image: np.ndarray, thorough: bool = False, locate: bool = True) -> list[dict]:
    """Every QR on the image, with a PNG crop of that code.

    Two codes that share a link stay separate when they sit in different places.
    OpenCV is the only reader. Apple Vision is not used.
    """
    if image is None or getattr(image, "size", 0) == 0 or min(image.shape[:2]) < 8:
        return []
    found: list[dict] = []

    def add_many(hits: list[dict], ox: float = 0, oy: float = 0, scale: float = 1.0) -> None:
        for hit in hits:
            payload = str(hit.get("payload") or "").strip()
            if not payload:
                continue
            box = _shift_box(hit.get("box"), ox, oy, scale)
            candidate = {"payload": payload, "png": hit.get("png") or b"", "box": box}
            for previous in found:
                if previous["payload"].lower() != payload.lower():
                    continue
                if box is None or previous.get("box") is None or _boxes_same(previous.get("box"), box):
                    if _prefer_hit(candidate, previous):
                        previous["png"] = candidate["png"]
                        if box is not None:
                            previous["box"] = box
                    break
            else:
                found.append(candidate)

    def scan(picture, ox: float = 0, oy: float = 0, scale: float = 1.0) -> None:
        # A full page is left to the PDF finder crops. OpenCV on that large
        # drawing is what made one PDF feel stuck.
        if max(picture.shape[:2]) <= 1800:
            add_many(_detector_qr_hits(picture), ox, oy, scale)

    scan(image)
    longest = max(image.shape[:2])
    shortest = min(image.shape[:2])
    # A wide strip is not a page. Read it in square windows so the 1800px page limit does not skip it.
    if not found and longest > 1800 and shortest <= 1000:
        _scan_strip_windows(image, found, scan)
    # Last try for a blurry small code: redraw it sharp. Run once, not on every piece.
    if not found and longest <= 1800 and shortest <= 1000:
        _redraw_into(image, found, add_many)
    # A small printed code is often under 300px in a logo or chunk. Scale that
    # crop up once; a full page is left alone so this does not run twelve times.
    if not found and shortest < 460 and longest <= 1400:
        scaled = _ensure_min_side(image, 520)
        if scaled is not image and max(scaled.shape[:2]) != longest:
            scale = float(scaled.shape[1]) / float(max(image.shape[1], 1))
            scan(scaled, scale=scale)

    # On a crop or logo, locate a code that was too small for the first look.
    # A full page is handled by the PDF scan, which re-renders the spot sharply
    # instead of stretching this picture.
    if locate and (thorough or not found) and longest <= 1800:
        _decode_finder_crops(image, found, scan)
    return found


def _redraw_into(image: np.ndarray, found: list[dict], add_many) -> None:
    height, width = image.shape[:2]
    if width > height * 1.6:
        # A wide chunk: the code sits at one end, so redraw only the two end squares.
        side = height
        pieces = [(image[:, width - side :], width - side), (image[:, :side], 0)]
    else:
        pieces = [(image, 0)]
    for piece, offset in pieces:
        hits = _redrawn_qr_hits(_to_gray(piece), _to_rgb(piece))
        if hits:
            add_many(hits, ox=offset)
            return


def _scan_strip_windows(image: np.ndarray, found: list[dict], scan) -> None:
    height, width = image.shape[:2]
    horizontal = width >= height
    side = min(height, width)
    length = max(height, width)
    start = max(0, length - side)
    # The code in a pay box sits at one end, so look at the two ends and the middle only.
    positions = [start, 0, start // 2]
    for offset in positions:
        if horizontal:
            window = image[:, offset : offset + side]
            ox, oy = offset, 0
        else:
            window = image[offset : offset + side, :]
            ox, oy = 0, offset
        fitted = _fit_image(window, 900)
        scale = float(fitted.shape[1]) / float(max(window.shape[1], 1))
        before = len(found)
        scan(fitted, ox=ox, oy=oy, scale=scale)
        if len(found) > before:
            return


def _decode_finder_crops(image: np.ndarray, found: list[dict], scan) -> None:
    try:
        boxes = finder_boxes(image)
    except Exception as exc:
        logger.debug("QR finder locate failed: %s", exc)
        return
    for x0, y0, x1, y1 in boxes:
        if any(_boxes_same((x0, y0, x1, y1), hit.get("box")) for hit in found):
            continue
        crop = image[y0:y1, x0:x1]
        if crop.size == 0 or min(crop.shape[:2]) < 12:
            continue
        boosted = _ensure_min_side(crop, 560)
        scale = float(boosted.shape[1]) / float(max(crop.shape[1], 1))
        scan(boosted, ox=x0, oy=y0, scale=scale)


def decode_qr_payloads(image: np.ndarray, thorough: bool = False) -> list[str]:
    """Decode QR payloads from a page render or an embedded image."""
    return [str(hit.get("payload") or "") for hit in decode_qr_hits(image, thorough) if hit.get("payload")]


def pixmap_to_numpy(pix) -> np.ndarray:
    samples = np.frombuffer(pix.samples, dtype=np.uint8)
    array = samples.reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        array = array[:, :, :3]
    elif pix.n == 1:
        array = array[:, :, 0]
    return np.ascontiguousarray(array.copy())


@lru_cache(maxsize=1)
def warmup() -> str:
    """Load OCR and QR detectors once so the first PDF is not delayed."""
    dummy = np.full((64, 64, 3), 255, dtype=np.uint8)
    try:
        decode_qr_payloads(dummy)
    except Exception as exc:
        logger.warning("QR warmup failed: %s", exc)
    try:
        recognize_text(dummy)
    except Exception as exc:
        logger.warning("OCR warmup failed: %s", exc)
    return ocr_backend()

