from typing import List, Tuple
import cv2, numpy as np
from doctr.models import ocr_predictor
from doctr.io import DocumentFile

_ocr = ocr_predictor('db_resnet50','crnn_vgg16_bn',pretrained=True)


def doctr_pairs_from_bgr(img_bgr: np.ndarray) -> List[Tuple[Tuple[int,int,int,int], str]]:
    # BGR -> RGB (docTR는 RGB 기대)
    rgb = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)
    # 메모리에서 PNG로 인코딩 → bytes
    ok, buf = cv2.imencode(".png", cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR))
    if not ok:
        return []
    png_bytes = buf.tobytes()

    # bytes는 from_images에서 지원됨
    doc = DocumentFile.from_images([png_bytes])  # ← 핵심! 경로/임시파일 불필요
    result = _ocr(doc)

    out = []
    if not result.pages:
        return out
    page = result.pages[0]
    ph, pw = page.dimensions

    for block in page.blocks:
        for line in block.lines:
            for word in line.words:
                txt = (getattr(word, "value", "") or "").strip()
                if not txt:
                    continue
                geom = getattr(word, "geometry", None)
                if not geom:
                    continue
                # (x_min, y_min, x_max, y_max) or ((x_min,y_min),(x_max,y_max))
                try:
                    x_min, y_min, x_max, y_max = geom
                except Exception:
                    try:
                        (x_min, y_min), (x_max, y_max) = geom
                    except Exception:
                        continue
                x1, y1 = int(x_min * pw), int(y_min * ph)
                x2, y2 = int(x_max * pw), int(y_max * ph)
                if x2 <= x1 or y2 <= y1:
                    continue
                out.append(((x1, y1, x2, y2), txt))
    return out