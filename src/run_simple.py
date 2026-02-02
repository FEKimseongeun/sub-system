# src/run_simple.py
from pathlib import Path
import cv2
import pandas as pd
from loguru import logger

from .logger import setup_logger
from .config import (
    PDF_DIR, IMG_DIR, OUT_CSV, OUT_XLSX,
    COLOR_REPR, USE_TROCR_FALLBACK
)
from .pdf_io import pdf_to_images
from .color_mask import non_black_mask, color_group_masks, save_mask_preview
from .ocr_doctr import doctr_pairs_from_bgr
from .tag_rules import classify

# (선택) TrOCR 재인식
if USE_TROCR_FALLBACK:
    from .ocr_trocr import trocr_recognize_bgr as trocr_recog
else:
    def trocr_recog(_):
        return None


def run_one_image(img_path: Path, page_no: int):
    bgr = cv2.imread(img_path.as_posix(), cv2.IMREAD_COLOR)
    if bgr is None:
        logger.warning(f"Cannot read image: {img_path}")
        return []

    # 1) (참고) 무채색 제외 전체 마스크 – 필요시 활용
    _ = non_black_mask(bgr)

    # 2) 색상 그룹 마스크
    masks = color_group_masks(bgr)
    save_mask_preview(img_path.stem, bgr, masks)

    rows = []

    for color_key, mask in masks.items():
        color_name = color_key  # 'red' / 'blue' / 'green' ...
        rep = COLOR_REPR.get(color_name, {"rgb": (0, 0, 0), "hex": "#000000"})
        group_rgb, group_hex = rep["rgb"], rep["hex"]

        # 색상 부분만 남긴 이미지
        colored = cv2.bitwise_and(bgr, bgr, mask=mask)

        # 3) 메모리 기반 docTR OCR
        try:
            pairs = doctr_pairs_from_bgr(colored)  # [( (x1,y1,x2,y2), text ), ...]
            logger.info(f"OCR pairs ({color_name}) sample={pairs[:2]!r}, count={len(pairs)}")
        except Exception as e:
            logger.error(f"OCR failed for color {color_name} on {img_path.name}: {e}")
            pairs = []

        # 4) (선택) TrOCR 재인식: 긴 라인번호 의심만
        for (x1, y1, x2, y2), text in pairs:
            txt = (text or "").strip()
            if not txt:
                continue
            # 너무 짧은 잡텍스트 컷
            if len(txt) == 1 and not txt.isdigit():
                continue

            if USE_TROCR_FALLBACK and (txt.count("-") >= 2 and len(txt) >= 12):
                crop = colored[y1:y2, x1:x2]
                try:
                    re_txt = trocr_recog(crop)
                    if re_txt and len(re_txt) >= len(txt) - 1:
                        txt = re_txt.strip()
                except Exception as e:
                    logger.debug(f"TrOCR fallback failed: {e}")

            tag_type = classify(txt)

            rows.append({
                "tag": txt,
                "type": tag_type,
                "color_group": color_name,
                "rgb": group_rgb,
                "hex": group_hex,
                "page": page_no,
                "x1": x1, "y1": y1, "x2": x2, "y2": y2
            })

    return rows


def main():
    setup_logger(Path("simple_pipeline.log"))
    OUT_CSV.mkdir(parents=True, exist_ok=True)
    OUT_XLSX.mkdir(parents=True, exist_ok=True)
    IMG_DIR.mkdir(parents=True, exist_ok=True)

    pdfs = sorted(PDF_DIR.glob("*.pdf"))
    if not pdfs:
        logger.error("Put PDF files into data/pdf/")
        return

    all_rows = []
    for pdf in pdfs:
        images = pdf_to_images(pdf)
        for i, img in enumerate(images, start=1):
            all_rows.extend(run_one_image(img, page_no=i))

    df = pd.DataFrame(all_rows)
    if not df.empty:
        # 동일 (tag, page, color_group) 중복 제거
        df = (df
              .sort_values(["tag", "page", "color_group", "x1", "y1"])
              .drop_duplicates(["tag", "page", "color_group"], keep="first"))

    csv_path = OUT_CSV / "colored_text_map.csv"
    xlsx_path = OUT_XLSX / "colored_text_map.xlsx"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    df.to_excel(xlsx_path, index=False)
    logger.info(f"Saved → {csv_path}")
    logger.info(f"Saved → {xlsx_path}")


if __name__ == "__main__":
    main()
