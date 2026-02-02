# -*- coding: utf-8 -*-
"""
Excel 좌표로 PDF에 type별 투명 마크업(annotation) 추가

- special / special_item : green
- line                  : yellow
- instrument            : blue
(채움 + 투명도 적용 → 뒤 텍스트 보이게)

사용:
python pdf_bbox_markup.py --pdf "../../data/PDF/master_1120.pdf" \
                          --excel "out/final_tags.xlsx" \
                          --out "out/master_1120_marked.pdf"
"""

from pathlib import Path
import argparse
import pandas as pd
import pymupdf as fitz  # PyMuPDF


# ---------------------------
# RGB(0~255) → float(0~1)
# ---------------------------
def rgb255_to_float(rgb):
    return tuple([c / 255.0 for c in rgb])


# type → (stroke_rgb255, fill_rgb255, opacity)
TYPE_STYLE_255 = {
    "special":      ((0, 180, 0),   (0, 255, 0),    0.25),
    "special_item": ((0, 180, 0),   (0, 255, 0),    0.25),  # alias
    "line":         ((200, 200, 0), (255, 255, 0),  0.25),
    "line_no":      ((200, 200, 0), (255, 255, 0),  0.25),  # line과 동일
    "instrument":   ((0, 90, 200),  (0, 140, 255),  0.25),
    "valve":        ((200, 0, 0),   (255, 80, 80),  0.25),  # 빨간색 계열
    "equipment":    ((128, 0, 128), (180, 0, 255),  0.25),  # 보라색 계열
}
DEFAULT_STYLE_255 = ((255, 0, 0), (255, 0, 0), 0.15)


def pick_bbox_columns(df: pd.DataFrame):
    """
    df에 어떤 bbox 컬럼이 있는지 확인해서 (x0,y0,x1,y1) 형태로 반환
    지원:
      - x0,y0,x1,y1
      - x1,y1,x2,y2
    """
    cols = {c.lower().strip() for c in df.columns}

    if {"x0", "y0", "x1", "y1"}.issubset(cols):
        return "x0", "y0", "x1", "y1"

    if {"x1", "y1", "x2", "y2"}.issubset(cols):
        return "x1", "y1", "x2", "y2"

    raise ValueError(
        "bbox 컬럼을 찾지 못했음. "
        "x0,y0,x1,y1 또는 x1,y1,x2,y2 중 하나는 있어야 합니다."
    )


def normalize_type(t: str) -> str:
    if not isinstance(t, str):
        return ""
    return t.strip().lower()


def add_rect_markup(page: fitz.Page, rect: fitz.Rect, ttype: str):
    """
    페이지에 사각 annotation 추가 (채움+투명도)
    PyMuPDF 버전 호환 고려
    """
    stroke255, fill255, opacity = TYPE_STYLE_255.get(ttype, DEFAULT_STYLE_255)
    stroke = rgb255_to_float(stroke255)
    fill   = rgb255_to_float(fill255)

    annot = page.add_rect_annot(rect)

    # 색상 (0~1 float로 넣어야 함)
    annot.set_colors(stroke=stroke, fill=fill)

    # 투명도
    annot.set_opacity(opacity)

    # 테두리 두께(버전별 API 차이 대비)
    try:
        # 신버전
        annot.set_border(width=0.5)
    except Exception:
        try:
            # 구버전(dict 방식)
            border = annot.border
            border["width"] = 0.5
            annot.set_border(border)
        except Exception:
            pass

    annot.update()


def markup_pdf_from_excel(pdf_path: Path, excel_path: Path, out_path: Path):
    df = pd.read_excel(excel_path)
    if df.empty:
        raise ValueError("엑셀 데이터가 비어있습니다.")

    df.columns = [c.strip() for c in df.columns]
    xA, yA, xB, yB = pick_bbox_columns(df)

    # page/type 필수
    if "page" not in [c.lower() for c in df.columns]:
        raise ValueError("필수 컬럼 없음: page")
    if "type" not in [c.lower() for c in df.columns]:
        raise ValueError("필수 컬럼 없음: type")

    df["page"] = df["page"].astype(int)
    df["type"] = df["type"].apply(normalize_type)

    doc = fitz.open(pdf_path.as_posix())

    for page_num, grp in df.groupby("page"):
        if page_num < 1 or page_num > doc.page_count:
            continue

        page = doc[page_num - 1]

        for _, row in grp.iterrows():
            ttype = row["type"]
            if ttype not in TYPE_STYLE_255:
                continue  # 요구한 타입만 마크업

            x0 = float(row[xA]); y0 = float(row[yA])
            x1 = float(row[xB]); y1 = float(row[yB])

            rect = fitz.Rect(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))

            # 작은 노이즈 박스는 스킵
            area = rect.width * rect.height
            if area < 4:
                continue

            add_rect_markup(page, rect, ttype)

    # 저장
    doc.save(out_path.as_posix())
    doc.close()
    print(f"✅ Marked PDF saved: {out_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True, help="원본 PDF 경로")
    ap.add_argument("--excel", required=True, help="좌표 엑셀(final_tags.xlsx) 경로")
    ap.add_argument("--out", required=True, help="저장할 마크업 PDF 경로")
    args = ap.parse_args()

    markup_pdf_from_excel(Path(args.pdf), Path(args.excel), Path(args.out))


if __name__ == "__main__":
    main()
