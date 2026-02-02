from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"
PDF_DIR = DATA_DIR / "pdf"
IMG_DIR = DATA_DIR / "images"
MASK_DIR = DATA_DIR / "masks"
OUT_DIR = PROJECT_ROOT / "outputs"
OUT_CSV = OUT_DIR / "csv"
OUT_XLSX = OUT_DIR / "excel"

DPI = 350  # PDF→이미지 해상도

# 도면에서 실제 쓰이는 대표 색(예시). 검정 제외.
# H:0-179, S/V:0-255
COLOR_RANGES = {
    "blue":  ((100, 80,  80), (130, 255, 255)),
    "red1":  ((0,   80,  80), (10,  255, 255)),   # red는 두 구간
    "red2":  ((170, 80,  80), (179, 255, 255)),
    "green": ((35,  60,  80), (85,  255, 255)),
}

# 리포팅용 대표색 (없으면 자동 대체)
COLOR_REPR = {
    "blue":  {"rgb": (0, 102, 204), "hex": "#0066CC"},
    "red":   {"rgb": (204, 0, 0),   "hex": "#CC0000"},
    "green": {"rgb": (0, 153, 0),   "hex": "#009900"},
}

# 검정/회색(무채색) 필터 기준: '검정 아님'을 판단할 때 사용
GRAY_S_MAX = 40   # 채도 S<=40 이하면 무채색(검정/회색)으로 간주
GRAY_V_MAX = 120  # 명도 V<=120 이하면 어두운(검정)으로 간주

# TrOCR 재인식 사용 여부(긴 라인번호 등)
USE_TROCR_FALLBACK = True

# ‘대표 시스템명’ 따로 뽑아야 하면 True로 바꾸고 title box 검출 로직 추가 가능
USE_TITLE_BOX = False  # 여기선 심플하게 OFF
