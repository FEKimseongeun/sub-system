import cv2
import numpy as np
from pathlib import Path
from .config import COLOR_RANGES, MASK_DIR, GRAY_S_MAX, GRAY_V_MAX

def ensure_dir(): MASK_DIR.mkdir(parents=True, exist_ok=True)

def non_black_mask(bgr: np.ndarray) -> np.ndarray:
    """
    검정/회색(무채색)을 제외하는 마스크 (유색만 True).
    """
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    h,s,v = cv2.split(hsv)
    # 유색 조건: 채도 S > GRAY_S_MAX AND (명도 V > GRAY_V_MAX OR 색상 성분이 충분)
    # 단순화: s>GRAY_S_MAX & v>GRAY_V_MAX
    mask = cv2.inRange(hsv,
                       (0, GRAY_S_MAX+1, GRAY_V_MAX+1),
                       (179, 255, 255))
    return mask  # 255: 유색, 0: 검정/회색

def color_group_masks(bgr: np.ndarray) -> dict[str, np.ndarray]:
    """
    컬러 그룹별 마스크. COLOR_RANGES 기반.
    red1/red2는 합쳐 red로 리턴.
    """
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    masks = {}
    for key, (low, high) in COLOR_RANGES.items():
        m = cv2.inRange(hsv, np.array(low, np.uint8), np.array(high, np.uint8))
        if key.startswith("red"):
            masks.setdefault("red", np.zeros_like(m))
            masks["red"] = cv2.bitwise_or(masks["red"], m)
        else:
            masks[key] = m

    # 노이즈 정리
    k = np.ones((3,3), np.uint8)
    for kname in list(masks.keys()):
        m = cv2.morphologyEx(masks[kname], cv2.MORPH_OPEN, k, iterations=1)
        m = cv2.dilate(m, k, iterations=1)
        masks[kname] = m
    return masks

def save_mask_preview(stem: str, bgr: np.ndarray, masks: dict[str, np.ndarray]):
    ensure_dir()
    for k, m in masks.items():
        prev = cv2.bitwise_and(bgr, bgr, mask=m)
        cv2.imwrite((MASK_DIR / f"{stem}_{k}.png").as_posix(), prev)
