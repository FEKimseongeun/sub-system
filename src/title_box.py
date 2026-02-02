import cv2
import numpy as np
from pathlib import Path
from typing import Optional
from .config import MIN_TITLE_BOX_AREA

def find_thick_boxes(color_img: np.ndarray) -> list[tuple[np.ndarray, tuple[int,int,int,int]]]:
    """
    굵은 박스 후보 탐지: (ROI이미지, bbox(x,y,w,h)) 리스트 반환
    color_img: 해당 색상 마스크 적용된 BGR
    """
    gray = cv2.cvtColor(color_img, cv2.COLOR_BGR2GRAY)
    # 에지→윤곽
    edges = cv2.Canny(gray, 50, 150)
    contours,_ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    boxes = []
    for cnt in contours:
        x,y,w,h = cv2.boundingRect(cnt)
        area = w*h
        if area < MIN_TITLE_BOX_AREA:
            continue
        # 사각형 근사
        peri = cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, 0.02*peri, True)
        if len(approx) == 4:
            roi = color_img[y:y+h, x:x+w].copy()
            boxes.append((roi, (x,y,w,h)))
    # 큰 박스 우선
    boxes.sort(key=lambda t: t[0].shape[0]*t[0].shape[1], reverse=True)
    return boxes

def ocr_title_from_boxes(roi_list: list[np.ndarray], ocr_fn) -> Optional[str]:
    """
    roi_list 각각에 대해 ocr_fn(이미지→문자열) 호출하여 가장 '타이틀다움' 문구 선택
    여기서는 가장 긴 문자열을 타이틀로 택한다(실무에 맞춰 규칙 보완 가능).
    """
    best = None
    best_len = 0
    for roi, _ in roi_list:
        text = ocr_fn(roi)
        text = (text or "").strip()
        # 태그 패턴(하이픈·숫자 위주)은 제외하고 단어 위주 문장 선호
        if len(text) > best_len and any(ch.isalpha() for ch in text) and " " in text:
            best = text
            best_len = len(text)
    return best
