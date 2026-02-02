from typing import List, Tuple, Optional
import numpy as np

BBox = Tuple[int,int,int,int]

def iou(a: BBox, b: BBox) -> float:
    ax1,ay1,ax2,ay2 = a; bx1,by1,bx2,by2 = b
    ix1,iy1 = max(ax1,bx1), max(ay1,by1)
    ix2,iy2 = min(ax2,bx2), min(ay2,by2)
    if ix2<=ix1 or iy2<=iy1: return 0.0
    inter = (ix2-ix1)*(iy2-iy1)
    area_a = (ax2-ax1)*(ay2-ay1)
    area_b = (bx2-bx1)*(by2-by1)
    return inter / (area_a + area_b - inter + 1e-6)

def center_dist(a: BBox, b: BBox) -> float:
    ax=(a[0]+a[2])/2; ay=(a[1]+a[3])/2
    bx=(b[0]+b[2])/2; by=(b[1]+b[3])/2
    return ((ax-bx)**2 + (ay-by)**2)**0.5

def match_symbol_to_text(
    symbol_boxes: List[Tuple[BBox, str]],  # [(bbox, symbol_class)]
    text_boxes: List[Tuple[BBox, str]],    # [(bbox, text)]
    max_dist: float = 120.0
) -> List[Tuple[str, str, BBox]]:
    """
    각 심볼에 가장 가까운 텍스트를 1:1 매칭 (동일 색상 그룹 내에서 호출)
    return: [(symbol_class, text, text_bbox)]
    """
    matched = []
    used = set()
    for sb, scls in symbol_boxes:
        best = None; best_score = 1e9; best_i = -1
        for i,(tb, txt) in enumerate(text_boxes):
            if i in used: continue
            d = center_dist(sb, tb)
            if d < best_score and d <= max_dist:
                best = (txt, tb); best_score = d; best_i = i
        if best is not None:
            used.add(best_i)
            matched.append((scls, best[0], best[1]))
    return matched
