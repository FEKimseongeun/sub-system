from pathlib import Path
import fitz
from .config import IMG_DIR, DPI

def pdf_to_images(pdf_path: Path) -> list[Path]:
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    out = []
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=DPI, alpha=False)
        p = IMG_DIR / f"{pdf_path.stem}_p{i:03d}.png"
        pix.save(p.as_posix())
        out.append(p)
    return out
