# src/ocr_trocr.py
from typing import Optional
import torch
import numpy as np
from PIL import Image
from transformers import TrOCRProcessor, VisionEncoderDecoderModel

_processor = TrOCRProcessor.from_pretrained("microsoft/trocr-base-printed")
_model = VisionEncoderDecoderModel.from_pretrained("microsoft/trocr-base-printed")
_model.eval()

def trocr_recognize_bgr(crop_bgr: np.ndarray) -> Optional[str]:
    rgb = crop_bgr[:, :, ::-1]
    pil = Image.fromarray(rgb)
    inputs = _processor(images=pil, return_tensors="pt")
    with torch.no_grad():
        gen = _model.generate(**inputs, max_length=96)
    text = _processor.batch_decode(gen, skip_special_tokens=True)[0]
    return text.strip() if text else None
