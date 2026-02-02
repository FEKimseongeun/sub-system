from loguru import logger
from pathlib import Path

def setup_logger(log_path: Path):
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logger.add(log_path, rotation="1 MB", level="INFO")
    return logger
