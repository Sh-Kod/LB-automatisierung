import logging
import yaml
from pathlib import Path

def erstelle_logger():
    with open("C:\\dcp_automatisierung\\config.yaml", "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    log_datei = config["logging"]["log_datei"]
    log_level = config["logging"].get("log_level", "INFO")
    Path(log_datei).parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=getattr(logging, log_level),
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_datei, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger("dcp_automatisierung")

logger = erstelle_logger()
