import yaml
from PIL import Image
import pytesseract

def lade_config():
    with open("C:\\dcp_automatisierung\\config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def lese_text_aus_bild(bildpfad):
    try:
        config = lade_config()
        pytesseract.pytesseract.tesseract_cmd = config["tesseract"]["pfad"]
        img = Image.open(bildpfad)
        text = pytesseract.image_to_string(img, lang="deu+eng")
        return text.strip()
    except Exception as e:
        print(f"OCR Fehler: {e}")
        return ""
