# utils.py
from PIL import Image, ImageEnhance, ImageFilter
import cv2
import numpy as np

# ---------- Dateiformate & Konfiguration ----------
bildformate = (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp")
pdfformate = (".pdf",)
output_txt_file = "treffer_ausgabe.txt"
max_seiten_pro_word_datei = 20


# ---------- Text-Utils ----------
def bereinige_zeile(zeile: str) -> str:
    return zeile.strip().replace("\t", " ")


def ist_treffer(zeile: str, suchbegriffe: list) -> list:
    zeile_lc = zeile.lower()
    return [wort for wort in suchbegriffe if wort in zeile_lc]


# ---------- Bild-Preprocessing ----------
def preprocess_pillow(image_path, contrast=2.0, resize_width=2000):
    """
    Optimierung mit Pillow:
    - Umwandeln in Graustufen
    - Kontrast erhöhen
    - Schärfen
    - Resize (Skalieren auf Zielbreite)
    - Binarisierung (Schwarz/Weiß)
    """
    img = Image.open(image_path).convert("L")

    # Kontrast erhöhen
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(contrast)

    # Schärfen
    img = img.filter(ImageFilter.SHARPEN)

    # Resize (Seitenverhältnis beibehalten)
    w_percent = resize_width / float(img.size[0])
    h_size = int(float(img.size[1]) * w_percent)
    img = img.resize((resize_width, h_size), Image.LANCZOS)

    # Binarisierung
    threshold = 128
    img = img.point(lambda x: 255 if x > threshold else 0, "L")

    return img


def preprocess_opencv(image_input):
    """
    Optimierung mit OpenCV:
    - Graustufen
    - Rauschen entfernen (MedianBlur)
    - Adaptive Binarisierung (Otsu)
    - Morphologische Operationen zur Rauschunterdrückung
    - Akzeptiert sowohl Pfad als auch PIL-Image
    """
    # Pfad oder numpy/PIL-Array
    if isinstance(image_input, str):
        img = cv2.imread(image_input, cv2.IMREAD_GRAYSCALE)
    elif isinstance(image_input, Image.Image):
        img = np.array(image_input.convert("L"))
    elif isinstance(image_input, np.ndarray):
        img = image_input
    else:
        raise ValueError("preprocess_opencv: ungültiger Input")

    img = cv2.medianBlur(img, 3)
    _, img = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    kernel = np.ones((1, 1), np.uint8)
    img = cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel)

    return Image.fromarray(img)
