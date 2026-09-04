import os
import pytesseract
import pymupdf

from PIL import Image, ImageEnhance, ImageFilter


# Windows Tesseract location
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


def preprocess_image(image: Image.Image) -> Image.Image:
    """
    Improve an image before OCR.
    """

    # Convert to RGB
    if image.mode != "RGB":
        image = image.convert("RGB")

    # Convert to grayscale
    image = image.convert("L")

    # Increase contrast
    image = ImageEnhance.Contrast(image).enhance(1.8)

    # Sharpen
    image = image.filter(ImageFilter.SHARPEN)

    return image


def ocr_image(image: Image.Image):
    """
    OCR an image and return text + word coordinates.
    """

    processed = preprocess_image(image)

    data = pytesseract.image_to_data(
        processed,
        output_type=pytesseract.Output.DICT,
        config="--psm 6"
    )

    words = []

    for i in range(len(data["text"])):

        text = data["text"][i].strip()

        if not text:
            continue

        confidence = float(data["conf"][i])

        words.append({
            "text": text,
            "x": int(data["left"][i]),
            "y": int(data["top"][i]),
            "width": int(data["width"][i]),
            "height": int(data["height"][i]),
            "confidence": confidence
        })

    full_text = " ".join(
        word["text"] for word in words
    )

    return {
        "text": full_text,
        "words": words
    }


def ocr_pdf(pdf_path: str):
    """
    OCR every page of a scanned PDF.
    """

    pdf = pymupdf.open(pdf_path)

    pages = []

    for page_number, page in enumerate(pdf):

        # Render PDF page at high resolution
        pixmap = page.get_pixmap(
            matrix=pymupdf.Matrix(2.5, 2.5),
            alpha=False
        )

        image = Image.frombytes(
            "RGB",
            [pixmap.width, pixmap.height],
            pixmap.samples
        )

        result = ocr_image(image)

        pages.append({
            "page": page_number + 1,
            "width": pixmap.width,
            "height": pixmap.height,
            "text": result["text"],
            "words": result["words"]
        })

    pdf.close()

    return pages