import os

import pymupdf
import pytesseract

from PIL import Image, ImageEnhance, ImageFilter


TESSERACT_PATH = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)

if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = (
        TESSERACT_PATH
    )


def preprocess_image(image):

    if image.mode != "RGB":
        image = image.convert("RGB")

    image = image.convert("L")

    image = ImageEnhance.Contrast(
        image
    ).enhance(1.8)

    image = image.filter(
        ImageFilter.SHARPEN
    )

    return image


def ocr_image(image):

    image = preprocess_image(image)

    data = pytesseract.image_to_data(
        image,
        output_type=pytesseract.Output.DICT,
        config="--psm 6"
    )

    words = []

    for i in range(len(data["text"])):

        text = data["text"][i].strip()

        if not text:
            continue

        try:
            confidence = float(
                data["conf"][i]
            )
        except:
            confidence = 0

        words.append({
            "text": text,
            "x": int(data["left"][i]),
            "y": int(data["top"][i]),
            "width": int(data["width"][i]),
            "height": int(data["height"][i]),
            "confidence": confidence,
        })

    return words


def ocr_pdf(pdf_path):

    pdf = pymupdf.open(pdf_path)

    pages = []

    for page_number, page in enumerate(
        pdf,
        start=1
    ):

        pixmap = page.get_pixmap(
            matrix=pymupdf.Matrix(3, 3),
            alpha=False
        )

        image = Image.frombytes(
            "RGB",
            [
                pixmap.width,
                pixmap.height
            ],
            pixmap.samples
        )

        words = ocr_image(image)

        pages.append({
            "page": page_number,
            "width": pixmap.width,
            "height": pixmap.height,
            "words": words,
        })

    pdf.close()

    return pages


def ocr_image_file(image_path):

    image = Image.open(
        image_path
    )

    words = ocr_image(image)

    width, height = image.size

    image.close()

    return [{
        "page": 1,
        "width": width,
        "height": height,
        "words": words,
    }]