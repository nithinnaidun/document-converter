import os

import pymupdf

from app.services.ocr_service import ocr_pdf


SUPPORTED_EXTENSIONS = {
    ".pdf": "pdf",
    ".doc": "word",
    ".docx": "word",
    ".xls": "excel",
    ".xlsx": "excel",
    ".ppt": "powerpoint",
    ".pptx": "powerpoint",
    ".jpg": "image",
    ".jpeg": "image",
    ".png": "image",
    ".txt": "text",
}


def get_extension(file_path: str):
    return os.path.splitext(file_path)[1].lower()


def detect_pdf_type(file_path: str):
    """
    Determine whether a PDF contains selectable text
    or is primarily scanned/image-based.
    """

    pdf = pymupdf.open(file_path)

    total_text = ""
    pages_with_text = 0

    for page in pdf:

        text = page.get_text("text").strip()

        if text:
            pages_with_text += 1
            total_text += text

    page_count = len(pdf)

    pdf.close()

    if page_count == 0:
        return {
            "is_scanned": False,
            "pages": 0,
            "text_pages": 0,
            "reason": "Empty PDF"
        }

    text_ratio = pages_with_text / page_count

    # If most pages contain meaningful text,
    # treat it as a text PDF.
    is_scanned = text_ratio < 0.5 or len(total_text.strip()) < 20

    return {
        "is_scanned": is_scanned,
        "pages": page_count,
        "text_pages": pages_with_text,
        "text_ratio": round(text_ratio, 3)
    }


def extract_text_pdf(file_path: str):
    """
    Extract selectable PDF text and coordinates.
    """

    pdf = pymupdf.open(file_path)

    pages = []

    for page_number, page in enumerate(pdf):

        text = page.get_text("text")

        words = page.get_text("words")

        page_words = []

        for word in words:

            x0, y0, x1, y1, text_value = word[:5]

            page_words.append({
                "text": text_value,
                "x": float(x0),
                "y": float(y0),
                "width": float(x1 - x0),
                "height": float(y1 - y0),
                "confidence": 100
            })

        pages.append({
            "page": page_number + 1,
            "text": text,
            "words": page_words
        })

    pdf.close()

    return pages


def analyze_pdf(file_path: str):

    pdf_info = detect_pdf_type(file_path)

    if pdf_info["is_scanned"]:

        pages = ocr_pdf(file_path)
        ocr_used = True

    else:

        pages = extract_text_pdf(file_path)
        ocr_used = False

    all_words = []

    for page in pages:
        all_words.extend(page.get("words", []))

    confidence_values = [
        word["confidence"]
        for word in all_words
        if word.get("confidence") is not None
    ]

    if confidence_values:
        average_confidence = sum(
            confidence_values
        ) / len(confidence_values)
    else:
        average_confidence = 0

    return {
        "file_type": "pdf",
        "pages": pdf_info["pages"],
        "is_scanned": pdf_info["is_scanned"],
        "ocr_used": ocr_used,
        "text_pages": pdf_info["text_pages"],
        "text_ratio": pdf_info["text_ratio"],
        "average_confidence": round(
            average_confidence,
            2
        ),
        "pages_data": pages
    }


def analyze_document(file_path: str):

    extension = get_extension(file_path)

    if extension not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported document type: {extension}"
        )

    document_type = SUPPORTED_EXTENSIONS[
        extension
    ]

    if document_type == "pdf":

        return analyze_pdf(file_path)

    return {
        "file_type": document_type,
        "extension": extension,
        "message": (
            "Detailed analyzer for this format "
            "will be added in the next stage."
        )
    }