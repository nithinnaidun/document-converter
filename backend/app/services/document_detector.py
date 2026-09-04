import os


FILE_TYPES = {
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
    ".csv": "csv",
}


def detect_document_type(file_path: str):

    extension = os.path.splitext(
        file_path
    )[1].lower()

    document_type = FILE_TYPES.get(
        extension
    )

    if not document_type:
        raise ValueError(
            f"Unsupported file type: {extension}"
        )

    return {
        "extension": extension,
        "type": document_type,
        "filename": os.path.basename(file_path),
    }