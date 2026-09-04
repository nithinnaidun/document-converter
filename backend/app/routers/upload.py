from fastapi import APIRouter, UploadFile, File, HTTPException
from pathlib import Path
import uuid
import shutil

router = APIRouter(
    prefix="/api",
    tags=["Upload"]
)

# Temporary upload directory
TEMP_DIR = Path("temp")
TEMP_DIR.mkdir(exist_ok=True)

# Maximum file size: 25 MB
MAX_FILE_SIZE = 25 * 1024 * 1024

# Supported file extensions
ALLOWED_EXTENSIONS = {
    ".pdf",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".ppt",
    ".pptx",
    ".jpg",
    ".jpeg",
    ".png",
    ".txt"
}


@router.post("/upload")
async def upload_file(file: UploadFile = File(...)):

    # Check filename
    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No file selected"
        )

    # Get extension
    extension = Path(file.filename).suffix.lower()

    # Check extension
    if extension not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=f"File type {extension} is not supported"
        )

    # Generate random filename
    file_id = str(uuid.uuid4())

    temp_filename = f"{file_id}{extension}"
    temp_path = TEMP_DIR / temp_filename

    # Save file temporarily
    try:
        with open(temp_path, "wb") as buffer:

            total_size = 0

            while True:
                chunk = await file.read(1024 * 1024)

                if not chunk:
                    break

                total_size += len(chunk)

                # Check file size
                if total_size > MAX_FILE_SIZE:
                    buffer.close()

                    if temp_path.exists():
                        temp_path.unlink()

                    raise HTTPException(
                        status_code=413,
                        detail="File size cannot exceed 25 MB"
                    )

                buffer.write(chunk)

    except HTTPException:
        raise

    except Exception as error:

        if temp_path.exists():
            temp_path.unlink()

        raise HTTPException(
            status_code=500,
            detail=f"Upload failed: {str(error)}"
        )

    return {
        "status": "success",
        "message": "File uploaded successfully",
        "file_id": file_id,
        "original_filename": file.filename,
        "extension": extension,
        "size": total_size
    }