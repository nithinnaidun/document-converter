from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from pathlib import Path
import asyncio


router = APIRouter(
    prefix="/api",
    tags=["Download"]
)

TEMP_DIR = Path("temp")


@router.get("/download/{file_id}")
async def download_file(file_id: str):

    # Look for the converted file
    matching_files = list(TEMP_DIR.glob(f"{file_id}.*"))

    if not matching_files:
        raise HTTPException(
            status_code=404,
            detail="File not found or already deleted"
        )

    file_path = matching_files[0]

    # Determine filename
    extension = file_path.suffix.lower()

    if extension == ".docx":
        download_name = "converted-document.docx"
    elif extension == ".pdf":
        download_name = "converted-document.pdf"
    else:
        download_name = file_path.name

    async def delete_after_download():
        # Give the browser time to start downloading
        await asyncio.sleep(10)

        try:
            if file_path.exists():
                file_path.unlink()

        except Exception as error:
            print(f"Could not delete file: {error}")

    # Start deletion task
    asyncio.create_task(delete_after_download())

    return FileResponse(
        path=file_path,
        filename=download_name,
        media_type="application/octet-stream"
    )