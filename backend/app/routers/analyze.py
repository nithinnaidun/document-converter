import os

from fastapi import APIRouter, HTTPException

from app.services.document_analyzer import analyze_document


router = APIRouter(
    prefix="/api/analyze",
    tags=["Document Analysis"]
)


TEMP_DIR = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "../../temp"
    )
)


def find_file(file_id: str):

    if not os.path.exists(TEMP_DIR):
        return None

    for filename in os.listdir(TEMP_DIR):

        if filename.startswith(file_id + "."):

            return os.path.join(
                TEMP_DIR,
                filename
            )

    return None


@router.post("/{file_id}")
def analyze(file_id: str):

    file_path = find_file(file_id)

    if not file_path:

        raise HTTPException(
            status_code=404,
            detail="File not found"
        )

    try:

        result = analyze_document(
            file_path
        )

        return {
            "status": "success",
            "file_id": file_id,
            "analysis": result
        }

    except Exception as error:

        raise HTTPException(
            status_code=500,
            detail=str(error)
        )