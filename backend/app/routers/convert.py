import os
import uuid

from fastapi import APIRouter, HTTPException

from app.services.pdf_service import (
    pdf_to_word,
    pdf_to_excel,
    pdf_to_jpg,
    image_to_pdf,
)

from app.services.word_service import word_to_pdf
from app.services.excel_service import excel_to_pdf
from app.services.powerpoint_service import powerpoint_to_pdf


router = APIRouter(
    prefix="/api/convert",
    tags=["Conversion"]
)

TEMP_DIR = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "../../temp"
    )
)


def get_input_file(file_id: str):
    if not os.path.exists(TEMP_DIR):
        return None

    for filename in os.listdir(TEMP_DIR):

        if filename.startswith(file_id + "."):
            return os.path.join(
                TEMP_DIR,
                filename
            )

    return None


def create_output(extension: str):
    output_id = str(uuid.uuid4())

    return (
        output_id,
        os.path.join(
            TEMP_DIR,
            output_id + extension
        )
    )


@router.post("/pdf-to-word/{file_id}")
def convert_pdf_to_word(file_id: str):

    input_path = get_input_file(file_id)

    if not input_path:
        raise HTTPException(
            status_code=404,
            detail="Uploaded file not found"
        )

    output_id, output_path = create_output(".docx")

    pdf_to_word(
        input_path,
        output_path
    )

    return {
        "status": "success",
        "file_id": output_id,
        "filename": "converted-document.docx",
        "format": "docx"
    }


@router.post("/pdf-to-excel/{file_id}")
def convert_pdf_to_excel(file_id: str):

    input_path = get_input_file(file_id)

    if not input_path:
        raise HTTPException(
            status_code=404,
            detail="Uploaded file not found"
        )

    output_id, output_path = create_output(".xlsx")

    pdf_to_excel(
        input_path,
        output_path
    )

    return {
        "status": "success",
        "file_id": output_id,
        "filename": "converted-document.xlsx",
        "format": "xlsx"
    }


@router.post("/pdf-to-jpg/{file_id}")
def convert_pdf_to_jpg(file_id: str):

    input_path = get_input_file(file_id)

    if not input_path:
        raise HTTPException(
            status_code=404,
            detail="Uploaded file not found"
        )

    output_id, output_path = create_output(".jpg")

    result = pdf_to_jpg(
        input_path,
        output_path
    )

    result_extension = os.path.splitext(result)[1]

    return {
        "status": "success",
        "file_id": output_id,
        "filename": os.path.basename(result),
        "format": result_extension.replace(".", "")
    }


@router.post("/word-to-pdf/{file_id}")
def convert_word_to_pdf(file_id: str):

    input_path = get_input_file(file_id)

    if not input_path:
        raise HTTPException(
            status_code=404,
            detail="Word file not found"
        )

    output_id, output_path = create_output(".pdf")

    word_to_pdf(
        input_path,
        output_path
    )

    return {
        "status": "success",
        "file_id": output_id,
        "filename": "converted-document.pdf",
        "format": "pdf"
    }


@router.post("/excel-to-pdf/{file_id}")
def convert_excel_to_pdf(file_id: str):

    input_path = get_input_file(file_id)

    if not input_path:
        raise HTTPException(
            status_code=404,
            detail="Excel file not found"
        )

    output_id, output_path = create_output(".pdf")

    excel_to_pdf(
        input_path,
        output_path
    )

    return {
        "status": "success",
        "file_id": output_id,
        "filename": "converted-document.pdf",
        "format": "pdf"
    }


@router.post("/ppt-to-pdf/{file_id}")
def convert_ppt_to_pdf(file_id: str):

    input_path = get_input_file(file_id)

    if not input_path:
        raise HTTPException(
            status_code=404,
            detail="PowerPoint file not found"
        )

    output_id, output_path = create_output(".pdf")

    powerpoint_to_pdf(
        input_path,
        output_path
    )

    return {
        "status": "success",
        "file_id": output_id,
        "filename": "converted-document.pdf",
        "format": "pdf"
    }


@router.post("/image-to-pdf/{file_id}")
def convert_image_to_pdf(file_id: str):

    input_path = get_input_file(file_id)

    if not input_path:
        raise HTTPException(
            status_code=404,
            detail="Image file not found"
        )

    output_id, output_path = create_output(".pdf")

    image_to_pdf(
        input_path,
        output_path
    )

    return {
        "status": "success",
        "file_id": output_id,
        "filename": "converted-document.pdf",
        "format": "pdf"
    }