import os
import shutil
import subprocess


def find_libreoffice():

    possible_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    for path in possible_paths:

        if os.path.exists(path):
            return path

    path = shutil.which(
        "soffice"
    )

    if path:
        return path

    raise RuntimeError(
        "LibreOffice is not installed."
    )


def office_to_pdf(
    input_path,
    output_directory
):

    soffice = find_libreoffice()

    os.makedirs(
        output_directory,
        exist_ok=True
    )

    command = [
        soffice,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        output_directory,
        input_path,
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=180
    )

    if result.returncode != 0:

        raise RuntimeError(
            result.stderr
            or "Office conversion failed"
        )

    filename = os.path.splitext(
        os.path.basename(input_path)
    )[0]

    output_path = os.path.join(
        output_directory,
        filename + ".pdf"
    )

    if not os.path.exists(
        output_path
    ):

        raise RuntimeError(
            "LibreOffice did not create "
            "the PDF."
        )

    return output_path