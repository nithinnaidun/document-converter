import os
import subprocess


def powerpoint_to_pdf(input_path: str, output_path: str):

    output_directory = os.path.dirname(output_path)

    command = [
        "soffice",
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
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(
            result.stderr or
            "PowerPoint to PDF conversion failed"
        )

    generated_file = os.path.join(
        output_directory,
        os.path.splitext(
            os.path.basename(input_path)
        )[0] + ".pdf"
    )

    if not os.path.exists(generated_file):
        raise RuntimeError(
            "LibreOffice did not create the PDF"
        )

    os.replace(
        generated_file,
        output_path
    )

    return output_path