import os
import re
import zipfile

import pymupdf
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from PIL import Image


DATE_PATTERN = re.compile(
    r"^\d{2}-\d{2}-\d{4}$"
)

AMOUNT_PATTERN = re.compile(
    r"^[\d,]+\.\d{2}$"
)


def clean_text(text):
    """Clean extracted PDF text."""

    if text is None:
        return ""

    text = str(text)

    text = text.replace("\u00a0", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def amount_to_number(value):
    """Convert Indian formatted amount to Excel number."""

    if not value:
        return None

    value = value.strip()

    if not AMOUNT_PATTERN.match(value):
        return None

    try:
        return float(value.replace(",", ""))
    except ValueError:
        return None


def get_page_words(page):
    """
    Get PDF words with coordinates.
    """

    raw_words = page.get_text(
        "words",
        sort=True
    )

    words = []

    for item in raw_words:

        if len(item) < 8:
            continue

        x0, y0, x1, y1, text = item[:5]

        text = clean_text(text)

        if not text:
            continue

        words.append({
            "text": text,
            "x0": float(x0),
            "y0": float(y0),
            "x1": float(x1),
            "y1": float(y1),
            "cx": (float(x0) + float(x1)) / 2,
            "cy": (float(y0) + float(y1)) / 2,
        })

    return words


def group_words_into_lines(words, tolerance=3.5):
    """
    Group PDF words into visual lines based on Y coordinate.
    """

    lines = []

    for word in sorted(
        words,
        key=lambda w: (w["cy"], w["x0"])
    ):

        placed = False

        for line in lines:

            if abs(word["cy"] - line["cy"]) <= tolerance:

                line["words"].append(word)

                # Recalculate line center
                line["cy"] = sum(
                    w["cy"] for w in line["words"]
                ) / len(line["words"])

                placed = True
                break

        if not placed:

            lines.append({
                "cy": word["cy"],
                "words": [word]
            })

    for line in lines:

        line["words"].sort(
            key=lambda w: w["x0"]
        )

        line["text"] = " ".join(
            w["text"] for w in line["words"]
        )

    return sorted(
        lines,
        key=lambda line: line["cy"]
    )


def find_table_header(lines):
    """
    Find the bank statement table header.
    """

    for index, line in enumerate(lines):

        text = line["text"].lower()

        if (
            "date" in text
            and "particulars" in text
            and "deposits" in text
            and "withdrawals" in text
            and "balance" in text
        ):
            return index, line

    return None, None


def get_column_positions(header_line):
    """
    Determine column positions from the actual PDF header.
    """

    positions = {}

    for word in header_line["words"]:

        text = word["text"].lower()

        if text == "date":
            positions["date"] = word["cx"]

        elif text == "particulars":
            positions["particulars"] = word["cx"]

        elif text == "deposits":
            positions["deposits"] = word["cx"]

        elif text == "withdrawals":
            positions["withdrawals"] = word["cx"]

        elif text == "balance":
            positions["balance"] = word["cx"]

    return positions


def detect_amount_columns(line, column_positions):
    """
    Find numeric amounts and assign them to the
    correct financial column using X coordinates.
    """

    deposits = None
    withdrawals = None
    balance = None

    amount_words = []

    for word in line["words"]:

        number = amount_to_number(
            word["text"]
        )

        if number is not None:

            amount_words.append({
                "word": word,
                "value": number
            })

    if not amount_words:
        return deposits, withdrawals, balance

    deposit_x = column_positions.get("deposits")
    withdrawal_x = column_positions.get("withdrawals")
    balance_x = column_positions.get("balance")

    for item in amount_words:

        x = item["word"]["cx"]
        value = item["value"]

        distances = {}

        if deposit_x is not None:
            distances["deposits"] = abs(
                x - deposit_x
            )

        if withdrawal_x is not None:
            distances["withdrawals"] = abs(
                x - withdrawal_x
            )

        if balance_x is not None:
            distances["balance"] = abs(
                x - balance_x
            )

        if not distances:
            continue

        closest = min(
            distances,
            key=distances.get
        )

        # Don't allow a value that is very far
        # from a financial column.
        if distances[closest] > 90:
            continue

        if closest == "deposits":
            deposits = value

        elif closest == "withdrawals":
            withdrawals = value

        elif closest == "balance":
            balance = value

    return deposits, withdrawals, balance


def extract_bank_statement_page(
    page,
    page_number
):
    """
    Extract transactions from one bank statement page.
    """

    words = get_page_words(page)

    lines = group_words_into_lines(words)

    header_index, header_line = find_table_header(
        lines
    )

    if header_line is None:
        return []

    column_positions = get_column_positions(
        header_line
    )

    transactions = []

    current_transaction = None

    for line in lines[header_index + 1:]:

        line_text = clean_text(
            line["text"]
        )

        lower_text = line_text.lower()

        # Ignore repeated page headers
        if (
            "date" in lower_text
            and "particulars" in lower_text
            and "balance" in lower_text
        ):
            continue

        # Ignore page number
        if re.match(
            r"^page\s+\d+$",
            lower_text
        ):
            continue

        # Opening balance
        if lower_text.startswith(
            "opening balance"
        ):

            amounts = [
                amount_to_number(w["text"])
                for w in line["words"]
                if amount_to_number(w["text"]) is not None
            ]

            if amounts:

                transactions.append({
                    "date": "",
                    "particulars": "Opening Balance",
                    "deposits": None,
                    "withdrawals": None,
                    "balance": amounts[-1],
                    "page": page_number
                })

            continue

        # Detect transaction start
        date_word = None

        for word in line["words"]:

            if DATE_PATTERN.match(
                word["text"]
            ):
                date_word = word
                break

        if date_word:

            # Save previous transaction
            if current_transaction:

                transactions.append(
                    current_transaction
                )

            current_transaction = {
                "date": date_word["text"],
                "particulars": "",
                "deposits": None,
                "withdrawals": None,
                "balance": None,
                "page": page_number
            }

        # Ignore anything before first transaction
        if current_transaction is None:
            continue

        # Detect financial amounts
        deposits, withdrawals, balance = (
            detect_amount_columns(
                line,
                column_positions
            )
        )

        if deposits is not None:
            current_transaction[
                "deposits"
            ] = deposits

        if withdrawals is not None:
            current_transaction[
                "withdrawals"
            ] = withdrawals

        if balance is not None:
            current_transaction[
                "balance"
            ] = balance

        # Build particulars
        particular_parts = []

        date_found = False

        for word in line["words"]:

            text = word["text"]

            if (
                date_word
                and text == date_word["text"]
                and not date_found
            ):
                date_found = True
                continue

            # Skip financial amounts
            if amount_to_number(text) is not None:
                continue

            # Ignore page labels
            if text.lower() == "page":
                continue

            # Ignore isolated page number
            if re.match(
                r"^\d+$",
                text
            ):
                # Don't aggressively remove numbers
                # unless this is clearly a page label.
                pass

            # Text from the particulars area
            particular_parts.append(text)

        if particular_parts:

            additional_text = clean_text(
                " ".join(particular_parts)
            )

            # Don't accidentally append the
            # repeated header.
            if additional_text:

                if current_transaction[
                    "particulars"
                ]:

                    current_transaction[
                        "particulars"
                    ] += " " + additional_text

                else:

                    current_transaction[
                        "particulars"
                    ] = additional_text

    # Save final transaction
    if current_transaction:

        transactions.append(
            current_transaction
        )

    return transactions


def clean_transactions(transactions):
    """
    Final cleaning of extracted transactions.
    """

    cleaned = []

    for transaction in transactions:

        particulars = clean_text(
            transaction.get(
                "particulars",
                ""
            )
        )

        # Remove accidental Chq text from
        # beginning/end spacing only.
        particulars = re.sub(
            r"\s+",
            " ",
            particulars
        ).strip()

        transaction[
            "particulars"
        ] = particulars

        cleaned.append(
            transaction
        )

    return cleaned


def format_excel_sheet(worksheet):
    """
    Professional Excel formatting.
    """

    header = worksheet[1]

    for cell in header:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )

    # Date
    for cell in worksheet["A"][1:]:

        cell.alignment = Alignment(
            vertical="top"
        )

    # Particulars
    for cell in worksheet["B"][1:]:

        cell.alignment = Alignment(
            wrap_text=True,
            vertical="top"
        )

    # Amount columns
    for column in ["C", "D", "E"]:

        for cell in worksheet[
            column
        ][1:]:

            cell.number_format = '#,##0.00'

            cell.alignment = Alignment(
                horizontal="right",
                vertical="top"
            )

    # Widths
    widths = {
        "A": 15,
        "B": 75,
        "C": 18,
        "D": 18,
        "E": 18
    }

    for column, width in widths.items():

        worksheet.column_dimensions[
            column
        ].width = width

    worksheet.row_dimensions[1].height = 25


def pdf_to_excel(
    input_path: str,
    output_path: str
):
    """
    Convert a bank statement PDF into
    structured Excel rows.

    This uses PDF coordinates rather than
    treating the document as plain text.
    """

    pdf = pymupdf.open(
        input_path
    )

    all_transactions = []

    for page_number, page in enumerate(
        pdf,
        start=1
    ):

        page_transactions = (
            extract_bank_statement_page(
                page,
                page_number
            )
        )

        all_transactions.extend(
            page_transactions
        )

    pdf.close()

    all_transactions = clean_transactions(
        all_transactions
    )

    # Create workbook
    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Transactions"

    # Header
    worksheet.append([
        "Date",
        "Particulars",
        "Deposits",
        "Withdrawals",
        "Balance"
    ])

    # Data
    for transaction in all_transactions:

        worksheet.append([
            transaction.get("date", ""),
            transaction.get(
                "particulars",
                ""
            ),
            transaction.get(
                "deposits"
            ),
            transaction.get(
                "withdrawals"
            ),
            transaction.get(
                "balance"
            )
        ])

    format_excel_sheet(
        worksheet
    )

    # Summary sheet
    summary = workbook.create_sheet(
        "Summary"
    )

    summary["A1"] = "Document"
    summary["B1"] = os.path.basename(
        input_path
    )

    summary["A2"] = "Transactions"
    summary["B2"] = len(
        all_transactions
    )

    summary["A3"] = "Total Deposits"

    summary["B3"] = sum(
        (
            transaction["deposits"]
            or 0
        )
        for transaction
        in all_transactions
    )

    summary["A4"] = "Total Withdrawals"

    summary["B4"] = sum(
        (
            transaction["withdrawals"]
            or 0
        )
        for transaction
        in all_transactions
    )

    summary["A5"] = "Pages Processed"

    # Calculate pages from transactions
    pages = {
        transaction["page"]
        for transaction
        in all_transactions
        if transaction.get("page")
    }

    summary["B5"] = len(pages)

    for cell in summary["1:1"]:
        cell.font = Font(
            bold=True
        )

    summary.column_dimensions[
        "A"
    ].width = 25

    summary.column_dimensions[
        "B"
    ].width = 30

    summary["B3"].number_format = '#,##0.00'
    summary["B4"].number_format = '#,##0.00'

    workbook.save(
        output_path
    )

    return output_path


def pdf_to_word(
    input_path: str,
    output_path: str
):
    """
    Basic text PDF → Word.
    """

    pdf = pymupdf.open(
        input_path
    )

    document = Document()

    for page_number, page in enumerate(
        pdf
    ):

        text = page.get_text(
            "text"
        )

        if text.strip():

            paragraphs = text.split(
                "\n"
            )

            for paragraph in paragraphs:

                paragraph = paragraph.strip()

                if paragraph:
                    document.add_paragraph(
                        paragraph
                    )

        if page_number < len(pdf) - 1:

            document.add_page_break()

    document.save(
        output_path
    )

    pdf.close()

    return output_path


def pdf_to_jpg(
    input_path: str,
    output_path: str
):
    """
    PDF → JPG.
    Multiple pages are returned as ZIP.
    """

    pdf = pymupdf.open(
        input_path
    )

    output_directory = os.path.dirname(
        output_path
    )

    base_name = os.path.splitext(
        os.path.basename(output_path)
    )[0]

    image_files = []

    for page_number, page in enumerate(
        pdf
    ):

        pixmap = page.get_pixmap(
            matrix=pymupdf.Matrix(2, 2),
            alpha=False
        )

        image_path = os.path.join(
            output_directory,
            f"{base_name}-page-{page_number + 1}.jpg"
        )

        pixmap.save(
            image_path
        )

        image_files.append(
            image_path
        )

    pdf.close()

    if len(image_files) == 1:

        os.replace(
            image_files[0],
            output_path
        )

        return output_path

    zip_path = output_path.replace(
        ".jpg",
        ".zip"
    )

    with zipfile.ZipFile(
        zip_path,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_file:

        for image_path in image_files:

            zip_file.write(
                image_path,
                os.path.basename(
                    image_path
                )
            )

    for image_path in image_files:

        try:
            os.remove(
                image_path
            )
        except OSError:
            pass

    return zip_path


def image_to_pdf(
    input_path: str,
    output_path: str
):
    """
    Image → PDF.
    """

    image = Image.open(
        input_path
    )

    if image.mode != "RGB":

        image = image.convert(
            "RGB"
        )

    image.save(
        output_path,
        "PDF",
        resolution=100.0
    )

    image.close()

    return output_path