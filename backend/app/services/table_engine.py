import re


DATE_PATTERNS = [
    r"^\d{2}-\d{2}-\d{4}$",
    r"^\d{2}-\d{2}-\d{2}$",
    r"^\d{2}/\d{2}/\d{4}$",
    r"^\d{2}/\d{2}/\d{2}$",
    r"^\d{2}-[A-Z]{3}-\d{2}$",
    r"^\d{2}-[A-Z]{3}-\d{4}$",
]


def is_date(value):

    value = value.strip()

    for pattern in DATE_PATTERNS:

        if re.match(
            pattern,
            value,
            re.IGNORECASE
        ):
            return True

    return False


def is_amount(value):

    value = value.strip()

    value = value.replace(
        ",",
        ""
    )

    value = value.replace(
        "₹",
        ""
    )

    value = value.replace(
        "$",
        ""
    )

    value = value.replace(
        "€",
        ""
    )

    value = value.replace(
        "£",
        ""
    )

    return bool(
        re.match(
            r"^-?\d+(\.\d+)?$",
            value
        )
    )


def group_words_into_lines(
    words,
    tolerance=5
):

    lines = []

    for word in sorted(
        words,
        key=lambda x: (
            x["y"],
            x["x"]
        )
    ):

        added = False

        for line in lines:

            if abs(
                word["y"] - line["y"]
            ) <= tolerance:

                line["words"].append(
                    word
                )

                added = True
                break

        if not added:

            lines.append({
                "y": word["y"],
                "words": [word]
            })

    for line in lines:

        line["words"].sort(
            key=lambda x: x["x"]
        )

        line["text"] = " ".join(
            word["text"]
            for word in line["words"]
        )

    return lines


def detect_header(lines):

    keywords = [
        "date",
        "description",
        "particular",
        "amount",
        "balance",
        "deposit",
        "withdraw",
        "debit",
        "credit",
        "value date",
        "ref",
    ]

    best_line = None
    best_score = 0

    for line in lines:

        text = line["text"].lower()

        score = 0

        for keyword in keywords:

            if keyword in text:
                score += 1

        if score > best_score:

            best_score = score
            best_line = line

    if best_score >= 2:
        return best_line

    return None


def create_generic_table(
    pages
):

    result = []

    for page in pages:

        lines = group_words_into_lines(
            page["words"]
        )

        header = detect_header(
            lines
        )

        if not header:
            continue

        header_words = header["words"]

        columns = []

        for word in header_words:

            columns.append({
                "name": word["text"],
                "x": word["x"]
            })

        for line in lines:

            if line is header:
                continue

            cells = [
                ""
                for _ in columns
            ]

            for word in line["words"]:

                distances = [
                    abs(
                        word["x"] -
                        column["x"]
                    )
                    for column in columns
                ]

                if not distances:
                    continue

                index = distances.index(
                    min(distances)
                )

                if cells[index]:

                    cells[index] += (
                        " " +
                        word["text"]
                    )

                else:

                    cells[index] = (
                        word["text"]
                    )

            if any(
                cell.strip()
                for cell in cells
            ):

                result.append(
                    cells
                )

    return result