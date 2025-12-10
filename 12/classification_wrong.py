import os
import json
from pathlib import Path

import xlwings as xw
from openai import OpenAI


# --- API key loading (as requested) ---
def load_api_key(path="openai_key.txt"):
    base = Path(__file__).parent
    abs_path = (base / path).expanduser().resolve()
    with abs_path.open("r", encoding="utf-8") as f:
        return f.read().strip()


os.environ["OPENAI_API_KEY"] = load_api_key()
client = OpenAI()


# --- OpenAI helper ---
def classify_row_with_openai(row_data: dict) -> str:
    """
    Call OpenAI to classify a single trip row into:
    'highly recommended', 'recommended', or 'not recommended'.
    """
    system_prompt = (
        "You are a travel expert. "
        "Classify each trip as 'highly recommended', 'recommended', or 'not recommended'. "
        "Use the review text and all other fields (satisfaction score, sentiment, spend, etc.). "
        "Return ONLY one of these exact strings."
    )

    user_content = (
        "Classify the following trip:\n\n"
        f"{json.dumps(row_data, ensure_ascii=False, indent=2)}\n\n"
        "Answer with exactly one of:\n"
        "- highly recommended\n"
        "- recommended\n"
        "- not recommended"
    )

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content},
        ],
        temperature=0.0,
    )

    category = response.choices[0].message.content.strip().lower()

    # Basic normalization / safety
    valid = {"highly recommended", "recommended", "not recommended"}
    if category not in valid:
        # Try to map roughly
        if "highly" in category:
            return "highly recommended"
        if "not" in category:
            return "not recommended"
        return "recommended"

    return category


# --- Main xlwings entry point ---
def main():
    """
    Read data from 'Travels' sheet, classify each row with OpenAI,
    and write the result into the 'Recommendation Category' column.
    """

    try:
        # When called from Excel via RunPython
        wb = xw.Book.caller()
    except Exception:
        # For testing outside Excel: open the workbook by filename
        wb = xw.Book("travels.xlsm")

    sheet = wb.sheets["Travels"]

    # Get the used range starting at A1
    data_range = sheet.range("A1").current_region
    values = data_range.value

    if not values or len(values) < 2:
        return  # no data

    headers = values[0]
    rows = values[1:]

    # Find index of the Recommendation Category column
    try:
        rec_col_idx = headers.index("Recommendation Category")
    except ValueError:
        raise RuntimeError("Column 'Recommendation Category' not found in header row.")

    # Loop over each data row
    for i, row in enumerate(rows):
        if row is None:
            continue

        # Ensure row is list-like & same length as headers
        if row is None or all(v is None for v in (row if isinstance(row, list) else [row])):
            continue

        # Normalize to list
        if not isinstance(row, (list, tuple)):
            row = [row]

        # Pad row to length of headers
        if len(row) < len(headers):
            row = list(row) + [None] * (len(headers) - len(row))

        row_dict = {h: row[idx] for idx, h in enumerate(headers)}

        # Classify using OpenAI
        category = classify_row_with_openai(row_dict)

        # Write classification back to the appropriate cell
        excel_row = data_range.row + 1 + i  # +1 because headers are in first row
        excel_col = data_range.column + rec_col_idx
        sheet.range(excel_row, excel_col).value = category

    # Optional: save workbook (comment out if you prefer manual saving)
    # wb.save()


if __name__ == "__main__":
    main()
