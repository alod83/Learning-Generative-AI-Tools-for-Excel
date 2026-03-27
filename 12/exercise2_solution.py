import os
import xlwings as xw
from openai import OpenAI

from pathlib import Path

def load_api_key(path="openai_key.txt"):
    base = Path(__file__).parent
    abs_path = (base / path).expanduser().resolve()
    with abs_path.open("r", encoding="utf-8") as f:
        return f.read().strip()

os.environ["OPENAI_API_KEY"] = load_api_key()

client = OpenAI()

def main():
    
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet[f"F1"].value = "Sentiment"
 
    row = 2
    while True:
        review = sheet[f"E{row}"].value
        if review is None:
            break

        prompt = (
            "Classify the sentiment of this travel review as "
            "'positive', 'negative', or 'mixed'. "
            "Return only the single word label.\n\n"
            f"Review: {review}"
        )

        response = client.responses.create(
            model="gpt-4.1-mini",
            input=prompt
        )

        label = response.output[0].content[0].text.strip()
        sheet[f"F{row}"].value = label

        row += 1

# Optional: allow testing from Python directly (outside Excel)
if __name__ == "__main__":
    
    xw.Book("films.xlsx").set_mock_caller()
    main()
