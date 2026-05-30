# Quote PDF → Excel

Web app that converts Haymans-format supplier PDF quotes into a tidy Excel
file with one block per line item (part number row + description row).

## Use online

Open the deployed Streamlit app, drag a quote PDF onto the uploader, click
**Download Excel file**.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then visit http://localhost:8501 in a browser.

## Command-line use

`quote_to_excel.py` also works standalone:

```bash
python quote_to_excel.py quote-425-348149-425.pdf
```

Produces `quote-425-348149-425.xlsx` next to the input.
