# Messy Data Cleaning Pipeline: Dedupes, Standardizes, Validates

A Python pipeline that takes a messy customer list and turns it into a clean,
standardized dataset, with a **full audit trail** of every change it made.

Includes a generator that creates realistic (fully synthetic) messy data, so you
can run the whole thing in under a minute.

## What it does

| Field | Before | After |
|---|---|---|
| Name | `SMITH, JOHN` / `  mary   o'brien ` / `Jon Smith3` | `John Smith` / `Mary O'Brien` / `Jon Smith` |
| Email | `JANE@GMIAL.COM` / `jane.doe.gmail.com` | `jane@gmail.com` / *(removed: invalid)* |
| Phone | `416.555.0187` / `+1 416 555 0187 ext 22` / `555-0187` | `(416) 555-0187` / `(416) 555-0187` / *(removed: invalid)* |
| Province/State | `Ontario` / `ont.` / `on` | `ON` |
| Date | `Mar 05, 2024` / `05-Mar-2024` / `03/05/2024` | `2024-03-05` |
| Amount | `$1,234.50` / `CAD 120` / `-50` | `1234.50` / `120.00` / *(removed: negative)* |
| Placeholders | `N/A`, `null`, `-` | blank |
| Duplicates | same customer entered 2-3 times | one record, blanks filled from the others |

## Outputs

Everything is written to `output/`:

- `clean_customers.csv`: the cleaned dataset (with `source_row` for traceability)
- `change_log.csv`: every edit: source row, column, old value, new value, why
- `rejected_rows.csv`: rows with no valid email *and* no valid phone, with the reason
- `cleaning_report.xlsx`: Summary, Clean Data, Change Log and Rejected Rows in one workbook

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python generate_messy_data.py     # creates data/messy_customers.csv (500 rows)
python clean_data.py              # cleans it
```

Use your own file: `python clean_data.py --input my_file.csv --output-dir results`.
It needs these columns: `customer_id, full_name, email, phone, city, province_state, signup_date, total_spend`.

## How duplicates are handled

Rows are linked when they share a **customer ID**, an **email**, or a **name + phone**.
For each group, the most complete row is kept and any blanks in it are filled from the
others. If two rows share an ID or email but have *different names* (for example, family
members sharing an email), they are **not** merged. They are flagged for review in the
change log instead.

## Assumptions and limitations

- Slashed dates are read as `MM/DD/YYYY`. Change `DATE_FORMATS` in `clean_data.py` if your data uses `DD/MM/YYYY`.
- Phone validation follows North American rules (10 digits).
- A duplicate can only be found if it shares an ID, a valid email, or a name + valid phone with another row. If everything linking two rows is invalid or blank, they stay separate.
- Rows are never silently dropped: duplicates and rejections are all listed in the change log.

## Tests

```bash
python test_clean_data.py
```

Covers every cleaner, the duplicate-safety rule, and a full end-to-end run (including
checking that cleaning the clean output a second time changes nothing).

## Tech

Python, pandas, openpyxl, regular expressions.

## Note

Built as a personal sample project. All data is synthetic: no real people, emails or
phone numbers (phone numbers use the reserved fictional 555-01xx range).
