# PL Payment & Date Automation

This Streamlit app automates the process of updating payment information from an updated payments Excel file to a target Excel file based on account numbers and dates.

## Requirements

- Python 3.x
- pandas
- openpyxl
- streamlit

Install dependencies:
```
pip install -r requirements.txt
```

## Usage

1. Run the app:
   ```
   streamlit run PLPayment&DateAuto.py
   ```

2. Upload the source Excel file containing updated payments and the tracker Excel file you want to fill.

3. Select the source columns for Account Number, Amount, and Payment Date.

4. Select the target worksheet and the target account number column.

6. If either workbook is password-protected, enter the password for that workbook.

7. Click the button and download the tracker workbook with weekly payment amounts and dates filled.

## Assumptions

- Source file has columns for Account Number, Amount, and Payment Date.
- Target file contains Account Number in the selected target column (default is the first column).
- The app auto-creates or reuses week columns for the detected month using real calendar weeks:
  - Week 1: days 1-7
  - Week 2: days 8-14
  - Week 3: days 15-21
  - Week 4: days 22-28
  - Week 5: days 29-end of month
- The target workbook is updated with "Payment Amount" and "Payment Date" columns for each week, plus a "Total Payment" column.

## Notes

- If an account is not found, a warning will be shown.
- Invalid date formats will show an error.
- The app processes the files in memory and provides a download link for the updated file.