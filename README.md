# Sky Report Generator

This Streamlit app creates one combined Sky LIVE working report from an approved audit export. UK, NI and ROI visits are included in the same workbook.

## Inputs

1. The approved Sky `audits_basic_data_export...csv`.
2. The most recent Sky LIVE workbook.
3. The Sky Sites export containing `internal_id` and `city`.
4. The latest UK and Ireland Sky account reference files used for Region, Territory, Record ID, Sky Reference Number, Premise ID and Pot.

The account-reference uploader accepts multiple CSV/XLSX/XLSM files. The app searches every worksheet for the required fields and combines matching Site ID, Account ID and postcode data. City is taken from the separate Sky Sites export.

## Outputs

- `Sky (DD.MM, ... ) LIVE.xlsx`: one combined working report containing raw data, formulas and updated history.

The app blocks normal generation if City, Region, Territory or Pot is missing for any audit and provides a diagnostic CSV. A clearly labelled override is available only for investigation.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```
