# SMC 4W Compiler & QC — web app

A small Streamlit app for the Gaza Site Management Cluster. The IM team uploads
the monthly 4W partner submissions; the app validates and compiles them and
returns the Power BI workbook, the cleaned `fact_activities.csv`, and a
readable error report.

## What it does

1. Upload the `YYYYMM_Activities_PARTNER.xlsx` files (one or many).
2. The app runs `compile_4w_v3.py` against the bundled reference data.
3. You get three downloads — the Power BI workbook (`SMC_4W_<period>_PowerBI.xlsx`),
   `fact_activities.csv`, and a ZIP of all 13 star-schema tables — plus an
   on-screen error report split into **Must fix** (errors), **Flag to partners**
   (warnings) and **Auto-corrected** (info), downloadable as an Excel file.

## Files

```
app.py                        the Streamlit app
compile_4w_v3.py              the compilation + validation pipeline
build_powerbi_workbook.py     packages the CSVs into the Power BI workbook
requirements.txt              Python dependencies
data/
  site_masterlist_*.csv       site population + site→agency baseline
  list_indicators_5w.csv      canonical Activity Index (25 indicators)
  COD_Gaza.xlsx               Common Operational Dataset (adm4 pcode fallback)
```

The three files in `data/` are the reference data. They are bundled so the team
only ever uploads partner 4Ws. When the masterlist or COD is updated, either
replace the file in `data/` (and redeploy) or use the **Reference data**
override uploaders in the sidebar for a one-off run.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy (Streamlit Community Cloud → link from smcopt.org)

1. Push this folder to a GitHub repo.
2. Go to https://share.streamlit.io → New app → point it at the repo and `app.py`.
3. Share the resulting URL, or link/embed it from smcopt.org.

> **Privacy note.** Gaza is an active conflict context and the 4W data is
> site-level. For anything beyond public summaries, deploy to a **private**
> Streamlit workspace or run it on an internal host rather than a public URL.

## Notes on the method

- Partner and month are read from the file **contents** (Organization Name
  column; in-file reporting month), not the filename — several partner files
  are consolidated multi-partner or multi-month containers.
- Cluster reach uses the agreed **Reading B** method (de-duplicated across
  overlapping activities and across partners).
- Neighborhood pcodes come from the masterlist, with a governorate-scoped COD
  name lookup filling rows whose site could not be resolved.
- The **Months to include** control defaults to March + April. Add other
  months there when you start a new cycle; rows outside the selected months are
  dropped.

Full methodology is in the companion `SMC_4W_Methodology_Note.html`.
