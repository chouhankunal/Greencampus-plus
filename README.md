# GreenCampus Plus

A beginner-friendly campus sustainability tracker built with Flask, SQLite, HTML, CSS, and Chart.js. It helps a college record electricity, water, and waste measurements, spot unusually high single-record usage, and compare resource totals across two periods.

## Run on Windows

1. Install Python 3.9 or newer and open this folder in VS Code.
2. Open **Terminal → New Terminal** and run:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install -r requirements.txt
   python app.py
   ```

   If PowerShell blocks activation, either run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` in that terminal and activate again, or use the environment's interpreter explicitly:

   ```powershell
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   .\.venv\Scripts\python.exe app.py
   ```

3. Open the local address printed in the terminal (normally `http://127.0.0.1:5000`).
4. Stop the server with **Ctrl+C**.

The SQLite database is created automatically at `instance/greencampus.db` the first time the app starts. Chart.js and the optional web fonts load from CDNs, so an internet connection is needed for those visual enhancements; the rest of the app runs locally.

Keep `Campus_Sustainability_Dataset_1000 (1).xlsx` beside `app.py`, then use **Import workbook dataset** on the dashboard. The import is repeatable; the first run inserts eligible source rows and later runs skip already imported IDs.

## Features

- Dashboard totals in kWh, litres, and kg; a daily Chart.js trend; recent records; and a simple rule-based sustainability score.
- Record entry for a campus area, resource, quantity, date, and optional note. The server checks the input before saving it.
- Searchable resource history with resource/date filters and confirmation before deletion.
- Idempotent demo-data loading. Sample rows are marked **DEMO** and visibly identified so they are not confused with measurements.
- Alerts for a single record above 1,000 kWh, 10,000 L, or 100 kg. These are simple fixed demonstration thresholds, not campus-specific targets; the alert states the threshold and reason.
- Practical recommendations for resources with recorded consumption.
- A period comparison that reports percentage increases/decreases from saved measurements. It does not convert usage into emissions or claim verified environmental savings.
- GreenCampus AutoPilot can simulate a three-resource batch every 30 seconds for a selected building. Each reading is marked **SIMULATED**, stored separately from the existing **DEMO** flag, and is not real sensor data. Some readings deliberately exceed the existing thresholds to demonstrate alerts.
- Import the bundled `Campus_Sustainability_Dataset_1000 (1).xlsx` with **Import workbook dataset** on the dashboard. Each source row becomes one DATASET electricity record (kWh), one water record (L), and one waste record (kg), retaining its historical date, building, Record_ID, occupancy, renewable energy, estimated CO₂ value, statuses, day type, and suggested action. Dataset readings are explicitly labeled, are not verified real campus measurements, and are never created by AutoPilot.
- Re-importing is safe: a unique dataset Record_ID/resource key prevents duplicate resource entries. The dashboard reports source rows imported/skipped; incomplete or invalid rows are skipped with a few row-level explanations. Resource History can filter by source and reveal the preserved workbook metadata.
- Impact comparisons can be filtered to a single source such as DATASET; the default combined view states that it combines sources.

The score starts at 100 and subtracts an equal share for each record that exceeds its resource threshold. It is an engagement indicator, not a scientific or verified sustainability rating. Resource totals aggregate saved records and can include MEASURED, DEMO, SIMULATED, and DATASET entries unless a source filter is selected.

AutoPilot monitoring state is stored in SQLite, so refreshing the dashboard does not create a second timer or reset the batch schedule. The open dashboard polls the server every 10 seconds; the server creates at most one batch when its 30-second deadline is reached. Leave the dashboard open for ongoing simulation. Stop Monitoring disables future batches.

## Project structure

```text
greencampus-plus/
├── app.py
├── requirements.txt
├── README.md
├── Campus_Sustainability_Dataset_1000 (1).xlsx
├── instance/                 # Created on first run; contains the SQLite database
├── static/
│   └── style.css
└── templates/
    ├── base.html
    ├── dashboard.html
    ├── history.html
    └── impact.html
```

## Demo presentation

Select **Load demo data** on the dashboard to add a few labeled sample readings, including one reading that triggers a water alert. Loading it more than once will not create duplicate demo rows. The records are saved in the same history table as real records, but have their own visible **DEMO** label and a separate dashboard count.
