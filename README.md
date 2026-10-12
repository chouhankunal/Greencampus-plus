# 🌱 GreenCampus Plus
### Campus Sustainability Intelligence

**Making campus resource consumption visible, measurable, and actionable.**

GreenCampus Plus is a campus sustainability tracking web application built with **Python Flask, SQLite, HTML, CSS, and Chart.js**. It helps educational institutions monitor electricity, water, and waste consumption, identify unusually high resource usage, and compare consumption across different periods.

The project demonstrates how a simple, data-driven system can help campuses make more informed resource-management decisions.

---

## 🎯 Problem Statement — KAUSHAL 8.0 | PS-04

### Sustainable Tech Without Limits

Educational campuses consume substantial resources every day, including electricity, water, and materials. Without accessible consumption records and clear comparisons, excessive usage can go unnoticed, making it difficult to identify potential waste and prioritize corrective action.

GreenCampus Plus addresses this challenge by providing a centralized dashboard for recording resource consumption, detecting unusually high readings, visualizing historical trends, and presenting practical sustainability recommendations.

**Our goal:** Transform resource-consumption data into actionable insights that encourage responsible resource management.

## ✨ Key Features

### 📊 Sustainability Dashboard
- View electricity consumption in kWh, water consumption in litres, and waste generation in kilograms.
- Visualize daily resource trends with interactive charts.
- Review recent records and a rule-based sustainability engagement score.
- Distinguish demonstration data from other recorded sources.

### ⚡ Resource Monitoring
- Record resource usage by campus building or area.
- Select a resource type, enter a quantity and date, and add optional notes.
- Validate input before saving records.
- Identify unusually high individual readings using configurable demonstration thresholds.

### 🚨 Waste Detection and Recommendations
- Flag individual readings that exceed predefined thresholds.
- Explain why each alert was triggered.
- Provide practical recommendations for improving resource management.
- Compare resource consumption across two periods.

### 📁 Historical Dataset Import
- Import records from the bundled Excel workbook.
- Preserve supported historical metadata, including building, date, occupancy, renewable energy, estimated CO₂ values, and suggested actions.
- Track imported source records and skip duplicate imports using unique source-record/resource keys.
- Filter history by data source.

### 🔄 GreenCampus AutoPilot
- Simulate periodic electricity, water, and waste readings for a selected building.
- Generate a three-resource batch every 30 seconds while monitoring is active.
- Persist monitoring state in SQLite to prevent duplicate timers after dashboard refreshes.
- Mark simulated readings separately from demonstration, dataset, and measured records.

**Note:** AutoPilot generates simulated readings; it does not connect to physical meters or IoT sensors.

---

## 🛠️ Technology Stack

| Technology | Purpose |
|---|---|
| Python | Application logic and data processing |
| Flask | Web framework and HTTP routes |
| SQLite | Persistent resource records and monitoring state |
| HTML5 | Dashboard structure |
| CSS3 | Responsive interface styling |
| JavaScript | Dynamic interactions and dashboard updates |
| Chart.js | Resource-consumption visualizations |
| Excel workbook | Historical demonstration dataset |

Chart.js and optional web fonts are loaded from external CDNs. An internet connection is required for those enhancements.

## 🏗️ Project Structure

```text
GreenCampusPlus/
├── app.py
├── requirements.txt
├── README.md
├── Campus_Sustainability_Dataset_1000 (1).xlsx
├── instance/
│   └── greencampus.db
├── static/
│   └── style.css
└── templates/
    ├── base.html
    ├── dashboard.html
    ├── history.html
    └── impact.html
```

The `instance/` directory and SQLite database are created automatically when the application initializes. They are runtime files rather than source-code files.

## 🚀 Getting Started

### Prerequisites

- Python 3.9 or newer, subject to compatibility with the installed dependencies.
- Git, if you intend to clone the repository.
- A modern web browser.
- An internet connection for CDN-hosted charts and optional fonts.

### 1. Clone the repository

Replace the placeholder with your GitHub repository URL.

```bash
git clone <YOUR_REPOSITORY_URL>
cd GreenCampusPlus
```

If you already have the project locally, open its folder in VS Code instead.

### 2. Create a virtual environment

**Windows PowerShell:**

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, run the following in the same terminal and activate again:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

Alternatively, use the environment's Python executable directly without activating it.

### 3. Install dependencies

```powershell
python -m pip install -r requirements.txt
```

### 4. Start the application

```powershell
python app.py
```

### 5. Open GreenCampus Plus

Visit:

```text
http://127.0.0.1:5000
```

Stop the development server with `Ctrl + C`.

---

## 📂 Importing the Historical Dataset

Keep the supplied workbook in the project root:

`Campus_Sustainability_Dataset_1000 (1).xlsx`

Then:

1. Start GreenCampus Plus.
2. Open the dashboard.
3. Select **Import workbook dataset**.
4. Review the import summary and skipped-row explanations.
5. Open Resource History to inspect imported records and their metadata.

Each eligible source row produces separate electricity, water, and waste records, subject to validation. The import process is designed to avoid duplicate entries when the same source rows are imported again.

**Data integrity:** Imported workbook values are explicitly identified as dataset records. They should not be interpreted as verified measurements from an actual campus.

## 🎬 Hackathon Demo Workflow

Use this sequence to demonstrate the application during evaluation:

1. **Load demo data** to populate the dashboard with labeled sample readings.
2. **Explore resource totals** and daily consumption trends.
3. **Trigger a resource alert** using the demonstration reading that exceeds the configured water threshold.
4. **Review recommendations** and explain how they relate to the recorded resource usage.
5. **Compare two periods** to inspect increases or decreases in consumption.
6. **Demonstrate AutoPilot** to show how periodic simulated readings appear in the dashboard and history.
7. **Inspect source labels** to distinguish DEMO, SIMULATED, DATASET, and MEASURED records.

Loading demonstration data repeatedly should not create duplicate demo records.

---

## 📐 Detection Rules and Calculation Limitations

GreenCampus Plus uses straightforward, rule-based logic to keep its results explainable.

| Resource | Demonstration alert threshold |
|---|---:|
| Electricity | Above 1,000 kWh per record |
| Water | Above 10,000 litres per record |
| Waste | Above 100 kg per record |

These thresholds are fixed demonstration values, not validated campus-specific targets. An alert indicates that a reading exceeds a configured threshold; it does not independently prove that waste occurred.

### Sustainability score

The dashboard starts with a score of 100 and subtracts an equal share for each record that exceeds its resource threshold. This score is an engagement indicator, not a scientific sustainability rating.

### Period comparison

The application compares saved consumption totals across selected periods and reports percentage changes. A percentage increase or decrease describes the recorded data; it does not establish the cause of that change.

### Environmental impact

Where imported metadata includes estimated CO₂ values, those values retain their source context. The application does not claim that its usage comparisons represent verified emissions reductions or actual financial savings.

### Data sources

Resource totals may combine MEASURED, DEMO, SIMULATED, and DATASET records. Use source filters where available to compare like with like.

---

## 🔐 Data and Security Notes

- Keep secrets, credentials, and environment-specific configuration out of version control.
- Do not commit a local database containing private or sensitive information.
- Use the Flask development server only for local development and demonstrations.
- Validate imported data and user inputs before relying on them for decisions.
- Treat simulated and sample readings as demonstration data rather than operational measurements.

## 🔮 Future Enhancements

Potential next steps include:

- Integration with verified smart-meter and IoT sensor data.
- Campus-specific resource baselines and configurable alert thresholds.
- Building-level dashboards and role-based access.
- Validated emission factors and documented environmental calculations.
- CSV and PDF reporting.
- Action tracking to evaluate whether recommended interventions reduce subsequent consumption.

These are proposed enhancements unless already implemented.

## 👥 Team and Contributions

Developed collaboratively for **KAUSHAL 8.0 — PS-04: Sustainable Tech Without Limits**.

Add team members, GitHub profiles, and individual contributions here.

## 📄 License

Specify the project's license before distributing the code publicly. If the repository does not yet have a license, its reuse permissions should not be assumed.

---

**GreenCampus Plus — Measure resource use. Identify opportunities. Support better decisions.**
