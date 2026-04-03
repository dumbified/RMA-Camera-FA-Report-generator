# RMA Camera Failure Analysis Report Generator

A desktop application for generating structured Failure Analysis (FA) reports for RMA camera units. It connects to a MySQL database, auto-fills form fields from screening records, and generates a formatted JIRA-ready report.

---

## Requirements

- Python 3.10 or higher
- MySQL Server 8.0 or higher (running locally or on a network)
- Git

---

## Installation

### Step 1: Clone the Repository

```bash
git clone https://github.com/dumbified/RMA-Camera-FA-Report-generator.git
cd RMA-Camera-FA-Report-generator
```

### Step 2: Install Dependencies

```bash
pip install -r requirements.txt
```

---

## Configuration

Open `fa_report/db_config.py` and update the database credentials to match your MySQL setup:

```python
return FaMySQLConfig(
    host="localhost",       # Your MySQL host
    port=3306,              # Your MySQL port
    user="root",            # Your MySQL username
    password="admin123",    # Your MySQL password
    database="fa_report",   # Your existing database name
    ...
)
```

## Database Setup (Run Once)

From the project root folder, run:

```bash
python -m fa_report.init_db
```

This will create the following tables in your MySQL database:
- `fa_report_sections`
- `fa_report_data`
- `fa_customer_request_templates`
- `fa_camh_base_sn`

> **Note:** The `rma_cam` and `fa_3.1_cam_component` tables are **not** created by this script. They must already exist in your database with the correct schema and data.

---

## Running the Application

From the project root folder, run:

```bash
python -m fa_report.app
```

The GUI will launch. On first use, click **Settings** to set the folder path where your VIT camera images are stored.

---

## Usage Overview

1. Enter a **VIT ID** in the form and press Enter — the app will auto-fill screening results from the database.
2. Fill in the remaining fields (Visual Check, PCBA ATS, FT result, FA result).
3. Click **Generate Report** to preview the formatted FA report.
4. Click **Copy report to clipboard** to paste directly into JIRA.

For batch processing, use the **Batch Queue** panel on the left:
1. Paste multiple VIT IDs (one per line) and click **Add to Queue**.
2. Select each VIT from the list to review and fill in their details.
3. Click **Batch Export (Copy to Clipboard)** to export all reports at once.

---

## Project Structure

```
RMA-Camera-FA-Report-generator/
├── fa_report/
│   ├── app.py                          # Main GUI application
│   ├── bridger.py                      # Image folder search & classification bridge
│   ├── image_classifier.py             # MobileViT v2 model + rule-based fallback
│   ├── db_config.py                    # MySQL connection configuration
│   ├── init_db.py                      # Database initialisation script
│   ├── mobilevitv2.pth                 # Trained model weights
│   ├── report_rules.json               # FA report generation rules
│   └── customer_request_templates.json # Customer request report templates
├── requirements.txt
└── README.md
```
