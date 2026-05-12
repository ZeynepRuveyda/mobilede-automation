# Mobile.de Automation Suite

A Streamlit application that automates the weekly KPI template generation workflow for Mobile.de exports.

## What it does

Upload your weekly detail export files + last week's template → get a ready-to-deliver `Template_<date>.xlsx` in one click.

### Pipeline Steps

| Step | Script | Input | Output |
|------|--------|-------|--------|
| 1 | `merge_original_exports.py` | Raw detail `.xlsx` exports | `merged_original_export_<date>.xlsx` |
| 2 | `reshape_merged_for_kpi_order.py` | Merged original + Template | `merged_exports_<date>.xlsx` |
| 3 | `build_from_exports_generic.py` | Reshaped export + Template | `deliver/Template_<date>.xlsx` |
| 4 *(optional)* | `update_template_from_crawl_export.py` | Crawl export + Template | Updated `Template_<date>.xlsx` |

## Project Structure

```
mobilede_automation/
├── app.py                          # Streamlit entry point (Dashboard)
├── pages/
│   ├── 1_pipeline.py               # Run Pipeline page
│   └── 2_preview.py                # Preview / inspect data
├── scripts/
│   ├── merge_original_exports.py
│   ├── reshape_merged_for_kpi_order.py
│   ├── build_from_exports_generic.py
│   └── update_template_from_crawl_export.py
├── requirements.txt
└── README.md
```

## Installation

```bash
git clone https://github.com/YOUR_USERNAME/mobilede-automation.git
cd mobilede-automation
pip install -r requirements.txt
streamlit run app.py
```

## Usage

1. Open the app in your browser (`http://localhost:8501`)
2. Navigate to **Run Pipeline**
3. Enter the **week date** in `DD.MM.YYYY` format (e.g. `28.04.2026`)
4. Upload your **detail export files** (one or more `.xlsx`)
5. Upload the **previous week's Template** file
6. *(Optional)* Upload a **Crawl Export** file
7. Click **Run Pipeline**
8. Download the generated `Template_<date>.xlsx`

## Template Column Structure

| Col | Field | Description |
|-----|-------|-------------|
| 1 | Year | Updated to current week's year |
| 2 | Month | Updated to current week's month |
| 3 | Day | Updated to current week's day |
| 5 | Dealer | Website/dealer name |
| 9 | KPI | KPI name |
| 10 | Value | Current week value (from exports) |
| 11 | Last week Value | Previous week's Value (auto-shifted) |

## KPI Mapping

Raw export KPI names are normalized to standard template format:

| Export name | Template KPI |
|-------------|-------------|
| `new` | New Listings |
| `used` | Used Listings |
| `new cars` | New Cars Listings |
| `used cars` | Used Cars Listings |
| `new motorbikes` | New Motorbikes Listings |
| `used motorbikes` | Used Motorbikes Listings |
| `new truck(s)` | New Trucks Listings |
| `used truck(s)` | Used Trucks Listings |
| `new caravan / new motorhomes` | New Motorhomes Listings |
| `used caravan / used motorhomes` | Used Motorhomes Listings |

## Requirements

- Python 3.10+
- streamlit
- pandas
- openpyxl
- xlrd
