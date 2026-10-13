# NovaRetail Quarterly Business Review Automation

AI-assisted executive presentation generation using **Google ADK**, **Gemini**, **Pandas**, **Matplotlib**, and **python-pptx**.

## Overview

This project automates NovaRetail's Quarterly Business Review (QBR) workflow. It analyzes an Excel dataset, generates data-grounded executive insights through a Google ADK agent, creates supporting charts, and populates a fixed PowerPoint template.

## Architecture

```text
Excel Dataset
      |
      v
analysis.py -> Verified Business Facts
      |
      v
Google ADK Agent + Gemini 3.5 Flash Lite
      |
      v
output/deck_content.json
      |
      +--> chart.py -> charts/*.png
      |
      v
build_deck.py + PowerPoint Template
      |
      v
output/NovaRetail_QBR.pptx
```

## Technology Stack

| Component | Technology |
|---|---|
| Agent framework | Google ADK |
| Model | Gemini 3.5 Flash Lite |
| Data processing | Pandas, OpenPyXL |
| Visualization | Matplotlib |
| Presentation automation | python-pptx |
| Structured validation | Pydantic |

## Project Structure

```text
Novaretail_capstone_project/
├── charts/                         # Generated chart images
├── novaretail_agent/               # Google ADK agent package
│   ├── __init__.py
│   ├── agent.py
│   └── .env                        # Local API configuration; do not commit
├── novaretail_icon_repo/           # Icons used in the PowerPoint deck
├── novaretail_venv/                # Local virtual environment; do not commit
├── output/
│   ├── deck_content.json           # Reviewed structured content from the agent
│   └── NovaRetail_QBR.pptx         # Final generated presentation
├── shape_name.txt                  # Extracted PowerPoint shape/placeholder names
├── analysis.py                     # KPI and business-fact calculations
├── build_deck.py                   # PowerPoint population and rendering
├── chart.py                        # Chart generation
├── inspect_template.py             # Inspects template shapes and names
├── models.py                       # Pydantic DeckContent schema
├── novaretail_dataset.xlsx         # Source workbook
├── novaretail_template.pptx        # Fixed presentation template
├── requirements.txt                # Python dependencies, if included
└── README.md
```

Generated folders, the virtual environment, `.env`, and Python cache files should be excluded through `.gitignore` as appropriate.

## Data and Analysis

The workbook contains three analytical areas:

- **Regional Performance:** revenue, units sold, returns rate, and marketing spend
- **Customer Experience:** NPS, CSAT, delivery days, and support tickets
- **Operational Risk:** stockout rate, return-processing days, warehouse utilization, and fraud incidents

Q2 2026 is treated as the current reporting period. Historical quarters are retained for trend analysis and charts. Quarter-over-quarter growth is calculated deterministically in Python.

## ADK Agent

The agent calls the business-facts tool before producing a validated `DeckContent` object containing:

- Executive summary
- Insight-led headlines
- Performance, customer, and risk bullets
- Recommendations
- Closing statement

The model generates narrative content only from verified facts. Python remains responsible for calculations, charts, and presentation rendering.

## PowerPoint Shape Mapping

`inspect_template.py` extracts the template's shape names and writes them to:

```text
shape_name.txt
```

`build_deck.py` uses these names to target dynamic elements without changing the fixed layout. Typical mappings include:

```text
summary_text
kpi1_value
kpi2_value
kpi3_value
title_2
insights_performance
chart_performance
title_3
insights_customer
chart_customer
title_4
insights_risk
chart_risk
recommendations_text
closing_statement
icon_1
icon_2
icon_3
icon_4
icon_5
```

To refresh the mapping after changing the template:

```bash
python inspect_template.py
```

## Setup

### 1. Create and activate a virtual environment

```bash
python -m venv novaretail_venv
```

Windows PowerShell:

```powershell
.\novaretail_venv\Scripts\Activate.ps1
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure the API key

Create `novaretail_agent/.env`:

```text
GOOGLE_API_KEY=YOUR_API_KEY
```

Never commit the `.env` file or API keys.

## Usage

### Inspect the PowerPoint template

```bash
python inspect_template.py
```

Review the generated `shape_name.txt` before updating placeholder mappings.

### Run the ADK agent

```bash
adk web
```

Open `http://localhost:8000`, select `novaretail_qbr_agent`, and request a QBR draft. Review the structured output and save the approved content as:

```text
output/deck_content.json
```

### Generate charts

```bash
python chart.py
```

Expected outputs include:

```text
charts/revenue.png
charts/customer.png
charts/stockout_rate_pct.png
```

### Generate the presentation

```bash
python build_deck.py
```

Final output:

```text
output/NovaRetail_QBR.pptx
```
## Future Enhancements

- Human-in-the-loop approval interface
- Automated JSON publication from session state
- Cloud deployment
- Multi-template support
- Automated content and layout validation

## Output

The final deliverable is `NovaRetail_QBR.pptx`, an executive QBR presentation generated from the NovaRetail dataset and fixed PowerPoint template.
