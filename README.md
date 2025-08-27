# Rapid ESG Data Insights

A Python-based project for analyzing Environmental, Social, and Governance (ESG) data using World Bank datasets. This project focuses on CO2 emissions per GDP analysis across 30 countries to identify sustainability leaders and laggards.

## Features

- Automated data retrieval from World Bank API
- CO2/GDP ratio analysis for 30 countries
- Excel export with pivot charts and visualizations
- One-page executive brief on ESG leaders vs laggards
- Interactive data visualizations

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Option 1: GUI Application (Recommended)
Launch the user-friendly graphical interface:
```bash
python3 esg_gui.py
```

Features of the GUI:
- 🖱️ Point-and-click interface with tabbed layout
- 📂 Browse and select output directory
- ⚙️ Configure analysis options (data source, output types)
- 📊 Real-time progress tracking with detailed logs
- 📈 Interactive results display with summary statistics
- 📉 Built-in chart viewer with multiple visualization types
- 🗂️ Integrated file management (open output folder)

### Option 2: Command Line

**Sample Data Version (Recommended for testing):**
```bash
python3 esg_analysis_sample.py
```

**Real World Bank Data Version:**
```bash
python3 esg_analysis_real.py
```

### Generated Outputs
- `esg_data_analysis.xlsx` - Excel workbook with data and pivot charts
- `esg_brief.txt` - One-page executive brief
- `visualizations/` - Folder with 4 PNG charts (scatter, performers, distribution, boxplot)

## Data Sources

- **World Bank Open Data API** (2024 Updated)
  - CO2 Emissions: `EN.GHG.CO2.PC.CE.AR5` (Carbon dioxide emissions per capita)
  - GDP Data: `NY.GDP.PCAP.CD` (GDP per capita, current US$)
  - Latest EDGAR v8.0 emissions database with AR5 GWP values
- **Sample Data Generator** (for testing and demonstration)
- CO2 emissions data
- GDP data
- Country metadata

## Project Structure

```
RapidESGDataInsights/
├── esg_analysis.py          # Main analysis script
├── data_processor.py        # Data processing utilities
├── visualizations/          # Generated charts and plots
├── requirements.txt         # Dependencies
└── README.md               # This file
```