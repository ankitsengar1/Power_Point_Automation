# PowerPoint Report Automation Using Python
## Overview

This project automates the generation of PowerPoint business reports using Python. It streamlines repetitive reporting tasks by transforming data into structured presentations containing KPI summaries, tables, charts, and business insights.

The goal is to reduce manual reporting effort, improve consistency, and make business reporting more efficient.

## Key Features

- Automated PowerPoint presentation generation.
- Data processing and transformation using Pandas.
- SQL-based data extraction and analysis.
- Dynamic KPI tables and charts.
- Automated formatting of PowerPoint slides.
- Reusable reporting workflows for recurring business reports.

## Tech Stack

- Python — Core programming and automation
- Pandas — Data cleaning and transformation
- SQL — Data extraction and querying
- python-pptx — PowerPoint presentation generation
- Matplotlib — Charts and data visualization

## Project Workflow

1. Extract data from SQL databases or input files.
2. Clean, transform, and prepare the data using Pandas.
3. Calculate KPIs and generate relevant charts.
4. Populate PowerPoint slides with tables, charts, and summaries.
5. Generate the final presentation for business reporting.

## Project Structure

Power_Point_Automation/
├── data/               # Input datasets 
├── src/                # Automation scripts
├── output/             # Generated PowerPoint reports
├── requirements.txt    # Project dependencies
└── README.md           # Project documentation

Update the folder structure above to match the actual repository.

## Installation

Clone the repository:

git clone <your-repository-url>
cd Power_Point_Automation

Create and activate a virtual environment:

python -m venv venv

On Windows:

venv\Scripts\activate

Install the dependencies:

pip install -r requirements.txt

## Usage

1. Configure the input data source and reporting parameters.
2. Run the relevant Python automation script.
3. Review the generated PowerPoint presentation in the output directory.

Refer to the project scripts and configuration files for the exact execution steps.

## Business Value

- Reduces repetitive manual reporting tasks.
- Improves consistency across recurring presentations.
- Simplifies KPI tracking and business performance analysis.
- Makes reporting workflows more reusable and scalable.

## Future Enhancements

- Support for multiple report templates.
- Automated email delivery of generated presentations.
- Scheduled report generation.
- Integration with additional data sources.

## Author

Ankit Sengar
---

*This project demonstrates the application of Python automation and data analytics to streamline business reporting workflows.*
