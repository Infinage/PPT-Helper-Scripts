# PowerPoint Automation Script for Windows

## Overview

This script automates the process of creating and updating editable PowerPoint presentations on Windows. It is designed for scenarios where you need to generate or update PowerPoint presentations with tables and charts that remain fully editable.

### Key Features
- **Table Management**: Directly update PowerPoint tables using `python-pptx` combined with XML manipulation for fine-grained control.
- **Chart Management**: Embed charts into PowerPoint presentations via dummy Excel files, and automate updates by:
  - Unzipping the PPTX file.
  - Modifying the embedded Excel files using `xlwings` to ensure Excel recalculates and updates the plots.
  - Updating the PowerPoint cache to reflect the changes in the Excel data, ensuring charts stay linked and editable.

## Prerequisites

This script is designed to run on Windows systems with PowerPoint installed. Before running the script, ensure you have the following Python libraries installed:

- `lxml`
- `pandas`
- `python-pptx`
- `xlwings`
- `pywin32`

## How It Works

### Table Handling

- The script uses the python-pptx library to manipulate tables directly in the PowerPoint file.
- For more advanced customization, it leverages XML manipulation to directly modify the table properties and content within the PPTX file structure.

### Chart Handling

- *Embedding Charts*: Charts are initially embedded into the PowerPoint presentation using dummy Excel files.
- *Updating Charts*: 
    - The script unzips the PPTX file and identifies the embedded Excel files.
    - It then uses xlwings to modify the Excel files, ensuring that any changes trigger Excel to recalculate and update the charts.
    - After the updates, the script modifies the PowerPoint's internal cache (via XML) to sync with the updated Excel data, ensuring that the charts reflect the latest data and remain editable.
