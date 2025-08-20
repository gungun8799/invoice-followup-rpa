# Invoice Follow-up RPA

## 🎯 Overview

Automated RPA system for exporting data from TIMS (Lotus) system and Power BI with file management capabilities.

## 📁 Project Structure

```
Invoice-follow-up-rpa/
├── tims_final.py                     # TIMS automation script
├── PBI_export.py                     # Power BI export automation
├── FINAL_SOLUTION_STATUS.md          # Project documentation
├── README.md                         # This file
└── downloads/                        # Downloaded export files
```

## 🚀 Quick Start

### TIMS Export
```bash
python tims_final.py
```

### Power BI Export
```bash
python PBI_export.py
```

## ✅ Features

### TIMS Automation
- **Automated Login**: Handles TIMS authentication automatically
- **Smart Date Processing**: Always exports yesterday's data (Bangkok timezone)
- **Robust Export Process**: Network interception ensures reliable downloads
- **File Management**: Automatic ZIP extraction and proper naming
- **Error Handling**: Comprehensive retry and recovery mechanisms

### Power BI Automation
- **Automated Export**: Extracts data from Power BI reports
- **File Processing**: Handles download and organization
- **Data Management**: Organizes exported files systematically

## 📊 Current Status

### ✅ Working Features
- **TIMS export automation**: Fully functional
- **Power BI export automation**: Ready for use
- **File download and extraction**: Reliable operation
- **Local file organization**: Automatic

## 📋 Requirements

- Python 3.x
- Playwright
- Required Python packages (install with requirements.txt if available)

## 🎉 Success Metrics

- **Time savings**: Automated daily export processes
- **Error reduction**: Elimination of manual export errors
- **Process reliability**: Consistent automation workflow
- **File organization**: Systematic file naming and storage