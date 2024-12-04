# Sales Analysis Tool 🚀

A Python-based tool for managing and analyzing sales data across different product categories (Pepsi, Water, Gatorade, Lipton, and Juice). 📊

## Features ⭐

- Data Input through copy-paste functionality 📋
- Automatic Excel sheet updates for all products 📝
- Comprehensive analysis with visualizations 📈
- Performance tracking and insights 📊
- Export capability for analysis graphs 💾

## Installation 💻

1. Download the latest installer from the releases section
2. Run the installer and follow the setup wizard
3. Launch the application from the Start Menu or Desktop shortcut

## Usage 🔨

1. Select your Excel file
2. Paste your sales data
3. Choose product(s) to analyze
4. View instant insights and analysis

## Building from Source 🛠️

Requirements:
- Python 3.12
- Required packages: pandas, openpyxl, matplotlib
- PyInstaller for creating executable
- Inno Setup for creating installer

Install required packages:
```
pip install pandas openpyxl matplotlib pyinstaller
```

Build steps:
1. Run PyInstaller:
```
pyinstaller sales_tool.spec
```

2. Create installer using Inno Setup with `setup.iss`

## Files Structure 📁
- `final.py`: Main application code
- `sales_tool.spec`: PyInstaller specification file
- `setup.iss`: Inno Setup script
- `icon.ico`: Application icon

## Credits 👨‍💻

**Developer:** Talah  
**Organization:** PepsiCo, Inc. 🥤  
**Department:** Sales Department  

This tool was developed to streamline sales data management and analysis processes for the PepsiCo Sales Division. 📈

## License ⚖️

© 2024 Talah Tanveer. All Rights Reserved.  
This software is proprietary and confidential.
