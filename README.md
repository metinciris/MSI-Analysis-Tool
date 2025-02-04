# MSI Analysis Tool

## Overview
This tool analyzes MSI (Microsatellite Instability) reports from Excel files and classifies them into different stability categories: **Stable (MSS)**, **Low (MS-Low)**, and **High (MS-High)**.

The script provides a graphical interface using **Tkinter**, allowing users to select Excel files for analysis and view the results in a structured format.

## Features
- **Excel File Selection:** Users can select multiple MSI report Excel files.
- **Automated Analysis:** Extracts and processes MSI status based on predefined thresholds.
- **Classification:** Determines MSI status as Stable (MSS), Low (MS-Low), or High (MS-High) based on unstable loci count.
- **User-friendly GUI:** Displays results in a large text window.
- **Copy Functionality:** Users can copy all results with a single click.
- **Exit Handling:** Closes the application and ensures clipboard content is retained.

## MSI Classification Criteria
- **Stable (MSS):** If **0 or 1 unstable loci** are detected.
- **Low (MS-Low):** If **2 or 3 unstable loci** are detected.
- **High (MS-High):** If **4 or more unstable loci** are detected.

## Installation
### Prerequisites
Ensure you have **Python 3.x** installed with the required dependencies:
```sh
pip install pandas openpyxl
```

## Usage
Run the script using Python:
```sh
python msi_analysis_excel.py
```
### Steps:
1. A message box will appear guiding you to select MSI report Excel files.
2. Choose the Excel files containing MSI data.
3. The results will be displayed in a large window.
4. You can copy all results using the **"Tümünü Kopyala"** button.
5. Click "Kapat" to exit.

## Example Output
```
VAKA-7-25  MS-stable
- Mikrosatellit Stabil (MS-Stable/MSS). Dokuz bölgenin hiç birinde instabilite saptanmadı.

VAKA-3-25  MS-stable
- Mikrosatellit Low (MS-Low/MSS). Dokuz bölgenin ikisinde (NR22(T)21, BAT34C4(A)18) instabilite saptandı.

VAKA-5-25  MS-stable
- Mikrosatellit High (MS-High/MSS). Dokuz bölgenin beşinde (BAT40(T)37, BAT26(A)27, NR21(A)21, NR22(T)21, MONO-27(T)27) instabilite saptandı.
```

## License
This project is licensed under the MIT License.


