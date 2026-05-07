#  Automated Woodwork Cutting List Generator

![CI Status](https://github.com/ELMILLO/AutomatedCalculator/actions/workflows/ci.yml/badge.svg)
![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![xlwings](https://img.shields.io/badge/xlwings-Enabled-green.svg)

##  Overview
This project automates the generation of precise cutting lists (Bill of Materials) for a professional woodworking shop. It transforms a highly manual, error-prone Excel estimation process into a streamlined, one-click automated pipeline. 

By applying strict business logic and dynamic dimensional calculations, the script ensures 100% accuracy in manufacturing measurements, eliminating material waste caused by human calculation errors.

##  Key Features & Business Logic
- **Algorithmic Dimensioning:** Dynamically calculates internal components (drawers, shelves, back panels) based on user-defined external dimensions.
- **Automated Edge Banding Allocation:** Intelligently assigns edge banding colors and materials, differentiating between structural components and exterior facades (doors/fronts).
- **Edge Case Handling & Validation:** Built-in validation prevents erroneous inputs (e.g., blocking depth inputs for 2D custom cuts).
- **Dynamic Data Cleaning:** Automatically filters out non-essential components (e.g., 'gola' cuts) and unneeded columns from the final production sheet.
- **Continuous Integration (CI):** Integrated with **GitHub Actions** for automated static code analysis (flake8) to maintain code quality and PEP-8 standards on every push.

##  Tech Stack
- **Language:** Python
- **Libraries:** `xlwings` (for seamless Excel-to-Python bidirectional communication), `pandas` (data extraction & testing)
- **CI/CD:** GitHub Actions
- **Interface:** Microsoft Excel Macro-Enabled Workbook (`.xlsm`)

##  The QA & SDET Perspective
From a Quality Assurance standpoint, this tool solves a critical business bottleneck:
1. **Defect Prevention:** Hardcoded formulas in Excel are easily broken by users. This Python backend locks the logic, making it tamper-proof.
2. **Reliability:** Standardizes outputs regardless of the user, ensuring the factory floor always receives the same standardized manufacturing template.
3. **Maintainability:** Business rules (like subtracting 15mm for a cabinet door) are centralized in the Python script, making future updates trivial compared to fixing hundreds of Excel cells.

##  How to Run
1. Ensure Python 3.10+ is installed.
2. Install the required dependencies: `pip install -r requirements.txt`
3. Open the `Calculadora.xlsm` file.
4. Input the desired furniture parameters in the UI sheet and click "Calcular Despieces".
5. The Python script will execute in the background and generate a formatted, production-ready sheet instantly.