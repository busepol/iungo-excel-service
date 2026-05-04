<p align="center">
  <img src="https://github.com/busepol/iungo-excel-service/blob/main/assets/REGALIDEA.png" alt="Regalidea Logo" width="250">
</p>

<h1 align="center">🚀 IUNGO Order Automator</h1>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/Flask-000000?style=for-the-badge&logo=flask&logoColor=white" />
  <img src="https://img.shields.io/badge/n8n-FF6D5A?style=for-the-badge&logo=n8n&logoColor=white" />
  <img src="https://img.shields.io/badge/Railway-131415?style=for-the-badge&logo=railway&logoColor=white" />
</p>

<p align="center">
  <strong>Transforming raw data extractions into professional, formula-driven IUNGO Excel Orders.</strong>
</p>

---

## 📖 Overview

The **IUNGO Order Automator** is a specialized microservice hosted on **Railway** that acts as the "processing brain" for the **Regalidea** automation ecosystem. 

When **n8n** extracts order data from varied sources (GABRIELLI, PAC 2000, etc.), it sends a JSON payload to this service. The service then dynamically builds a pixel-perfect, protected Excel file from scratch, featuring live formulas and professional styling.

## ✨ Key Features

*   **🎯 Intelligent Code Mapping**: Automatically distinguishes between Customer Article Codes and Supplier Codes to ensure data integrity.
*   **🧮 Live Excel Formulas**: Injects `=SUM()` and discounted price formulas directly into cells, ensuring totals update if values are changed post-generation.
*   **🛡️ Enterprise-Grade Protection**: Automatically locks admin fields (N. Documento, Totals) while highlighting and unlocking "Yellow Zones" for user input.
*   **🎨 IUNGO Professional Styling**: Precise adherence to the IUNGO color palette (`#203764` Dark Blue) and layout standards.
*   **📑 Instruction Sheet**: Automatically generates a secondary "ISTRUZIONI" tab to guide customers through the order process.

## 🏗️ The Workflow

1.  **Extraction**: `n8n` cleanses data from incoming order formats.
2.  **Request**: `n8n` sends a POST request to this service on Railway.
3.  **Generation**: `pp.py` builds the `.xlsx` file using `openpyxl`.
4.  **Completion**: The binary file is returned to `n8n` for delivery to Google Drive or the final customer.

---

## 🚀 Deployment & Installation

### Railway Setup
This repository is optimized for one-click deployment on Railway.
1. Connect your GitHub repository to **Railway.app**.
2. Railway will automatically detect the `requirements.txt` and `pp.py`.
3. The service will expose the `/generate` endpoint.

### Local Development
```bash
# Clone the repository
git clone [https://github.com/busepol/iungo-excel-service.git](https://github.com/busepol/iungo-excel-service.git)

# Install dependencies
pip install -r requirements.txt

# Start the Flask server
python pp.py
