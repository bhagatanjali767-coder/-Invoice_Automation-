# Automated Invoice Processing System Using Python and RPA Techniques

An end-to-end Robotic Process Automation (RPA) solution designed to streamline invoice data extraction, minimize manual data entry errors, and accelerate financial workflows. The system handles diverse invoice formats, extracts core transaction parameters, and consolidates data into structured ledgers.

## Live Deployment
The operational cloud-hosted application is accessible here:  
[InvoiceAI Live Dashboard](https://invoice-automation-dashboard.onrender.com)

---

## Features
- Multi-Format Processing: Support for invoice documents in PDF (.pdf), Word (.docx), and image formats (.png, .jpg, .jpeg).
- Automated OCR Pipeline: Tesseract OCR implementation utilizing image preprocessing (grayscale conversion and rescaling) to improve extraction accuracy on low-resolution scans.
- Regex Data Parsing: Pattern-matching heuristics to identify key transaction metadata including Invoice Number, Date, Total Amount, and Payment Status.
- Ledger Generation: Aggregates processed invoice data into a unified DataFrame and exports a timestamped Excel ledger (.xlsx) on demand.
- Analytics Interface: Web-based frontend dashboard presenting real-time KPIs, transaction volume counters, and payment status tracking.

---

## Technologies Used
- Backend Architecture: Python, FastAPI, Uvicorn
- OCR and Document Parsing: Tesseract OCR, Pillow (PIL), PyPDF2, python-docx, docx2txt
- Data Engineering: Pandas, OpenPyXL
- Frontend Interface: HTML5, CSS3 (Tailwind layout), Vanilla JavaScript (Fetch API)

---

## Repository Structure
```text
Invoice_Automation/
│
├── backend/
│   ├── main.py            # FastAPI application and extraction engine
│   └── requirements.txt   # Backend dependencies
│
├── frontend/
│   ├── index.html         # Interactive UI Dashboard
│   └── [static assets]    # Stylesheets and client-side scripts
│
├── uploads/               # Temporary server directory for raw files
└── processed_data/        # Output directory for generated Excel ledgers
