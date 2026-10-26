\# CNC Part Intelligence



CNC Part Intelligence is a lightweight, LLM-powered application developed for the SPEC'26 AI Hackathon. It accepts a part number and combines evidence from:



\- An Excel manufacturing record

\- A PDF engineering drawing

\- An IGES 3D model



The system extracts manufacturing facts, analyzes model geometry, cross-checks dimensions, flags disagreements, and answers questions using Gemini with file-level citations.



\## Level Achieved



Level 3 with stretch-oriented file upload support:



\- Level 1: Excel manufacturing record lookup and plain-language questions

\- Level 2: PDF engineering drawing extraction with source citations

\- Level 3: IGES geometry extraction and cross-source validation

\- Stretch capability: Users can upload different Excel, PDF, and IGES files



\## Main Features



\- Part-number-based Excel lookup

\- Multi-sheet Excel searching

\- PDF text and engineering-fact extraction

\- IGES geometry analysis through FreeCAD

\- Bounding-box, surface-area, and volume reporting

\- Cross-source dimensional validation

\- Gemini-powered natural-language questions

\- Exact source citations

\- Automatic switching between available Gemini models

\- Deterministic offline fallback

\- Missing-information refusal instead of guessing

\- Support for jury-supplied files



\## Technology Stack



\- Python 3

\- Streamlit

\- pandas

\- openpyxl and xlrd

\- PyMuPDF

\- FreeCAD / FreeCADCmd

\- Google Gemini API

\- python-dotenv



\## Project Structure



```text

CNC\_Part\_Analyzer/

├── app.py

├── llm\_helper.py

├── extract\_iges.py

├── requirements.txt

├── README.md

├── demo\_dataset/

│   ├── manufacturing\_records.xlsx

│   ├── 2280\_engineering\_drawing.pdf

│   ├── 2280\_model.iges

│   ├── 2280\_model.FCStd

│   └── 2280\_geometry.json

├── screenshots/

└── .env

