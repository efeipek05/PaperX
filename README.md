
# PaperX

PaperX is a Word (.docx) to LaTeX (.tex) report automation system.

It converts structured Word reports into LaTeX and allows you to generate
professional PDF outputs using LaTeX.

This guide explains how to use PaperX from scratch.


ENGLISH
----------------------------------------------------------

WHAT DOES PAPERX DO?

<img width="412" height="214" alt="image" src="https://github.com/user-attachments/assets/9c71fc8f-fd76-4d0f-88bc-c2906ba73228" />



1) REQUIRED PROGRAMS (INSTALL FIRST)
------------------------------------------------------------

1. Python (Required)
   Recommended: Python 3.11 (64-bit)
   Download: https://www.python.org/downloads/
   During installation: CHECK "Add Python to PATH"

   Test:
   python --version

2. MiKTeX (Required for PDF generation)
   Download: https://miktex.org/download

   Test:
   pdflatex --version

3. Pandoc (Required for equation conversion)
   Download: https://pandoc.org/installing.html

   Test:
   pandoc --version


2) INSTALL REQUIRED PYTHON PACKAGES
------------------------------------------------------------

Option A — Manual Installation
- Open the project folder in VS Code.
- Open a terminal.
- Install dependencies manually:

  pip install -r requirements.txt

  or

  pip install python-docx
  pip install numpy pandas matplotlib openpyxl


Option B — Automatic Setup (Recommended)
- Open the project folder in VS Code
- Open the integrated terminal
- Run:

  python setup.py

Then (VS Code – do this once):
Press Ctrl + Shift + P
Select Python: Select Interpreter
Choose .venv


3) PROJECT FILE STRUCTURE
------------------------------------------------------------

<img width="491" height="417" alt="image" src="https://github.com/user-attachments/assets/491365c3-fd1e-479a-b0d8-32ccbe87ef22" />


Accepted image formats are PNG, JPG, and JPEG.


4) HOW TO USE PAPERX (STEP BY STEP)
------------------------------------------------------------

STEP 1 – Generate Cover Page (Optional but Recommended)
-
Run:
python PaperX_cover.py

<img width="237" height="161" alt="image" src="https://github.com/user-attachments/assets/6c08066a-0af0-4ea6-b1b5-103b1ff78e97" />

To use PaperX_cover.py, the document logo file must be named logo_en.

It generates:
cover.tex


STEP 2 – Generate Graphs (Optional)
-
If you want to create plots from Excel:
Place your .xlsx file inside the project folder.

Run:
python PaperX_plots.py

It will:

<img width="353" height="160" alt="image" src="https://github.com/user-attachments/assets/2a179783-e817-411d-91b7-c7ce163531c6" />

You can write the axis names at the top of the column.
The column must start between rows 0 and 50.


<img width="445" height="142" alt="image" src="https://github.com/user-attachments/assets/14257f06-8151-4e9e-9b5a-adb7554df071" />


STEP 3 – Convert Word to LaTeX
-
Place your .docx file inside the project folder.

Run:
python PaperX_report.py

The script will ask:
• Language (tr/en)
• Feature selections (figures, tables, equations, etc.)

It generates:
content.tex
toc.tex


STEP 4 – Generate PDF
-

Compile LaTeX:

pdflatex main.tex

(Compile twice for correct TOC and references.)

PaperX – Structured Academic Report Automation
