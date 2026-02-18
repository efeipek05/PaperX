
# PaperX

PaperX is a Word (.docx) to LaTeX (.tex) report automation system.

It converts structured Word reports into LaTeX and allows you to generate
professional PDF outputs using LaTeX.

This guide explains how to use PaperX from scratch.

===========================================
ENGLISH
===========================================

WHAT DOES PAPERX DO?

• Converts .docx to LaTeX
• Automatically processes headings
• Converts Word equations via Pandoc
• Handles figures and tables
• Generates table of contents
• Supports automatic section numbering
• Allows automatic cover page generation
• Supports automatic graph generation from Excel

------------------------------------------------------------
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

------------------------------------------------------------
2) INSTALL REQUIRED PYTHON PACKAGES
------------------------------------------------------------

Run inside the project folder:

pip install python-docx
pip install numpy pandas matplotlib openpyxl

------------------------------------------------------------
3) PROJECT FILE STRUCTURE
------------------------------------------------------------

PaperX/
│
├── PaperX_report.py      ← MAIN REPORT ENGINE
├── PaperX_cover.py       ← COVER PAGE GENERATOR
├── PaperX_plots.py       ← GRAPH GENERATOR (Excel → PNG)
├── main.tex
├── content.tex
├── toc.tex
├── your_file.docx
│
└── assets/
    ├── image1.png
    ├── image2.png
    ├── logo_tr.png
    ├── logo_en.png
    └── plots/

------------------------------------------------------------
4) HOW TO USE PAPERX (STEP BY STEP)
------------------------------------------------------------

STEP 1 – Generate Cover Page (Optional but Recommended)

Run:

python PaperX_cover.py

It will ask:
• Language
• Course code
• Experiment title
• Submission date
• Group member information

It generates:
cover.tex

------------------------------------------------------------

STEP 2 – Generate Graphs (Optional)

If you want to create plots from Excel:

python PaperX_plots.py

It will:
• Ask Excel file name
• Ask number of plots
• Ask polynomial degree
• Ask column letters
• Generate PNG files inside assets/plots/
• Create plots_meta.json

------------------------------------------------------------

STEP 3 – Convert Word to LaTeX

Place your .docx file inside the project folder.

Run:

python PaperX_report.py

The script will ask:
• Language (tr/en)
• Feature selections (figures, tables, equations, etc.)

It generates:
content.tex
toc.tex

------------------------------------------------------------

STEP 4 – Generate PDF

Compile LaTeX:

pdflatex main.tex
pdflatex main.tex

(Compile twice for correct TOC and references.)

============================================================
TÜRKÇE
============================================================

PAPERX NE YAPAR?

• .docx dosyasını LaTeX’e çevirir
• Başlıkları otomatik işler
• Word denklemlerini Pandoc ile dönüştürür
• Görselleri ve tabloları işler
• İçindekiler üretir
• Bölüm numaralandırması yapar
• Otomatik kapak sayfası oluşturur
• Excel’den grafik üretir

------------------------------------------------------------
1) GEREKLİ PROGRAMLAR (ÖNCE KURUN)
------------------------------------------------------------

1. Python (Zorunlu)
   Önerilen: Python 3.11 (64-bit)
   İndirme: https://www.python.org/downloads/
   Kurulum sırasında "Add Python to PATH" işaretleyin.

   Kontrol:
   python --version

2. MiKTeX (PDF için zorunlu)
   İndirme: https://miktex.org/download

   Kontrol:
   pdflatex --version

3. Pandoc (Denklem dönüşümü için zorunlu)
   İndirme: https://pandoc.org/installing.html

   Kontrol:
   pandoc --version

------------------------------------------------------------
2) GEREKLİ PYTHON PAKETLERİ
------------------------------------------------------------

Proje klasöründe:

pip install python-docx
pip install numpy pandas matplotlib openpyxl

------------------------------------------------------------
3) PROJE DOSYA YAPISI
------------------------------------------------------------

PaperX/
│
├── PaperX_report.py
├── PaperX_cover.py
├── PaperX_plots.py
├── main.tex
├── content.tex
├── toc.tex
├── dosyaniz.docx
│
└── assets/
    ├── image1.png
    ├── image2.png
    ├── logo_tr.png
    ├── logo_en.png
    └── plots/

------------------------------------------------------------
4) PAPERX NASIL KULLANILIR?
------------------------------------------------------------

ADIM 1 – Kapak Oluşturma (Opsiyonel)

python PaperX_cover.py

Dil, ders kodu, deney adı ve grup bilgileri sorulur.
cover.tex oluşturulur.

------------------------------------------------------------

ADIM 2 – Grafik Oluşturma (Opsiyonel)

python PaperX_plots.py

Excel’den grafik üretir ve assets/plots/ içine PNG dosyaları koyar.

------------------------------------------------------------

ADIM 3 – Word → LaTeX

python PaperX_report.py

content.tex ve toc.tex oluşturulur.

------------------------------------------------------------

ADIM 4 – PDF Üretme

pdflatex main.tex
pdflatex main.tex

(İçindekiler için iki kez derleyin.)

------------------------------------------------------------

PaperX – Structured Academic Report Automation
