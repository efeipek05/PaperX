
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

Run inside the project folder:

pip install python-docx
pip install numpy pandas matplotlib openpyxl


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
pdflatex main.tex

(Compile twice for correct TOC and references.)



TÜRKÇE
===================================

PAPERX NE YAPAR?

<img width="358" height="198" alt="image" src="https://github.com/user-attachments/assets/77fb1ada-dbd8-48a7-8114-2796bd967295" />



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


2) GEREKLİ PYTHON PAKETLERİ
------------------------------------------------------------

Proje klasöründe:

pip install python-docx
pip install numpy pandas matplotlib openpyxl


3) PROJE DOSYA YAPISI
------------------------------------------------------------
<img width="179" height="407" alt="image" src="https://github.com/user-attachments/assets/28dd3b85-cc8d-4f33-8868-63804e2cdaec" />


4) PAPERX NASIL KULLANILIR?
------------------------------------------------------------

ADIM 1 – Kapak Oluşturma (Opsiyonel)
-

python PaperX_cover.py

Dil, ders kodu, deney adı ve grup bilgileri sorulur.
cover.tex oluşturulur.



ADIM 2 – Grafik Oluşturma (Opsiyonel)
-

python PaperX_plots.py

Excel’den grafik üretir ve assets/plots/ içine PNG dosyaları koyar.



ADIM 3 – Word → LaTeX
-

python PaperX_report.py

content.tex ve toc.tex oluşturulur.



ADIM 4 – PDF Üretme
-

pdflatex main.tex
pdflatex main.tex

(İçindekiler için iki kez derleyin.)



PaperX – Structured Academic Report Automation
