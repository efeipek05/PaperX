TÜRKÇE
===================================

PAPERX NE YAPAR?

<img width="358" height="198" alt="image" src="https://github.com/user-attachments/assets/77fb1ada-dbd8-48a7-8114-2796bd967295" />

Bu kılavuz, PaperX’i sıfırdan nasıl kullanacağınızı açıklar.

----------------------------------------------------------

1) GEREKLİ PROGRAMLAR (ÖNCE KURUN)
------------------------------------------------------------

1. Python (Zorunlu)
   Önerilen: Python 3.11 (64-bit)
   İndirme: https://www.python.org/downloads/
   Kurulum sırasında: "Add Python to PATH" seçeneğini işaretleyin.

   Test:
   python --version

2. MiKTeX (PDF üretimi için zorunlu)
   İndirme: https://miktex.org/download

   Test:
   pdflatex --version

3. Pandoc (Denklem dönüşümü için zorunlu)
   İndirme: https://pandoc.org/installing.html

   Test:
   pandoc --version


2) GEREKLİ PYTHON PAKETLERİ
------------------------------------------------------------

Proje klasörü içinde çalıştırın:

pip install python-docx
pip install numpy pandas matplotlib openpyxl


3) PROJE DOSYA YAPISI
------------------------------------------------------------
<img width="179" height="407" alt="image" src="https://github.com/user-attachments/assets/28dd3b85-cc8d-4f33-8868-63804e2cdaec" />

Desteklenen görsel formatları: PNG, JPG ve JPEG.


4) PAPERX NASIL KULLANILIR? (ADIM ADIM)
------------------------------------------------------------

ADIM 1 – Kapak Oluşturma (Opsiyonel fakat önerilir)
-
Çalıştırın:

python PaperX_cover.py

PaperX_cover.py dosyasını kullanabilmek için, belge logosu dosyasının adı logo_en olmalıdır.

Desteklenen formatlar: PNG, JPG ve JPEG.

Oluşturulan dosya:
cover.tex


ADIM 2 – Grafik Oluşturma (Opsiyonel)
-
Excel dosyanızı (.xlsx) proje klasörüne yerleştirin.

Çalıştırın:

python PaperX_plots.py

Script şunları sorar:

<img width="553" height="159" alt="image" src="https://github.com/user-attachments/assets/7b07eaff-ff23-4521-83ab-ffe3d664bde1" />


Grafik oluştururken:

- Eksen isimlerini sütunun en üstüne yazabilirsiniz.
- Veri sütunu 0–50 satır aralığında başlamalıdır.

<img width="422" height="123" alt="image" src="https://github.com/user-attachments/assets/a6282bc6-b92a-41b5-9af2-cc3eb7a0f2b6" />


ADIM 3 – Word → LaTeX Dönüştürme
-
.docx dosyanızı proje klasörüne yerleştirin.

Çalıştırın:

python PaperX_report.py

Script sizden şunları ister:

- Dil seçimi (tr/en)
- Özellik seçimleri (şekiller, tablolar, denklemler vb.)

Oluşturulan dosyalar:
content.tex
toc.tex


ADIM 4 – PDF Üretme
-

LaTeX dosyasını derleyin:

pdflatex main.tex

(İçindekiler ve referansların doğru oluşması için iki kez derleyin.)
