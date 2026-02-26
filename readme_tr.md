# PaperX

PaperX, Word (.docx) belgelerini LaTeX (.tex) formatına dönüştüren bir rapor otomasyon sistemidir. Yapılandırılmış Word raporlarını LaTeX’e çevirir ve LaTeX kullanarak profesyonel PDF çıktıları üretmenizi sağlar. Bu rehber, PaperX’i sıfırdan nasıl kullanacağınızı açıklar.

TÜRKÇE
----------------------------------------------------------

## PAPERX NE YAPAR?

(README'deki görsel)

----------------------------------------------------------
## 1) GEREKLİ PROGRAMLAR (ÖNCE KURULUM)

1. Python (Gerekli)  
   Önerilen: Python 3.11 (64-bit)  
   İndir: https://www.python.org/downloads/  
   Kurulum sırasında: "Add Python to PATH" seçeneğini işaretleyin  
   Test: `python --version`

2. MiKTeX (PDF üretimi için gerekli)  
   İndir: https://miktex.org/download  
   Test: `pdflatex --version`

3. Pandoc (Denklem dönüşümü için gerekli)  
   İndir: https://pandoc.org/installing.html  
   Test: `pandoc --version`

----------------------------------------------------------
## 2) GEREKLİ PYTHON PAKETLERİNİ KUR

### Seçenek A — Manuel Kurulum
- Proje klasörünü VS Code ile açın.
- Terminali açın.
- Bağımlılıkları manuel kurun:

`pip install -r requirements.txt`

veya

`pip install python-docx`  
`pip install numpy pandas matplotlib openpyxl`

### Seçenek B — Otomatik Kurulum (Önerilen)
- Proje klasörünü VS Code ile açın
- VS Code içindeki terminali açın
- Çalıştırın:

`python setup.py`

Sonra (VS Code – bunu bir kere yapın):
- `Ctrl + Shift + P`
- **Python: Select Interpreter**
- `.venv` seçin

----------------------------------------------------------
## 3) PROJE DOSYA YAPISI

(README'deki görsel)

Kabul edilen görsel formatları: PNG, JPG ve JPEG.

----------------------------------------------------------
## 4) PAPERX NASIL KULLANILIR (ADIM ADIM)

### ADIM 1 – Kapak Sayfası Oluştur (Opsiyonel ama önerilir)
- Çalıştırın: `python PaperX_cover.py`

(README'deki görsel)

PaperX_cover.py kullanmak için, doküman logosu dosyasının adı `logo_en` olmalıdır.  
Ürettiği çıktı: `cover.tex`

### ADIM 2 – Grafik Oluştur (Opsiyonel)
- Excel’den plot üretmek istiyorsanız:
  `.xlsx` dosyanızı proje klasörünün içine koyun.

Çalıştırın: `python PaperX_plots.py`

Bu script:

(README'deki görseller)

Notlar:
- Eksen isimlerini sütunun en üstüne yazabilirsiniz.
- Sütun başlangıcı 0 ile 50. satır aralığında olmalıdır.

### ADIM 3 – Word’ü LaTeX’e Dönüştür
- `.docx` dosyanızı proje klasörünün içine koyun.
- Çalıştırın: `python PaperX_report.py`

Script sizden şunları ister:
- Dil seçimi (tr/en)
- Özellik seçimleri (figürler, tablolar, denklemler vb.)

Ürettiği dosyalar:
- `content.tex`
- `toc.tex`

### ADIM 4 – PDF Oluştur
- LaTeX derleyin: `pdflatex main.tex`
  (İçindekiler ve referansların doğru çıkması için iki kez derleyin.)

PaperX – Yapılandırılmış Akademik Rapor Otomasyonu
