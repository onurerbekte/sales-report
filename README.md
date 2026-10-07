# Satış analizi ve Excel otomasyonu / Sales analysis and Excel automation
**Kurgusal demo proje / Fictional demo project.**

## Türkçe
pandas ile CSV satışlarını doğrular, ürün bazında gruplar ve XlsxWriter ile Türkçe/İngilizce Excel raporları oluşturur. Altı kurgusal kayıt vardır. Excel'de kaynak adet/fiyatı değiştirince formüller ve bağlı grafik güncellenir. pandas özeti dosya üretiminde hesaplanır; yeni ürün veya satır eklenirse scripti yeniden çalıştır. Tutarlar TL, vergi/indirim hariç brüt satıştır.

Python 3.12+ sanal ortamında `python -m pip install -r requirements.txt`, ardından:
```powershell
python report.py
python -m pytest -q
```
Çıktılar: `reports/sales_tr.xlsx`, `reports/sales_en.xlsx`. Yeni veri: `python report.py --input data/demo_sales.csv --output reports`. CSV kolonları: `order_id,date,product,quantity,unit_price`; tarih YYYY-MM-DD; fiyat ondalık noktayla, en fazla iki basamak. Kimlikler satır başına benzersiz olmalı. Geçersiz veriler sessizce sıfıra çevrilmez; script durur.

## English
Validates sales CSV data with pandas, groups products, and creates Turkish/English Excel reports with XlsxWriter. The six entries are fictional. Source quantity/price changes recalculate Excel formulas and the linked chart; rerun the script after adding new products or rows. Amounts are gross TRY sales before tax/discounts. Install requirements in a Python 3.12+ virtual environment and run the commands above.

Output paths and CSV schema are listed above. Dates use YYYY-MM-DD; prices use a decimal point and at most two decimal places. Entry IDs must be unique. Invalid input stops processing rather than becoming a plausible zero. Formula caches are populated; actual Excel desktop recalculation was not exercised in this environment.
