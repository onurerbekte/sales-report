# Doğrulama / Verification
`pytest`: **10 passed**. CSV doğrulama, pandas toplamları, formüller, önbellek değerleri, tarihler, grafik ve metin güvenliği test edildi / CSV validation, pandas totals, formulas, cached values, dates, chart, and literal text tested.

Örnek veri: 6 kayıt, 21 adet, 2.097,50 TL / Demo data: 6 entries, 21 units, 2,097.50 TRY. İki Excel dosyası üretildi / Two Excel files generated.

Artifact Tool ile iki dosyanın her iki sayfası render edilip görsel incelendi; formül hata taraması sıfır sonuç verdi. Bir kaynak adedi 3 → 4 değiştirildiğinde brüt toplam 2.097,50 → 2.228,00 oldu; giriş geri alındı / Both sheets of both workbooks rendered and visually inspected with Artifact Tool; no formula errors matched. Changing one source quantity from 3 to 4 updated gross total from 2,097.50 to 2,228.00; restored afterward.

Microsoft Excel masaüstü uygulamasında açma/yeniden hesaplama denenmedi / Opening and recalculation in desktop Microsoft Excel were not tested.
