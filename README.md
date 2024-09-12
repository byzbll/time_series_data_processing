# time_series_data_processing
### Proje Tanımı
Bu proje, bir excel dosyasından çeşitli verileri okur, verileri analiz eder ve özetler. Veriler, belirli periyotlarla özetlenir ve sonuçlar yeni bir excel dosyasına kaydedilir. Bu süreç, verilerin analizini ve raporlamasını kolaylaştırmak amacıyla yapılır.

### Gerekli Kütüphaneler
- pandas: Veri işleme ve analizi için kullanılır.
- openpyxl: Excel dosyalarını işlemek için kullanılır.
- re: Düzenli ifadeler ile metin işleme için kullanılır.

### Kodun Açıklaması
1. Excel Dosyasını Yükleme: Kod, belirtilen bir Excel dosyasını (***.xlsx) yükler ve belirli bir sayfayı okur.
2. Hyperlink Bilgilerinin Çekilmesi: Excel dosyasındaki D sütunundaki her hücreden hyperlink bilgilerini ayıklar. Eğer hücrede bir bağlantı varsa, bu bağlantıdan belirli bilgileri çıkarır.
3. Verilerin İşlenmesi:
- df1, df2 ve df3 adında üç farklı sayfadan veri yüklenir.
- Timestamps (zaman damgaları) datetime formatına dönüştürülür.
- Veriler belirlenen periyotlarla özetlenir.
- Veriler, zaman indeksine göre birleştirilir ve boş hücreler bir önceki geçerli değerle doldurulur (forward fill).
4. Sonuçların Kaydedilmesi: İşlenmiş veriler sonuc.xlsx adlı bir Excel dosyasına kaydedilir. Zaman bilgileri 'Yıl-Ay-Gün Saat:Dakika' formatında sunulur.

### Önemli Kavramlar
- Hyperlink: Excel hücresinde bulunan ve belirli bir URL'ye işaret eden bağlantıdır. Kodda bu bağlantılar içinden belirli bilgiler çıkarılır.
- Resampling: Verilerin belirli bir periyotta özetlenmesi işlemidir. Örneğin, verilerin 15 saniyelik, 1 dakikalık veya 3 dakikalık periyotlarla özetlenmesi.
- Forward Fill: Boş hücreleri, bir önceki geçerli değer ile doldurma yöntemidir.
- Zaman Damgası(timestamp): Verilerin toplandığı veya işlendiği belirli bir zaman anını temsil eder. Zaman damgaları, veri analizi sırasında verilerin sıralanmasını ve zaman içindeki değişimini izlemeyi sağlar. Kodda, zaman damgaları datetime formatına dönüştürülür ve veriler bu zaman damgalarına göre özetlenir.




