# Weekly Report App

Haftalık rapor hazırlama ve gönderme uygulaması.

## Kurulum

1. `WeeklyReportApp-Setup.exe` dosyasını çalıştırın
2. Kurulum sihirbazını takip edin
3. Kurulum tamamlandığında masaüstünde bir kısayol oluşturulacak

## İlk Kullanım

1. Uygulamayı ilk kez çalıştırdığınızda aşağıdaki bilgileri girmeniz istenecek:
   - Ad Soyad
   - Gmail adresi
   - Hedef email adresi (raporun gönderileceği adres)
   - Word şablon dosyası

2. Gmail hesabınız için "App Password" oluşturmanız gerekiyor:
   - Gmail hesabınıza giriş yapın
   - Sağ üst köşedeki profil resminize tıklayın
   - "Google Hesabını Yönet" seçeneğine tıklayın
   - Sol menüden "Güvenlik" seçeneğine tıklayın
   - "2 Adımlı Doğrulama"yı açın
   - "Uygulama Şifreleri" bölümüne gidin
   - "Uygulama Seç" dropdown'ından "Diğer" seçeneğini seçin
   - İsim olarak "WeeklyReportApp" yazın
   - "Oluştur" butonuna tıklayın
   - Verilen 16 haneli şifreyi kaydedin

## Word Şablonu

Word şablonunuzda aşağıdaki yer tutucuları kullanabilirsiniz:
- {Ad} - Adınız için
- {Tarih} - Tarih için
- {TamamlananFaaliyetler} - Tamamlanan faaliyetler için
- {DevamEdenFaaliyetler} - Devam eden faaliyetler için
- {PlanlananFaaliyetler} - Planlanan faaliyetler için

## Otomatik Çalışma

Uygulama her Cuma saat 18:00'de otomatik olarak açılacak ve rapor girmenizi isteyecektir.

## Ayarları Güncelleme

Uygulamayı açtığınızda "Settings" butonuna tıklayarak ayarlarınızı güncelleyebilirsiniz.

## Kaldırma

1. Windows Denetim Masası > Programlar ve Özellikler
2. "Weekly Report App"i seçin
3. "Kaldır" butonuna tıklayın 