# Yerel URL Kontrol Aracı

Excel ve metin dosyalarındaki web adreslerini kontrol eden, bilgisayarınızda çalışan basit bir web arayüzüdür. Bir dosya seçip taramayı başlattığınızda adreslere HTTP isteği gönderir; yanıt kodunu, yönlendirme zincirini ve son adresi gösterir. Sayfa başlığı ve kısa içerik özeti de sonuçlara eklenir.

## Özellikler

- XLSX, XLSM, XLS, ODS, CSV, TSV ve TXT dosyalarını açar. Tablo dosyalarında ilk sayfanın A sütununu, CSV/TSV dosyalarında ilk sütunu; TXT dosyalarında her satırı okur.
- Yalnızca arayüzden seçtiğiniz dosyayı tarar. Dosya başına sınır 10 MB ve 100 adrestir.
- HTTP yanıtlarını ve yönlendirmeleri kaydeder; örneğin `301 → 200`.
- Sonuçları durum kodu, erişim durumu, başlık ve açıklama üzerinden filtrelemeye, XLSX olarak indirmeye izin verir.

`401` ve `403`, sunucunun yanıt verdiğini ancak erişimin kimlik doğrulaması veya izin nedeniyle kısıtlandığını belirtir. Bir HTTP yanıtı almak, sayfanın tüm işlevlerinin sorunsuz çalıştığını garanti etmez. Ağ, VPN, güvenlik duvarı ve hedef sitenin kuralları sonuçları etkileyebilir.

## Gereksinimler

- Python 3.10 veya üzeri
- Güncel bir web tarayıcısı

## Windows'ta çalıştırma

Proje klasöründe Komut İstemi (CMD) açın ve şu komutları çalıştırın:

```cmd
py -3 -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python run.py
```

## Linux'ta çalıştırma

Terminali proje klasöründe açın:

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
python run.py
```

Uygulama tarayıcıyı otomatik açar. Açılmazsa terminalde yazan `http://127.0.0.1:...` adresini tarayıcıya girin. Durdurmak için terminalde `Ctrl+C` kullanın.

## Yerel çalışma ve gizlilik

Arayüz sunucusu yalnızca `127.0.0.1` üzerinde çalışır; kurum ağına veya internete açılmaz. Uygulama bulut servisi kullanmaz. Yüklenen dosya bellekte işlenir ve diske kaydedilmez. Tarama başladığında bilgisayarınız seçtiğiniz adreslere istek gönderir; yerel/özel ağ adresleri güvenlik amacıyla engellenir.

GitHub'a kaynak kodu yüklerken kişisel tarama dosyalarını eklemeyin. `.gitignore`, tablo dosyalarını, sanal ortamı ve derleme çıktılarını Git üzerinden yüklemeye karşı hariç tutar.
