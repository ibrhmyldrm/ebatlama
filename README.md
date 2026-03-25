# 🪵 Atölyemhanem — Ahşap Kesim Listesi & Maliyet Hesaplama

Atölye projelerinde kullanmak üzere geliştirilmiş, kesim listelerini hazırlamana
ve maliyeti hesaplamana yarayan masaüstü uygulaması.

📺 [YouTube Kanalı](https://www.youtube.com/@atolyemhanem) &nbsp;|&nbsp; 📸 [Instagram](https://www.instagram.com/atolyemhanem/)

---

## 💻 Kullanım — EXE (Kolay Yol)

[Releases](../../releases) sayfasından `AtolyemHanem.exe` dosyasını indir, çalıştır. Python gerekmez.

---

## 🚀 Kaynak Koddan Çalıştırma

### 1. Python yüklü mü kontrol et
```
python --version
```
Python 3.8 veya üzeri gerekli.

### 2. Gerekli paketleri yükle
```
pip install -r requirements.txt
```

### 3. Çalıştır
```
python app.py
```

Uygulama otomatik olarak masaüstü penceresinde açılır.

---

## 📱 Tarayıcıdan Açmak (Aynı Wi-Fi)

`app.py` içinde `webview.start()` satırını yorum satırına alıp Flask'ı normal modda çalıştırırsan:

- **Bilgisayar:** http://localhost:5000
- **Telefon:** http://BİLGİSAYARIN_IP:5000
  - IP bulmak için: `ipconfig` (Windows) veya `ifconfig` (Mac/Linux)

---

## ✨ Özellikler

- ✅ Kesim listesi girişi (parça adı, boy, en, adet, bantlanan kenarlar)
- ✅ Standart levha boyutları + manuel giriş
- ✅ **Maximal Rectangles + WAF** algoritması ile otomatik yerleşim optimizasyonu
- ✅ Fire oranı hesaplama & görsel fire alanları
- ✅ Maliyet hesaplama: levha, kesim, PVC bant, nakliye
- ✅ Görsel levha yerleşim planı (parça numaraları, ölçüler, bant kenarları)
- ✅ **Excel** export (levha planı dahil)
- ✅ **PDF** export (her levha ayrı sayfa)
- ✅ Proje kaydet / yükle (JSON)
- ✅ Excel'den parça listesi import
- ✅ Masaüstü pencere (PyWebView)

---

## 📦 Gereksinimler

```
flask>=2.3.0
openpyxl>=3.1.0
pywebview>=4.0.0
reportlab>=4.0.0
pillow>=10.0.0
```

---

## 🔧 EXE Derleme

```
pip install pyinstaller
pyinstaller ebatlama.spec
```

`dist/AtolyemHanem.exe` oluşur.

---

## 💡 İpuçları

- Bant girişinde UK-1/UK-2 = uzun kenar, KK-1/KK-2 = kısa kenar
- Levha planında turuncu çizgiler = bantlı kenar, taralı alan = fire
- Her parçanın üstünde en, solunda boy ölçüsü yazar
- Proje JSON olarak kaydedilip sonra yüklenebilir
