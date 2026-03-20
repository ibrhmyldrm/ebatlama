# 🪵 PORE — Ahşap Kesim Listesi & Maliyet Hesaplama

Atölye projelerinde kullanmak üzere geliştirilmiş, poreye (veya kendi atölyene)
gönderilecek kesim listelerini hazırlamana ve maliyeti hesaplamana yarayan
masaüstü + mobil uyumlu web uygulaması.

---

## 🚀 Kurulum (İlk kez)

### 1. Python yüklü mü kontrol et
```
python --version
```
Python 3.8 veya üzeri gerekli.

### 2. Gerekli paketleri yükle
```
pip install -r requirements.txt
```

---

## ▶️ Çalıştırma

```
python app.py
```

Ardından tarayıcında aç:
- **Bilgisayar:** http://localhost:5000
- **Telefon (aynı Wi-Fi'de ise):** http://BİLGİSAYARIN_IP:5000
  - Bilgisayarının IP'sini bulmak için: `ipconfig` (Windows) veya `ifconfig` (Mac/Linux)

---

## 📱 Özellikler

- ✅ Kesim listesi girişi (parça adı, en, boy, adet, bantlanan kenarlar)
- ✅ Standart levha boyutları + manuel giriş
- ✅ Guillotine algoritması ile otomatik yerleşim optimizasyonu
- ✅ Fire oranı hesaplama
- ✅ Maliyet: levha, kesim, PVC bant, nakliye
- ✅ Görsel levha yerleşim planı (hangi parça hangi levhada)
- ✅ Renkli Excel çıktısı
- ✅ Mobil uyumlu arayüz

---

## 🔧 VSCode'da Açmak

1. VSCode'da klasörü aç: `Dosya > Klasör Aç > pore_app`
2. Terminal aç: `Ctrl + J`
3. `python app.py` yaz ve Enter
4. Tarayıcıda http://localhost:5000 aç

---

## 💡 İpuçları

- Bant girişinde birden fazla kenar seçebilirsin
- "Levha Yerleşim Planı"nda turuncu çizgiler = bantlı kenar
- Excel dosyası tarih-saatli şekilde kaydedilir
- Aynı Wi-Fi'deyse telefonundan da açabilirsin
