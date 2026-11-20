# Yapay Zeka Modelleri A/B Testi

Üç yerel büyük dil modelinin, farklı **temperature**, **top-p** ve **top-k** ayarlarında
100 genel kültür sorusuna verdiği cevapların doğruluk açısından karşılaştırılması.

## Dosyalar

| Dosya | İçerik |
|---|---|
| `AB_Testi_Sonuclar.xlsx` | Sonuçlar: 2100 cevap, puanlar, özet tablolar ve grafik |
| `sorgula.py` | Testi baştan sona çalıştıran betik (Excel'i oluşturur, soruları modellere sorar, cevapları yazar) |

## Deney Tasarımı

**Modeller** (Ollama ile yerel olarak, Q4_K_M sıkıştırma):

| Kod | Model |
|---|---|
| M1 | Qwen3.5 9B (`qwen3.5:9b-q4_K_M`) |
| M2 | Gemma 4 12B (`gemma4:12b-it-q4_K_M`) |
| M3 | Ministral 3 8B Instruct (`ministral-3:8b-instruct-2512-q4_K_M`) |

**Parametre ayarları:** A3 temel ayardır. A1–A2 ve A6–A7 temperature'ın etkisini,
A4 top-p'nin, A5 top-k'nın etkisini ölçer. A7, modelin hata yapmaya başladığı sınırı
test eden uç ayardır.

| Ayar | Açıklama | Temperature | Top-p | Top-k |
|---|---|---|---|---|
| A1 | En düşük çeşitlilik (deterministik) | 0 | 1.0 | 1 |
| A2 | Düşük temperature | 0.2 | 0.9 | 40 |
| A3 | Temel (referans) | 0.7 | 0.9 | 40 |
| A4 | Düşük top-p | 0.7 | 0.5 | 40 |
| A5 | Düşük top-k | 0.7 | 0.9 | 5 |
| A6 | Yüksek temperature | 1.2 | 0.9 | 40 |
| A7 | Çok yüksek temperature | 1.5 | 0.95 | 50 |

**Sorular:** 7 kategoride 100 soru (coğrafya, tarih, bilim, matematik, edebiyat ve sanat,
teknoloji, spor ve genel kültür). Toplam 3 model × 7 ayar × 100 soru = **2100 cevap**.

**Ortak koşullar:** Tüm modellerde aynı sistem mesajı ("kısa ve net cevapla"),
en fazla 256 token, sabit seed (42), düşünme modu kapalı.

**Puanlama:** Her cevap elle 1 (doğru) veya 0 (yanlış) olarak puanlandı. Doğru anahtar
kelimeyi içerse bile uydurma bilgi barındıran cevaplar 0 sayıldı. Yazım hataları ve
yuvarlamalar (ör. 42,195 km yerine 42 km) kabul edildi. Sınırda kalan kararların
gerekçeleri Excel'deki "Not" sütunundadır.

## Sonuçlar

Ayar bazında doğruluk oranı:

| Ayar | Qwen3.5 9B | Gemma 4 12B | Ministral 3 8B | Ortalama |
|---|---|---|---|---|
| A1 (T=0) | %92 | %98 | %95 | %95,0 |
| A2 (T=0.2) | %92 | %97 | %94 | %94,3 |
| A3 (T=0.7, temel) | %83 | %97 | %90 | %90,0 |
| A4 (top-p 0.5) | %87 | %97 | %94 | %92,7 |
| A5 (top-k 5) | %86 | %97 | %91 | %91,3 |
| A6 (T=1.2) | %79 | %97 | %91 | %89,0 |
| A7 (T=1.5) | %70 | %96 | %84 | %83,3 |
| **Genel** | **%84,1** | **%97,0** | **%91,3** | |

**Bulgular:**

- **Temperature arttıkça doğruluk düşüyor.** Etki en belirgin Qwen'de: T=0'da %92 olan
  doğruluk T=1.5'te %70'e iniyor. Yüksek temperature'daki hataların çoğu akıcı Türkçe ama
  yanlış ya da uydurma cevaplar (ör. "Türkiye'nin başkenti İstanbul", "Su 0°C'de kaynar").
- **Top-p ve top-k'yı düşürmek doğruluğu artırıyor.** Temel ayara (A3) göre top-p'yi 0.5'e
  (A4) veya top-k'yı 5'e (A5) indirmek, düşük olasılıklı kelimeleri eleyerek Qwen ve
  Ministral'de doğruluğu yükseltiyor.
- **Gemma 4 ayarlardan neredeyse hiç etkilenmiyor** (%96–%98). Model bu sorularda cevaptan
  çok emin olduğu için örnekleme ayarları seçimini değiştirmiyor.
- **En çok yanlış yapılan sorular Türkiye'ye özgü bilgiler:** Çalıkuşu'nun yazarı,
  Türkiye'nin en uzun nehri, TBMM'nin açılış tarihi, Nasreddin Hoca'nın ilçesi. Qwen'in
  coğrafya ve edebiyat kategorilerindeki doğruluğu %62,9'da kalıyor.
- **Sonuçlar tekrarlanabilir.** Sabit seed sayesinde aynı ayarla yapılan çalıştırmalar
  birebir aynı cevapları üretiyor.

## Testi Tekrar Çalıştırma

1. [Ollama](https://ollama.com)'yı kur ve modelleri indir:
   ```
   ollama pull qwen3.5:9b-q4_K_M
   ollama pull gemma4:12b-it-q4_K_M
   ollama pull ministral-3:8b-instruct-2512-q4_K_M
   ```
2. Gerekli Python kütüphanesini kur:
   ```
   pip install openpyxl
   ```
3. Betiği **boş bir klasörde** çalıştır (aynı adlı dolu bir Excel varsa betik onu doldurmaya devam eder):
   ```
   python sorgula.py --limit 5   # kısa deneme
   python sorgula.py             # tüm 2100 soru
   ```
   Betik `AB_Testi_Sonuclar.xlsx` dosyasını oluşturur ve cevapları yazar.
4. Excel'de "Sonuçlar" sayfasındaki "Puan (1/0)" sütununu doldur. "Özet" sayfası ve grafik
   kendiliğinden güncellenir.

Test ortamı: RTX 5070 Ti (laptop, 12 GB VRAM), Windows 11. Soru başına ortalama süre yaklaşık 2,2 saniye.
