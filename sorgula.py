"""
Yapay Zeka Modelleri A/B Testi
==============================
3 yerel model (Ollama) x 7 parametre ayarı x 100 soru = 2100 cevap.

Betik:
  1. Excel dosyası yoksa oluşturur (Özet, Sonuçlar, Ayarlar, Sorular sayfaları).
  2. Boş kalan her satır için soruyu ilgili modele, ilgili ayarla sorar.
  3. Cevabı ve süreyi Excel'e yazar. Puanlama (1/0) elle yapılır; Özet sayfası
     puanlara göre kendiliğinden hesaplanır.

Gereksinimler:
  - Ollama açık olmalı ve modeller indirilmiş olmalı:
        ollama pull qwen3.5:9b-q4_K_M
        ollama pull gemma4:12b-it-q4_K_M
        ollama pull ministral-3:8b-instruct-2512-q4_K_M
  - pip install openpyxl

Kullanım (Excel çalışırken KAPALI olmalı):
    python sorgula.py                  -> dosyayı oluşturur / boş satırları doldurur
    python sorgula.py --limit 5        -> deneme: sadece 5 satır
    python sorgula.py --model M1       -> sadece bir model (M1, M2, M3)
    python sorgula.py --ayar A3        -> sadece bir ayar (A1 ... A7)
    python sorgula.py --yeniden        -> dolu satırları da baştan sorar
    python sorgula.py --dosya X.xlsx   -> başka dosya adı

Ctrl+C ile durdurursan o ana kadarki cevaplar kaydedilir; tekrar çalıştırınca
kaldığı yerden devam eder.
"""
import argparse
import json
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

# ----------------------------------------------------------------- AYARLAR
KLASOR = Path(__file__).parent
EXCEL = KLASOR / "AB_Testi_Sonuclar.xlsx"
OLLAMA = "http://localhost:11434"

SISTEM_MESAJI = "Soruyu Türkçe, kısa ve net cevapla. Sadece cevabı yaz, açıklama yapma."
MAX_TOKEN = 256      # cevap uzunluğu sınırı (tüm modellerde aynı)
SEED = 42            # tekrarlanabilirlik için sabit tohum
KAYIT_ARALIGI = 25   # kaç cevapta bir Excel kaydedilsin

MODELLER = [
    ("M1", "qwen3.5:9b-q4_K_M"),
    ("M2", "gemma4:12b-it-q4_K_M"),
    ("M3", "ministral-3:8b-instruct-2512-q4_K_M"),
]

# (kod, açıklama, temperature, top_p, top_k)
AYARLAR = [
    ("A1", "En düşük çeşitlilik (deterministik)", 0, 1.0, 1),
    ("A2", "Düşük temperature", 0.2, 0.9, 40),
    ("A3", "Temel (referans)", 0.7, 0.9, 40),
    ("A4", "Düşük top-p", 0.7, 0.5, 40),
    ("A5", "Düşük top-k", 0.7, 0.9, 5),
    ("A6", "Yüksek temperature", 1.2, 0.9, 40),
    ("A7", "Çok yüksek temperature", 1.5, 0.95, 50),
]

# (kategori, soru, doğru cevap)
SORULAR = [
    ('Coğrafya', "Türkiye'nin başkenti neresidir?", 'Ankara'),
    ('Coğrafya', 'Dünyanın en büyük okyanusu hangisidir?', 'Büyük Okyanus (Pasifik)'),
    ('Coğrafya', "Türkiye'nin en yüksek dağı hangisidir?", 'Ağrı Dağı'),
    ('Coğrafya', "Japonya'nın başkenti neresidir?", 'Tokyo'),
    ('Coğrafya', "Türkiye'nin en büyük gölü hangisidir?", 'Van Gölü'),
    ('Coğrafya', 'Nil Nehri hangi kıtadadır?', 'Afrika'),
    ('Coğrafya', 'Yüzölçümü en büyük ülke hangisidir?', 'Rusya'),
    ('Coğrafya', "Avustralya'nın başkenti neresidir?", 'Kanberra'),
    ('Coğrafya', 'Türkiye kaç coğrafi bölgeye ayrılır?', '7'),
    ('Coğrafya', 'İstanbul Boğazı hangi iki denizi birbirine bağlar?', 'Karadeniz ve Marmara Denizi'),
    ('Coğrafya', 'Dünyanın en yüksek dağı hangisidir?', 'Everest'),
    ('Coğrafya', "Kanada'nın başkenti neresidir?", 'Ottava'),
    ('Coğrafya', 'Amazon yağmur ormanlarının büyük kısmı hangi ülkededir?', 'Brezilya'),
    ('Coğrafya', 'Tamamı Türkiye sınırları içinde kalan en uzun nehir hangisidir?', 'Kızılırmak'),
    ('Coğrafya', 'Dünyanın en büyük sıcak çölü hangisidir?', 'Sahra Çölü'),
    ('Tarih', 'Türkiye Cumhuriyeti hangi yıl ilan edildi?', '1923'),
    ('Tarih', 'İstanbul hangi yıl fethedildi?', '1453'),
    ('Tarih', "Türkiye Cumhuriyeti'nin ilk cumhurbaşkanı kimdir?", 'Mustafa Kemal Atatürk'),
    ('Tarih', 'Malazgirt Meydan Muharebesi hangi yıl yapıldı?', '1071'),
    ('Tarih', 'Birinci Dünya Savaşı hangi yıl başladı?', '1914'),
    ('Tarih', 'İkinci Dünya Savaşı hangi yıl sona erdi?', '1945'),
    ('Tarih', 'Fransız İhtilali hangi yıl başladı?', '1789'),
    ('Tarih', "Ay'a ayak basan ilk insan kimdir?", 'Neil Armstrong'),
    ('Tarih', "Osmanlı Devleti'nin kurucusu kimdir?", 'Osman Bey (Osman Gazi)'),
    ('Tarih', 'Berlin Duvarı hangi yıl yıkıldı?', '1989'),
    ('Tarih', "1492'de Amerika kıtasına ulaşan kâşif kimdir?", 'Kristof Kolomb'),
    ('Tarih', 'TBMM hangi tarihte açıldı?', '23 Nisan 1920'),
    ('Tarih', 'Kurtuluş Savaşı sonrası imzalanan ve yeni Türk devletini uluslararası alanda tanıtan barış antlaşması hangisidir?', 'Lozan Antlaşması'),
    ('Tarih', "Gize'deki en büyük piramit hangi firavun adına yapılmıştır?", 'Keops (Khufu)'),
    ('Tarih', "Hareketli harfli matbaayı Avrupa'da geliştiren kişi kimdir?", 'Johannes Gutenberg'),
    ('Bilim', 'Suyun kimyasal formülü nedir?', 'H2O'),
    ('Bilim', 'Deniz seviyesinde saf su kaç santigrat derecede kaynar?', '100 °C'),
    ('Bilim', 'Güneş sistemindeki en büyük gezegen hangisidir?', 'Jüpiter'),
    ('Bilim', "Güneş'e en yakın gezegen hangisidir?", 'Merkür'),
    ('Bilim', 'İnsan vücudundaki en büyük organ hangisidir?', 'Deri'),
    ('Bilim', 'Yetişkin bir insanda yirmilik dişler dahil kaç diş bulunur?', '32'),
    ('Bilim', 'Altının kimyasal sembolü nedir?', 'Au'),
    ('Bilim', 'Bitkilerin güneş ışığını kullanarak besin üretmesine ne denir?', 'Fotosentez'),
    ('Bilim', "Işığın boşluktaki hızı yaklaşık kaç km/s'dir?", 'Yaklaşık 300.000 km/s'),
    ('Bilim', 'Kanı vücuda pompalayan organ hangisidir?', 'Kalp'),
    ('Bilim', 'Periyodik tablodaki ilk element hangisidir?', 'Hidrojen'),
    ('Bilim', 'Evrensel çekim yasasını formüle eden bilim insanı kimdir?', 'Isaac Newton'),
    ('Bilim', 'Görelilik (izafiyet) teorisini geliştiren bilim insanı kimdir?', 'Albert Einstein'),
    ('Bilim', 'DNA kısaltmasının açılımı nedir?', 'Deoksiribonükleik asit'),
    ('Bilim', "'Kızıl Gezegen' olarak bilinen gezegen hangisidir?", 'Mars'),
    ('Bilim', 'Atmosferin (havanın) büyük kısmını oluşturan gaz hangisidir?', 'Azot (Nitrojen)'),
    ('Bilim', 'Yetişkin insan iskeletinde kaç kemik vardır?', '206'),
    ('Bilim', 'Penisilini kim keşfetmiştir?', 'Alexander Fleming'),
    ('Bilim', 'Elektrik akımının SI birimi nedir?', 'Amper'),
    ('Bilim', 'Ses hangi ortamda yayılamaz?', 'Boşlukta (vakum)'),
    ('Matematik', '12 × 12 kaçtır?', '144'),
    ('Matematik', 'Bir üçgenin iç açıları toplamı kaç derecedir?', '180'),
    ('Matematik', "144'ün karekökü kaçtır?", '12'),
    ('Matematik', 'Pi sayısının ilk iki ondalık basamağa kadar değeri nedir?', '3,14'),
    ('Matematik', 'Bir saatte kaç saniye vardır?', '3600'),
    ('Matematik', "1'den 10'a kadar olan tam sayıların toplamı kaçtır?", '55'),
    ('Matematik', 'Asal olan tek çift sayı hangisidir?', '2'),
    ('Matematik', "Bir kenarı 5 cm olan karenin alanı kaç cm²'dir?", '25'),
    ('Matematik', '2 üzeri 10 kaçtır?', '1024'),
    ('Matematik', "100'ün yüzde 15'i kaçtır?", '15'),
    ('Matematik', 'Dairenin çevresinin formülü nedir?', '2πr'),
    ('Matematik', 'Altıgenin kaç kenarı vardır?', '6'),
    ('Matematik', '0,5 sayısı kesir olarak nasıl yazılır?', '1/2'),
    ('Matematik', 'Romen rakamlarıyla 10 nasıl yazılır?', 'X'),
    ('Matematik', 'Artık yıl olmayan bir yılda kaç gün vardır?', '365'),
    ('Edebiyat ve Sanat', "'Çalıkuşu' romanının yazarı kimdir?", 'Reşat Nuri Güntekin'),
    ('Edebiyat ve Sanat', "İstiklal Marşı'nın şairi kimdir?", 'Mehmet Akif Ersoy'),
    ('Edebiyat ve Sanat', "'Romeo ve Juliet' kimin eseridir?", 'William Shakespeare'),
    ('Edebiyat ve Sanat', 'Mona Lisa tablosunu kim yapmıştır?', 'Leonardo da Vinci'),
    ('Edebiyat ve Sanat', "'Suç ve Ceza' romanının yazarı kimdir?", 'Fyodor Dostoyevski'),
    ('Edebiyat ve Sanat', "Nobel Edebiyat Ödülü'nü alan ilk Türk yazar kimdir?", 'Orhan Pamuk'),
    ('Edebiyat ve Sanat', "'Don Kişot' romanının yazarı kimdir?", 'Miguel de Cervantes'),
    ('Edebiyat ve Sanat', "'Sefiller' romanının yazarı kimdir?", 'Victor Hugo'),
    ('Edebiyat ve Sanat', "'Neşeye Övgü' bölümüyle bilinen 9. Senfoni'nin bestecisi kimdir?", 'Ludwig van Beethoven'),
    ('Edebiyat ve Sanat', 'Nasreddin Hoca hangi ilçeyle özdeşleşmiştir?', 'Akşehir'),
    ('Teknoloji', 'CPU kısaltmasının açılımı nedir?', 'Central Processing Unit (Merkezi İşlem Birimi)'),
    ('Teknoloji', "World Wide Web'i (WWW) kim icat etmiştir?", 'Tim Berners-Lee'),
    ('Teknoloji', "Microsoft'u Paul Allen ile birlikte kuran kişi kimdir?", 'Bill Gates'),
    ('Teknoloji', '1 bayt kaç bittir?', '8'),
    ('Teknoloji', 'HTML kısaltmasının açılımı nedir?', 'HyperText Markup Language'),
    ('Teknoloji', 'Python programlama dilini kim geliştirmiştir?', 'Guido van Rossum'),
    ('Teknoloji', 'Android işletim sistemi hangi şirkete aittir?', 'Google'),
    ('Teknoloji', 'RAM kısaltmasının açılımı nedir?', 'Random Access Memory'),
    ('Teknoloji', 'İkili sistemde 1 kilobayt kaç bayttır?', '1024'),
    ('Teknoloji', "iPhone'u üreten şirket hangisidir?", 'Apple'),
    ('Spor ve Genel Kültür', 'Bir futbol takımı sahada kaç oyuncuyla oynar?', '11'),
    ('Spor ve Genel Kültür', 'Olimpiyat bayrağında kaç halka vardır?', '5'),
    ('Spor ve Genel Kültür', 'Basketbolda bir takım sahada kaç oyuncuyla oynar?', '5'),
    ('Spor ve Genel Kültür', 'Satrançta oyun başında her oyuncunun kaç taşı vardır?', '16'),
    ('Spor ve Genel Kültür', 'FIFA Dünya Kupası kaç yılda bir düzenlenir?', '4'),
    ('Spor ve Genel Kültür', 'Türk bayrağında hangi semboller bulunur?', 'Ay ve yıldız'),
    ('Spor ve Genel Kültür', 'Gökkuşağında geleneksel olarak kaç renk sayılır?', '7'),
    ('Spor ve Genel Kültür', 'En az gün sayısına sahip ay hangisidir?', 'Şubat'),
    ('Spor ve Genel Kültür', "Türkiye'nin para birimi nedir?", 'Türk lirası'),
    ('Spor ve Genel Kültür', "Japonya'nın para birimi nedir?", 'Yen'),
    ('Spor ve Genel Kültür', 'Cumhuriyet Bayramı hangi tarihte kutlanır?', '29 Ekim'),
    ('Spor ve Genel Kültür', 'Bir maraton koşusu kaç kilometredir?', '42,195 km'),
    ('Spor ve Genel Kültür', 'Bir düzine kaç adettir?', '12'),
    ('Spor ve Genel Kültür', 'Ana dil olarak en çok konuşulan dil hangisidir?', 'Mandarin Çincesi'),
    ('Spor ve Genel Kültür', 'Dünya kendi ekseni etrafındaki bir turunu yaklaşık kaç saatte tamamlar?', '24 saat')]

# Sonuçlar sayfası sütunları (1 = A)
SUTUNLAR = ["Kayıt No", "Model Kodu", "Model Adı", "Ayar Kodu", "Temperature", "Top-p", "Top-k",
            "Soru No", "Kategori No", "Kategori", "Soru", "Doğru Cevap", "Model Cevabı",
            "Puan (1/0)", "Not", "Süre (sn)", "Sonuç"]
C_MODEL, C_AYAR, C_SORU_NO, C_CEVAP, C_SURE = 2, 4, 8, 13, 16

# ------------------------------------------------------- EXCEL OLUŞTURMA
F = "Arial"
YAZI, KALIN = Font(name=F), Font(name=F, bold=True)
BASLIK_YAZI = Font(name=F, bold=True, color="FFFFFF")
BASLIK_DOLGU = PatternFill("solid", fgColor="1F4E78")
GIRDI = PatternFill("solid", fgColor="FFF2CC")
_k = Side(style="thin", color="BFBFBF")
CERCEVE = Border(left=_k, right=_k, top=_k, bottom=_k)
SAR = Alignment(wrap_text=True, vertical="top")


def baslik(ws, satir, basliklar, sutun=1):
    for i, b in enumerate(basliklar, sutun):
        c = ws.cell(satir, i, b)
        c.font, c.fill, c.border = BASLIK_YAZI, BASLIK_DOLGU, CERCEVE
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")


def excel_olustur(yol):
    kategoriler = list(dict.fromkeys(k for k, _, _ in SORULAR))
    kat_no = {k: i + 1 for i, k in enumerate(kategoriler)}
    wb = Workbook()
    O = wb.active
    O.title = "Özet"

    # --- Sonuçlar: model > ayar > soru sırasıyla 2100 satır
    R = wb.create_sheet("Sonuçlar")
    baslik(R, 1, SUTUNLAR)
    n = 0
    for m_kod, m_ad in MODELLER:
        for a_kod, _, t, p, k in AYARLAR:
            for s_no, (kat, soru, cevap) in enumerate(SORULAR, 1):
                n += 1
                r = n + 1
                degerler = [n, m_kod, m_ad, a_kod, t, p, k, s_no, kat_no[kat], kat, soru, cevap,
                            None, None, None, None,
                            f'=IF(N{r}="","",IF(N{r}=1,"Doğru","Yanlış"))']
                for j, v in enumerate(degerler, 1):
                    R.cell(r, j, v).font = YAZI
                for j in (13, 14, 15):
                    R.cell(r, j).fill = GIRDI
                R.cell(r, 11).alignment = SAR
                R.cell(r, 13).alignment = SAR
    son = n + 1
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="1",
                        allow_blank=True, error="Sadece 1 (doğru) veya 0 (yanlış)", errorTitle="Geçersiz puan")
    R.add_data_validation(dv)
    dv.add(f"N2:N{son}")
    R.conditional_formatting.add(f"N2:N{son}", CellIsRule(operator="equal", formula=["1"],
                                 fill=PatternFill("solid", fgColor="C6EFCE")))
    R.conditional_formatting.add(f"N2:N{son}", CellIsRule(operator="equal", formula=["0"],
                                 fill=PatternFill("solid", fgColor="FFC7CE")))
    for j, w in enumerate([8, 8, 30, 8, 11, 8, 8, 8, 9, 18, 40, 22, 40, 9, 30, 9, 9], 1):
        R.column_dimensions[L(j)].width = w
    R.freeze_panes = "M2"
    R.auto_filter.ref = f"A1:Q{son}"

    # --- Ayarlar
    A = wb.create_sheet("Ayarlar")
    A["A1"] = "Modeller ve Parametre Ayarları"
    A["A1"].font = Font(name=F, bold=True, size=14)
    baslik(A, 3, ["Model Kodu", "Model Adı (Ollama etiketi)"])
    for i, satir in enumerate(MODELLER, 4):
        for j, v in enumerate(satir, 1):
            c = A.cell(i, j, v); c.font = YAZI; c.border = CERCEVE
    baslik(A, 9, ["Ayar Kodu", "Açıklama", "Temperature", "Top-p", "Top-k"])
    for i, satir in enumerate(AYARLAR, 10):
        for j, v in enumerate(satir, 1):
            c = A.cell(i, j, v); c.font = YAZI; c.border = CERCEVE
    notlar = [
        "Deney tasarımı: A3 temel ayardır. A1–A2 ve A6–A7 temperature'ın etkisini, A4 top-p'nin, "
        f"A5 top-k'nın etkisini ölçer. Her ayar {len(MODELLER)} modelde {len(SORULAR)} soruyla denenir "
        f"(toplam {n} cevap).",
        f"Ortak koşullar: aynı sistem mesajı ('Kısa ve net cevapla'), en fazla {MAX_TOKEN} token, "
        f"seed {SEED}, düşünme modu kapalı, Q4_K_M sıkıştırma, Ollama.",
        "Puanlama: uydurma bilgi içeren cevaplar, doğru anahtar kelimeyi içerse bile 0 sayılır. "
        "Yazım hataları ve yuvarlamalar kabul edilir. Gerekçeler 'Sonuçlar' sayfasındaki 'Not' sütununa yazılır.",
    ]
    for i, t in enumerate(notlar, 18):
        A.cell(i, 1, t).font = Font(name=F, italic=True)
    for c, w in zip("ABCDE", [10, 40, 12, 9, 9]):
        A.column_dimensions[c].width = w

    # --- Sorular
    S = wb.create_sheet("Sorular")
    baslik(S, 1, ["Soru No", "Kategori", "Soru", "Doğru Cevap"])
    for i, (kat, soru, cevap) in enumerate(SORULAR, 2):
        for j, v in enumerate((i - 1, kat, soru, cevap), 1):
            c = S.cell(i, j, v); c.font = YAZI; c.alignment = SAR; c.border = CERCEVE
    for c, w in zip("ABCD", [8, 20, 60, 30]):
        S.column_dimensions[c].width = w
    S.freeze_panes = "A2"

    # --- Özet (formüller yalnızca kod/numara eşleştirir; Türkçe metin karşılaştırması yok)
    aralik = lambda s: f"Sonuçlar!${s}$2:${s}${son}"
    B, D, N, I, P = aralik("B"), aralik("D"), aralik("N"), aralik("I"), aralik("P")
    O["A1"] = "Özet – Yapay Zeka Modelleri A/B Testi"
    O["A1"].font = Font(name=F, bold=True, size=14)
    O["A2"] = "Doğruluk = Puan (1/0) ortalaması. Puanları 'Sonuçlar' sayfasında girince bu sayfa kendiliğinden güncellenir."
    O["A2"].font = Font(name=F, italic=True)

    O.cell(4, 1, "1) Ayar bazında doğruluk oranı").font = KALIN
    baslik(O, 5, ["Ayar", "Açıklama", "Temp.", "Top-p", "Top-k"]
           + [f"{m} – {ad.split(':')[0]}" for m, ad in MODELLER] + ["Ortalama"])
    r0 = 6
    for i, ayar in enumerate(AYARLAR):
        r = r0 + i
        for j, v in enumerate(ayar, 1):
            O.cell(r, j, v).font = YAZI
        for j, (m, _) in enumerate(MODELLER, 6):
            c = O.cell(r, j, f'=IFERROR(SUMIFS({N},{B},"{m}",{D},$A{r})/COUNTIFS({B},"{m}",{D},$A{r},{N},">=0"),"")')
            c.number_format = "0.0%"; c.font = YAZI
        c = O.cell(r, 9, f'=IFERROR(AVERAGE(F{r}:H{r}),"")')
        c.number_format = "0.0%"; c.font = KALIN
        for j in range(1, 10):
            O.cell(r, j).border = CERCEVE
        O.cell(r, 11, f"{ayar[0]} (T={ayar[2]}, p={ayar[3]}, k={ayar[4]})").font = Font(name=F, size=8, color="808080")
    O.cell(5, 11, "Grafik etiketi").font = Font(name=F, size=8, color="808080")
    rs = r0 + len(AYARLAR) - 1

    r = rs + 3
    O.cell(r, 1, "2) Model bazında genel doğruluk").font = KALIN
    baslik(O, r + 1, ["Model Kodu", "Model Adı", "Doğruluk"])
    for i, (m, ad) in enumerate(MODELLER):
        rr = r + 2 + i
        O.cell(rr, 1, m).font = YAZI; O.cell(rr, 2, ad).font = YAZI
        c = O.cell(rr, 3, f'=IFERROR(SUMIFS({N},{B},$A{rr})/COUNTIFS({B},$A{rr},{N},">=0"),"")')
        c.number_format = "0.0%"; c.font = KALIN
        for j in range(1, 4):
            O.cell(rr, j).border = CERCEVE

    r = r + 7
    O.cell(r, 1, "3) Kategoriye göre doğruluk").font = KALIN
    baslik(O, r + 1, ["Kat. No", "Kategori"] + [m for m, _ in MODELLER])
    for i, kat in enumerate(kategoriler):
        rr = r + 2 + i
        O.cell(rr, 1, i + 1).font = YAZI; O.cell(rr, 2, kat).font = YAZI
        for j, (m, _) in enumerate(MODELLER, 3):
            c = O.cell(rr, j, f'=IFERROR(SUMIFS({N},{B},"{m}",{I},$A{rr})/COUNTIFS({B},"{m}",{I},$A{rr},{N},">=0"),"")')
            c.number_format = "0.0%"; c.font = YAZI
        for j in range(1, 6):
            O.cell(rr, j).border = CERCEVE

    r = r + 2 + len(kategoriler) + 2
    O.cell(r, 1, "4) Ortalama cevap süresi (sn)").font = KALIN
    baslik(O, r + 1, ["Ayar", "Açıklama"] + [m for m, _ in MODELLER])
    for i, ayar in enumerate(AYARLAR):
        rr = r + 2 + i
        O.cell(rr, 1, ayar[0]).font = YAZI; O.cell(rr, 2, ayar[1]).font = YAZI
        for j, (m, _) in enumerate(MODELLER, 3):
            c = O.cell(rr, j, f'=IFERROR(AVERAGEIFS({P},{B},"{m}",{D},$A{rr}),"")')
            c.number_format = "0.00"; c.font = YAZI
        for j in range(1, 6):
            O.cell(rr, j).border = CERCEVE
    for c, w in zip("ABCDEFGHIJK", [9, 36, 8, 8, 8, 16, 16, 16, 11, 2, 24]):
        O.column_dimensions[c].width = w
    O.row_dimensions[5].height = 30

    g = BarChart()
    g.type, g.grouping = "col", "clustered"
    g.title, g.y_axis.title = "Ayar bazında doğruluk oranı", "Doğruluk"
    g.y_axis.number_format = "0%"
    g.y_axis.scaling.min, g.y_axis.scaling.max = 0.5, 1.0
    g.add_data(Reference(O, min_col=6, max_col=5 + len(MODELLER), min_row=5, max_row=rs), titles_from_data=True)
    g.set_categories(Reference(O, min_col=11, min_row=r0, max_row=rs))
    g.roundedCorners = False
    g.legend.overlay = g.title.overlay = g.y_axis.title.overlay = False
    g.x_axis.delete = g.y_axis.delete = False
    g.height, g.width = 9, 26
    O.add_chart(g, "M4")

    wb.save(yol)
    print(f"Excel oluşturuldu: {yol.name} ({n} satır)")


# ------------------------------------------------------------- OLLAMA
def istek(yol, veri, zaman_asimi=300):
    req = urllib.request.Request(OLLAMA + yol, data=json.dumps(veri).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=zaman_asimi) as r:
        return json.loads(r.read().decode("utf-8"))


def temizle(metin):
    metin = re.sub(r"<think>.*?</think>", "", metin, flags=re.S)  # olası düşünme bloğu
    metin = ILLEGAL_CHARACTERS_RE.sub("", metin)                  # Excel'in kabul etmediği karakterler
    return metin.strip()


def sor(model, soru, t, p, k):
    veri = {
        "model": model,
        "messages": [{"role": "system", "content": SISTEM_MESAJI},
                     {"role": "user", "content": soru}],
        "stream": False,
        "think": False,  # düşünme modu kapalı
        "keep_alive": "10m",
        "options": {"temperature": t, "top_p": p, "top_k": int(k),
                    "seed": SEED, "num_predict": MAX_TOKEN},
    }
    try:
        cevap = istek("/api/chat", veri)
    except urllib.error.HTTPError as e:
        hata = e.read().decode("utf-8", "ignore")
        if "think" in hata.lower():  # model 'think' ayarını tanımıyorsa onsuz dene
            veri.pop("think")
            cevap = istek("/api/chat", veri)
        else:
            raise RuntimeError(f"Ollama hatası ({e.code}): {hata}")
    return temizle(cevap["message"]["content"])


def bellekten_cikar(model):
    try:
        istek("/api/generate", {"model": model, "keep_alive": 0}, 60)
    except Exception:
        pass


def kaydet(wb, yol):
    while True:
        try:
            wb.save(yol)
            return
        except PermissionError:
            input("Excel dosyası açık görünüyor. Kapatıp Enter'a bas...")


# --------------------------------------------------------------- ANA AKIŞ
def main():
    ap = argparse.ArgumentParser(description="Yapay zeka modelleri A/B testi")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--model", default="")
    ap.add_argument("--ayar", default="")
    ap.add_argument("--yeniden", action="store_true")
    ap.add_argument("--dosya", default="")
    a = ap.parse_args()
    yol = KLASOR / a.dosya if a.dosya else EXCEL

    if not yol.exists():
        excel_olustur(yol)
    try:
        urllib.request.urlopen(OLLAMA + "/api/tags", timeout=5)
    except Exception:
        sys.exit("Ollama'ya bağlanılamadı. Ollama açık mı? (görev çubuğunda simgesi olmalı)")

    model_adi = dict(MODELLER)
    ayar = {kod: (t, p, k) for kod, _, t, p, k in AYARLAR}
    wb = load_workbook(yol)  # formüller korunur
    R = wb["Sonuçlar"]

    isler = []
    for r in range(2, R.max_row + 1):
        m, k = R.cell(r, C_MODEL).value, R.cell(r, C_AYAR).value
        if not m or (a.model and m != a.model) or (a.ayar and k != a.ayar):
            continue
        if R.cell(r, C_CEVAP).value not in (None, "") and not a.yeniden:
            continue
        isler.append(r)
    if a.limit:
        isler = isler[:a.limit]
    print(f"{len(isler)} satır sorulacak.")

    onceki, sayac, basla = None, 0, time.time()
    try:
        for i, r in enumerate(isler, 1):
            m_kod, a_kod = R.cell(r, C_MODEL).value, R.cell(r, C_AYAR).value
            s_no = int(R.cell(r, C_SORU_NO).value)
            model = model_adi[m_kod]
            t, p, k = ayar[a_kod]
            if model != onceki:
                if onceki:
                    bellekten_cikar(onceki)
                print(f"\n=== {model} yükleniyor ===")
                onceki = model
            t0 = time.time()
            try:
                cevap = sor(model, SORULAR[s_no - 1][1], t, p, k)
            except Exception as e:
                print(f"  [satır {r}] HATA: {e}")
                continue
            sure = time.time() - t0
            R.cell(r, C_CEVAP).value = cevap
            R.cell(r, C_SURE).value = round(sure, 2)
            sayac += 1
            print(f"[{i}/{len(isler)}] {m_kod} {a_kod} S{s_no} ({sure:.1f} sn): {cevap[:70]!r}")
            if sayac % KAYIT_ARALIGI == 0:
                kaydet(wb, yol)
    except KeyboardInterrupt:
        print("\nDurduruldu, kaydediliyor...")
    finally:
        kaydet(wb, yol)
        if onceki:
            bellekten_cikar(onceki)
    print(f"\nBitti: {sayac} cevap yazıldı, {(time.time() - basla) / 60:.1f} dk sürdü.")
    print("Sonraki adım: 'Sonuçlar' sayfasında 'Puan (1/0)' sütununu doldur; Özet kendiliğinden hesaplanır.")


if __name__ == "__main__":
    main()
