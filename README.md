# LC Waikiki Depo–Mağaza Dağıtım Süreci: Çok Depolu Araç Rotalama ve Maliyet Analizi

İstanbul Üniversitesi-Cerrahpaşa · Mühendislik Fakültesi · Endüstri Mühendisliği Bölümü
Bitirme Projesi

---

## Projenin Amacı

Bu çalışma, LC Waikiki'nin depo–mağaza dağıtım sürecini seçilmiş bir **pilot mağaza ağı**
(6 mevcut depo + 204 mağaza) üzerinde ele almaktadır. Amaç, dış kaynaklı (outsource)
dağıtım yapısı ile şirket kontrollü (in-house) dağıtımın birlikte değerlendirildiği
**hibrit bir dağıtım senaryosunu** rota planlama ve maliyet analizi açısından incelemektir.

Mağazalar, en yakın mevcut depoya **200 km mesafe eşiği** uygulanarak in-house aday ve
outsource kalan mağazalar olarak sınıflandırılmış; in-house aday mağazalar için **Çok Depolu
Kapasite Kısıtlı Araç Rotalama Problemi (MDCVRP)** çerçevesinde Python/OR-Tools ile rota
çözümü üretilmiştir. Matematiksel formülasyonun kısıt yapısı, GAMS ortamında küçük ölçekli
(2 depo, 8 mağaza) bir örnek üzerinde kavramsal olarak sınanmıştır — bu, 204 mağazalık
pilot çözümün sayısal bir doğrulaması değildir (bkz. aşağıdaki "GAMS karşılaştırması" notu).

---

## Temel Sonuçlar (Pilot 204 Mağaza Bazında)

| Gösterge | Değer |
|---|---|
| In-house aday mağaza | 99 |
| Outsource kalan mağaza | 105 |
| Toplam rota mesafesi | 1.086,4 km |
| Aktif araç sayısı | 6 |
| Ortalama araç doluluğu | %82,5 |
| Maksimum rota süresi | 536 dk (540 dk sınırının altında) |
| Pilot hibrit günlük maliyet | 99.414 TL |
| Pilot tam outsource maliyet | 114.750 TL |
| Günlük tasarruf | 15.336 TL (%13,4) |
| Yatırım geri dönüş süresi | ≈ 6,5 yıl |

**GAMS karşılaştırması (küçük ölçekli, gösterim amaçlı):** 2 depo/8 mağazalık bir "nano"
örnekte GAMS/CPLEX kesin çözümü (85,40 km) ile OR-Tools sezgisel çözümü (87,75 km) arasındaki
rota mesafesi farkı **%2,8**'dir. **Önemli çekince:** GAMS modeli toplam *parasal* maliyeti
(mesafe + araç sabit gideri + yatırım maliyeti) minimize ederken, OR-Tools çözümü yalnızca
toplam *mesafeyi* minimize etmektedir; iki çözücü farklı amaç fonksiyonlarını optimize
etmektedir. Bu nedenle %2,8'lik yakınlık, OR-Tools sezgiselinin GAMS optimaline istatistiksel
bir "doğrulaması" olarak sunulamaz; yalnızca küçük ölçekte iki farklı yaklaşımın rastlantısal
olarak yakın rota yapıları ürettiğini gösterir. Ayrıca bu örnek 204 mağazalık gerçek pilot veri
setiyle bağlantılı değildir ve GAMS'ın x.L çözüm değerleri elle (manuel) Python'a girilmiştir.
Ayrıntı için `03_GAMS_Model/gams_ortools_karsilastirma.py` dosyasındaki kod içi not ve
`06_Tez/LC Waikiki Bitirme - son hal.docx` EK-8'e bakınız.

> Not: Bu çalışma bir pilot senaryo düzeyinde karar destek analizidir; kesin bir yatırım
> kararı değildir. Geri dönüş süresinin firma beklentisinin (3 yıl) üzerinde kalması nedeniyle
> sonuçlar temkinli yorumlanmalıdır.

---

## Klasör Yapısı

```
LCWaikiki-MDCVRP-Bitirme/
│
├── README.md                    ← bu dosya
│
├── 01_Veri/                     Girdi verileri
│   ├── LCW_Lokasyonlar.xlsx           (210 lokasyon: 6 depo + 204 mağaza)
│   └── LCW_Mesafe_Sure_Matrisi.xlsx   (210×210 OSRM mesafe/süre matrisi)
│
├── 02_Kod/                      Python kaynak kodları (çalışma sırasıyla)
│   ├── 1_mesafe_matrisi.py            OSRM ile mesafe/süre matrisi üretimi
│   ├── 2_magaza_secimi.py           200 km eşiği, in-house/outsource ayrımı
│   └── 3_mdcvrp_cozum.py            OR-Tools MDCVRP çözümü + maliyet analizi
│
├── 03_GAMS_Model/               Matematiksel model ve doğrulama
│   ├── LCW_MDCVRP.gms                MDCVRP formülasyonu (GAMS)
│   └── gams_ortools_karsilastirma.py GAMS–OR-Tools doğrulama karşılaştırması
│
├── 04_Sonuclar/                 Nihai çıktı
│   └── SONUC_rota_maliyet.xlsx        Rota, maliyet ve duyarlılık sonuçları
│
├── 05_Arsiv/                    Çalışmanın gelişim süreci (nihai olmayan)
│   ├── eski_excel_ciktilari/         Ara Excel çıktıları
│   ├── eski_kodlar/                  Önceki kod sürümleri
│   ├── mvp_testleri/                 İlk prototipler
│   ├── yedekler/                     Güvenlik kopyaları
│   └── ARSIV_README.md
│
└── 06_Tez/                      Nihai tez dosyası
    └── LC Waikiki Bitirme - son hal.docx
```

---

## Yöntem (Çalışma Akışı)

1. **Veri hazırlama** — 210 lokasyonun enlem/boylam koordinatları (`01_Veri`).
2. **Mesafe/süre matrisi** — OSRM Table API ile 210×210 karayolu matrisi (`1_mesafe_matrisi.py`).
3. **Mesafe eşiği** — Her mağaza en yakın depoya göre sınıflandırılır; ≤200 km in-house aday (`2_magaza_secimi.py`).
4. **Rota çözümü** — In-house aday mağazalar için OR-Tools (SAVINGS + Guided Local Search) ile MDCVRP (`3_mdcvrp_cozum.py`).
5. **Maliyet analizi** — In-house, outsource ve hibrit senaryolar pilot 204 mağaza bazında karşılaştırılır.
6. **Kavramsal sınama** — GAMS ile küçük (2 depo, 8 mağaza) bir örnekte kısıt yapısı test edilir; OR-Tools ile rota mesafesi %2,8 yakın çıkar, ancak farklı amaç fonksiyonları (parasal vs. mesafe) nedeniyle bu bir "doğrulama" değildir (`03_GAMS_Model`).
7. **Duyarlılık analizi** — Koli fiyatı, yakıt fiyatı, talep vb. parametreler için senaryolar.

---

## Kurulum ve Çalıştırma

Gereksinimler: Python 3.10+, ve şu paketler:

```
pip install pandas openpyxl numpy ortools requests
```

Kodlar proje kök klasörüne göre **göreli yol** kullanır; klasör yapısı korunduğu sürece
herhangi bir makinede çalışır. Çalışma sırası:

```
python 02_Kod/1_mesafe_matrisi.py      # (matris zaten 01_Veri'de mevcut, tekrar üretmek isteğe bağlı)
python 02_Kod/2_magaza_secimi.py       # in-house/outsource sınıflandırması
python 02_Kod/3_mdcvrp_cozum.py        # rota + maliyet → 04_Sonuclar/SONUC_rota_maliyet.xlsx
```

GAMS modeli için GAMS 53+ gereklidir (`03_GAMS_Model/LCW_MDCVRP.gms`).

---

## Kullanılan Yöntem ve Araçlar

Çok Depolu Kapasite Kısıtlı Araç Rotalama Problemi (MDCVRP) · Google OR-Tools Routing Solver
(SAVINGS + Guided Local Search) · GAMS/CPLEX (doğrulama) · OSRM (mesafe/süre matrisi) ·
Python (pandas, openpyxl, numpy).

---

## Kapsam ve Sınırlılıklar

Çalışma pilot mağaza ağıyla sınırlıdır. Talep sabit, filo homojen varsayılmış; gerçek zamanlı
trafik, AVM teslimat zaman pencereleri ve araç sahiplik maliyetlerinin (bakım, sigorta,
amortisman) tamamı modele dahil edilmemiştir. Sonuçlar bu nedenle bir karar destek çıktısı
olarak değerlendirilmelidir.
