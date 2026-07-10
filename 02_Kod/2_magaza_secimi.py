# -*- coding: utf-8 -*-
"""
LCW Mağaza Seçimi — Eşik Bazlı Maliyet Simülasyonu
Hangi mağazaların in-house, hangilerinin outsource olması gerektiğini analiz eder.
"""

import math
import pandas as pd
import numpy as np
from pathlib import Path

# ─── DOSYA YOLLARI ────────────────────────────────────────────────────────────
_PROJE_KOK  = Path(__file__).resolve().parent.parent
LOK_PATH    = _PROJE_KOK / "01_Veri" / "LCW_Lokasyonlar.xlsx"
MATRIX_PATH = _PROJE_KOK / "01_Veri" / "LCW_Mesafe_Sure_Matrisi.xlsx"

# ─── MALİYET PARAMETRELERİ ────────────────────────────────────────────────────
NUM_DEPOTS              = 6
STORE_DEMAND            = 150       # koli/mağaza
VEHICLE_CAPACITY        = 3000      # koli/araç
STORES_PER_VEHICLE      = 20        # araç başı mağaza varsayımı
ROTA_FAKTORU            = 2.5       # depo mesafesi → rota çarpanı
YAKIT_FIYATI_TL         = 43.0
YAKIT_TUKETIM_LT_KM     = 0.35
ARAC_GUNLUK_SABIT_TL    = 2500
SURUCU_GUNLUK_TL        = 1500
OUTSOURCE_KOLI_FIYAT_TL = 3.75
TOPLAM_MAGAZA_SAYISI    = 533       # tüm LCW mağazaları
PILOT_MAGAZA_SAYISI     = 204       # bu verisetindeki mağaza sayısı (6-209 arası)

ESIKLER = [100, 150, 200, 250, 300, 400, 500, 750, 1000]

SEP  = "─" * 80
SEP2 = "═" * 80


# ─── 1. VERİ YÜKLEME ─────────────────────────────────────────────────────────

print()
print(SEP2)
print("  LCW MAĞAZA SEÇİMİ — EŞİK BAZLI MALİYET SİMÜLASYONU (Gerçek Depolar)")
print(SEP2)

print("\n[1/4] Lokasyon verisi okunuyor...")
df_lok = pd.read_excel(LOK_PATH)
print(f"  Toplam satır: {len(df_lok)}")
print(f"  İlk 6 (depolar): {list(df_lok['Mağaza Adı'][:6])}")

print("\n[2/4] Mesafe matrisi okunuyor...")
df_dist = pd.read_excel(MATRIX_PATH, sheet_name="Mesafe (km)", index_col=0)
print(f"  Matris boyutu: {df_dist.shape[0]} × {df_dist.shape[1]}")

dist_matrix = df_dist.fillna(0).values  # 210×210

# ─── 2. EN YAKIN DEPO MESAFESİ ───────────────────────────────────────────────

print("\n[3/4] Her mağaza için en yakın depo mesafesi hesaplanıyor...")

depo_idx   = list(range(NUM_DEPOTS))                        # [0..5]
magaza_idx = list(range(NUM_DEPOTS, dist_matrix.shape[0]))  # [6..209]

depo_adlari = list(df_lok["Mağaza Adı"][:NUM_DEPOTS])

en_yakin_mesafe = []
en_yakin_depo   = []

for m in magaza_idx:
    dists   = [dist_matrix[m][d] for d in depo_idx]
    min_d   = min(dists)
    min_dep = depo_idx[dists.index(min_d)]
    en_yakin_mesafe.append(min_d)
    en_yakin_depo.append(min_dep)

en_yakin_mesafe = np.array(en_yakin_mesafe)
en_yakin_depo   = np.array(en_yakin_depo)

print(f"  Mağaza sayısı : {len(magaza_idx)}")
print(f"  Mesafe — Min  : {en_yakin_mesafe.min():.1f} km")
print(f"  Mesafe — Maks : {en_yakin_mesafe.max():.1f} km")
print(f"  Mesafe — Ort. : {en_yakin_mesafe.mean():.1f} km")

# ─── 3. MESAFE BANT DAĞILIMI ─────────────────────────────────────────────────

print("\n[4/4] Sonuçlar hesaplanıyor...")

bantlar = [
    ("0 – 100 km",    0,   100),
    ("100 – 200 km",  100, 200),
    ("200 – 300 km",  200, 300),
    ("300 – 500 km",  300, 500),
    ("500 – 750 km",  500, 750),
    ("750 – 1000 km", 750, 1000),
    ("1000+ km",      1000, 99999),
]

print()
print(SEP2)
print("  BÖLÜM 1 — EN YAKIN DEPO MESAFESİ DAĞILIMI (204 Mağaza)")
print(SEP2)
print(f"  {'Mesafe Bandı':<18} {'Mağaza':>7} {'Yüzde':>8} {'Kümülatif':>11}")
print(f"  {SEP}")

kumulatif = 0
for etiket, alt, ust in bantlar:
    if ust == 99999:
        sayisi = int((en_yakin_mesafe >= alt).sum())
    else:
        sayisi = int(((en_yakin_mesafe >= alt) & (en_yakin_mesafe < ust)).sum())
    yuzde     = sayisi / PILOT_MAGAZA_SAYISI * 100
    kumulatif += sayisi
    bar        = "█" * int(yuzde / 2)
    print(f"  {etiket:<18} {sayisi:>7}  {yuzde:>6.1f}%  {kumulatif:>10}  {bar}")

# ─── 4. DEPO BAŞI YOĞUNLUK (TÜM MAĞAZALAR) ───────────────────────────────────

print()
print(SEP2)
print("  BÖLÜM 2 — DEPO BAŞI MAĞAZA YOĞUNLUĞU (Tüm 204 Mağaza, Sınırsız Eşik)")
print(SEP2)
print(f"  {'Depo':<28} {'Mağaza':>7} {'Ort. (km)':>11} {'Min (km)':>10} {'Maks (km)':>11}")
print(f"  {SEP}")

for d in depo_idx:
    mask      = en_yakin_depo == d
    mesafeler = en_yakin_mesafe[mask]
    if len(mesafeler) > 0:
        print(f"  {depo_adlari[d]:<28} {len(mesafeler):>7}  "
              f"{np.mean(mesafeler):>9.1f}  {mesafeler.min():>9.1f}  {mesafeler.max():>10.1f}")
    else:
        print(f"  {depo_adlari[d]:<28} {'0':>7}  {'–':>9}  {'–':>9}  {'–':>10}")

# ─── 5. MALİYET SİMÜLASYONU ──────────────────────────────────────────────────

# DUZELTME (kod incelemesi sonrasi): referans "tam outsource" maliyeti
# TOPLAM_MAGAZA_SAYISI (533, tum LCW agi) yerine PILOT_MAGAZA_SAYISI (204)
# uzerinden hesaplanir. Aksi halde hibrit maliyet 204 magazalik pilot
# kapsamindan, referans ise 533 magazalik tum ag kapsamindan gelir; bu da
# Bolum 3.6.4'te tanimlanan "kapsam uyumu" ilkesini ihlal eder ve tasarruf
# oranini olduğundan fazla gosterir (orn. 200 km esiginde %71 yerine %25).
tam_outsource = PILOT_MAGAZA_SAYISI * STORE_DEMAND * OUTSOURCE_KOLI_FIYAT_TL

print()
print(SEP2)
print("  BÖLÜM 3 — EŞİK BAZLI MALİYET SİMÜLASYONU")
print(f"  Referans: Tam Outsource ({PILOT_MAGAZA_SAYISI} mağaza × {STORE_DEMAND} koli × "
      f"{OUTSOURCE_KOLI_FIYAT_TL} TL) = {tam_outsource:>10,.0f} TL/gün  [kapsam: pilot 204]")
print(SEP2)

baslik = (
    f"  {'Eşik':>7}  {'İH Mağ':>7}  {'Out Mağ':>7}  {'Araç':>5}  "
    f"{'İH Maliyet':>12}  {'Out Kısım':>11}  {'Hibrit':>12}  "
    f"{'Tasarruf':>12}  {'Oran':>7}"
)
print(baslik)
print(f"  {SEP}")

sim_sonuclar = []

for esik in ESIKLER:
    inhouse_mask  = en_yakin_mesafe <= esik
    inhouse_count = int(inhouse_mask.sum())
    out_count     = PILOT_MAGAZA_SAYISI - inhouse_count

    if inhouse_count == 0:
        inhouse_maliyet = 0
        arac_sayisi     = 0
    else:
        ort_depo_mesafe = float(en_yakin_mesafe[inhouse_mask].mean())
        ort_rota        = ort_depo_mesafe * ROTA_FAKTORU
        toplam_km       = ort_rota * (inhouse_count / STORES_PER_VEHICLE)
        arac_sayisi     = math.ceil(inhouse_count * STORE_DEMAND / VEHICLE_CAPACITY)
        yakit           = toplam_km * YAKIT_TUKETIM_LT_KM * YAKIT_FIYATI_TL
        sabit           = arac_sayisi * ARAC_GUNLUK_SABIT_TL
        personel        = arac_sayisi * SURUCU_GUNLUK_TL
        inhouse_maliyet = yakit + sabit + personel

    out_kisim     = out_count * STORE_DEMAND * OUTSOURCE_KOLI_FIYAT_TL
    hibrit_toplam = inhouse_maliyet + out_kisim
    tasarruf      = tam_outsource - hibrit_toplam
    tasarruf_oran = (tasarruf / tam_outsource) * 100

    sim_sonuclar.append({
        "esik":          esik,
        "inhouse_count": inhouse_count,
        "out_count":     out_count,
        "arac":          arac_sayisi,
        "inhouse_mal":   inhouse_maliyet,
        "out_kisim":     out_kisim,
        "hibrit":        hibrit_toplam,
        "tasarruf":      tasarruf,
        "oran":          tasarruf_oran,
    })

    print(
        f"  {esik:>5} km  {inhouse_count:>7}  {out_count:>7}  {arac_sayisi:>5}  "
        f"{inhouse_maliyet:>12,.0f}  {out_kisim:>11,.0f}  {hibrit_toplam:>12,.0f}  "
        f"{tasarruf:>+12,.0f}  {tasarruf_oran:>+6.1f}%"
    )

# ─── 6. MAKSİMUM TASARRUF & ÖZET ─────────────────────────────────────────────

en_iyi  = max(sim_sonuclar, key=lambda x: x["tasarruf"])
pozitif = [s for s in sim_sonuclar if s["tasarruf"] > 0]

print()
print(SEP2)
print("  BÖLÜM 4 — DEPO BAŞI DAĞILIM (≤ Maksimum Tasarruf Eşiği)")
print(SEP2)

esik_best = en_iyi["esik"]
inhouse_best_mask = en_yakin_mesafe <= esik_best

print(f"  Eşik: {esik_best} km  |  In-house mağaza: {en_iyi['inhouse_count']}  |  Outsource: {en_iyi['out_count']}")
print()
print(f"  {'Depo':<28} {'Mağaza (≤eşik)':>15} {'Ort. (km)':>11} {'Maks (km)':>11}")
print(f"  {SEP}")

for d in depo_idx:
    mask_d    = (en_yakin_depo == d) & inhouse_best_mask
    mesafeler = en_yakin_mesafe[mask_d]
    if len(mesafeler) > 0:
        print(f"  {depo_adlari[d]:<28} {len(mesafeler):>15}  "
              f"{np.mean(mesafeler):>9.1f}  {mesafeler.max():>10.1f}")
    else:
        print(f"  {depo_adlari[d]:<28} {'0':>15}  {'–':>9}  {'–':>10}")

print()
print(SEP2)
print("  BÖLÜM 5 — ÖZET VE KARAR ÖNERİSİ")
print(SEP2)
print(f"  Tam Outsource Günlük Maliyet : {tam_outsource:>12,.0f} TL/gün")
print()
print(f"  ★ Maksimum tasarruf eşiği    : {en_iyi['esik']} km")
print(f"    In-house mağaza sayısı      : {en_iyi['inhouse_count']} / {PILOT_MAGAZA_SAYISI}")
print(f"    Outsource mağaza sayısı     : {en_iyi['out_count']} / {PILOT_MAGAZA_SAYISI}")
print(f"    Gereken araç sayısı         : {en_iyi['arac']}")
print(f"    In-house günlük maliyet     : {en_iyi['inhouse_mal']:>12,.0f} TL/gün")
print(f"    Outsource kısım maliyeti    : {en_iyi['out_kisim']:>12,.0f} TL/gün")
print(f"    Hibrit günlük toplam        : {en_iyi['hibrit']:>12,.0f} TL/gün")
print(f"    Günlük tasarruf             : {en_iyi['tasarruf']:>+12,.0f} TL/gün")
print(f"    Yıllık tasarruf (300 gün)   : {en_iyi['tasarruf']*300:>+12,.0f} TL/yıl")
print(f"    Tasarruf oranı              : %{en_iyi['oran']:>+.1f}")
print()

if pozitif:
    ilk = min(pozitif, key=lambda x: x["esik"])
    print(f"  İlk pozitif tasarruf eşiği   : {ilk['esik']} km  "
          f"(+{ilk['tasarruf']:,.0f} TL/gün, %{ilk['oran']:.1f})")
else:
    print("  [BİLGİ] Test edilen hiçbir eşikte pozitif tasarruf elde edilemedi.")
    print("  Mevcut parametrelerle tam outsource (3,75 TL/koli) her eşikten daha ucuz.")
    print()
    print("  Pozitif tasarruf için gereken başabaş outsource birim fiyatı:")
    for s in sim_sonuclar:
        if s["inhouse_count"] > 0:
            x = s["inhouse_mal"] / (
                (TOPLAM_MAGAZA_SAYISI - s["out_count"]) * STORE_DEMAND
            )
            print(f"    Eşik {s['esik']:>4} km → başabaş 3PL fiyatı: {x:.2f} TL/koli")

print()
print(SEP2)
print("  Analiz tamamlandı.")
print(SEP2)
print()
