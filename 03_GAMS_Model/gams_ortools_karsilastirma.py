# -*- coding: utf-8 -*-
"""
GAMS (nano model, 2 depo + 8 mağaza) vs OR-Tools karşılaştırması.
Aynı mesafe matrisi, aynı talep yapısı, aynı kapasite kısıtları.

ONEMLI METODOLOJIK NOT (kod incelemesi sonrasi eklendi):
GAMS modeli (LCW_MDCVRP.gms) toplam PARASAL maliyeti minimize eder
(mesafe/yakit + arac sabit gideri F_cost + yatirim maliyeti W_cost).
Bu script'teki OR-Tools cozumu ise SADECE MESAFEYI minimize eder; sabit/
yatirim maliyeti terimi yoktur. Dolayisiyla iki cozucu FARKLI amac
fonksiyonlarini optimize etmektedir - bu nedenle asagidaki kiyaslama
yalnizca "ayni rota yapisina ne kadar yakin sonuc buldular" sorusuna
(mesafe bazinda) cevap verir; GAMS'in tam parasal Z degeriyle dogrudan
kiyaslanamaz. Bu ornekte arac sayisi da (GAMS: 1, OR-Tools: asagida
hesaplanir) tesadufen ortusmustur; sabit maliyet terimi olmadan OR-Tools
prensipte daha fazla arac kullanmayi da secebilirdi. Bu kiyaslamayi genis
olcekli (204 magazali) problem icin tekrarlarken bu farki goz onunde
bulundurun; asagidaki kod artik arac sayisi uyusmazligini da ayrica
uyari olarak yazdirir.
GAMS x.L cozumu (gams_rota_hesapla icindeki gams_kenarlar) GAMS/CPLEX
ciktisindan elle (manuel) Python'a girilmistir; buyuk olcekli/tekrarlanan
dogrulamalarda GDX dosyasinin programatik okunmasi (orn. gdxpds veya
GAMS Python API) onerilir.
"""
from ortools.constraint_solver import routing_enums_pb2, pywrapcp
import math

# ─── Mesafe matrisi (LCW_GAMS_loader.gms'den, km) ───────────────────────────
# Satır/Sütun sırası: s001 s002 s003 s004 s005 s006 s007 s008 s009 s010
DIST_KM = [
    [0,     9.38,  7.77,  2.03,  18.39, 33.73, 13.90, 2.94,  18.39, 5.88 ],  # s001
    [9.09,  0,     3.89,  7.50,  12.54, 23.22, 19.94, 6.74,  12.54, 11.10],  # s002
    [7.00,  3.70,  0,     5.41,  12.71, 28.05, 12.87, 4.65,  12.71, 8.14 ],  # s003
    [2.55,  7.33,  5.71,  0,     16.33, 31.67, 12.75, 2.11,  16.33, 4.73 ],  # s004
    [17.44, 12.94, 12.23, 15.85, 0,     20.17, 23.57, 15.08, 0.00,  19.45],  # s005
    [34.16, 23.31, 28.95, 32.57, 20.51, 0,     30.83, 31.80, 20.51, 33.31],  # s006
    [14.47, 14.16, 12.55, 13.30, 23.17, 31.04, 0,     12.39, 23.17, 8.23 ],  # s007
    [2.36,  7.10,  5.48,  2.14,  16.11, 31.45, 12.38, 0,     16.11, 4.36 ],  # s008
    [17.44, 12.94, 12.23, 15.85, 0.00,  20.17, 23.57, 15.08, 0,     19.45],  # s009
    [7.08,  11.54, 8.40,  5.91,  20.55, 33.57, 7.85,  5.00,  20.55, 0    ],  # s010
]

ISIMLER = ["s001(Depo-Panelvan)","s002(Depo-Kamyon)",
           "s003","s004","s005","s006","s007","s008","s009","s010"]

# ─── GAMS parametreleriyle aynı talep yapısı ────────────────────────────────
# q(I): s003..s010 için ORD(I) mod 3 kuralı
# s003(3 mod 3=0)→80, s004(4 mod 3=1)→150, s005(5 mod 3=2)→320
# s006(6 mod 3=0)→80, s007(7 mod 3=1)→150, s008(8 mod 3=2)→320
# s009(9 mod 3=0)→80, s010(10 mod 3=1)→150
TALEPLER = [0, 0, 80, 150, 320, 80, 150, 320, 80, 150]   # 10 düğüm

# Araç kapasiteleri: k1=Panelvan (1000), k2=Kamyon (2000)
KAPASITELER = [1000, 2000]
NUM_ARAC    = 2
STARTS      = [0, 1]   # k1: s001, k2: s002
ENDS        = [0, 1]

# OR-Tools metre cinsinden integer matris
DIST_M = [[int(d * 1000) for d in satir] for satir in DIST_KM]


def gams_rota_hesapla():
    """GAMS sonucundan (x.L değerleri) rotayı ve mesafeyi hesapla."""
    # x.L değerleri → k2 (Kamyon) kullanıyor
    # s001→s004, s004→s010, s010→s007, s007→s006,
    # s006→s009, s009→s005, s005→s003, s003→s008, s008→s001
    gams_kenarlar = [
        ("s001","s004"), ("s004","s010"), ("s010","s007"),
        ("s007","s006"), ("s006","s009"), ("s009","s005"),
        ("s005","s003"), ("s003","s008"), ("s008","s001"),
    ]
    isim_idx = {isim.split("(")[0]: i for i, isim in enumerate(ISIMLER)}
    idx = {f"s{i+1:03d}": i for i in range(10)}

    toplam = 0.0
    rota = []
    for (a, b) in gams_kenarlar:
        ia, ib = idx[a], idx[b]
        toplam += DIST_KM[ia][ib]
        if not rota:
            rota.append(a)
        rota.append(b)
    return rota, round(toplam, 2)


def ortools_coz():
    """OR-Tools ile aynı küçük MDCVRP'yi çöz."""
    manager = pywrapcp.RoutingIndexManager(
        len(DIST_M), NUM_ARAC, STARTS, ENDS
    )
    routing = pywrapcp.RoutingModel(manager)

    def mesafe_cb(from_i, to_i):
        return DIST_M[manager.IndexToNode(from_i)][manager.IndexToNode(to_i)]

    transit_cb = routing.RegisterTransitCallback(mesafe_cb)
    routing.SetArcCostEvaluatorOfAllVehicles(transit_cb)

    def talep_cb(from_i):
        return TALEPLER[manager.IndexToNode(from_i)]

    talep_idx = routing.RegisterUnaryTransitCallback(talep_cb)
    routing.AddDimensionWithVehicleCapacity(
        talep_idx, 0, KAPASITELER, True, "Kapasite"
    )

    params = pywrapcp.DefaultRoutingSearchParameters()
    params.first_solution_strategy = (
        routing_enums_pb2.FirstSolutionStrategy.SAVINGS
    )
    params.local_search_metaheuristic = (
        routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
    )
    params.time_limit.seconds = 30

    cozum = routing.SolveWithParameters(params)
    if not cozum:
        return None, None, None

    rotalar = []
    toplam_m = 0
    for v in range(NUM_ARAC):
        idx_cur = routing.Start(v)
        rota_v  = []
        while not routing.IsEnd(idx_cur):
            node = manager.IndexToNode(idx_cur)
            rota_v.append(f"s{node+1:03d}")
            idx_nxt = cozum.Value(routing.NextVar(idx_cur))
            toplam_m += DIST_M[manager.IndexToNode(idx_cur)][manager.IndexToNode(idx_nxt)]
            idx_cur = idx_nxt
        rota_v.append(f"s{manager.IndexToNode(idx_cur)+1:03d}")
        if len(rota_v) > 2:   # boş rota gösterme
            rotalar.append((v, rota_v))

    toplam_km = round(toplam_m / 1000, 2)
    obj_m = cozum.ObjectiveValue()
    return rotalar, toplam_km, obj_m


# ─── Çalıştır ────────────────────────────────────────────────────────────────
print("=" * 65)
print("  GAMS ROTA SONUCU (x.L değerlerinden)")
print("=" * 65)
gams_rota, gams_km = gams_rota_hesapla()
print(f"  Kullanılan araç : k2 (Kamyon, kapasite 2000)")
print(f"  Çıkış deposu    : s001 (Istanbul - Fatih)")
print(f"  Rota            : {' → '.join(gams_rota)}")
print(f"  Toplam mesafe   : {gams_km} km")
toplam_talep = sum(TALEPLER[2:])
print(f"  Toplam yük      : {toplam_talep} / 2000 koli")
print(f"  GAMS Z (obj)    : 8351.52")
print(f"  Bileşenler      : {gams_km} km × 40 × 0.22 = {round(gams_km*8.8,2)} + 1600 (depo sabit) + 6000 (Kamyon W_cost) = {round(gams_km*8.8+7600,2)}")

print()
print("=" * 65)
print("  OR-TOOLS ROTA SONUCU")
print("=" * 65)
rotalar, ort_km, obj_m = ortools_coz()
if rotalar is None:
    print("  Çözüm bulunamadı!")
else:
    ort_toplam_km = 0
    for v, r in rotalar:
        arac_ismi = "k1 (Panelvan)" if v == 0 else "k2 (Kamyon)"
        print(f"  Araç {v+1} ({arac_ismi}):")
        print(f"    Rota    : {' → '.join(r)}")
        dugumler = [int(x[1:]) - 1 for x in r]
        mesafe_v = sum(DIST_KM[dugumler[i]][dugumler[i+1]] for i in range(len(dugumler)-1))
        mesafe_v = round(mesafe_v, 2)
        yuk_v    = sum(TALEPLER[int(x[1:])-1] for x in r if x not in ("s001","s002"))
        print(f"    Mesafe  : {mesafe_v} km")
        print(f"    Yük     : {yuk_v} / {KAPASITELER[v]} koli")
        ort_toplam_km += mesafe_v
    ort_toplam_km = round(ort_toplam_km, 2)
    print(f"\n  OR-Tools toplam mesafe: {ort_toplam_km} km")

print()
print("=" * 65)
print("  KARŞILAŞTIRMA TABLOSU")
print("=" * 65)
if rotalar is not None:
    fark_km  = round(ort_toplam_km - gams_km, 2)
    fark_pct = round(abs(fark_km) / gams_km * 100, 1) if gams_km > 0 else 0
    print(f"  {'Metrik':<35} {'GAMS':>10} {'OR-Tools':>10} {'Fark':>8}")
    print(f"  {'-'*63}")
    print(f"  {'Toplam rota mesafesi (km)':<35} {gams_km:>10.2f} {ort_toplam_km:>10.2f} {fark_km:>+8.2f}")

    gams_arac  = 1
    ort_arac   = len(rotalar)
    print(f"  {'Kullanılan araç sayısı':<35} {gams_arac:>10} {ort_arac:>10} {ort_arac-gams_arac:>+8}")
    if ort_arac != gams_arac:
        print("  [UYARI] Arac sayisi GAMS ve OR-Tools arasinda farkli: "
              "amac fonksiyonlari (parasal vs. mesafe) farkli oldugu icin "
              "km bazli kiyaslama tek basina yeterli degildir, yorumda dikkatli olun.")
    print(f"  {'Ziyaret edilen mağaza':<35} {'8':>10} {'8':>10} {'0':>8}")
    print(f"  {'Fark (km)':<35} {'—':>10} {'—':>10} {fark_km:>+8.2f}")
    print(f"  {'Fark (%)':<35} {'—':>10} {'—':>10} {fark_pct:>7.1f}%")
    print()
    if fark_pct <= 5:
        print("  ► OR-Tools çözümü GAMS optimaline ≤%5 yakın.")
        print("    Sezgisel çözüm, kesin optimale eşdeğer kabul edilebilir.")
    elif fark_pct <= 15:
        print(f"  ► OR-Tools çözümü GAMS optimalinin %{fark_pct:.1f} üzerinde.")
        print("    Sezgisel çözüm kabul edilebilir yakınlıkta.")
    else:
        print(f"  ► OR-Tools çözümü GAMS optimalinin %{fark_pct:.1f} üzerinde.")
        print("    Fark büyük — sezgisel parametre ayarı gerekebilir.")
