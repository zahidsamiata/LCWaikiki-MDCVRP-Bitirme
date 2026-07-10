import pandas as pd
import numpy as np
from ortools.constraint_solver import routing_enums_pb2
from ortools.constraint_solver import pywrapcp

def main():
    print("==============================================================")
    print("🚀 [BİLGİ] LCW OR-Tools GENİŞ ÖLÇEKLİ (204) VRP MOTORU BAŞLATILIYOR...")
    print("==============================================================\n")

    # 1. GERÇEK EXCEL VERİLERİNİ OKUMA
    try:
        print("[İŞLEM] 204x204 Devasa Mesafe Matrisi Excel'den Okunuyor...")
        # (1) takısı olmayan, bilgisayarındaki orijinal dosya ismini kullanıyoruz
        df_dist = pd.read_excel('01_Gercek_Veriler/mesafe_matrisi_204.xlsx', index_col=0)
        
        # OR-Tools tam sayı (integer) sever. Mesafeleri metreye çevirip tam sayı yapıyoruz.
        distance_matrix = (df_dist.fillna(0).values * 1000).astype(int).tolist()
        lokasyon_isimleri = df_dist.columns.tolist()
        num_locations = len(lokasyon_isimleri)
        print(f"✅ Matris Başarıyla Yüklendi! Toplam Nokta Sayısı: {num_locations}")
        
    except Exception as e:
        print(f"❌ [HATA] Excel dosyaları okunurken bir problem oluştu!\nDetay: {e}")
        print("Lütfen '01_Gercek_Veriler' klasöründe 'LCW_Mesafe_Sure_Matrisi.xlsx' dosyasının olduğundan emin olun.")
        return

    # 2. BÜYÜK MODEL İÇİN FİLO VE TALEP YAPILANDIRMASI
    data = {}
    data['distance_matrix'] = distance_matrix
    
    # GAMS modelindeki gibi: İlk 6 nokta DEPO, kalanlar MAĞAZA
    # Büyük veri testinde kolaylık olması için 10 adet güçlü Kamyon sahaya sürüyoruz.
    data['num_vehicles'] = 10
    
    # 10 aracın çıkış depoları (0'dan 5'e kadar olan indeksler depoları temsil eder)
    data['starts'] = [0, 0, 1, 1, 2, 2, 3, 3, 4, 5]
    data['ends'] = [0, 0, 1, 1, 2, 2, 3, 3, 4, 5] # Depolarına geri dönecekler

    # 204 Lokasyon için kısıt oluşturma: 
    # Depoların talebi 0'dır. Mağazalar için ortalama 150 koli talep varsayıyoruz.
    demands = [0] * 6 # 6 adet depo için talep = 0
    for i in range(6, num_locations):
        demands.append(150) # Her mağaza için standart 150 koli talep
    data['demands'] = demands

    # Her bir aracın koli taşıma kapasitesi (Standart Kamyon: 3000 Koli)
    data['vehicle_capacities'] = [3000] * data['num_vehicles']

    # 3. OR-TOOLS YÖNLENDİRME YÖNETİCİSİ
    manager = pywrapcp.RoutingIndexManager(len(data['distance_matrix']),
                                           data['num_vehicles'],
                                           data['starts'],
                                           data['ends'])
    routing = pywrapcp.RoutingModel(manager)

    # 4. MESAFE (MALİYET) FONKSİYONU
    def distance_callback(from_index, to_index):
        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        return data['distance_matrix'][from_node][to_node]

    transit_callback_index = routing.RegisterTransitCallback(distance_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)

    # 5. KAPASİTE KISITI EKLENMESİ
    def demand_callback(from_index):
        from_node = manager.IndexToNode(from_index)
        return data['demands'][from_node]

    demand_callback_index = routing.RegisterUnaryTransitCallback(demand_callback)
    routing.AddDimensionWithVehicleCapacity(
        demand_callback_index,
        0,  # null capacity slack
        data['vehicle_capacities'],
        True,  # start cumul to zero
        'Capacity'
    )

    # 6. ARAMA PARAMETRELERİ (Büyük veri olduğu için Meta-Sezgisel devrede!)
    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = (routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC)
    search_parameters.local_search_metaheuristic = (routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH)
    search_parameters.time_limit.seconds = 10  # Devasa ağ için algoritmanın 10 saniye çalışmasına izin veriyoruz

    # 7. MODELİ ÇÖZ
    print("\n⏳ [İŞLEM] Google OR-Tools Meta-Sezgisel Algoritması Optimize Ediyor (Max 10 saniye)...")
    solution = routing.SolveWithParameters(search_parameters)

    # 8. SONUÇLARI RAPORLA
    if solution:
        print("\n🏆 ✅ [BAŞARILI] 204 LOKASYON İÇİN OPTİMAL SEZGİSEL ROTALAR ÜRETİLDİ!\n")
        toplam_mesafe = 0
        toplam_yuk = 0
        aktif_arac_sayisi = 0
        
        for vehicle_id in range(data['num_vehicles']):
            index = routing.Start(vehicle_id)
            if routing.IsEnd(solution.Value(routing.NextVar(index))):
                continue # Eğer araç depodan hiç çıkmadıysa rapora yazma
                
            aktif_arac_sayisi += 1
            plan_output = f"🚛 Araç {vehicle_id + 1} Rotalama Planı (Depo İndeksi: {data['starts'][vehicle_id]}):\n   "
            route_distance = 0
            route_load = 0
            durak_sayisi = 0
            
            while not routing.IsEnd(index):
                node_index = manager.IndexToNode(index)
                route_load += data['demands'][node_index]
                plan_output += f"{lokasyon_isimleri[node_index]} -> "
                previous_index = index
                index = solution.Value(routing.NextVar(index))
                route_distance += routing.GetArcCostForVehicle(previous_index, index, vehicle_id)
                durak_sayisi += 1
            
            node_index = manager.IndexToNode(index)
            plan_output += f"{lokasyon_isimleri[node_index]}\n"
            print(plan_output)
            print(f"   📍 Rota Uzunluğu: {route_distance / 1000:.2f} km | 📦 Ziyaret Edilen Mağaza: {durak_sayisi-1} | Toplam Yük: {route_load} Koli\n")
            toplam_mesafe += route_distance
            toplam_yuk += route_load
            
        print("==============================================================")
        print(f"📊 ÖZET FİLO RAPORU (LARGE-SCALE MDCVRP)")
        print("==============================================================")
        print(f"🔹 Görev Alan Aktif Araç Sayısı : {aktif_arac_sayisi} / {data['num_vehicles']}")
        print(f"🔹 Toplam Kat Edilen Mesafe     : {toplam_mesafe / 1000:.2f} km")
        print(f"🔹 Dağıtılan Toplam Ürün Miktarı: {toplam_yuk} Koli")
        print("==============================================================")
    else:
        print("❌ [HATA] Bu kısıtlar altında geçerli bir rota kombinasyonu bulunamadı!")

if __name__ == '__main__':
    main()