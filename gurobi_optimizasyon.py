import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import gurobipy as gp
from gurobipy import GRB
from scipy import stats
import warnings
warnings.filterwarnings('ignore')

PROJE_VERI_DOSYA = 'PROJE VERİ.xlsx'
CIKTI_DOSYA = 'rapor.xlsx'

df_prosesler_oncelik = pd.read_excel(PROJE_VERI_DOSYA, sheet_name='PROSESLER-ÖNCELİK', header=1)
df_cizelge_sure = pd.read_excel(PROJE_VERI_DOSYA, sheet_name='ÇİZELGE PROSES SÜRE')
df_urun_tarih = pd.read_excel(PROJE_VERI_DOSYA, sheet_name='ÜRÜN İSTENEN TARİH')
df_malzemeler = pd.read_excel(PROJE_VERI_DOSYA, sheet_name='MALZEMELER VE TERMİN')

proses_isimleri = [col for col in df_cizelge_sure.columns if col != 'SANTRAL KODU']

santral_kodlari = df_cizelge_sure['SANTRAL KODU'].dropna().tolist()

malzeme_df_clean = df_malzemeler[df_malzemeler['Unnamed: 2'].notna()]
malzeme_kodlari = malzeme_df_clean['Unnamed: 2'].tolist()
if 'MALZEME KOD' in malzeme_kodlari:
    malzeme_kodlari.remove('MALZEME KOD')

np.random.seed(42)
gecikme_verileri = []
for malzeme in malzeme_kodlari:
    for i in range(100):
        siparis_gun = np.random.randint(1, 300)
        siparis_tarih = datetime(2025, 1, 1) + timedelta(days=siparis_gun)

        if malzeme in ['R01', 'R10']:
            gecikme = max(-5, np.random.normal(0, 3))
        elif malzeme in ['R02', 'R03', 'R04']:
            gecikme = max(-2, np.random.normal(30, 15))
        elif malzeme in ['R05', 'R06', 'R07']:
            gecikme = max(-1, np.random.exponential(35))
        elif malzeme in ['R09']:
            gecikme = max(-1, np.random.exponential(5))
        elif malzeme in ['R12', 'R13']:
            gecikme = max(-2, np.random.exponential(20))
        elif malzeme in ['R14', 'R15', 'R16', 'R19']:
            gecikme = max(-2, np.random.triangular(-5, 30, 60))
        elif malzeme in ['R17', 'R18', 'R20', 'R21']:
            gecikme = max(-2, np.random.normal(5, 8))
        elif malzeme in ['R27']:
            gecikme = max(-2, np.random.exponential(50))
        elif malzeme in ['R29', 'R31']:
            gecikme = max(-1, np.random.exponential(20))
        else:
            gecikme = max(-2, np.random.normal(10, 15))

        gercek_tarih = siparis_tarih + timedelta(days=int(gecikme))
        durum = 'erken gelmiş' if gecikme < 0 else ('zamanında gelmiş' if gecikme == 0 else 'gecikme')

        gecikme_verileri.append({
            'MALZEME': malzeme,
            'SIPARIS_TARIHI': siparis_tarih,
            'GERCEKLESEN_TARIH': gercek_tarih,
            'GECIKME_GUN': int(gecikme),
            'DURUM': durum
        })

df_gecikme = pd.DataFrame(gecikme_verileri)

dagilim_parametreleri = {}
for malzeme in malzeme_kodlari:
    malz_data = df_gecikme[df_gecikme['MALZEME'] == malzeme]['GECIKME_GUN'].values

    if malzeme in ['R01', 'R02', 'R03', 'R04', 'R10', 'R17', 'R18', 'R20', 'R21']:
        try:
            mu, sigma = stats.norm.fit(malz_data)
            dagilim_parametreleri[malzeme] = {'tip': 'NORMAL DAĞILIM', 'mu': mu, 'sigma': sigma}
        except:
            dagilim_parametreleri[malzeme] = {'tip': 'NORMAL DAĞILIM', 'mu': np.mean(malz_data), 'sigma': np.std(malz_data)}
    elif malzeme in ['R05', 'R06', 'R07', 'R09', 'R12', 'R13', 'R27', 'R29', 'R31']:
        try:
            loc, scale = stats.expon.fit(malz_data)
            dagilim_parametreleri[malzeme] = {'tip': 'EXPONENTIAL (ÜSTEL) DAĞILIM', 'loc': loc, 'scale': scale}
        except:
            dagilim_parametreleri[malzeme] = {'tip': 'EXPONENTIAL (ÜSTEL) DAĞILIM', 'loc': min(malz_data), 'scale': np.mean(malz_data) - min(malz_data)}
    elif malzeme in ['R14', 'R15', 'R16', 'R19']:
        a_min = float(np.min(malz_data))
        a_max = float(np.max(malz_data))
        a_mode = float(np.median(malz_data))
        dagilim_parametreleri[malzeme] = {'tip': 'TRIANGULAR (ÜÇGENSEL) DAĞILIM', 'min': a_min, 'mode': a_mode, 'max': a_max}
    else:
        dagilim_parametreleri[malzeme] = {'tip': 'DETERMİNİSTİC DAĞILIM', 'value': np.mean(malz_data)}

santral_bom = {}
for santral in santral_kodlari[:30]:
    santral_malzemeler = []

    if 'A' in santral:
        santral_malzemeler = ['R01', 'R02', 'R03', 'R04', 'R05', 'R06', 'R07', 'R09', 'R10', 'R12', 'R13', 'R14', 'R15', 'R16', 'R17', 'R19', 'R20', 'R21']
    elif 'S' in santral:
        santral_malzemeler = ['R01', 'R02', 'R03', 'R04', 'R05', 'R09', 'R10', 'R12', 'R16', 'R17', 'R19', 'R21', 'R27']
    else:
        santral_malzemeler = ['R01', 'R02', 'R10', 'R12', 'R17', 'R19', 'R21']

    santral_bom[santral] = santral_malzemeler

hizlandirma_secenekleri = {
    'MİNİVAN YÜKLEME': {'gun': 3, 'maliyet': 1000},
    'EXPRESS YÜKLEME': {'gun': 2, 'maliyet': 2000},
    'UÇAK KARGO': {'gun': 1, 'maliyet': 4000}
}

musteri_onemi = {}
np.random.seed(100)
for sant in santral_kodlari:
    if np.random.rand() < 0.4:
        musteri_onemi[sant] = 'MÜŞTERİ TESLİM TARİHİ ÖNEMLİ'
    else:
        musteri_onemi[sant] = 'MALİYET BASKISI ÖNEMLİ'

sla_parametreleri = {}
for sant in santral_kodlari:
    musteri_row = df_urun_tarih[df_urun_tarih['SANTRAL KODU'] == sant]
    if not musteri_row.empty:
        musteri_tarihi = pd.to_datetime(musteri_row.iloc[0]['MÜŞTERİNİN İSTEDİĞİ TARİH'])
        bugun = datetime.now()
        teslim_kalan_gun = (musteri_tarihi - bugun).days

        cizelge_row = df_cizelge_sure[df_cizelge_sure['SANTRAL KODU'] == sant]
        if not cizelge_row.empty:
            uretim_suresi = cizelge_row[proses_isimleri].sum(axis=1, skipna=True).values[0]
        else:
            uretim_suresi = 60

        if musteri_onemi[sant] == 'MÜŞTERİ TESLİM TARİHİ ÖNEMLİ':
            ceza_katsayisi = 3
            musteri_onemi_katsayisi = 3
        else:
            ceza_katsayisi = 1
            musteri_onemi_katsayisi = 1

        w1 = 0.4
        w2 = 0.3
        w3 = 0.3

        uretim_suresi_norm = uretim_suresi / teslim_kalan_gun if teslim_kalan_gun > 0 else 1.0

        sla_skoru = w1 * uretim_suresi_norm + w2 * ceza_katsayisi + w3 * musteri_onemi_katsayisi

        sla_parametreleri[sant] = {
            'teslim_kalan_gun': max(1, teslim_kalan_gun),
            'uretim_suresi': uretim_suresi,
            'ceza_katsayisi': ceza_katsayisi,
            'musteri_onemi_katsayisi': musteri_onemi_katsayisi,
            'sla_skoru': sla_skoru,
            'w1': w1,
            'w2': w2,
            'w3': w3
        }

senaryo_santraller = santral_kodlari[:10]

np.random.seed(200)
# 10 santralden 5 tanesi geciksin (oranı koruyalım)
gec_kalan_santraller = np.random.choice(senaryo_santraller, size=5, replace=False).tolist()

santral_malzeme_gecikme = {}
for sant in senaryo_santraller:
    if sant in gec_kalan_santraller:
        if sant in santral_bom:
            bom = santral_bom[sant]
            kritik_malzemeler = ['R05', 'R06', 'R07', 'R09']
            gec_kalacak_malzemeler = [m for m in bom if m in kritik_malzemeler]

            if not gec_kalacak_malzemeler:
                gec_kalacak_malzemeler = bom[:2] if len(bom) >= 2 else bom

            for malz in gec_kalacak_malzemeler:
                if malz in dagilim_parametreleri:
                    dag_info = dagilim_parametreleri[malz]
                    if 'mu' in dag_info:
                        gecikme_gun = int(max(0, np.random.normal(dag_info['mu'], dag_info.get('sigma', 5))))
                    elif 'scale' in dag_info:
                        gecikme_gun = int(max(0, np.random.exponential(dag_info['scale'])))
                    else:
                        gecikme_gun = int(max(0, np.random.normal(20, 10)))
                else:
                    gecikme_gun = int(max(0, np.random.normal(20, 10)))

                santral_malzeme_gecikme[(sant, malz)] = gecikme_gun

def ciz_optimizasyon_yap(santraller, senaryo_adi='NORMAL', delay_scenario=None):
    model = gp.Model(f'HVAC_{senaryo_adi}')
    model.setParam('OutputFlag', 0)

    baslangic_zaman = {}
    gecikme_var = {}
    expedite_karar = {}

    for santral in santraller:
        for proses in proses_isimleri:
            baslangic_zaman[(santral, proses)] = model.addVar(vtype=GRB.CONTINUOUS, name=f'start_{santral}_{proses}')
        
        gecikme_var[santral] = model.addVar(vtype=GRB.CONTINUOUS, name=f'delay_{santral}')

        if senaryo_adi == 'RECOVERY' and santral in santral_bom:
            for malz in santral_bom[santral]:
                expedite_karar[(santral, malz)] = model.addVar(vtype=GRB.BINARY, name=f'expedite_{santral}_{malz}')

    # Süreleri önceden çekelim
    proses_sureleri = {}
    for santral in santraller:
        cizelge_row = df_cizelge_sure[df_cizelge_sure['SANTRAL KODU'] == santral]
        if not cizelge_row.empty:
            for proses in proses_isimleri:
                sure = cizelge_row[proses].values[0]
                if pd.notna(sure) and sure > 0:
                    proses_sureleri[(santral, proses)] = sure
                else:
                    proses_sureleri[(santral, proses)] = 0
        else:
            for proses in proses_isimleri:
                proses_sureleri[(santral, proses)] = 0

    # Kısıtlar
    # Kısıtlar
    # Öncelik matrisini düzgün kullanalım
    # df_prosesler_oncelik'te 'OPERASYON ADI' sütunu satırları belirtir.
    # Sütunlar ise proses isimleridir.
    
    # Proses isimlerinin dataframe'deki karşılıklarını bulalım
    # Bazen boşluklar farklı olabilir, strip edelim
    df_prosesler_oncelik.columns = [c.strip() if isinstance(c, str) else c for c in df_prosesler_oncelik.columns]
    df_prosesler_oncelik['OPERASYON ADI'] = df_prosesler_oncelik['OPERASYON ADI'].apply(lambda x: x.strip() if isinstance(x, str) else x)
    
    constraint_count = 0
    for santral in santraller:
        for proses1 in proses_isimleri:
            for proses2 in proses_isimleri:
                if proses1 == proses2:
                    continue
                
                # Matriste proses1 satırını bul
                row = df_prosesler_oncelik[df_prosesler_oncelik['OPERASYON ADI'] == proses1]
                if not row.empty:
                    # Sütun olarak proses2'ye bak
                    if proses2 in row.columns:
                        val = row.iloc[0][proses2]
                        if val == 1:
                            # proses1 -> proses2 (proses2, proses1'i bekliyor)
                            # Start[p2] >= Start[p1] + Sure[p1]
                            sure1 = proses_sureleri[(santral, proses1)]
                            model.addConstr(baslangic_zaman[(santral, proses2)] >= baslangic_zaman[(santral, proses1)] + sure1)

    for santral in santraller:
        musteri_row = df_urun_tarih[df_urun_tarih['SANTRAL KODU'] == santral]
        if not musteri_row.empty:
            due_date_dt = pd.to_datetime(musteri_row.iloc[0]['MÜŞTERİNİN İSTEDİĞİ TARİH'])
            bugun = datetime.now()
            due_date_gun = (due_date_dt - bugun).days

            son_proses = proses_isimleri[-1]
            sure_son = proses_sureleri[(santral, son_proses)]
            # bitis_zaman[son] yerine baslangic_zaman[son] + sure[son]
            model.addConstr(gecikme_var[santral] >= (baslangic_zaman[(santral, son_proses)] + sure_son) - due_date_gun)
            model.addConstr(gecikme_var[santral] >= 0)

    for santral in santraller:
        musteri_row = df_urun_tarih[df_urun_tarih['SANTRAL KODU'] == santral]
        if not musteri_row.empty:
            due_date_dt = pd.to_datetime(musteri_row.iloc[0]['MÜŞTERİNİN İSTEDİĞİ TARİH'])
            bugun = datetime.now()
            due_date_gun = (due_date_dt - bugun).days

            son_proses = proses_isimleri[-1]
            sure_son = proses_sureleri[(santral, son_proses)]
            # bitis_zaman[son] yerine baslangic_zaman[son] + sure[son]
            model.addConstr(gecikme_var[santral] >= (baslangic_zaman[(santral, son_proses)] + sure_son) - due_date_gun)
            model.addConstr(gecikme_var[santral] >= 0)

    # --- Malzeme Gecikmesi ve Expedite Etkisi ---
    ilk_proses = proses_isimleri[0]
    
    for santral in santraller:
        if santral in santral_bom:
            for malz in santral_bom[santral]:
                if delay_scenario and (santral, malz) in delay_scenario:
                    gecikme_gun = delay_scenario.get((santral, malz), 0)
                else:
                    gecikme_gun = 0
                
                if gecikme_gun > 0:
                    if senaryo_adi == 'RECOVERY':
                        if (santral, malz) in expedite_karar:
                             model.addConstr(baslangic_zaman[(santral, ilk_proses)] >= gecikme_gun * (1 - expedite_karar[(santral, malz)]))
                        else:
                             model.addConstr(baslangic_zaman[(santral, ilk_proses)] >= gecikme_gun)
                    else:
                        model.addConstr(baslangic_zaman[(santral, ilk_proses)] >= gecikme_gun)

    # --- SLA ve X, Y Katsayılarının Hesaplanması (Her Santral İçin) ---
    santral_katsayilar = {}
    for sant in santraller:
        if sant in sla_parametreleri:
            sla_info = sla_parametreleri[sant]
            # SLA skoru 0-3 arasında, bunu normalize edelim
            sla_norm = sla_info['sla_skoru'] / 3.0 if sla_info['sla_skoru'] > 0 else 0.5
        else:
            sla_norm = 0.5 # Varsayılan
        
        # Bütçe faktörü (sabit varsayıldı veya parametrik olabilir)
        butce_norm = 0.5 
        
        toplam_payda = sla_norm + butce_norm
        if toplam_payda == 0:
            X = 0.5
            Y = 0.5
        else:
            X = sla_norm / toplam_payda
            Y = butce_norm / toplam_payda
            
        santral_katsayilar[sant] = {'X': X, 'Y': Y}

    # --- Amaç Fonksiyonu ---
    # Her santralin kendi X (Müşteri/Zaman) ve Y (Maliyet) ağırlığına göre maliyet minimize edilir
    
    obj_terms = []
    
    for sant in santraller:
        X_val = santral_katsayilar[sant]['X']
        Y_val = santral_katsayilar[sant]['Y']
        
        # Gecikme Maliyeti (Birim gecikme maliyeti 500 TL varsayıldı)
        gecikme_cost = gecikme_var[sant] * 500
        
        # Hızlandırma Maliyeti
        hizlandirma_cost = 0
        if senaryo_adi == 'RECOVERY' and sant in santral_bom:
            for malz in santral_bom[sant]:
                if (sant, malz) in expedite_karar:
                    min_cost = min([hizlandirma_secenekleri[k]['maliyet'] for k in hizlandirma_secenekleri])
                    hizlandirma_cost += expedite_karar[(sant, malz)] * min_cost
                    
                    d_gun = 0
                    if delay_scenario and (sant, malz) in delay_scenario:
                        d_gun = delay_scenario.get((sant, malz), 0)
                    if d_gun > 40:
                        model.addConstr(expedite_karar[(sant, malz)] == 1)
        
        # Ağırlıklı Toplam
        obj_terms.append(X_val * gecikme_cost + Y_val * hizlandirma_cost)

    model.setObjective(gp.quicksum(obj_terms), GRB.MINIMIZE)

    model.optimize()

    sonuclar = []
    expedite_sayisi = 0
    if model.status == GRB.OPTIMAL or model.status == GRB.SUBOPTIMAL:
        for santral in santraller:
            X_val = santral_katsayilar[santral]['X']
            Y_val = santral_katsayilar[santral]['Y']
            
            for proses in proses_isimleri:
                try:
                    bas_gun = min(max(0, baslangic_zaman[(santral, proses)].X), 365)
                    # bitis_gun hesapla
                    sure = proses_sureleri.get((santral, proses), 0)
                    bit_gun = min(max(0, bas_gun + sure), 365)

                    bugun = datetime.now()
                    try:
                        bas_tarih = bugun + timedelta(days=bas_gun)
                        bas_str = bas_tarih.strftime('%Y-%m-%d')
                    except:
                        bas_str = f'GÜN {int(bas_gun)}'

                    try:
                        bit_tarih = bugun + timedelta(days=bit_gun)
                        bit_str = bit_tarih.strftime('%Y-%m-%d')
                    except:
                        bit_str = f'GÜN {int(bit_gun)}'

                    sonuclar.append({
                        'SANTRAL': santral,
                        'PROSES': proses,
                        'BAŞLANGIÇ': bas_str,
                        'BİTİŞ': bit_str,
                        'X_KATSAYI': round(X_val, 3),
                        'Y_KATSAYI': round(Y_val, 3)
                    })
                except:
                    pass

    
    expedite_values = {}
    if model.status == GRB.OPTIMAL or model.status == GRB.SUBOPTIMAL:
        if senaryo_adi == 'RECOVERY':
            for sant in santraller:
                if sant in santral_bom:
                    for malz in santral_bom[sant]:
                        if (sant, malz) in expedite_karar:
                            val = 1 if expedite_karar[(sant, malz)].X > 0.5 else 0
                            if val == 1:
                                expedite_sayisi += 1
                            expedite_values[(sant, malz)] = val

    return sonuclar, expedite_sayisi, santral_katsayilar, expedite_values

# --- SENARYOLARIN OLUŞTURULMASI (DİNAMİK YAKLAŞIM) ---
# 1. Base Plan: Beklenen sürelerle (Gecikme Yok varsayımı)
cizelge_base, _, katsayilar_normal, _ = ciz_optimizasyon_yap(senaryo_santraller, 'BASE_PLAN', delay_scenario={})

# 2. Realization (Disrupted): Rastgele gecikmeler gerçekleşti
cizelge_disrupted, _, _, _ = ciz_optimizasyon_yap(senaryo_santraller, 'DISRUPTED', delay_scenario=santral_malzeme_gecikme)

# 3. Recovery (Rescheduled): Gecikmelere karşı expedite kararı
cizelge_recovery, exp_sayisi, katsayilar_final, expedite_kararlar_var = ciz_optimizasyon_yap(senaryo_santraller, 'RECOVERY', delay_scenario=santral_malzeme_gecikme)

# --- DETAYLI HIZLANDIRMA RAPORU OLUŞTURMA ---
expedite_detay_listesi = []
for (sant, malz), karar_val in expedite_kararlar_var.items():
    if karar_val == 1: # Eğer expedite kararı verildiyse (Artık integer dönüyor)
        # Hangi metod seçildiğini bulalım (en ucuz olan varsayımıyla)
        secilen_metod = 'MİNİVAN YÜKLEME' # Varsayılan en ucuz
        min_cost = 1000
        
        # Gecikme süresine bakalım
        gecikme_gun = santral_malzeme_gecikme.get((sant, malz), 0)
        
        expedite_detay_listesi.append({
            'SANTRAL': sant,
            'MALZEME': malz,
            'GEÇİKME SÜRESİ (Gün)': gecikme_gun,
            'KARAR': 'HIZLANDIRMA YAPILDI',
            'METOD': secilen_metod,
            'EK MALİYET (TL)': min_cost,
            'KAZANILAN SÜRE (Gün)': 3 # Minivan ile varsayılan 3 gün kazanım (veya gecikme kadar)
        })

# --- RAPOR VERİLERİNİN HAZIRLANMASI ---
ornek_santral_data = []
ornek_santral_kod = senaryo_santraller[0]
ornek_santral_data.append({'BİLGİ': f'Örneğin aşağıdaki santral için teslim tarihi', 'DEĞER': ''})
ornek_santral_data.append({'BİLGİ': ornek_santral_kod, 'DEĞER': ''})
ornek_santral_data.append({'BİLGİ': 'BOM LİST:', 'DEĞER': 'GECİKME OLASILIK DAĞILIMI:'})

if ornek_santral_kod in santral_bom:
    for malz in santral_bom[ornek_santral_kod]:
        if malz in dagilim_parametreleri:
            dag_info = dagilim_parametreleri.get(malz, {})
            dag_tip = dag_info.get('tip', 'BİLİNMİYOR')
            ornek_santral_data.append({'BİLGİ': malz, 'DEĞER': dag_tip})

satis_listesi_data = []
for santral in santral_kodlari:
    musteri_tarih_row = df_urun_tarih[df_urun_tarih['SANTRAL KODU'] == santral]
    if not musteri_tarih_row.empty:
        musteri_tarih = musteri_tarih_row.iloc[0]['MÜŞTERİNİN İSTEDİĞİ TARİH']
        onem = musteri_onemi.get(santral, 'MALİYET BASKISI ÖNEMLİ')
        satis_listesi_data.append({
            'SANTRAL KODU': santral,
            'MÜŞTERİNİN İSTEDİĞİ TARİH': musteri_tarih,
            'ÖNEM': onem
        })

hizlandirma_data = [{
    'EXPRESS YÜKLEME': f"{hizlandirma_secenekleri['EXPRESS YÜKLEME']['gun']} gün / {hizlandirma_secenekleri['EXPRESS YÜKLEME']['maliyet']} TL",
    'MİNİVAN YÜKLEME': f"{hizlandirma_secenekleri['MİNİVAN YÜKLEME']['gun']} gün / {hizlandirma_secenekleri['MİNİVAN YÜKLEME']['maliyet']} TL",
    'UÇAK KARGO': f"{hizlandirma_secenekleri['UÇAK KARGO']['gun']} gün / {hizlandirma_secenekleri['UÇAK KARGO']['maliyet']} TL"
}]

kisitlar_data = []
kisitlar_data.append({'KISIT TÜRÜ': 'Operasyon Öncelikleri', 'AÇIKLAMA': 'PROSESLER-ÖNCELİK matrisinden gelen proses sıralama kısıtları', 'ÖRNEK': 'SAC KESİM → SAC CNC → SAC BÜKÜM'})
kisitlar_data.append({'KISIT TÜRÜ': 'Malzeme Termin Süreleri', 'AÇIKLAMA': 'Her malzemenin tedarik süresi ve gecikme olasılıkları', 'ÖRNEK': 'R01(SAC): 12-15 hafta, R05(FAN): 8-10 hafta'})
kisitlar_data.append({'KISIT TÜRÜ': 'Müşteri Teslim Tarihleri', 'AÇIKLAMA': 'Her santral için belirlenen müşteri teslim tarihi kısıtı', 'ÖRNEK': 'Gecikme = max(0, BitişTarihi - MüşteriTarihi)'})
kisitlar_data.append({'KISIT TÜRÜ': 'Proses Süreleri', 'AÇIKLAMA': 'Her santral-proses kombinasyonu için belirlenen süre', 'ÖRNEK': 'BitişZamanı = BaşlangıçZamanı + ProsesSüresi'})
kisitlar_data.append({'KISIT TÜRÜ': 'SLA Kritikliği', 'AÇIKLAMA': 'Müşteri önemine göre ağırlıklandırma (w1=0.4, w2=0.3, w3=0.3)', 'ÖRNEK': 'SLA_skor = w1*(üretim/teslim) + w2*ceza + w3*müşteri_önemi'})
kisitlar_data.append({'KISIT TÜRÜ': 'Amaç Fonksiyonu', 'AÇIKLAMA': 'Min Z = Σ (X_i * GecikmeMaliyeti_i + Y_i * HızlandırmaMaliyeti_i)', 'ÖRNEK': 'X+Y=1, X=SLA/(SLA+Bütçe)'})

hesaplama_data = []
for sant in senaryo_santraller:
    if sant in sla_parametreleri:
        sla = sla_parametreleri[sant]
    else:
        sla = {}
    
    katsayi = katsayilar_normal.get(sant, {'X': 0.5, 'Y': 0.5})
    X_norm = katsayi['X']
    Y_norm = katsayi['Y']
    gecikme_durumu = 'VAR' if sant in gec_kalan_santraller else 'YOK'
    hesaplama_data.append({
        'SANTRAL': sant,
        'SLA_SKORU': round(sla.get('sla_skoru', 0), 3) if sant in sla_parametreleri else 0,
        'X (normalize)': round(X_norm, 3),
        'Y (normalize)': round(Y_norm, 3),
        'MALZEME GECİKME': gecikme_durumu,
        'MÜŞTERİ ÖNEMİ': musteri_onemi.get(sant, '')
    })

df_cizelge_base = pd.DataFrame(cizelge_base)
df_cizelge_disrupted = pd.DataFrame(cizelge_disrupted)
df_cizelge_recovery = pd.DataFrame(cizelge_recovery)

df_cizelge_combined = pd.concat([
    df_cizelge_base.assign(SENARYO='1. BASE PLAN (Beklenen)'),
    df_cizelge_disrupted.assign(SENARYO='2. DISRUPTED (Gecikmeli)'),
    df_cizelge_recovery.assign(SENARYO='3. RECOVERY (Hızlandırmalı)')
], ignore_index=True)

with pd.ExcelWriter(CIKTI_DOSYA, engine='openpyxl') as writer:
    df_gecikme.to_excel(writer, sheet_name='GEÇMİŞ VERİ ANALİZİ', index=False)
    # AKADEMİK DETAYLAR sayfası kaldırıldı (Revize isteği üzerine)
    pd.DataFrame(ornek_santral_data).to_excel(writer, sheet_name='ÖRNEK SANTRAL DETAY', index=False)
    pd.DataFrame(satis_listesi_data).to_excel(writer, sheet_name='SATIŞ VERİLERİ', index=False)
    pd.DataFrame(hizlandirma_data).to_excel(writer, sheet_name='HIZLANDIRMA SEÇENEKLERİ', index=False)
    # Yeni eklenen detaylı hızlandırma kararları sayfası
    if expedite_detay_listesi:
        pd.DataFrame(expedite_detay_listesi).to_excel(writer, sheet_name='HIZLANDIRMA KARARLARI', index=False)
    else:
        pd.DataFrame([{'MESAJ': 'Hızlandırma Yapılan Santral Yok'}]).to_excel(writer, sheet_name='HIZLANDIRMA KARARLARI', index=False)
        
    df_cizelge_combined.to_excel(writer, sheet_name='DİNAMİK ÇİZELGE SONUÇLARI', index=False)
    pd.DataFrame(kisitlar_data).to_excel(writer, sheet_name='KISITLAR', index=False)
    pd.DataFrame(hesaplama_data).to_excel(writer, sheet_name='SLA VE KATSAYILAR', index=False)
    malzeme_df_clean.to_excel(writer, sheet_name='MALZEMELER', index=False)
    df_prosesler_oncelik.to_excel(writer, sheet_name='PROSESLER', index=False)
    df_cizelge_sure[df_cizelge_sure['SANTRAL KODU'].isin(senaryo_santraller)].to_excel(writer, sheet_name='ÇİZELGE PROSES SÜRE', index=False)

print('Optimizasyon tamamlandı (Dinamik Senaryolar: Base -> Disrupted -> Recovery).')
print(f'Sonuçlar {CIKTI_DOSYA} dosyasına kaydedildi.')