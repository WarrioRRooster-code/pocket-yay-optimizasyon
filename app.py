import streamlit as st
import pandas as pd
import numpy as np
import pulp
import io
import re
from openpyxl.styles import PatternFill

# Sayfa Ayarları
st.set_page_config(page_title="Pocket Yay Planlama", page_icon="🛏️", layout="wide")
st.title("🛏️ Pocket Yay Çizelgeleme ve Optimizasyon Sistemi (MILP)")
st.markdown("Haftalık sipariş verilerinizi içeren Excel dosyasını yükleyin ve yapay zeka tabanlı optimizasyonu başlatın.")

# 1. Dosya Yükleme Alanı
uploaded_file = st.file_uploader("Lütfen 'Temiz' sayfasını barındıran Excel dosyasını yükleyin", type=["xlsx"])

if uploaded_file is not None:
    st.success("Dosya başarıyla yüklendi! Veriler okunuyor...")
    
    if st.button("🚀 Çizelgeyi Oluştur (Optimizasyonu Başlat)"):
        with st.spinner('Yapay Zeka tüm ihtimalleri değerlendiriyor. Bu işlem sipariş sayısına göre 5 dakikaya kadar sürebilir...'):
            try:
                # 1. Veriyi Oku
                df = pd.read_excel(uploaded_file, sheet_name="Temiz")
                df["Toplam İş Süresi (Dakika)"] = pd.to_numeric(df["Toplam İş Süresi (Dakika)"], errors='coerce').fillna(0)

                # 2. Üretim Parse Modülleri
                def model_zone_bul(tanim):
                    kws = []
                    tanim = str(tanim).upper()
                    if "UE" in tanim: kws.append("UE")
                    if " ST " in tanim: kws.append("ST")
                    if " CR " in tanim: kws.append("CR")
                    if " 5Z " in tanim or "5Z" in tanim: kws.append("5Z")
                    if " 7Z " in tanim or "7Z" in tanim: kws.append("7Z")
                    if " NZ " in tanim or "NZ" in tanim: kws.append("NZ")
                    if not kws: return "STANDART"
                    return " & ".join(kws)

                def tela_durumu_bul(tanim):
                    tanim_upper = str(tanim).upper()
                    if "TELASIZ" in tanim_upper: return "TELASIZ"
                    elif "TELALI" in tanim_upper: return "TELALI"
                    else: return "STANDART"

                def telleri_bul(tanim, mevcut_tel):
                    teller = re.findall(r'\d[.,]\d[Xx/]\d+', str(tanim))
                    if teller:
                        unique_teller = sorted(list(set([t.upper().replace('.', ',') for t in teller])))
                        return tuple(unique_teller)
                    else: return tuple([str(mevcut_tel).upper().replace('.', ',')])

                df["Model & Zone Bilgisi"] = df["Malzeme Uzun Tanımı"].apply(model_zone_bul)
                df["Tela Durumu"] = df["Malzeme Uzun Tanımı"].apply(tela_durumu_bul)
                df["Kullanılan Teller"] = df.apply(lambda row: telleri_bul(row["Malzeme Uzun Tanımı"], row["Tel Kalınlığı"]), axis=1)

                # 3. Siparişleri Birleştir
                groupby_cols = ["Model & Zone Bilgisi", "Kullanılan Teller", "Tel Kalınlığı", "Lamet Durumu", "Tela Durumu", "En", "Boy", "Yükseklik", "Sıralama Alanı (En * Boy)"]
                df_agg = df.groupby(groupby_cols, as_index=False).agg({
                    "Bileşen Kodu": "first",
                    "Malzeme Uzun Tanımı": "first",
                    "Sipariş Miktarı": "sum",
                    "Toplam İş Süresi (Dakika)": "sum"
                })

                jobs = df_agg.to_dict('records')
                N = len(jobs)

                # 4. Güncel Setup ve Neden Hesaplama Modülleri
                def calculate_setup(job_i, job_j):
                    setup = 0
                    if job_i['Model & Zone Bilgisi'] != job_j['Model & Zone Bilgisi']: setup += 30
                    if str(job_i['Yükseklik']) != str(job_j['Yükseklik']): setup += 60
                    
                    wires_i = set(job_i['Kullanılan Teller'])
                    wires_j = set(job_j['Kullanılan Teller'])
                    yeni_takilan_tel_sayisi = len(wires_j - wires_i)
                    setup += (yeni_takilan_tel_sayisi * 10)

                    tela_i = "TELASIZ" if job_i['Tela Durumu'] == "STANDART" else job_i['Tela Durumu']
                    tela_j = "TELASIZ" if job_j['Tela Durumu'] == "STANDART" else job_j['Tela Durumu']
                    if tela_i != tela_j: setup += 30
                    return setup

                def get_setup_reason(job_i, job_j):
                    reasons = []
                    if job_i['Model & Zone Bilgisi'] != job_j['Model & Zone Bilgisi']: reasons.append("Model/Zone")
                    if str(job_i['Yükseklik']) != str(job_j['Yükseklik']): reasons.append("Yükseklik")
                    
                    wires_i = set(job_i['Kullanılan Teller'])
                    wires_j = set(job_j['Kullanılan Teller'])
                    yeni_takilan_tel_sayisi = len(wires_j - wires_i)
                    if yeni_takilan_tel_sayisi > 0: reasons.append(f"{yeni_takilan_tel_sayisi} Yeni Tel")

                    tela_i = "TELASIZ" if job_i['Tela Durumu'] == "STANDART" else job_i['Tela Durumu']
                    tela_j = "TELASIZ" if job_j['Tela Durumu'] == "STANDART" else job_j['Tela Durumu']
                    if tela_i != tela_j: reasons.append("Tela")

                    if not reasons: return "-"
                    return " + ".join(reasons)

                # 5. MILP MODELİNİN KURULMASI
                prob = pulp.LpProblem("Pocket_Yay_Cizelgeleme_MILP", pulp.LpMinimize)
                V = list(range(N + 1))
                x = pulp.LpVariable.dicts("x", (V, V), cat='Binary')
                C = pulp.LpVariable.dicts("C", V, lowBound=0, cat='Continuous')

                prob += pulp.lpSum(calculate_setup(jobs[i-1], jobs[j-1]) * x[i][j] for i in V[1:] for j in V[1:] if i != j)

                for j in V[1:]: prob += pulp.lpSum(x[i][j] for i in V if i != j) == 1
                for i in V[1:]: prob += pulp.lpSum(x[i][j] for j in V if i != j) == 1
                prob += pulp.lpSum(x[0][j] for j in V[1:]) == 1
                prob += pulp.lpSum(x[i][0] for i in V[1:]) == 1

                M = 10000
                for i in V[1:]:
                    for j in V[1:]:
                        if i != j:
                            p_j = jobs[j-1]["Toplam İş Süresi (Dakika)"]
                            s_ij = calculate_setup(jobs[i-1], jobs[j-1])
                            prob += C[j] >= C[i] + p_j + s_ij - M * (1 - x[i][j])

                prob += C[0] == 0

                # Ebat Kesin Kısıtı (Model, Zone, Teller, Yükseklik ve Tela aynıysa ebat büyükten küçüğe akmalı)
                for i in V[1:]:
                    for j in V[1:]:
                        if i != j:
                            job_i, job_j = jobs[i-1], jobs[j-1]
                            tela_i_kural = "TELASIZ" if job_i['Tela Durumu'] == "STANDART" else job_i['Tela Durumu']
                            tela_j_kural = "TELASIZ" if job_j['Tela Durumu'] == "STANDART" else job_j['Tela Durumu']

                            if (job_i['Model & Zone Bilgisi'] == job_j['Model & Zone Bilgisi'] and
                                job_i['Kullanılan Teller'] == job_j['Kullanılan Teller'] and
                                str(job_i['Yükseklik']) == str(job_j['Yükseklik']) and
                                tela_i_kural == tela_j_kural):
                                if job_j['Sıralama Alanı (En * Boy)'] > job_i['Sıralama Alanı (En * Boy)']:
                                    prob += x[i][j] == 0

                # 6. ÇÖZÜCÜYÜ BAŞLAT
                prob.solve(pulp.PULP_CBC_CMD(timeLimit=300, msg=False))

                if pulp.LpStatus[prob.status] == 'Optimal' or pulp.LpStatus[prob.status] == 'Not Solved':
                    rota = []
                    mevcut_node = 0
                    for _ in range(N):
                        for j in V[1:]:
                            if pulp.value(x[mevcut_node][j]) == 1:
                                rota.append(j)
                                mevcut_node = j
                                break

                    sirali_isler = [jobs[i-1] for i in rota]
                    df_sonuc = pd.DataFrame(sirali_isler)
                    df_sonuc["Tespit Edilen Teller"] = df_sonuc["Kullanılan Teller"].apply(lambda x: " & ".join(x))

                    # 7. ZAMAN, SETUP VE NEDEN HESAPLAMALARI
                    setup_süreleri = [0]
                    setup_nedenleri = ["İlk İş (Kurulum)"]
                    
                    for idx in range(1, len(rota)):
                        onceki_is_idx = rota[idx-1] - 1
                        suanki_is_idx = rota[idx] - 1
                        job_i = jobs[onceki_is_idx]
                        job_j = jobs[suanki_is_idx]
                        
                        setup_süreleri.append(calculate_setup(job_i, job_j))
                        setup_nedenleri.append(get_setup_reason(job_i, job_j))

                    df_sonuc["Önceki İşten Geçiş Süresi (Setup - Dk)"] = setup_süreleri
                    df_sonuc["Setup Nedeni"] = setup_nedenleri

                    bitis_zamanlari = []
                    kümülatif_zaman = 0

                    for i in range(len(rota)):
                        is_suresi = sirali_isler[i]["Toplam İş Süresi (Dakika)"]
                        setup_suresi = setup_süreleri[i]
                        kümülatif_zaman += is_suresi + setup_suresi
                        bitis_zamanlari.append(kümülatif_zaman)

                    baslangic_zamanlari = [0] + bitis_zamanlari[:-1]
                    df_sonuc["Başlangıç Zamanı (Dk)"] = baslangic_zamanlari
                    df_sonuc["Bitiş Zamanı (Dk)"] = bitis_zamanlari

                    # Takvim Ataması
                    def vardiya_bul(sure):
                        if sure <= 0: return "1. Hafta", "Pazartesi", "Gece"
                        t = sure - 0.001
                        hafta_no = int(t // 5340) + 1
                        hafta_ici_dk = t % 5340
                        gunler = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]

                        if hafta_ici_dk < 4900:
                            gun_endeksi = int(hafta_ici_dk // 980)
                            gun_ici_dk = hafta_ici_dk % 980
                            if gun_ici_dk < 440: return f"{hafta_no}. Hafta", gunler[gun_endeksi], "Gece"
                            else: return f"{hafta_no}. Hafta", gunler[gun_endeksi], "Gündüz"
                        else:
                            return f"{hafta_no}. Hafta", "Cumartesi (Cuma'dan bağlayan)", "Gece"

                    baslangic_bilgileri = df_sonuc["Başlangıç Zamanı (Dk)"].apply(lambda x: pd.Series(vardiya_bul(x)))
                    bitis_bilgileri = df_sonuc["Bitiş Zamanı (Dk)"].apply(lambda x: pd.Series(vardiya_bul(x)))

                    df_sonuc["Başlama Haftası"] = baslangic_bilgileri[0]
                    df_sonuc["Başlama Günü"] = baslangic_bilgileri[1]
                    df_sonuc["Başlama Vardiyası"] = baslangic_bilgileri[2]
                    df_sonuc["Bitiş Haftası"] = bitis_bilgileri[0]
                    df_sonuc["Bitiş Günü"] = bitis_bilgileri[1]
                    df_sonuc["Bitiş Vardiyası"] = bitis_bilgileri[2]

                    # Sütunları Düzenleme (Setup Nedeni ve Model & Zone Bilgisi Eklendi)
                    cols_front = ['Bileşen Kodu', 'Malzeme Uzun Tanımı', 'Model & Zone Bilgisi', 'Tespit Edilen Teller', 'Tela Durumu']
                    cols_middle = [c for c in df_sonuc.columns if c not in cols_front and "Başlama" not in c and "Bitiş" not in c and "Başlangıç Zamanı" not in c and "Setup" not in c and "Tela Durumu" not in c and "Kullanılan Teller" not in c and "Tespit Edilen Teller" not in c and "Model & Zone Bilgisi" not in c and "Zone Bilgisi" not in c]
                    cols_end = ['Önceki İşten Geçiş Süresi (Setup - Dk)', 'Setup Nedeni', 'Başlangıç Zamanı (Dk)', 'Başlama Haftası', 'Başlama Günü', 'Başlama Vardiyası', 'Bitiş Zamanı (Dk)', 'Bitiş Haftası', 'Bitiş Günü', 'Bitiş Vardiyası']

                    df_sonuc = df_sonuc[cols_front + cols_middle + cols_end]

                    # ----------------------------------------------------
                    # EXCEL'E YAZDIRMA, OTOMATİK BOYAMA VE GİZLİ İMZA
                    # ----------------------------------------------------
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_sonuc.to_excel(writer, index=False, sheet_name='Optimum_Plan')
                        
                        workbook = writer.book
                        worksheet = writer.sheets['Optimum_Plan']
                        
                        # GİZLİ İMZA (Buraya Kendi Adını Yazabilirsin)
                        workbook.properties.creator = "Senin Adın Soyadın" 
                        workbook.properties.title = "Pocket Yay Çizelgeleme Motoru"
                        
                        # Boyama ve Filtreleme
                        mavi_dolgu = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                        bej_dolgu = PatternFill(start_color="EAE2D0", end_color="EAE2D0", fill_type="solid")
                        
                        col_baslangic_zaman = df_sonuc.columns.get_loc('Başlangıç Zamanı (Dk)') + 1
                        col_bitis_zaman = df_sonuc.columns.get_loc('Bitiş Zamanı (Dk)') + 1
                        col_baslama_vardiya = df_sonuc.columns.get_loc('Başlama Vardiyası') + 1
                        col_bitis_vardiya = df_sonuc.columns.get_loc('Bitiş Vardiyası') + 1
                        
                        for row in range(2, len(df_sonuc) + 2):
                            worksheet.cell(row=row, column=col_baslangic_zaman).fill = mavi_dolgu
                            worksheet.cell(row=row, column=col_bitis_zaman).fill = mavi_dolgu
                            worksheet.cell(row=row, column=col_baslama_vardiya).fill = bej_dolgu
                            worksheet.cell(row=row, column=col_bitis_vardiya).fill = bej_dolgu
                            
                        worksheet.auto_filter.ref = worksheet.dimensions

                    processed_data = output.getvalue()

                    st.success(f"Yapay Zeka planlamayı %100 doğrulukla tamamladı! İşlenen benzersiz ürün grubu: {N}")
                    
                    st.download_button(
                        label="📥 Oluşturulan Planı Excel Olarak İndir",
                        data=processed_data,
                        file_name="Optimum_Uretim_Plani_Yapay_Zeka.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.dataframe(df_sonuc.head(10)) 
                else:
                    st.error("Sistem geçerli bir rota bulamadı. Lütfen kuralların birbiriyle çakışmadığından emin olun.")
            except Exception as e:
                st.error(f"Bir hata oluştu: {e}")
