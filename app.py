import streamlit as st
import pandas as pd
import numpy as np
import pulp
import io

# Sayfa Ayarları
st.set_page_config(page_title="Pocket Yay Planlama", page_icon="🛏️", layout="wide")
st.title("🛏️ Pocket Yay Çizelgeleme ve Optimizasyon Sistemi")
st.markdown("Haftalık sipariş verilerinizi içeren Excel dosyasını yükleyin ve optimizasyonu başlatın.")

# 1. Dosya Yükleme Alanı
uploaded_file = st.file_uploader("Lütfen 'Temiz' sayfasını barındıran Excel dosyasını yükleyin", type=["xlsx"])

if uploaded_file is not None:
    st.success("Dosya başarıyla yüklendi! Veriler okunuyor...")
    
    if st.button("🚀 Çizelgeyi Oluştur (Optimizasyonu Başlat)"):
        with st.spinner('Matematiksel model kuruluyor ve çözülüyor. Bu işlem sipariş sayısına göre birkaç dakika sürebilir...'):
            try:
                # Veriyi Oku
                df = pd.read_excel(uploaded_file, sheet_name="Temiz")
                df["Toplam İş Süresi (Dakika)"] = pd.to_numeric(df["Toplam İş Süresi (Dakika)"], errors='coerce').fillna(0)

                # Siparişleri Birleştir (Modeli Küçült)
                groupby_cols = ["Tel Kalınlığı", "Zone Bilgisi", "Lamet Durumu", "En", "Boy", "Yükseklik", "Sıralama Alanı (En * Boy)"]
                df_agg = df.groupby(groupby_cols, as_index=False).agg({
                    "Bileşen Kodu": "first",
                    "Malzeme Uzun Tanımı": "first",
                    "Sipariş Miktarı": "sum",
                    "Toplam İş Süresi (Dakika)": "sum"
                })
                jobs = df_agg.to_dict('records')
                N = len(jobs)

                # Setup Matrisi Fonksiyonu
                def calculate_setup(job_i, job_j):
                    setup = 0
                    if job_i['Zone Bilgisi'] != job_j['Zone Bilgisi']: setup += 8
                    if job_i['Lamet Durumu'] != job_j['Lamet Durumu']: setup += 2
                    if str(job_i['Tel Kalınlığı']) != str(job_j['Tel Kalınlığı']): setup += 5 
                    return setup

                # MILP Modeli
                prob = pulp.LpProblem("Pocket_Yay_Cizelgeleme", pulp.LpMinimize)
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

                for i in V[1:]:
                    for j in V[1:]:
                        if i != j:
                            job_i, job_j = jobs[i-1], jobs[j-1]
                            if (job_i['Tel Kalınlığı'] == job_j['Tel Kalınlığı'] and 
                                job_i['Zone Bilgisi'] == job_j['Zone Bilgisi'] and 
                                job_i['Lamet Durumu'] == job_j['Lamet Durumu']):
                                if job_j['Sıralama Alanı (En * Boy)'] > job_i['Sıralama Alanı (En * Boy)']:
                                    prob += x[i][j] == 0

                # Çözücü (Maks 300 saniye)
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
                    
                    # Başlangıç ve Bitiş Sürelerinin Hesaplanması
                    bitis_zamanlari = [pulp.value(C[i]) for i in rota]
                    baslangic_zamanlari = [0] + bitis_zamanlari[:-1]
                    
                    df_sonuc["Başlangıç Zamanı (Dk)"] = baslangic_zamanlari
                    df_sonuc["Bitiş Zamanı (Dk)"] = bitis_zamanlari
                    
                    # Takvim ve Vardiya Ataması (Yeni Gece-Gündüz Düzeni)
                    def vardiya_bul(sure):
                        if sure <= 0: return "1. Hafta", "Pazartesi", "Gece" # Gün Gece ile başlıyor
                        
                        t = sure - 0.001 
                        hafta_no = int(t // 5340) + 1
                        hafta_ici_dk = t % 5340
                        gunler = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
                        
                        if hafta_ici_dk < 4900:
                            gun_endeksi = int(hafta_ici_dk // 980)
                            gun_ici_dk = hafta_ici_dk % 980
                            
                            # Yeni Mesai Kurgusu: İlk 440dk Gece, Sonraki 540dk Gündüz
                            if gun_ici_dk < 440:
                                vardiya = "Gece"
                                gun = gunler[gun_endeksi]
                            else:
                                vardiya = "Gündüz"
                                gun = gunler[gun_endeksi]
                        else:
                            # 4900 sonrası Cuma gecesinden Cumartesi sabahına sarkan vardiya
                            vardiya = "Gece"
                            gun = "Cumartesi (Cuma'dan bağlayan)"
                            
                        return f"{hafta_no}. Hafta", gun, vardiya

                    # Başlangıç ve Bitiş için Fonksiyonu Ayrı Ayrı Uygulama
                    baslangic_bilgileri = df_sonuc["Başlangıç Zamanı (Dk)"].apply(lambda x: pd.Series(vardiya_bul(x)))
                    bitis_bilgileri = df_sonuc["Bitiş Zamanı (Dk)"].apply(lambda x: pd.Series(vardiya_bul(x)))

                    df_sonuc["Başlama Haftası"] = baslangic_bilgileri[0]
                    df_sonuc["Başlama Günü"] = baslangic_bilgileri[1]
                    df_sonuc["Başlama Vardiyası"] = baslangic_bilgileri[2]
                    
                    df_sonuc["Bitiş Haftası"] = bitis_bilgileri[0]
                    df_sonuc["Bitiş Günü"] = bitis_bilgileri[1]
                    df_sonuc["Bitiş Vardiyası"] = bitis_bilgileri[2]
                    
                    # Sütunları Düzenleme (Sıralama)
                    cols_front = ['Bileşen Kodu', 'Malzeme Uzun Tanımı']
                    cols_middle = [c for c in df_sonuc.columns if c not in cols_front and "Başlama" not in c and "Bitiş" not in c and "Başlangıç Zamanı" not in c]
                    cols_end = ['Başlangıç Zamanı (Dk)', 'Başlama Haftası', 'Başlama Günü', 'Başlama Vardiyası', 'Bitiş Zamanı (Dk)', 'Bitiş Haftası', 'Bitiş Günü', 'Bitiş Vardiyası']
                    
                    df_sonuc = df_sonuc[cols_front + cols_middle + cols_end]

                    # Sonucu İndirilebilir Hale Getirme
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_sonuc.to_excel(writer, index=False, sheet_name='Optimum_Plan')
                    processed_data = output.getvalue()

                    st.success(f"Optimizasyon başarıyla tamamlandı! Toplam işlenen satır grubu: {N}")
                    
                    st.download_button(
                        label="📥 Oluşturulan Planı Excel Olarak İndir",
                        data=processed_data,
                        file_name="Optimum_Uretim_Plani.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.dataframe(df_sonuc.head(10)) 
                else:
                    st.error("Sistem geçerli bir rota bulamadı. Lütfen verilerinizi kontrol edin.")
            except Exception as e:
                st.error(f"Bir hata oluştu: {e}")
