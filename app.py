import streamlit as st
import pandas as pd
import numpy as np
import pulp
import io

# Sayfa Ayarları
st.set_page_config(page_title="Pocket Yay Planlama", page_icon="🛏️", layout="wide")
st.title("🛏️ Pocket Yay Çizelgeleme ve Optimizasyon Sistemi (MILP)")
st.markdown("Haftalık sipariş verilerinizi içeren Excel dosyasını yükleyin ve optimizasyonu başlatın.")

# 1. Dosya Yükleme Alanı
uploaded_file = st.file_uploader("Lütfen 'Temiz' sayfasını barındıran Excel dosyasını yükleyin", type=["xlsx"])

if uploaded_file is not None:
    st.success("Dosya başarıyla yüklendi! Veriler okunuyor...")
    
    if st.button("🚀 Çizelgeyi Oluştur (Optimizasyonu Başlat)"):
        with st.spinner('Matematiksel model kuruluyor ve çözülüyor. Bu işlem sipariş sayısına göre 5 dakikaya kadar sürebilir...'):
            try:
                # 1. Veriyi Oku
                df = pd.read_excel(uploaded_file, sheet_name="Temiz")
                df["Toplam İş Süresi (Dakika)"] = pd.to_numeric(df["Toplam İş Süresi (Dakika)"], errors='coerce').fillna(0)

                # 2. Tanım Sütunundan Tela Durumunu Otomatik Çıkarma
                def tela_durumu_bul(tanim):
                    tanim_upper = str(tanim).upper()
                    if "TELASIZ" in tanim_upper:
                        return "TELASIZ"
                    elif "TELALI" in tanim_upper:
                        return "TELALI"
                    else:
                        return "STANDART"

                df["Tela Durumu"] = df["Malzeme Uzun Tanımı"].apply(tela_durumu_bul)

                # 3. Siparişleri Birleştir (Tela Durumu Gruplamaya Eklendi)
                groupby_cols = ["Tel Kalınlığı", "Zone Bilgisi", "Lamet Durumu", "Tela Durumu", "En", "Boy", "Yükseklik", "Sıralama Alanı (En * Boy)"]
                df_agg = df.groupby(groupby_cols, as_index=False).agg({
                    "Bileşen Kodu": "first",
                    "Malzeme Uzun Tanımı": "first",
                    "Sipariş Miktarı": "sum",
                    "Toplam İş Süresi (Dakika)": "sum"
                })
                
                jobs = df_agg.to_dict('records')
                N = len(jobs)

                # 4. Setup Hesaplama (Tüm Kısıtlar ve Cezalar)
                def calculate_setup(job_i, job_j):
                    setup = 0
                    if job_i['Zone Bilgisi'] != job_j['Zone Bilgisi']: setup += 8
                    if str(job_i['Tel Kalınlığı']) != str(job_j['Tel Kalınlığı']): setup += 5
                    if job_i['Lamet Durumu'] != job_j['Lamet Durumu']: setup += 3
                    if job_i['Tela Durumu'] != job_j['Tela Durumu']: setup += 4
                    return setup

                # 5. MILP MODELİNİN KURULMASI
                prob = pulp.LpProblem("Pocket_Yay_Cizelgeleme_MILP", pulp.LpMinimize)
                V = list(range(N + 1))
                x = pulp.LpVariable.dicts("x", (V, V), cat='Binary')
                C = pulp.LpVariable.dicts("C", V, lowBound=0, cat='Continuous')

                # Amaç Fonksiyonu (Toplam Setup'ı Minimize Et)
                prob += pulp.lpSum(calculate_setup(jobs[i-1], jobs[j-1]) * x[i][j] for i in V[1:] for j in V[1:] if i != j)

                # Rota Kısıtları
                for j in V[1:]: prob += pulp.lpSum(x[i][j] for i in V if i != j) == 1
                for i in V[1:]: prob += pulp.lpSum(x[i][j] for j in V if i != j) == 1
                prob += pulp.lpSum(x[0][j] for j in V[1:]) == 1
                prob += pulp.lpSum(x[i][0] for i in V[1:]) == 1

                # Alt Tur Engelleme ve Zaman Kısıtları (MTZ)
                M = 10000
                for i in V[1:]:
                    for j in V[1:]:
                        if i != j:
                            p_j = jobs[j-1]["Toplam İş Süresi (Dakika)"]
                            s_ij = calculate_setup(jobs[i-1], jobs[j-1])
                            prob += C[j] >= C[i] + p_j + s_ij - M * (1 - x[i][j])

                prob += C[0] == 0

                # YENİ MÜHENDİS KISITI: Ebat (Hacim) Kesin Kısıtı
                for i in V[1:]:
                    for j in V[1:]:
                        if i != j:
                            job_i, job_j = jobs[i-1], jobs[j-1]
                            # Tel, Zone, Lamet ve Tela tamamen aynıysa (Aynı modelse) büyük olan KESİNLİKLE önce gelmeli
                            if (job_i['Tel Kalınlığı'] == job_j['Tel Kalınlığı'] and
                                job_i['Zone Bilgisi'] == job_j['Zone Bilgisi'] and
                                job_i['Lamet Durumu'] == job_j['Lamet Durumu'] and
                                job_i['Tela Durumu'] == job_j['Tela Durumu']):

                                if job_j['Sıralama Alanı (En * Boy)'] > job_i['Sıralama Alanı (En * Boy)']:
                                    prob += x[i][j] == 0

                # 6. ÇÖZÜCÜYÜ BAŞLAT
                prob.solve(pulp.PULP_CBC_CMD(timeLimit=300, msg=False))

                if pulp.LpStatus[prob.status] == 'Optimal' or pulp.LpStatus[prob.status] == 'Not Solved':
                    # Çözümden rotayı çıkar
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

                    # 7. ZAMAN, SETUP VE VARDİYA HESAPLAMALARI
                    # Setup sürelerini açıkça sütunda göster
                    setup_süreleri = [0]
                    for idx in range(1, len(rota)):
                        onceki_is_idx = rota[idx-1] - 1
                        suanki_is_idx = rota[idx] - 1
                        setup_süreleri.append(calculate_setup(jobs[onceki_is_idx], jobs[suanki_is_idx]))

                    df_sonuc["Önceki İşten Geçiş Süresi (Setup - Dk)"] = setup_süreleri

                    bitis_zamanlari = []
                    kümülatif_zaman = 0

                    # Hesaplama döngüsü (Süre + Setup eklenecek)
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

                    # Sütunları Düzenleme
                    cols_front = ['Bileşen Kodu', 'Malzeme Uzun Tanımı', 'Tela Durumu']
                    cols_middle = [c for c in df_sonuc.columns if c not in cols_front and "Başlama" not in c and "Bitiş" not in c and "Başlangıç Zamanı" not in c and "Setup" not in c and "Tela Durumu" not in c]
                    cols_end = ['Önceki İşten Geçiş Süresi (Setup - Dk)', 'Başlangıç Zamanı (Dk)', 'Başlama Haftası', 'Başlama Günü', 'Başlama Vardiyası', 'Bitiş Zamanı (Dk)', 'Bitiş Haftası', 'Bitiş Günü', 'Bitiş Vardiyası']

                    df_sonuc = df_sonuc[cols_front + cols_middle + cols_end]

                    # Excel İndirme İşlemi
               # Excel İndirme İşlemi
output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df_sonuc.to_excel(writer, index=False, sheet_name='Optimum_Plan')
    
    # GİZLİ İMZA: Excel dosyasının dijital özelliklerine adını kazıma
    workbook = writer.book
    workbook.properties.creator = "Abdullah Kerem Göktaş" 
    workbook.properties.title = "Pocket Yay Çizelgeleme Motoru"

processed_data = output.getvalue()

st.success(f"Optimizasyon başarıyla tamamlandı! Toplam işlenen satır grubu: {N}")
                    
                    st.download_button(
                        label="📥 Oluşturulan Planı Excel Olarak İndir",
                        data=processed_data,
                        file_name="Optimum_Uretim_Plani_MILP.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.dataframe(df_sonuc.head(10)) 
                else:
                    st.error("Sistem geçerli bir rota bulamadı. Lütfen kısıtların birbiriyle çakışmadığından emin olun.")
            except Exception as e:
                st.error(f"Bir hata oluştu: {e}")
