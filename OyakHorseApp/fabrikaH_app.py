import os
import re
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Finansal Performans Özeti", layout="wide")

st.title("📊 Motor Perimetresi Finansal Performans Özeti (K€)")

# 🗓️ Kronolojik Ay Tanım Haritası
AYLAR_MAP = {
    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
    7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
}

def ay_bilgisi_al(dosya_adi):
    match = re.search(r'M(\d{2})', dosya_adi, re.IGNORECASE)
    if match:
        ay_no = int(match.group(1))
        if 1 <= ay_no <= 12:
            return ay_no, AYLAR_MAP[ay_no]
    return 99, dosya_adi

def excel_oku(file_source, dosya_adi, sheet_type='Masse'):
    ay_no, ay_adi = ay_bilgisi_al(dosya_adi)
    try:
        xl = pd.ExcelFile(file_source)
    except Exception:
        return None
        
    sheet_names = xl.sheet_names
    
    target_sheet = None
    for s in sheet_names:
        if sheet_type.lower() in s.lower():
            target_sheet = s
            break
            
    if not target_sheet:
        return None

    try:
        if sheet_type == 'Masse':
            df = pd.read_excel(file_source, sheet_name=target_sheet, skiprows=5)
            df = df.iloc[:, 1:7]
            df.columns = ['Kalem', 'Reference', 'Ref_Semi_Adjusted', 'Parity_Impact', 'Performance', 'Actual']
            df = df.dropna(subset=['Kalem'])
            df = df[~df['Kalem'].astype(str).str.lower().str.contains('nature')]
            df['Performance'] = pd.to_numeric(df['Performance'], errors='coerce').fillna(0)
            df['Ay_Sira'] = ay_no
            df['Ay'] = ay_adi
            return df[['Kalem', 'Performance', 'Ay', 'Ay_Sira']]

        elif sheet_type == 'Cost Center':
            df = pd.read_excel(file_source, sheet_name=target_sheet, skiprows=1)
            df = df.iloc[:, 2:8]
            df.columns = ['Cost_Center', 'Nature', 'Ref', 'Ref_Adj', 'Parity', 'Performance']
            df = df.dropna(subset=['Cost_Center'])
            df['Cost_Center'] = df['Cost_Center'].ffill()
            df = df[df['Cost_Center'].astype(str).str.startswith('TY') & ~df['Cost_Center'].astype(str).str.contains('Toplam')]
            df['Performance'] = pd.to_numeric(df['Performance'], errors='coerce').fillna(0)
            df['Ay_Sira'] = ay_no
            df['Ay'] = ay_adi
            return df[['Cost_Center', 'Performance', 'Ay', 'Ay_Sira']]

        elif sheet_type == 'FIP':
            df = pd.read_excel(file_source, sheet_name=target_sheet, skiprows=2)
            df = df.iloc[:, [2, 5, 9]]
            df.columns = ['Cost_Center', 'Masraf_Kalemi', 'Performance']
            df['Cost_Center'] = df['Cost_Center'].ffill()
            df = df.dropna(subset=['Masraf_Kalemi'])
            df = df[~df['Masraf_Kalemi'].astype(str).str.contains('Toplam|Account')]
            df['Performance'] = pd.to_numeric(df['Performance'], errors='coerce').fillna(0)
            df['Ay_Sira'] = ay_no
            df['Ay'] = ay_adi
            return df[['Cost_Center', 'Masraf_Kalemi', 'Performance', 'Ay', 'Ay_Sira']]
            
    except Exception:
        return None

def trend_ve_yorum_uret(row, ay_kolonlari):
    if len(ay_kolonlari) < 2:
        return "🟡 Veri Yetersiz", "Yetersiz Ay"
    
    son_ay = row[ay_kolonlari[-1]]
    onceki_ay = row[ay_kolonlari[-2]]
    
    if son_ay > 50:
        return "🔴 En kritik merkez", "Bütçe aşımı yüksek risk oluşturuyor."
    elif son_ay < -50:
        return "🟢 Güçlü performans", "Ciddi tasarruf / olumlu sapma sağlandı."
    elif son_ay < onceki_ay and son_ay < 0:
        return "🟢 İyileşme / Tasarruf artıyor", "Performans olumlu yönde gelişiyor."
    elif son_ay > onceki_ay and son_ay > 0:
        return "🔴 Bozulma var / Artan Sapma", "Performans olumsuza kayıyor."
    elif abs(son_ay - onceki_ay) < 10:
        return "🟡 Kontrol altında", "Stabil seyir devam ediyor."
    else:
        return "🟠 Takip edilmeli", "Aylık dalgalanma mevcut."

def pivot_tablo_olustur(veriler, index_col):
    if not veriler:
        return None, []
    df_concat = pd.concat(veriler, ignore_index=True)
    
    aylar_sirali = df_concat[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
    
    piv = df_concat.pivot_table(index=index_col, columns='Ay', values='Performance', aggfunc='sum', fill_value=0)
    piv = piv.reindex(columns=[col for col in aylar_sirali if col in piv.columns])
    
    ay_listesi = list(piv.columns)
    trendler, yorumlar = [], []
    for _, row in piv.iterrows():
        tr, yr = trend_ve_yorum_uret(row, ay_listesi)
        trendler.append(tr)
        yorumlar.append(yr)
        
    piv['Sonuç / Trend'] = trendler
    piv['Yönetici Analiz Yorumu'] = yorumlar
    
    # TOPLAM satırı ekleme
    numeric_cols = piv.select_dtypes(include='number').columns
    piv.loc['TOPLAM'] = 0
    for col in numeric_cols:
        piv.loc['TOPLAM', col] = piv[col].iloc[:-1].sum()
    piv.loc['TOPLAM', 'Sonuç / Trend'] = '-'
    piv.loc['TOPLAM', 'Yönetici Analiz Yorumu'] = 'Genel Toplam'

    return piv, df_concat

# --- SIDEBAR (Sadece Yükleme Alanı) ---
st.sidebar.header("📁 Excel Yükleme Paneli")
uploaded_files = st.sidebar.file_uploader("Aylık Excel Dosyalarını Yükleyin (Örn: M06, M07, M08)", type=["xlsx", "xls"], accept_multiple_files=True)

masse_list, cc_list, fip_list = [], [], []

if uploaded_files:
    for u_file in uploaded_files:
        try:
            m = excel_oku(u_file, u_file.name, 'Masse')
            c = excel_oku(u_file, u_file.name, 'Cost Center')
            f = excel_oku(u_file, u_file.name, 'FIP')
            
            if m is not None: masse_list.append(m)
            if c is not None: cc_list.append(c)
            if f is not None: fip_list.append(f)
        except Exception as e:
            st.sidebar.error(f"Hata ({u_file.name}): {e}")

# --- ANA EKRAN ---
if masse_list or cc_list or fip_list:
    piv_m, df_m_all = pivot_tablo_olustur(masse_list, 'Kalem')
    piv_c, df_c_all = pivot_tablo_olustur(cc_list, 'Cost_Center')
    piv_f, df_f_all = pivot_tablo_olustur(fip_list, 'Masraf_Kalemi')  # Orijinal genel FIP tablosu

    # 4 SEKMELİ YAPI
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 Ana Kalemler", 
        "🏢 Cost Center Özeti", 
        "🛠️ Alt Masraf Kalemleri (FIP)", 
        "🔍 Cost Center FIP Detayı"
    ])

    with tab1:
        if piv_m is not None:
            st.subheader("Ana Kalemler Performans Özeti (K€)")
            st.dataframe(
                piv_m.style.format({col: "{:+,.1f}" for col in piv_m.columns if col not in ['Sonuç / Trend', 'Yönetici Analiz Yorumu']}), 
                width="stretch"
            )
            
            df_m_chart = df_m_all.sort_values('Ay_Sira')
            fig_m = px.line(
                df_m_chart, x='Ay', y='Performance', color='Kalem', markers=True,
                title="Ana Kalemlerin Kronolojik Değişim Grafiği (K€)"
            )
            st.plotly_chart(fig_m, width="stretch")

    with tab2:
        if piv_c is not None:
            st.subheader("Cost Center (Masraf Yeri) Yönetici Özeti")
            st.dataframe(
                piv_c.style.format({col: "{:+,.1f}" for col in piv_c.columns if col not in ['Sonuç / Trend', 'Yönetici Analiz Yorumu']}), 
                width="stretch"
            )
            
            df_c_chart = df_c_all.sort_values('Ay_Sira')
            fig_c = px.bar(
                df_c_chart, x='Cost_Center', y='Performance', color='Ay', barmode='group',
                title="Cost Center'ların Aylık Performans Karşılaştırma Grafiği (K€)"
            )
            st.plotly_chart(fig_c, width="stretch")

    with tab3:
        # Ekran görüntüsündeki orijinal Alt Masraf Kalemleri (FIP) Genel Görünümü
        if piv_f is not None:
            st.subheader("Alt Masraf Kalemleri Detaylı Analizi")
            st.dataframe(
                piv_f.style.format({col: "{:+,.1f}" for col in piv_f.columns if col not in ['Sonuç / Trend', 'Yönetici Analiz Yorumu']}), 
                width="stretch"
            )
            
            df_f_chart = pd.concat(fip_list, ignore_index=True).sort_values('Ay_Sira') if fip_list else pd.DataFrame()
            if not df_f_chart.empty:
                fig_f = px.bar(
                    df_f_chart, x='Performance', y='Masraf_Kalemi', color='Ay', orientation='h',
                    title="Alt Masraf Kalemleri Dağılım Grafiği (K€)"
                )
                st.plotly_chart(fig_f, width="stretch")

    with tab4:
        # Yeni eklenen Cost Center bazlı filtreleme sekmesi
        df_f_raw = pd.concat(fip_list, ignore_index=True) if fip_list else pd.DataFrame()
        if not df_f_raw.empty:
            st.subheader("🛠️ Cost Center Bazlı Alt Masraf Kalemleri Arama ve İnceleme")
            
            unique_ccs = sorted(df_f_raw['Cost_Center'].dropna().unique().tolist())
            selected_cc = st.selectbox("İncelemek İstediğiniz Cost Center Numarasını Seçin:", unique_ccs)
            
            if selected_cc:
                df_filtered = df_f_raw[df_f_raw['Cost_Center'] == selected_cc]
                
                aylar_sirali = df_filtered[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
                
                piv_fip_cc = df_filtered.pivot_table(
                    index='Masraf_Kalemi', 
                    columns='Ay', 
                    values='Performance', 
                    aggfunc='sum', 
                    fill_value=0
                )
                piv_fip_cc = piv_fip_cc.reindex(columns=[col for col in aylar_sirali if col in piv_fip_cc.columns])
                
                # FIP CC Tablosu için TOPLAM satırı
                numeric_cols_fip = piv_fip_cc.select_dtypes(include='number').columns
                piv_fip_cc.loc['TOPLAM'] = 0
                for col in numeric_cols_fip:
                    piv_fip_cc.loc['TOPLAM', col] = piv_fip_cc[col].iloc[:-1].sum()
                
                st.markdown(f"**{selected_cc}** Numaralı Merkeze Ait FIP Kalemleri Performans Özeti (K€)")
                st.dataframe(
                    piv_fip_cc.style.format({col: "{:+,.1f}" for col in piv_fip_cc.columns}),
                    width="stretch"
                )
                
                fig_f_cc = px.bar(
                    df_filtered, x='Performance', y='Masraf_Kalemi', color='Ay', orientation='h',
                    title=f"{selected_cc} - Alt Masraf Kalemleri Dağılım Grafiği (K€)"
                )
                st.plotly_chart(fig_f_cc, width="stretch")
        else:
            st.info("FIP verisi bulunamadı.")

else:
    st.info("💡 Karşılaştırma tabloları ve trend grafiklerinin oluşması için sol panelden aylık Excel dosyalarınızı yükleyin.")