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
            df = df.iloc[:, 1:9]
            df.columns = ['Department', 'Cost_Center', 'Nature', 'Ref', 'Ref_Adj', 'Parity', 'Performance', 'Actual']
            df['Cost_Center'] = df['Cost_Center'].ffill()
            df = df.dropna(subset=['Nature'])
            df = df[df['Cost_Center'].astype(str).str.startswith('TY') & ~df['Cost_Center'].astype(str).str.contains('Toplam') & ~df['Nature'].astype(str).str.contains('Toplam')]
            df['Performance'] = pd.to_numeric(df['Performance'], errors='coerce').fillna(0)
            df['Ay_Sira'] = ay_no
            df['Ay'] = ay_adi
            return df[['Cost_Center', 'Nature', 'Performance', 'Ay', 'Ay_Sira']]

        elif sheet_type == 'FIP':
            df = pd.read_excel(file_source, sheet_name=target_sheet, skiprows=2)
            df = df.iloc[:, [2, 5, 9]]
            df.columns = ['Cost_Center', 'Masraf_Kalemi', 'Performance']
            df['Cost_Center'] = df['Cost_Center'].ffill()
            df = df.dropna(subset=['Masraf_Kalemi'])
            df = df[~df['Masraf_Kalemi'].astype(str).str.contains('Toplam|Account')]
            
            def grup_turu_bul(kalem_adi):
                match = re.search(r'([0-9][A-Za-z])', str(kalem_adi))
                if match:
                    return match.group(1).upper()
                return "DİĞER"
                
            df['Grup'] = df['Masraf_Kalemi'].apply(grup_turu_bul)
            df['Performance'] = pd.to_numeric(df['Performance'], errors='coerce').fillna(0)
            df['Ay_Sira'] = ay_no
            df['Ay'] = ay_adi
            return df[['Cost_Center', 'Masraf_Kalemi', 'Grup', 'Performance', 'Ay', 'Ay_Sira']]
            
    except Exception:
        return None

def trend_ve_yorum_uret(row, ay_kolonlari):
    if len(ay_kolonlari) < 2:
        return '➔ Takip Edilmeli', "Yetersiz Ay"
    
    son_ay = row[ay_kolonlari[-1]]
    onceki_ay = row[ay_kolonlari[-2]]
    
    if son_ay > 50:
        return '▼ Yüksek Risk / Bütçe Aşımı', "Maliyetler bütçenin oldukça üzerine çıktı."
    elif son_ay < -50:
        return '▲ Güçlü Tasarruf', "Ciddi oranda olumlu sapma / tasarruf sağlandı."
    elif son_ay < onceki_ay and son_ay < 0:
        return '▲ İyileşme Eğilimi', "Tasarruf miktarı artıyor, performans olumlu."
    elif son_ay > onceki_ay and son_ay > 0:
        return '▼ Maliyet Artışı', "Harcamalarda olumsuz yönde artış var."
    elif abs(son_ay - onceki_ay) < 10:
        return '➔ Kontrol Altında', "Stabil bir maliyet seyri izleniyor."
    else:
        return '➔ Takip Edilmeli', "Dönemsel dalgalanma gözleniyor."

def renk_stili_ver(val):
    val_str = str(val)
    if "▲" in val_str:
        return 'color: #2e7d32; font-weight: bold;'
    elif "▼" in val_str:
        return 'color: #c62828; font-weight: bold;'
    elif "➔" in val_str:
        return 'color: #ef6c00; font-weight: bold;'
    return ''

def formatli_tablo_goster(piv_df):
    format_dict = {col: "{:+,.1f}" for col in piv_df.columns if col not in ['Sonuç / Trend', 'Yönetici Analiz Yorumu']}
    
    styler = piv_df.style.format(format_dict, na_rep="0.0")
    if 'Sonuç / Trend' in piv_df.columns:
        try:
            styler = styler.map(renk_stili_ver, subset=['Sonuç / Trend'])
        except AttributeError:
            styler = styler.applymap(renk_stili_ver, subset=['Sonuç / Trend'])
        
    st.dataframe(styler, width="stretch")

def pivot_tablo_olustur(veriler, index_col):
    if not veriler:
        return None, []
    df_concat = pd.concat(veriler, ignore_index=True)
    if df_concat.empty:
        return None, []
    
    aylar_sirali = df_concat[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
    
    piv = df_concat.pivot_table(index=index_col, columns='Ay', values='Performance', aggfunc='sum', fill_value=0)
    piv = piv.reindex(columns=[col for col in aylar_sirali if col in piv.columns])
    
    piv.columns.name = None
    piv.index.name = index_col
    
    # Kesin sıralama ve TOPLAM satırının eklenmesi
    if index_col == 'Kalem':
        istenilen_sira = ['MOD', 'MOS', 'FIP', 'TAXE', 'DEPRECIATION', 'DEPRECİATİON']
        aktif_indexler = [idx for idx in piv.index if str(idx).upper() != 'TOPLAM']
        
        def sort_key(idx_val):
            val_upper = str(idx_val).upper()
            for i, k in enumerate(istenilen_sira):
                if k in val_upper:
                    return i
            return 50
            
        sirali_indexler = sorted(aktif_indexler, key=sort_key)
        piv = piv.reindex(sirali_indexler)
        
        # TOPLAM satırını her koşulda en alta hesaplayıp ekliyoruz
        numeric_cols = piv.select_dtypes(include='number').columns
        piv.loc['TOPLAM'] = 0
        for col in numeric_cols:
            if len(piv) > 1:
                piv.loc['TOPLAM', col] = piv.drop('TOPLAM', errors='ignore')[col].sum()
    else:
        numeric_cols = piv.select_dtypes(include='number').columns
        piv.loc['TOPLAM'] = 0
        for col in numeric_cols:
            if len(piv) > 1:
                piv.loc['TOPLAM', col] = piv.drop('TOPLAM', errors='ignore')[col].sum()

    ay_listesi = list(piv.columns)
    trendler, yorumlar = [], []
    for _, row in piv.iterrows():
        tr, yr = trend_ve_yorum_uret(row, ay_listesi)
        trendler.append(tr)
        yorumlar.append(yr)
        
    piv['Sonuç / Trend'] = trendler
    piv['Yönetici Analiz Yorumu'] = yorumlar
    
    toplam_row = piv.loc['TOPLAM']
    top_tr, top_yr = trend_ve_yorum_uret(toplam_row, ay_listesi)
    piv.loc['TOPLAM', 'Sonuç / Trend'] = top_tr
    piv.loc['TOPLAM', 'Yönetici Analiz Yorumu'] = f"Genel Toplam - {top_yr}"

    return piv, df_concat

# --- SIDEBAR ---
st.sidebar.header("📁 Excel Yükleme Paneli")
uploaded_files = st.sidebar.file_uploader("Aylık Excel Dosyalarını Yükleyin (Örn: M06, M07, M08)", type=["xlsx", "xls"], accept_multiple_files=True)

masse_list, cc_list, fip_list = [], [], []

if uploaded_files:
    for u_file in uploaded_files:
        try:
            m = excel_oku(u_file, u_file.name, 'Masse')
            c = excel_oku(u_file, u_file.name, 'Cost Center')
            f = excel_oku(u_file, u_file.name, 'FIP')
            
            if m is not None and not m.empty: masse_list.append(m)
            if c is not None and not c.empty: cc_list.append(c)
            if f is not None and not f.empty: fip_list.append(f)
        except Exception as e:
            st.sidebar.error(f"Hata ({u_file.name}): {e}")

# --- ANA EKRAN ---
if masse_list or cc_list or fip_list:
    piv_m, df_m_all = pivot_tablo_olustur(masse_list, 'Kalem')
    piv_c, df_c_all = pivot_tablo_olustur(cc_list, 'Cost_Center')
    piv_f, df_f_all = pivot_tablo_olustur(fip_list, 'Masraf_Kalemi')

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📋 Ana Kalemler", 
        "🏢 Cost Center Özeti", 
        "📑 Cost Center Detayları", 
        "🔍 Cost Center FIP Detayı", 
        "📂 FIP Grup Analizi (3A, 3B vb.)", 
        "🛠️ Alt Masraf Kalemleri (FIP)"
    ])

    with tab1:
        if piv_m is not None:
            st.subheader("Ana Kalemler Performans Özeti (K€)")
            formatli_tablo_goster(piv_m)
            
            if not df_m_all.empty:
                df_m_chart = df_m_all.sort_values('Ay_Sira')
                
                # Tablodaki sıra (TOPLAM hariç) grafiğe birebir aktarılıyor
                sabit_kalem_sirasi = [k for k in piv_m.index if k != 'TOPLAM']
                
                fig_m = px.bar(
                    df_m_chart, x='Kalem', y='Performance', color='Ay', barmode='group',
                    title="Ana Kalemlerin Aylara Göre Performansı (+ / - Değerler)",
                    template="plotly_white",
                    category_orders={"Kalem": sabit_kalem_sirasi}
                )
                fig_m.update_layout(
                    height=420,
                    margin=dict(l=10, r=10, t=40, b=10),
                    font=dict(size=11, color="#333"),
                    bargap=0.2,
                    bargroupgap=0.05,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title=None),
                    xaxis=dict(showgrid=False, tickangle=-20),
                    yaxis=dict(showgrid=True, gridcolor='#f2f2f2', zeroline=True, zerolinewidth=1.5, zerolinecolor='#666')
                )
                st.plotly_chart(fig_m, width="stretch")

    with tab2:
        if piv_c is not None:
            st.subheader("🏢 Cost Center (Masraf Yeri) Yönetici Özeti")
            formatli_tablo_goster(piv_c)
            st.markdown("---")
            if not df_c_all.empty:
                df_c_toplam = df_c_all.groupby('Cost_Center')['Performance'].sum().reset_index()
                df_c_toplam = df_c_toplam.sort_values('Performance', ascending=True)
                
                fig_c = px.bar(
                    df_c_toplam, x='Performance', y='Cost_Center', orientation='h',
                    title="Cost Center Bazlı Net Performans Dağılımı (K€)",
                    template="plotly_white"
                )
                fig_c.update_layout(
                    height=380,
                    margin=dict(l=10, r=10, t=40, b=10),
                    font=dict(size=11, color="#333"),
                    xaxis=dict(showgrid=True, gridcolor='#f2f2f2', zeroline=True, zerolinewidth=1.5, zerolinecolor='#666'),
                    yaxis=dict(showgrid=False)
                )
                st.plotly_chart(fig_c, width="stretch")

    with tab3:
        df_c_raw = pd.concat(cc_list, ignore_index=True) if cc_list else pd.DataFrame()
        if not df_c_raw.empty:
            st.subheader("📑 Cost Center Alt Kırılım Detayları")
            unique_cc_summary = sorted(df_c_raw['Cost_Center'].dropna().unique().tolist())
            selected_cc_summary = st.selectbox("İncelemek İstediğiniz Cost Center Numarasını Seçin:", unique_cc_summary, key="cc_summary_selectbox")
            
            if selected_cc_summary:
                df_cc_filtered = df_c_raw[df_c_raw['Cost_Center'] == selected_cc_summary]
                aylar_sirali_cc = df_cc_filtered[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
                
                piv_cc_nature = df_cc_filtered.pivot_table(
                    index='Nature', columns='Ay', values='Performance', aggfunc='sum', fill_value=0
                )
                piv_cc_nature = piv_cc_nature.reindex(columns=[col for col in aylar_sirali_cc if col in piv_cc_nature.columns])
                piv_cc_nature.columns.name = None
                piv_cc_nature.index.name = 'Nature'
                
                numeric_cols_cc = piv_cc_nature.select_dtypes(include='number').columns
                piv_cc_nature.loc['TOPLAM'] = 0
                for col in numeric_cols_cc:
                    if len(piv_cc_nature) > 1:
                        piv_cc_nature.loc['TOPLAM', col] = piv_cc_nature.drop('TOPLAM', errors='ignore')[col].sum()
                
                ay_listesi_cc = list(piv_cc_nature.columns)
                t_list, y_list = [], []
                for _, row in piv_cc_nature.iterrows():
                    tr, yr = trend_ve_yorum_uret(row, ay_listesi_cc)
                    t_list.append(tr)
                    y_list.append(yr)
                piv_cc_nature['Sonuç / Trend'] = t_list
                piv_cc_nature['Yönetici Analiz Yorumu'] = y_list
                
                top_row_cc = piv_cc_nature.loc['TOPLAM']
                t_tr, t_yr = trend_ve_yorum_uret(top_row_cc, ay_listesi_cc)
                piv_cc_nature.loc['TOPLAM', 'Sonuç / Trend'] = t_tr
                piv_cc_nature.loc['TOPLAM', 'Yönetici Analiz Yorumu'] = f"Toplam - {t_yr}"
                    
                st.markdown(f"**{selected_cc_summary}** Numaralı Merkeze Ait Alt Kalemler (K€)")
                formatli_tablo_goster(piv_cc_nature)
                
                df_cc_toplam_nature = df_cc_filtered.groupby('Nature')['Performance'].sum().reset_index()
                fig_cc_nature = px.bar(
                    df_cc_toplam_nature, x='Performance', y='Nature', orientation='h',
                    title=f"{selected_cc_summary} - Alt Kalemlerin Net Dağılımı (K€)",
                    template="plotly_white"
                )
                fig_cc_nature.update_layout(
                    height=350,
                    margin=dict(l=10, r=10, t=40, b=10),
                    font=dict(size=11, color="#333"),
                    yaxis={'categoryorder':'total ascending', 'showgrid': False},
                    xaxis=dict(showgrid=True, gridcolor='#f2f2f2', zeroline=True, zerolinewidth=1.5, zerolinecolor='#666')
                )
                st.plotly_chart(fig_cc_nature, width="stretch")
        else:
            st.info("Cost Center detay verisi bulunamadı.")

    with tab4:
        df_f_raw = pd.concat(fip_list, ignore_index=True) if fip_list else pd.DataFrame()
        if not df_f_raw.empty:
            st.subheader("🛠️ Cost Center Bazlı Alt Masraf Kalemleri")
            unique_ccs = sorted(df_f_raw['Cost_Center'].dropna().unique().tolist())
            selected_cc = st.selectbox("İncelemek İstediğiniz Cost Center Numarasını Seçin:", unique_ccs, key="cc_fip_select")
            
            if selected_cc:
                df_filtered = df_f_raw[df_f_raw['Cost_Center'] == selected_cc]
                aylar_sirali = df_filtered[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
                
                piv_fip_cc = df_filtered.pivot_table(
                    index='Masraf_Kalemi', columns='Ay', values='Performance', aggfunc='sum', fill_value=0
                )
                piv_fip_cc = piv_fip_cc.reindex(columns=[col for col in aylar_sirali if col in piv_fip_cc.columns])
                piv_fip_cc.columns.name = None
                piv_fip_cc.index.name = 'Masraf_Kalemi'
                
                numeric_cols_fip = piv_fip_cc.select_dtypes(include='number').columns
                piv_fip_cc.loc['TOPLAM'] = 0
                for col in numeric_cols_fip:
                    if len(piv_fip_cc) > 1:
                        piv_fip_cc.loc['TOPLAM', col] = piv_fip_cc.drop('TOPLAM', errors='ignore')[col].sum()
                
                ay_listesi_fip = list(piv_fip_cc.columns)
                t_fip_list, y_fip_list = [], []
                for _, row in piv_fip_cc.iterrows():
                    tr, yr = trend_ve_yorum_uret(row, ay_listesi_fip)
                    t_fip_list.append(tr)
                    y_fip_list.append(yr)
                piv_fip_cc['Sonuç / Trend'] = t_fip_list
                piv_fip_cc['Yönetici Analiz Yorumu'] = y_fip_list

                top_row_fip = piv_fip_cc.loc['TOPLAM']
                t_tr_f, t_yr_f = trend_ve_yorum_uret(top_row_fip, ay_listesi_fip)
                piv_fip_cc.loc['TOPLAM', 'Sonuç / Trend'] = t_tr_f
                piv_fip_cc.loc['TOPLAM', 'Yönetici Analiz Yorumu'] = f"Toplam - {t_yr_f}"

                st.markdown(f"**{selected_cc}** Numaralı Merkeze Ait FIP Kalemleri (K€)")
                formatli_tablo_goster(piv_fip_cc)
                
                df_f_toplam_cc = df_filtered.groupby('Masraf_Kalemi')['Performance'].sum().reset_index()
                fig_f_cc = px.bar(
                    df_f_toplam_cc, x='Performance', y='Masraf_Kalemi', orientation='h',
                    title=f"{selected_cc} - FIP Kalemleri Sıralaması (K€)",
                    template="plotly_white"
                )
                fig_f_cc.update_layout(
                    height=400,
                    margin=dict(l=10, r=10, t=40, b=10),
                    font=dict(size=11, color="#333"),
                    yaxis={'categoryorder':'total ascending', 'showgrid': False},
                    xaxis=dict(showgrid=True, gridcolor='#f2f2f2', zeroline=True, zerolinewidth=1.5, zerolinecolor='#666')
                )
                st.plotly_chart(fig_f_cc, width="stretch")
        else:
            st.info("FIP verisi bulunamadı.")

    with tab5:
        df_f_raw_group = pd.concat(fip_list, ignore_index=True) if fip_list else pd.DataFrame()
        if not df_f_raw_group.empty:
            st.subheader("📂 FIP Grup Bazlı İnceleme (3A, 3B vb.)")
            
            unique_gruplar = sorted(df_f_raw_group['Grup'].dropna().unique().tolist())
            selected_grup = st.selectbox("İncelemek İstediğiniz FIP Grubunu Seçin:", unique_gruplar, key="grup_select_tab5")
            
            if selected_grup:
                df_grup_filtered = df_f_raw_group[df_f_raw_group['Grup'] == selected_grup]
                aylar_sirali_grup = df_grup_filtered[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
                
                piv_grup_detay = df_grup_filtered.pivot_table(
                    index='Masraf_Kalemi', columns='Ay', values='Performance', aggfunc='sum', fill_value=0
                )
                piv_grup_detay = piv_grup_detay.reindex(columns=[col for col in aylar_sirali_grup if col in piv_grup_detay.columns])
                piv_grup_detay.columns.name = None
                piv_grup_detay.index.name = 'Masraf_Kalemi'
                
                numeric_cols_gd = piv_grup_detay.select_dtypes(include='number').columns
                piv_grup_detay.loc['TOPLAM'] = 0
                for col in numeric_cols_gd:
                    if len(piv_grup_detay) > 1:
                        piv_grup_detay.loc['TOPLAM', col] = piv_grup_detay.drop('TOPLAM', errors='ignore')[col].sum()

                ay_listesi_g = list(piv_grup_detay.columns)
                tg_list, yg_list = [], []
                for _, row in piv_grup_detay.iterrows():
                    tr, yr = trend_ve_yorum_uret(row, ay_listesi_g)
                    tg_list.append(tr)
                    yg_list.append(yr)
                piv_grup_detay['Sonuç / Trend'] = tg_list
                piv_grup_detay['Yönetici Analiz Yorumu'] = yg_list
                
                top_row_g = piv_grup_detay.loc['TOPLAM']
                g_tr, g_yr = trend_ve_yorum_uret(top_row_g, ay_listesi_g)
                piv_grup_detay.loc['TOPLAM', 'Sonuç / Trend'] = g_tr
                piv_grup_detay.loc['TOPLAM', 'Yönetici Analiz Yorumu'] = f"Grup Toplamı - {g_yr}"
                    
                st.markdown(f"**{selected_grup}** Grubuna Ait Detaylar")
                formatli_tablo_goster(piv_grup_detay)
                
                df_grup_toplam = df_grup_filtered.groupby('Masraf_Kalemi')['Performance'].sum().reset_index()
                fig_grup = px.bar(
                    df_grup_toplam, x='Performance', y='Masraf_Kalemi', orientation='h',
                    title=f"{selected_grup} Grubu Kalemleri (K€)",
                    template="plotly_white"
                )
                fig_grup.update_layout(
                    height=380,
                    margin=dict(l=10, r=10, t=40, b=10),
                    font=dict(size=11, color="#333"),
                    yaxis={'categoryorder':'total ascending', 'showgrid': False},
                    xaxis=dict(showgrid=True, gridcolor='#f2f2f2', zeroline=True, zerolinewidth=1.5, zerolinecolor='#666')
                )
                st.plotly_chart(fig_grup, width="stretch")
        else:
            st.info("FIP grup verisi bulunamadı.")

    with tab6:
        if piv_f is not None:
            st.subheader("Alt Masraf Kalemleri Detaylı Analizi")
            formatli_tablo_goster(piv_f)
            df_f_chart = pd.concat(fip_list, ignore_index=True) if fip_list else pd.DataFrame()
            if not df_f_chart.empty:
                df_f_toplam_all = df_f_chart.groupby('Masraf_Kalemi')['Performance'].sum().reset_index()
                fig_f = px.bar(
                    df_f_toplam_all, x='Performance', y='Masraf_Kalemi', orientation='h',
                    title="Tüm Alt Masraf Kalemleri Genel Dağılımı (K€)",
                    template="plotly_white"
                )
                fig_f.update_layout(
                    height=420,
                    margin=dict(l=10, r=10, t=40, b=10),
                    font=dict(size=11, color="#333"),
                    yaxis={'categoryorder':'total ascending', 'showgrid': False},
                    xaxis=dict(showgrid=True, gridcolor='#f2f2f2', zeroline=True, zerolinewidth=1.5, zerolinecolor='#666')
                )
                st.plotly_chart(fig_f, width="stretch")

else:
    st.info("💡 Karşılaştırma tabloları ve trend grafiklerinin oluşması için sol panelden aylık Excel dosyalarınızı yükleyin.")
