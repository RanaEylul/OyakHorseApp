import os
import re
import pandas as pd
import streamlit as st
import plotly.express as px

# 📁 Tamamen İzole Edilmiş Ayrı Kalıcı Depolama Klasörleri
CUMUL_DIR = "uploaded_cumul_files"
MONTHLY_DIR = "uploaded_monthly_files"

for d in [CUMUL_DIR, MONTHLY_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

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

def loupes_sayfasini_oku_tum_aylar(saved_files_list):
    loupes_list = []
    # Doğru bütçe ve reel indeks eşleştirmeleri: B4D (5, 6), HR16 (10, 11), HR18 (15, 16)
    proje_sutunlari = [
        ("B4D", 5, 6),
        ("HR16", 10, 11),
        ("HR18", 15, 16)
    ]
    
    for f_path in saved_files_list:
        f_name = os.path.basename(f_path)
        ay_no, ay_adi = ay_bilgisi_al(f_name)
        try:
            xl = pd.ExcelFile(f_path)
            loupes_sheet = None
            for s in xl.sheet_names:
                if 'loupe' in s.lower():
                    loupes_sheet = s
                    break
            if not loupes_sheet:
                continue
                
            df = pd.read_excel(f_path, sheet_name=loupes_sheet)
            for proje, col_bud, col_reel in proje_sutunlari:
                try:
                    bud_val = pd.to_numeric(df.iloc[3, col_bud], errors='coerce') or 0.0
                    reel_val = pd.to_numeric(df.iloc[3, col_reel], errors='coerce') or 0.0
                    loupes_list.append({
                        "Ay": ay_adi,
                        "Ay_Sira": ay_no,
                        "Proje": proje,
                        "BUDGET": bud_val,
                        "REEL": reel_val
                    })
                except Exception:
                    pass
        except Exception:
            continue
            
    if not loupes_list:
        return pd.DataFrame(), []
        
    df_l_all = pd.DataFrame(loupes_list)
    aylar_sirali = df_l_all[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
    
    piv_reel = df_l_all.pivot_table(index='Proje', columns='Ay', values='REEL', aggfunc='sum', fill_value=0)
    piv_bud = df_l_all.pivot_table(index='Proje', columns='Ay', values='BUDGET', aggfunc='sum', fill_value=0)
    
    piv_reel = piv_reel.reindex(columns=[c for c in aylar_sirali if c in piv_reel.columns])
    piv_bud = piv_bud.reindex(columns=[c for c in aylar_sirali if c in piv_bud.columns])
    
    l_comb = pd.DataFrame(index=piv_reel.index)
    for col in piv_reel.columns:
        l_comb[f"{col} - REEL"] = piv_reel[col]
        l_comb[f"{col} - BUDGET"] = piv_bud[col]
        l_comb[f"{col} - Sapma"] = piv_reel[col] - piv_bud[col]
        
    t_list, y_list = [], []
    for _, row in l_comb.iterrows():
        last_ay = aylar_sirali[-1] if aylar_sirali else None
        if last_ay and f"{last_ay} - Sapma" in row:
            fark = row[f"{last_ay} - Sapma"]
            if fark > 0.1:
                t_list.append("▼ Bütçe Aşımı")
                y_list.append(f"{last_ay} ayında bütçe aşıldı.")
            elif fark < -0.1:
                t_list.append("▲ Tasarruf")
                y_list.append(f"{last_ay} ayında tasarruf sağlandı.")
            else:
                t_list.append("➔ Kontrol Altında")
                y_list.append("Bütçeye paralel.")
        else:
            t_list.append("➔ Takipte")
            y_list.append("Veri inceleniyor.")
            
    l_comb['Sonuç / Trend'] = t_list
    l_comb['Yönetici Analiz Yorumu'] = y_list
    return l_comb.reset_index(), aylar_sirali

def vtu_sayfasini_oku_tum_aylar(saved_files_list):
    vtu_aylik_ozet_list = []
    vtu_aylik_detay_list = []
    
    mappings = [
        (0, "B4D"),
        (6, "HR 16"),
        (12, "HR 18")
    ]
    
    for f_path in saved_files_list:
        f_name = os.path.basename(f_path)
        ay_no, ay_adi = ay_bilgisi_al(f_name)
        
        try:
            xl = pd.ExcelFile(f_path)
            vtu_sheet = None
            for s in xl.sheet_names:
                if 'vtu' in s.lower():
                    vtu_sheet = s
                    break
            if not vtu_sheet:
                continue
                
            df = pd.read_excel(f_path, sheet_name=vtu_sheet)
            
            for base_col, hat_adi in mappings:
                sub_df = df.iloc[3:8, base_col+3:base_col+6].copy()
                sub_df.columns = ['NATURE', 'REEL', 'BUDGET']
                sub_df = sub_df.dropna(subset=['NATURE'])
                sub_df['Hat'] = hat_adi
                sub_df['Ay'] = ay_adi
                sub_df['Ay_Sira'] = ay_no
                sub_df['REEL'] = pd.to_numeric(sub_df['REEL'], errors='coerce').fillna(0)
                sub_df['BUDGET'] = pd.to_numeric(sub_df['BUDGET'], errors='coerce').fillna(0)
                vtu_aylik_ozet_list.append(sub_df)
                
                detay_df = df.iloc[13:34, [base_col+1, base_col+2, base_col+3, base_col+4, base_col+5]].copy()
                detay_df.columns = ['KOD / BAŞLIK', 'GM (TANIM)', 'NATURE', 'REEL', 'BUDGET']
                detay_df = detay_df.dropna(subset=['KOD / BAŞLIK', 'GM (TANIM)', 'NATURE'], how='all')
                detay_df['NATURE'] = detay_df['NATURE'].fillna('')
                detay_df['Hat'] = hat_adi
                detay_df['Ay'] = ay_adi
                detay_df['Ay_Sira'] = ay_no
                detay_df['REEL'] = pd.to_numeric(detay_df['REEL'], errors='coerce').fillna(0)
                detay_df['BUDGET'] = pd.to_numeric(detay_df['BUDGET'], errors='coerce').fillna(0)
                vtu_aylik_detay_list.append(detay_df)
        except Exception:
            continue

    if not vtu_aylik_ozet_list:
        return {}

    df_ozet_all = pd.concat(vtu_aylik_ozet_list, ignore_index=True)
    df_detay_all = pd.concat(vtu_aylik_detay_list, ignore_index=True)
    aylar_sirali = df_ozet_all[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
    
    hierarchical_blocks = {}
    for hat_adi in ["B4D", "HR 16", "HR 18"]:
        o_hat = df_ozet_all[df_ozet_all['Hat'] == hat_adi]
        d_hat = df_detay_all[df_detay_all['Hat'] == hat_adi]
        
        if o_hat.empty and d_hat.empty:
            continue
            
        piv_o_reel = o_hat.pivot_table(index='NATURE', columns='Ay', values='REEL', aggfunc='sum', fill_value=0)
        piv_o_bud = o_hat.pivot_table(index='NATURE', columns='Ay', values='BUDGET', aggfunc='sum', fill_value=0)
        piv_o_reel = piv_o_reel.reindex(columns=[c for c in aylar_sirali if c in piv_o_reel.columns])
        piv_o_bud = piv_o_bud.reindex(columns=[c for c in aylar_sirali if c in piv_o_bud.columns])
        
        o_comb = pd.DataFrame(index=piv_o_reel.index)
        for col in piv_o_reel.columns:
            o_comb[f"{col} - REEL"] = piv_o_reel[col]
            o_comb[f"{col} - BUDGET"] = piv_o_bud[col]
            o_comb[f"{col} - Sapma"] = piv_o_reel[col] - piv_o_bud[col]
            
        t_ozet, y_ozet = [], []
        for _, row in o_comb.iterrows():
            last_ay = aylar_sirali[-1] if aylar_sirali else None
            if last_ay and f"{last_ay} - Sapma" in row:
                fark = row[f"{last_ay} - Sapma"]
                if fark > 5:
                    t_ozet.append("▼ Bütçe Aşımı"); y_ozet.append(f"{last_ay} ayında bütçe aşıldı.")
                elif fark < -5:
                    t_ozet.append("▲ Güçlü Tasarruf"); y_ozet.append(f"{last_ay} ayında olumlu tasarruf.")
                else:
                    t_ozet.append("➔ Kontrol Altında"); y_ozet.append("Bütçeye paralel seyrediyor.")
            else:
                t_ozet.append("➔ Takip Edilmeli"); y_ozet.append("Veri inceleniyor.")
                
        o_comb['Sonuç / Trend'] = t_ozet
        o_comb['Yönetici Analiz Yorumu'] = y_ozet
        o_comb = o_comb.reset_index()

        piv_d_reel = d_hat.pivot_table(index=['KOD / BAŞLIK', 'GM (TANIM)', 'NATURE'], columns='Ay', values='REEL', aggfunc='sum', fill_value=0)
        piv_d_bud = d_hat.pivot_table(index=['KOD / BAŞLIK', 'GM (TANIM)', 'NATURE'], columns='Ay', values='BUDGET', aggfunc='sum', fill_value=0)
        piv_d_reel = piv_d_reel.reindex(columns=[c for c in aylar_sirali if c in piv_d_reel.columns])
        piv_d_bud = piv_d_bud.reindex(columns=[c for c in aylar_sirali if c in piv_d_bud.columns])
        
        d_comb = pd.DataFrame(index=piv_d_reel.index)
        for col in piv_d_reel.columns:
            d_comb[f"{col} - REEL"] = piv_d_reel[col]
            d_comb[f"{col} - BUDGET"] = piv_d_bud[col]
            d_comb[f"{col} - Sapma"] = piv_d_reel[col] - piv_d_bud[col]
            
        t_detay, y_detay = [], []
        for _, row in d_comb.iterrows():
            last_ay = aylar_sirali[-1] if aylar_sirali else None
            if last_ay and f"{last_ay} - Sapma" in row:
                fark = row[f"{last_ay} - Sapma"]
                if fark > 2:
                    t_detay.append("▼ Maliyet Artışı"); y_detay.append(f"{last_ay} harcaması yüksek.")
                elif fark < -2:
                    t_detay.append("▲ İyileşme"); y_detay.append(f"{last_ay} tasarruflu.")
                else:
                    t_detay.append("➔ Stabil"); y_detay.append("Kontrol altında.")
            else:
                t_detay.append("➔ Takip Edilmeli"); y_detay.append("Veri inceleniyor.")
                
        d_comb['Sonuç / Trend'] = t_detay
        d_comb['Yönetici Analiz Yorumu'] = y_detay
        d_comb = d_comb.reset_index()

        hierarchical_blocks[hat_adi] = {
            "ozet": o_comb,
            "detay": d_comb
        }

    return hierarchical_blocks

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

veri_tipi = st.sidebar.radio(
    "Görüntülenecek Veri Modunu Seçin:",
    ["📊 Kümülatif (Cumul)", "📅 Aylık (Monthly)"],
    key="veri_tipi_secim"
)

is_cumul_mode = "Kümülatif" in veri_tipi

if is_cumul_mode:
    uploaded_files_cumul = st.sidebar.file_uploader("Kümülatif Dosyaları Yükleyin", type=["xlsx", "xls"], accept_multiple_files=True, key="up_c")
    if uploaded_files_cumul:
        for u_file in uploaded_files_cumul:
            if "monthly" in u_file.name.lower():
                st.sidebar.error("⚠️ Bu bir Aylık dosyasıdır!")
                continue
            yeni_ay_no, yeni_ay_adi = ay_bilgisi_al(u_file.name)
            if any(ay_bilgisi_al(m)[0] == yeni_ay_no for m in os.listdir(CUMUL_DIR) if yeni_ay_no != 99):
                st.sidebar.warning(f"⚠️ **{yeni_ay_adi}** zaten mevcut!")
            else:
                with open(os.path.join(CUMUL_DIR, u_file.name), "wb") as f: f.write(u_file.getbuffer())
                st.sidebar.success("✅ Eklendi.")
    saved_files = [os.path.join(CUMUL_DIR, f) for f in os.listdir(CUMUL_DIR) if f.lower().endswith(('.xlsx', '.xls'))]
    if saved_files:
        st.sidebar.markdown("---")
        for f_path in saved_files:
            f_name = os.path.basename(f_path)
            c1, c2 = st.sidebar.columns([4, 1])
            c1.text(f"• {f_name}")
            if c2.button("❌", key=f"del_c_{f_name}"): os.remove(f_path); st.rerun()
        if st.sidebar.button("🗑️ Tümünü Temizle", key="clr_c"):
            for f in saved_files: os.remove(f)
            st.rerun()
else:
    uploaded_files_monthly = st.sidebar.file_uploader("Aylık Dosyaları Yükleyin", type=["xlsx", "xls"], accept_multiple_files=True, key="up_m")
    if uploaded_files_monthly:
        for u_file in uploaded_files_monthly:
            if "cumul" in u_file.name.lower():
                st.sidebar.error("⚠️ Bu bir Kümülatif dosyasıdır!")
                continue
            yeni_ay_no, yeni_ay_adi = ay_bilgisi_al(u_file.name)
            if any(ay_bilgisi_al(m)[0] == yeni_ay_no for m in os.listdir(MONTHLY_DIR) if yeni_ay_no != 99):
                st.sidebar.warning(f"⚠️ **{yeni_ay_adi}** zaten mevcut!")
            else:
                with open(os.path.join(MONTHLY_DIR, u_file.name), "wb") as f: f.write(u_file.getbuffer())
                st.sidebar.success("✅ Eklendi.")
    saved_files = [os.path.join(MONTHLY_DIR, f) for f in os.listdir(MONTHLY_DIR) if f.lower().endswith(('.xlsx', '.xls'))]
    if saved_files:
        st.sidebar.markdown("---")
        for f_path in saved_files:
            f_name = os.path.basename(f_path)
            c1, c2 = st.sidebar.columns([4, 1])
            c1.text(f"• {f_name}")
            if c2.button("❌", key=f"del_m_{f_name}"): os.remove(f_path); st.rerun()
        if st.sidebar.button("🗑️ Tümünü Temizle", key="clr_m"):
            for f in saved_files: os.remove(f)
            st.rerun()

masse_list, cc_list, fip_list = [], [], []
if saved_files:
    for file_path in saved_files:
        file_name = os.path.basename(file_path)
        try:
            m = excel_oku(file_path, file_name, 'Masse')
            c = excel_oku(file_path, file_name, 'Cost Center')
            f = excel_oku(file_path, file_name, 'FIP')
            if m is not None and not m.empty: masse_list.append(m)
            if c is not None and not c.empty: cc_list.append(c)
            if f is not None and not f.empty: fip_list.append(f)
        except Exception: pass

# --- ANA EKRAN ---
if masse_list or cc_list or fip_list:
    st.info(f"Aktif Görüntülenen Mod: **{veri_tipi}**")
    
    piv_m, df_m_all = pivot_tablo_olustur(masse_list, 'Kalem')
    piv_c, df_c_all = pivot_tablo_olustur(cc_list, 'Cost_Center')
    piv_f, df_f_all = pivot_tablo_olustur(fip_list, 'Masraf_Kalemi')

    if not is_cumul_mode:
        tabs = st.tabs([
            "📋 Ana Kalemler", "🏢 Cost Center Özeti", "📑 Cost Center Detayları", 
            "🔍 Cost Center FIP Detayı", "📂 FIP Grup Analizi", "🛠️ Alt Masraf Kalemleri", "📊 VTU Özet Analizi", "🔍 Loupes Analizi"
        ])
        tab1, tab2, tab3, tab4, tab5, tab6, tab_vtu, tab_loupes = tabs
    else:
        tabs = st.tabs([
            "📋 Ana Kalemler", "🏢 Cost Center Özeti", "📑 Cost Center Detayları", 
            "🔍 Cost Center FIP Detayı", "📂 FIP Grup Analizi", "🛠️ Alt Masraf Kalemleri"
        ])
        tab1, tab2, tab3, tab4, tab5, tab6 = tabs
        tab_vtu = None
        tab_loupes = None

    with tab1:
        if piv_m is not None:
            st.subheader(f"Ana Kalemler Performans Özeti - {veri_tipi} (K€)")
            formatli_tablo_goster(piv_m)

    with tab2:
        if piv_c is not None:
            st.subheader(f"🏢 Cost Center Yönetici Özeti - {veri_tipi}")
            formatli_tablo_goster(piv_c)

    with tab3:
        df_c_raw = pd.concat(cc_list, ignore_index=True) if cc_list else pd.DataFrame()
        if not df_c_raw.empty:
            st.subheader("📑 Cost Center Alt Kırılım Detayları")
            unique_cc_summary = sorted(df_c_raw['Cost_Center'].dropna().unique().tolist())
            selected_cc_summary = st.selectbox("Cost Center Seçin:", unique_cc_summary, key="cc_sum")
            if selected_cc_summary:
                df_cc_filtered = df_c_raw[df_c_raw['Cost_Center'] == selected_cc_summary]
                aylar_sirali_cc = df_cc_filtered[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
                piv_cc_nature = df_cc_filtered.pivot_table(index='Nature', columns='Ay', values='Performance', aggfunc='sum', fill_value=0)
                piv_cc_nature = piv_cc_nature.reindex(columns=[col for col in aylar_sirali_cc if col in piv_cc_nature.columns])
                piv_cc_nature.loc['TOPLAM'] = piv_cc_nature.sum()
                t_list, y_list = [], []
                for _, row in piv_cc_nature.iterrows():
                    tr, yr = trend_ve_yorum_uret(row, list(piv_cc_nature.columns))
                    t_list.append(tr); y_list.append(yr)
                piv_cc_nature['Sonuç / Trend'] = t_list
                piv_cc_nature['Yönetici Analiz Yorumu'] = y_list
                formatli_tablo_goster(piv_cc_nature)

    with tab4:
        df_f_raw = pd.concat(fip_list, ignore_index=True) if fip_list else pd.DataFrame()
        if not df_f_raw.empty:
            st.subheader("🛠️ Cost Center Bazlı Alt Masraf Kalemleri")
            unique_ccs = sorted(df_f_raw['Cost_Center'].dropna().unique().tolist())
            selected_cc = st.selectbox("Cost Center Seçin:", unique_ccs, key="cc_fip")
            if selected_cc:
                df_filtered = df_f_raw[df_f_raw['Cost_Center'] == selected_cc]
                aylar_sirali = df_filtered[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
                piv_fip_cc = df_filtered.pivot_table(index='Masraf_Kalemi', columns='Ay', values='Performance', aggfunc='sum', fill_value=0)
                piv_fip_cc = piv_fip_cc.reindex(columns=[col for col in aylar_sirali if col in piv_fip_cc.columns])
                piv_fip_cc.loc['TOPLAM'] = piv_fip_cc.sum()
                t_list, y_list = [], []
                for _, row in piv_fip_cc.iterrows():
                    tr, yr = trend_ve_yorum_uret(row, list(piv_fip_cc.columns))
                    t_list.append(tr); y_list.append(yr)
                piv_fip_cc['Sonuç / Trend'] = t_list
                piv_fip_cc['Yönetici Analiz Yorumu'] = y_list
                formatli_tablo_goster(piv_fip_cc)

    with tab5:
        df_f_raw_group = pd.concat(fip_list, ignore_index=True) if fip_list else pd.DataFrame()
        if not df_f_raw_group.empty:
            st.subheader("📂 FIP Grup Bazlı İnceleme")
            unique_gruplar = sorted(df_f_raw_group['Grup'].dropna().unique().tolist())
            selected_grup = st.selectbox("FIP Grubu Seçin:", unique_gruplar, key="grup_sel")
            if selected_grup:
                df_grup_filtered = df_f_raw_group[df_f_raw_group['Grup'] == selected_grup]
                aylar_sirali_grup = df_grup_filtered[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
                piv_grup_detay = df_grup_filtered.pivot_table(index='Masraf_Kalemi', columns='Ay', values='Performance', aggfunc='sum', fill_value=0)
                piv_grup_detay = piv_grup_detay.reindex(columns=[col for col in aylar_sirali_grup if col in piv_grup_detay.columns])
                piv_grup_detay.loc['TOPLAM'] = piv_grup_detay.sum()
                t_list, y_list = [], []
                for _, row in piv_grup_detay.iterrows():
                    tr, yr = trend_ve_yorum_uret(row, list(piv_grup_detay.columns))
                    t_list.append(tr); y_list.append(yr)
                piv_grup_detay['Sonuç / Trend'] = t_list
                piv_grup_detay['Yönetici Analiz Yorumu'] = y_list
                formatli_tablo_goster(piv_grup_detay)

    with tab6:
        if piv_f is not None:
            st.subheader("Alt Masraf Kalemleri Detaylı Analizi")
            formatli_tablo_goster(piv_f)

    # --- VTU ÖZET ANALİZİ ---
    if tab_vtu is not None:
        with tab_vtu:
            st.subheader("📊 VTU Proje ve Hat Detayları (Çoklu Ay Karşılaştırmalı)")
            
            tum_vtu_bloklari = vtu_sayfasini_oku_tum_aylar(saved_files)
                
            if tum_vtu_bloklari:
                secilen_hat = st.selectbox(
                    "İncelemek İstediğiniz Ana Hattı Seçin (B4D, HR 16, HR 18 vb.):", 
                    list(tum_vtu_bloklari.keys()),
                    key="vtu_hat_select"
                )
                
                if secilen_hat:
                    veri_paket = tum_vtu_bloklari[secilen_hat]
                    
                    st.markdown(f"### 📌 **{secilen_hat}** - Aylara Göre Üst Özet Tablo")
                    format_dict = {col: "{:+,.2f}" if "Sapma" in col else "{:,.2f}" for col in veri_paket["ozet"].columns if col not in ['NATURE', 'Sonuç / Trend', 'Yönetici Analiz Yorumu']}
                    
                    ozet_styler = veri_paket["ozet"].style.format(format_dict, na_rep="0.0")
                    if 'Sonuç / Trend' in veri_paket["ozet"].columns:
                        try:
                            ozet_styler = ozet_styler.map(renk_stili_ver, subset=['Sonuç / Trend'])
                        except AttributeError:
                            ozet_styler = ozet_styler.applymap(renk_stili_ver, subset=['Sonuç / Trend'])
                    st.dataframe(ozet_styler, width="stretch", hide_index=True)
                    
                    st.markdown(f"### 🔍 **{secilen_hat}** - Aylara Göre Alt Detay Operasyonları")
                    format_dict_detay = {col: "{:+,.2f}" if "Sapma" in col else "{:,.2f}" for col in veri_paket["detay"].columns if col not in ['KOD / BAŞLIK', 'GM (TANIM)', 'NATURE', 'Sonuç / Trend', 'Yönetici Analiz Yorumu']}
                    
                    detay_styler = veri_paket["detay"].style.format(format_dict_detay, na_rep="0.0")
                    if 'Sonuç / Trend' in veri_paket["detay"].columns:
                        try:
                            detay_styler = detay_styler.map(renk_stili_ver, subset=['Sonuç / Trend'])
                        except AttributeError:
                            detay_styler = detay_styler.applymap(renk_stili_ver, subset=['Sonuç / Trend'])
                    st.dataframe(detay_styler, width="stretch", hide_index=True)
            else:
                st.info("Yüklenen aylık dosyalar içinde VTU verisi bulunamadı.")

    # --- LOUPES ANALİZİ ---
    if tab_loupes is not None:
        with tab_loupes:
            st.subheader("🔍 Loupes Sayfası Çoklu Ay Karşılaştırma ve Proje Özeti (B4D, HR16, HR18)")
            df_loupes_comb, aylar_listesi = loupes_sayfasini_oku_tum_aylar(saved_files)
            
            if not df_loupes_comb.empty:
                format_dict_loupes = {col: "{:+,.2f}" if "Sapma" in col else "{:,.2f}" for col in df_loupes_comb.columns if col not in ['Proje', 'Sonuç / Trend', 'Yönetici Analiz Yorumu']}
                
                loupes_styler = df_loupes_comb.style.format(format_dict_loupes, na_rep="0.0")
                if 'Sonuç / Trend' in df_loupes_comb.columns:
                    try:
                        loupes_styler = loupes_styler.map(renk_stili_ver, subset=['Sonuç / Trend'])
                    except AttributeError:
                        loupes_styler = loupes_styler.applymap(renk_stili_ver, subset=['Sonuç / Trend'])
                st.dataframe(loupes_styler, width="stretch", hide_index=True)
            else:
                st.info("Yüklenen dosyalarda Loupes sayfası verisi bulunamadı.")

else:
    st.info(f"💡 `{veri_tipi}` modunda analiz tablolarının oluşması için sol panelden ilgili Excel dosyalarınızı yükleyin.")
