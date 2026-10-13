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

def loupes_sayfasini_oku(file_source, dosya_adi):
    ay_no, ay_adi = ay_bilgisi_al(dosya_adi)
    try:
        xl = pd.ExcelFile(file_source)
        loupes_sheet = None
        for s in xl.sheet_names:
            if 'loupe' in s.lower():
                loupes_sheet = s
                break
        if not loupes_sheet:
            return pd.DataFrame()
            
        df = pd.read_excel(file_source, sheet_name=loupes_sheet)
        
        records = []
        proje_sutunlari = [
            ("B4D", 3, 4, 5),
            ("HR16", 9, 10, 11),
            ("HR18", 14, 15, 16)
        ]
        
        for proje, col_isim, col_bud, col_reel in proje_sutunlari:
            try:
                bud_val = pd.to_numeric(df.iloc[3, col_bud], errors='coerce') or 0.0
                reel_val = pd.to_numeric(df.iloc[3, col_reel], errors='coerce') or 0.0
                records.append({
                    "Ay": ay_adi,
                    "Ay_Sira": ay_no,
                    "Proje": proje,
                    "BUDGET": bud_val,
                    "REEL": reel_val
                })
            except Exception:
                pass
                
        return pd.DataFrame(records)
    except Exception:
        return pd.DataFrame()

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
        return {}, {}

    df_ozet_all = pd.concat(vtu_aylik_ozet_list, ignore_index=True)
    df_detay_all = pd.concat(vtu_aylik_detay_list, ignore_index=True)
    
    # Tamamlanan kısım: Benzersiz ayları kronolojik sıraya göre listeleme
    aylar_sirali = df_ozet_all[['Ay', 'Ay_Sira']].drop_duplicates().sort_values('Ay_Sira')['Ay'].tolist()
    
    return {
        "ozet": df_ozet_all,
        "detay": df_detay_all,
        "aylar_sirali": aylar_sirali
    }
