# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import openpyxl
import re
import os
import ast
import warnings
import json
import logging
import sys
from datetime import datetime, timedelta

import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from google.oauth2.service_account import Credentials

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ============================================================
# CẤU HÌNH
# ============================================================

BASE_DIR        = r"D:\Rooftop_IPos_data"
CREDENTIALS_FILE = r"C:\Users\PC\credentials.json"  # <-- đường dẫn file JSON service account

SHEET_URL = "https://docs.google.com/spreadsheets/d/1MNEKiD28QdjXwqg7VRyso1bbfMrvZxROF7ZTWK14uE8/edit?usp=sharing"

# ============================================================
# LOGGING
# ============================================================

os.makedirs(BASE_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(os.path.join(BASE_DIR, "analysis.log"), encoding="utf-8"),
    ]
)
log = logging.getLogger(__name__)

# ============================================================
# KẾT NỐI GOOGLE SHEETS
# ============================================================

def connect_gsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc

def append_to_sheet(gc, spreadsheet, worksheet_index, new_data):
    """Lấy data cũ + append data mới vào sheet"""
    worksheet = spreadsheet.get_worksheet(worksheet_index)
    old_data  = get_as_dataframe(worksheet, evaluate_formulas=True).dropna(how='all')
    final     = pd.concat([old_data, new_data], ignore_index=True)
    final     = final.replace([np.inf, -np.inf], 0).fillna(0)
    set_with_dataframe(worksheet, final)
    log.info(f"   ✅ Sheet {worksheet_index}: {len(new_data)} dòng mới (tổng: {len(final)})")

# ============================================================
# ĐỌC FILE EXCEL
# ============================================================

def read_latest_file(folder, prefix, suffix=".xlsx"):
    """Đọc file của ngày hôm qua, nếu không có thì lấy file mới nhất"""
    yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
    filepath  = os.path.join(folder, f"{prefix}{yesterday}{suffix}")

    if os.path.exists(filepath):
        df = pd.read_excel(filepath)
        log.info(f"✅ Đọc {len(df)} rows từ {os.path.basename(filepath)}")
        return df

    # Lấy file mới nhất
    files = sorted([f for f in os.listdir(folder) if f.startswith(prefix) and f.endswith(suffix)])
    if files:
        filepath = os.path.join(folder, files[-1])
        df = pd.read_excel(filepath)
        log.info(f"✅ Đọc file mới nhất: {files[-1]} ({len(df)} rows)")
        return df

    raise FileNotFoundError(f"Không có file nào trong {folder} với prefix '{prefix}'")

# ============================================================
# I. ĐỌC DATA
# ============================================================

def load_data():
    log.info("📂 Đọc data từ D:\\Rooftop_IPos_data\\...")

    # Hóa đơn theo thời gian
    Hoadontheothoigian = read_latest_file(
        os.path.join(BASE_DIR, "Hoa_don"), "hoa_don_"
    )
    Hoadontheothoigian = Hoadontheothoigian[[
        '_store_uid','tran_id','origin_tran_id','tran_no','created_at',
        'start_hour','start_minute','end_hour','end_minute','table_name',
        'discount_extra_amount','extra_data','total_amount'
    ]]

    # Chi tiết hóa đơn (sale_detail)
    Chitiethoadon = read_latest_file(
        os.path.join(BASE_DIR, "sale_detail"), "sale_detail_"
    )
    Chitiethoadon = Chitiethoadon[[
        'tran_id','tran_no','table_name','store_uid','peo_count',
        'item_name','item_type_name','quantity','price','unit_id',
        'amount_origin','store_name'
    ]]

    # Nhật ký order (sale_change_log)
    Nhatkyorder = read_latest_file(
        os.path.join(BASE_DIR, "sale_change_log"), "sale_"
    )
    Nhatkyorder = Nhatkyorder[['created_at','tran_id','log_type','table_name','change_data']]

    # Hủy món
    Huymon = read_latest_file(
        os.path.join(BASE_DIR, "Huy_mon"), "huy_mon_"
    )

    return Hoadontheothoigian, Chitiethoadon, Nhatkyorder, Huymon

# ============================================================
# II. CLEAN DATA - HÓA ĐƠN
# ============================================================

def process_hoadon(Hoadontheothoigian, Chitiethoadon):
    log.info("🔄 Xử lý Hóa đơn...")

    # Tách JSON extra_data
    def extract_fields(x):
        try:
            data = ast.literal_eval(x)
            return pd.Series({
                'customer_phone': data.get('customer_phone'),
                'customer_name': data.get('customer_name'),
                'Membership_Type_Name': data.get('Membership_Type_Name')
            })
        except:
            return pd.Series([None, None, None])

    Hoadontheothoigian[['customer_phone','customer_name','Membership_Type_Name']] = \
        Hoadontheothoigian['extra_data'].apply(extract_fields)

    Hoadontheothoigian['Giờ vào'] = (
        Hoadontheothoigian['start_hour'].astype(str).str.zfill(2) + ':' +
        Hoadontheothoigian['start_minute'].astype(str).str.zfill(2)
    )
    Hoadontheothoigian['Giờ ra'] = (
        Hoadontheothoigian['end_hour'].astype(str).str.zfill(2) + ':' +
        Hoadontheothoigian['end_minute'].astype(str).str.zfill(2)
    )
    Hoadontheothoigian = Hoadontheothoigian.rename(columns={
        '_store_uid':'Mã cửa hàng','tran_id':'Mã hóa đơn','origin_tran_id':'Mã hóa đơn gốc',
        'tran_no':'Số hóa đơn','created_at':'Ngày','table_name':'Bàn',
        'discount_extra_amount':'Giảm giá','total_amount':'Tổng hóa đơn',
        'customer_phone':'SĐT','customer_name':'Tên Khách','Membership_Type_Name':'Loại thành viên'
    })
    Hoadontheothoigian.loc[Hoadontheothoigian['Mã hóa đơn gốc'].notna(),'Mã hóa đơn'] = Hoadontheothoigian['Mã hóa đơn gốc']
    Hoadontheothoigian = Hoadontheothoigian.drop(
        columns=['extra_data','start_hour','start_minute','end_hour','end_minute']
    )

    # Chi tiết hóa đơn
    Chitiethoadon = Chitiethoadon.rename(columns={
        'tran_id':'Mã hóa đơn','tran_no':'Số hóa đơn','table_name':'Bàn',
        'store_uid':'Mã cửa hàng','peo_count':'Số khách','item_name':'Tên hàng',
        'item_type_name':'Mã hàng hóa','quantity':'Số lượng','price':'Đơn giá',
        'unit_id':'Đơn vị tính','amount_origin':'Thành tiền','store_name':'Cửa hàng',
    })

    # Merge
    merged_df_1 = pd.merge(
        Chitiethoadon, Hoadontheothoigian,
        on=['Mã hóa đơn','Mã cửa hàng','Bàn','Số hóa đơn'], how='right'
    )
    merged_df_1 = merged_df_1[~merged_df_1['Tên hàng'].str.contains('Không lấy Topping', na=False)]

    # Loại khách hàng
    merged_df_1['Loại khách hàng'] = None
    mask = merged_df_1.index.notnull()
    merged_df_1.loc[mask, 'Loại khách hàng'] = np.select(
        condlist=[
            (merged_df_1.loc[mask,'SĐT'].isna() | (merged_df_1.loc[mask,'SĐT'] == '')) &
            (merged_df_1.loc[mask,'Loại thành viên'].isna() | (merged_df_1.loc[mask,'Loại thành viên'] == '')),
            (merged_df_1.loc[mask,'SĐT'].notna() & (merged_df_1.loc[mask,'SĐT'] != '')) &
            (merged_df_1.loc[mask,'Loại thành viên'].isna() |
             (merged_df_1.loc[mask,'Loại thành viên'] == '') |
             (merged_df_1.loc[mask,'Loại thành viên'] == 'Thành viên mặc định'))
        ],
        choicelist=['Khách lẻ','Khách mới'],
        default='Khách quay lại'
    )

    merged_df_1 = merged_df_1.sort_values(by=['Số hóa đơn','Ngày','Tên hàng'])
    merged_df_1['Mark'] = merged_df_1.groupby(['Số hóa đơn','Tên hàng','Số lượng']).cumcount() + 1

    return merged_df_1

# ============================================================
# III. CLEAN DATA - NHẬT KÝ ORDER
# ============================================================

def process_nhatky(Nhatkyorder_raw):
    log.info("🔄 Xử lý Nhật ký order...")

    def process_row(row):
        try:
            data = ast.literal_eval(row['change_data'])
            created_at = row['created_at']
            log_type = row['log_type']
            table_name = row['table_name']
            employee_name = data.get('employee_name')
            store_uid = data.get('store_uid')
            tran_id = data.get('tran_id')
            origin_tran_id = data.get('origin_tran_id')
            tran_no = data.get('tran_no')
            extra_data = data.get('extra_data', {})
            message_modify_table = extra_data.get('message_modify_table')
            items = data.get('sale_detail', [])
            rows = []
            for item in items:
                rows.append({
                    'created_at': created_at, 'log_type': log_type, 'table_name': table_name,
                    'employee_name': employee_name, 'store_uid': store_uid,
                    'tran_id': tran_id, 'origin_tran_id': origin_tran_id, 'tran_no': tran_no,
                    'message_modify_table': message_modify_table,
                    'item_type': 'MAIN', 'item_name': item.get('item_name'), 'parent_item': None,
                    'price': item.get('price'), 'quantity': item.get('quantity'), 'unit_id': item.get('unit_id')
                })
                for tp in (item.get('toppings') or []):
                    rows.append({
                        'created_at': created_at, 'log_type': log_type, 'table_name': table_name,
                        'employee_name': employee_name, 'store_uid': store_uid,
                        'tran_id': tran_id, 'origin_tran_id': origin_tran_id, 'tran_no': tran_no,
                        'message_modify_table': message_modify_table,
                        'item_type': 'TOPPING', 'item_name': tp.get('item_name'), 'parent_item': item.get('item_name'),
                        'price': tp.get('price'), 'quantity': tp.get('quantity', item.get('quantity')), 'unit_id': tp.get('unit_id')
                    })
            return rows
        except:
            return []

    result = Nhatkyorder_raw.apply(process_row, axis=1).explode()
    result = result[result.apply(lambda x: isinstance(x, dict))]
    result_df = pd.DataFrame(result.tolist()).reset_index(drop=True)
    result_df['line_revenue'] = result_df['price'] * result_df['quantity']
    result_df = result_df.rename(columns={
        'created_at':'Thời gian','log_type':'Loại log','table_name':'Bàn',
        'employee_name':'Nhân viên','store_uid':'Mã cửa hàng','tran_id':'Mã hóa đơn',
        'origin_tran_id':'Mã hóa đơn gốc','tran_no':'Số hóa đơn',
        'message_modify_table':'Ghi chú','item_name':'Tên hàng',
        'price':'Đơn giá','quantity':'Số lượng','unit_id':'Đơn vị tính','line_revenue':'Thành tiền'
    })
    result_df = result_df.drop(columns=['item_type','parent_item'], errors='ignore')
    Nhatkyorder = result_df

    # Lọc log cần thiết
    Nhatkyorder = Nhatkyorder[Nhatkyorder['Loại log'].isin(['SALE_MERGE_ORDER','SALE_CHANGE','SALE_SPLIT_ORDER','-'])]

    # Nhóm 1: Gộp/Tách đơn
    target_hoadon_list = Nhatkyorder[
        Nhatkyorder['Loại log'].str.contains('SALE_MERGE_ORDER|SALE_SPLIT_ORDER', case=False, na=False)
    ][['Số hóa đơn']].drop_duplicates()

    Nhatkyorder_gop_tach = Nhatkyorder.merge(target_hoadon_list, on='Số hóa đơn', how='inner')
    Nhatkyorder_gop_tach = Nhatkyorder_gop_tach.sort_values(['Số hóa đơn','Thời gian']).reset_index(drop=True)

    # Xử lý gộp đơn
    Nhatkyorder_gop_tach['Mã hóa đơn sau khi gộp bàn'] = None
    mask = Nhatkyorder_gop_tach['Ghi chú'].str.contains('[Gộp đơn]', case=False, na=False)

    def extract_new_invoice_gop(text):
        if pd.isna(text): return None
        # Regex mới: lấy mã giữa "tạo thành hóa đơn" và " -"
        match = re.search(r'hóa đơn bị hủy bởi vì gộp vào\s+(.+?)\s+-', text)
        if match:
            return match.group(1).strip()
        # Fallback regex
        match = re.search(r'hóa đơn bị hủy bởi vì gộp vào\s+([A-Z0-9]+)', text)
        return match.group(1).strip() if match else None

    Nhatkyorder_gop_tach.loc[mask, 'Mã hóa đơn sau khi gộp bàn'] = \
        Nhatkyorder_gop_tach.loc[mask, 'Ghi chú'].apply(extract_new_invoice_gop)
    Nhatkyorder_gop_tach['Mã hóa đơn sau khi gộp bàn'] = \
        Nhatkyorder_gop_tach['Mã hóa đơn sau khi gộp bàn'].replace('', None)
    Nhatkyorder_gop_tach['Mã hóa đơn sau khi gộp bàn'] = (
        Nhatkyorder_gop_tach.groupby('Số hóa đơn')['Mã hóa đơn sau khi gộp bàn']
        .transform(lambda x: x.ffill().bfill())
    )
    Nhatkyorder_gop_tach = Nhatkyorder_gop_tach[
        ~Nhatkyorder_gop_tach['Ghi chú'].str.contains('[Gộp đơn]', na=False, regex=False)
    ]

    mapping = dict(zip(Nhatkyorder_gop_tach['Mã hóa đơn'], Nhatkyorder_gop_tach['Mã hóa đơn sau khi gộp bàn']))

    def find_final_code(code):
        visited = set()
        while pd.notna(code) and code in mapping and mapping[code] not in visited:
            visited.add(code)
            next_code = mapping.get(code)
            if pd.isna(next_code): break
            code = next_code
        return code

    Nhatkyorder_gop_tach['Mã hóa đơn cuối cùng'] = Nhatkyorder_gop_tach['Mã hóa đơn'].apply(find_final_code)
    sohd_map = (Nhatkyorder_gop_tach.drop_duplicates(subset=['Mã hóa đơn'], keep='first')
                .set_index('Mã hóa đơn')['Số hóa đơn'])
    Nhatkyorder_gop_tach['Số hóa đơn cuối'] = Nhatkyorder_gop_tach['Mã hóa đơn cuối cùng'].map(sohd_map)
    Nhatkyorder_gop_tach['Số hóa đơn cuối'] = Nhatkyorder_gop_tach['Số hóa đơn cuối'].fillna(
        Nhatkyorder_gop_tach['Số hóa đơn']
    )
    Nhatkyorder_gop_tach = Nhatkyorder_gop_tach.drop(columns=['Số hóa đơn'], errors='ignore')
    Nhatkyorder_gop_tach = Nhatkyorder_gop_tach.rename(columns={'Số hóa đơn cuối':'Số hóa đơn'})
    Nhatkyorder_gop_tach = Nhatkyorder_gop_tach.drop(
        columns=['Mã hóa đơn','Mã hóa đơn sau khi gộp bàn'], errors='ignore'
    )
    Nhatkyorder_gop_tach = Nhatkyorder_gop_tach.rename(columns={'Mã hóa đơn cuối cùng':'Mã hóa đơn'})
    cols_order = ['Mã hóa đơn','Số hóa đơn','Thời gian','Bàn','Nhân viên','Loại log',
                  'Ghi chú','Tên hàng','Số lượng','Mã hóa đơn gốc','Đơn giá','Đơn vị tính',
                  'Thành tiền','Mã cửa hàng']
    Nhatkyorder_gop_tach = Nhatkyorder_gop_tach.reindex(columns=cols_order)
    Nhatkyorder_gop_tach = Nhatkyorder_gop_tach.sort_values(['Số hóa đơn','Thời gian']).reset_index(drop=True)

    # Xử lý tách đơn
    ds_hoadon_tach = Nhatkyorder_gop_tach[
        Nhatkyorder_gop_tach['Loại log'].str.contains('SALE_SPLIT_ORDER', case=False, na=False)
    ]['Số hóa đơn'].unique()
 
    Nhatkyorder_tachdon = Nhatkyorder_gop_tach[
        Nhatkyorder_gop_tach['Số hóa đơn'].isin(ds_hoadon_tach)
    ].sort_values(['Thời gian','Số hóa đơn']).reset_index(drop=True)
 
    Nhatkyorder_con_lai = Nhatkyorder_gop_tach[
        ~Nhatkyorder_gop_tach['Số hóa đơn'].isin(ds_hoadon_tach)
    ].sort_values(['Số hóa đơn','Thời gian']).reset_index(drop=True)
 
    # Tách đơn - xử lý
    Nhatkyorder_tachdon['Mã hóa đơn sau khi tách bàn'] = None
    mask_tach = (
        Nhatkyorder_tachdon['Ghi chú'].str.contains(r'\[Tách đơn\]', case=False, na=False) &
        Nhatkyorder_tachdon['Ghi chú'].str.contains('các món được tách và tạo thành hóa đơn', case=False, na=False)
    )
 
    def extract_new_invoice_tach(text):
        if pd.isna(text): return None
        # Regex mới: lấy mã giữa "tạo thành hóa đơn" và " -"
        match = re.search(r'tạo thành hóa đơn\s+(.+?)\s+-', text)
        if match:
            return match.group(1).strip()
        # Fallback regex
        match = re.search(r'tạo thành hóa đơn\s+([A-Z0-9]+)', text)
        return match.group(1).strip() if match else None
 
    Nhatkyorder_tachdon.loc[mask_tach, 'Mã hóa đơn sau khi tách bàn'] = \
        Nhatkyorder_tachdon.loc[mask_tach, 'Ghi chú'].apply(extract_new_invoice_tach)
    Nhatkyorder_tachdon['Mã hóa đơn sau khi tách bàn'] = \
        Nhatkyorder_tachdon['Mã hóa đơn sau khi tách bàn'].replace('', None)
 
    mask_tach2 = Nhatkyorder_tachdon['Loại log'] == 'SALE_SPLIT_ORDER'
    Nhatkyorder_tachdon.loc[mask_tach2, 'Mã hóa đơn sau khi tách bàn'] = (
        Nhatkyorder_tachdon.groupby('Mã hóa đơn')['Mã hóa đơn sau khi tách bàn']
        .transform(lambda x: x.ffill().bfill())
    )
    Nhatkyorder_tachdon.loc[~mask_tach2, 'Mã hóa đơn sau khi tách bàn'] = None
 
    mapping_df = Nhatkyorder_tachdon[Nhatkyorder_tachdon['Mã hóa đơn sau khi tách bàn'].notna()][
        ['Mã hóa đơn','Mã hóa đơn sau khi tách bàn']
    ].drop_duplicates()
    child_to_parent_map = dict(zip(mapping_df['Mã hóa đơn sau khi tách bàn'], mapping_df['Mã hóa đơn']))
 
    def find_ultimate_root(invoice_code, p_map, cache):
        visited = []
        current = invoice_code
    
        while current in p_map:
            if current in cache:
                root = cache[current]
                break
        
            if current in visited:
                root = current
                break
        
            visited.append(current)
            current = p_map[current]
        else:
            root = current

        for node in visited:
            cache[node] = root
    
        return root

    cache = {}

    Nhatkyorder_tachdon['Group_ID_Goc'] = Nhatkyorder_tachdon['Mã hóa đơn'].map(
        lambda x: find_ultimate_root(x, child_to_parent_map, cache)
    )
    Nhatkyorder_tachdon = Nhatkyorder_tachdon.sort_values(
        by=['Thời gian','Group_ID_Goc'], ascending=[True,True]
    ).reset_index(drop=True)
 
    lookup_dict = dict(zip(Nhatkyorder_tachdon['Mã hóa đơn'], Nhatkyorder_tachdon['Số hóa đơn']))
 
    def map_full_invoice(row):
        if row['Loại log'] == 'SALE_SPLIT_ORDER':
            return lookup_dict.get(row['Mã hóa đơn sau khi tách bàn'], None)
        return None
 
    Nhatkyorder_tachdon['Số hóa đơn tách bàn'] = Nhatkyorder_tachdon.apply(map_full_invoice, axis=1)
 
    # Xóa dòng SALE_SPLIT_ORDER mà không tìm được hóa đơn tách (NaN) - theo code Colab mới
    Nhatkyorder_tachdon = Nhatkyorder_tachdon[
        ~(
            (Nhatkyorder_tachdon['Loại log'] == 'SALE_SPLIT_ORDER') &
            (Nhatkyorder_tachdon['Số hóa đơn tách bàn'].isna())
        )
    ]
 
    Nhatkyorder_tachdon = Nhatkyorder_tachdon[
        ~Nhatkyorder_tachdon['Ghi chú'].str.contains(r'hóa đơn được tạo mới', case=False, na=False)
    ]
    Nhatkyorder_tachdon['Ma_Sort_Phu'] = Nhatkyorder_tachdon['Mã hóa đơn sau khi tách bàn'].astype(str).str[-4:]
    Nhatkyorder_tachdon = Nhatkyorder_tachdon.sort_values(
        by=['Group_ID_Goc','Thời gian','Ma_Sort_Phu']
    ).reset_index(drop=True).drop(columns=['Ma_Sort_Phu'])
 
    Nhatkyorder_tachdon['Số lượng'] = pd.to_numeric(Nhatkyorder_tachdon['Số lượng'], errors='coerce').fillna(0)
    Nhatkyorder_tachdon['Tên hàng'] = Nhatkyorder_tachdon['Tên hàng'].astype(str).str.strip()
    Nhatkyorder_tachdon = Nhatkyorder_tachdon.replace({pd.NA: None, np.nan: None, "": None, "null": None})
    Nhatkyorder_tachdon = Nhatkyorder_tachdon.sort_values(['Group_ID_Goc','Thời gian']).reset_index(drop=True)
 
    final_rows = []
    for group_id, df_group in Nhatkyorder_tachdon.groupby('Group_ID_Goc'):
        try:
            root_id = df_group[df_group['Mã hóa đơn'].str.contains(str(group_id), na=False)]['Mã hóa đơn'].iloc[0]
        except:
            root_id = df_group.iloc[0]['Mã hóa đơn']
 
        current_inventory = []
        processed_indices = set()
 
        for idx, row in df_group.iterrows():
            if idx in processed_indices: continue
            if row['Loại log'] == 'SALE_CHANGE':
                if row['Mã hóa đơn'] == root_id:
                    if row['Số lượng'] > 0:
                        current_inventory.append({'data': row.to_dict(), 'Tên hàng': row['Tên hàng'], 'Số lượng': row['Số lượng']})
                    else:
                        qty_to_void = abs(row['Số lượng'])
                        for item in reversed(current_inventory):
                            if item['Tên hàng'] == row['Tên hàng'] and item['Số lượng'] > 0:
                                void = min(item['Số lượng'], qty_to_void)
                                item['Số lượng'] -= void
                                qty_to_void -= void
                            if qty_to_void <= 0: break
                else:
                    final_rows.append(row)
            elif row['Loại log'] == 'SALE_SPLIT_ORDER' and row['Mã hóa đơn sau khi tách bàn'] is not None:
                hd_con_target_shd = row['Số hóa đơn tách bàn']
                hd_con_target = row['Mã hóa đơn sau khi tách bàn']
                thoi_gian_tach = row['Thời gian']
                df_block_tach = df_group[
                    (df_group['Loại log'] == 'SALE_SPLIT_ORDER') &
                    (df_group['Số hóa đơn tách bàn'] == hd_con_target_shd) &
                    (df_group['Mã hóa đơn sau khi tách bàn'] == hd_con_target) &
                    (df_group['Thời gian'] == thoi_gian_tach)
                ]
                processed_indices.update(df_block_tach.index.tolist())
                staying_items = df_block_tach.groupby('Tên hàng')['Số lượng'].sum().abs().to_dict()
                new_inventory_for_root = []
                for item in current_inventory:
                    mon_name = item['Tên hàng']
                    sl_trong_kho = item['Số lượng']
                    if sl_trong_kho <= 0: continue
                    sl_o_lai_yeu_cau = staying_items.get(mon_name, 0)
                    if sl_trong_kho > sl_o_lai_yeu_cau:
                        sl_tach_di = sl_trong_kho - sl_o_lai_yeu_cau
                        new_split_row = pd.Series(item['data'])
                        new_split_row['Mã hóa đơn'] = hd_con_target
                        new_split_row['Số hóa đơn'] = hd_con_target_shd
                        new_split_row['Số lượng'] = sl_tach_di
                        new_split_row['Mã hóa đơn sau khi tách bàn'] = None
                        final_rows.append(new_split_row)
                        if sl_o_lai_yeu_cau > 0:
                            item['Số lượng'] = sl_o_lai_yeu_cau
                            new_inventory_for_root.append(item)
                            staying_items[mon_name] = 0
                    else:
                        new_inventory_for_root.append(item)
                        staying_items[mon_name] -= sl_trong_kho
                current_inventory = new_inventory_for_root
 
        for item in current_inventory:
            if item['Số lượng'] > 0:
                final_root_row = pd.Series(item['data'])
                final_root_row['Số lượng'] = item['Số lượng']
                final_root_row['Mã hóa đơn sau khi tách bàn'] = None
                final_rows.append(final_root_row)
 
    Nhatkyorder_tachdon = pd.DataFrame(final_rows).reset_index(drop=True)
    Nhatkyorder_tachdon = Nhatkyorder_tachdon.replace({np.nan: None})
     # ✅ Thêm dòng này để đảm bảo cột luôn tồn tại
    if 'Số hóa đơn tách bàn' not in Nhatkyorder_tachdon.columns:
        Nhatkyorder_tachdon['Số hóa đơn tách bàn'] = None
    mask_not_nan = Nhatkyorder_tachdon['Số hóa đơn tách bàn'].notna()
    Nhatkyorder_tachdon.loc[mask_not_nan, 'Số hóa đơn'] = Nhatkyorder_tachdon.loc[mask_not_nan, 'Số hóa đơn tách bàn']
    Nhatkyorder_tachdon = Nhatkyorder_tachdon.drop(
        columns=['Group_ID_Goc','Mã hóa đơn sau khi tách bàn','Số hóa đơn tách bàn'], errors='ignore'
    ).reset_index(drop=True)
 
    Nhatkyorder_nhom_1 = Nhatkyorder_tachdon

    # Nhóm 2: Chỉ sửa đơn
    mask_hop_le = Nhatkyorder['Loại log'].str.contains('SALE_CHANGE', case=False, na=False)
    hoa_don_co_hanh_dong_khac = Nhatkyorder.loc[~mask_hop_le, 'Số hóa đơn'].unique()
    Nhatkyorder_nhom_2 = (
        Nhatkyorder[~Nhatkyorder['Số hóa đơn'].isin(hoa_don_co_hanh_dong_khac)]
        .sort_values(['Số hóa đơn','Thời gian']).reset_index(drop=True)
    )

    # Gộp
    Nhatkyorder = pd.concat([Nhatkyorder_nhom_1, Nhatkyorder_nhom_2], ignore_index=True)
    Nhatkyorder = Nhatkyorder.sort_values(by=['Mã hóa đơn','Thời gian'], ascending=True)

    # Xử lý thời gian
    Nhatkyorder['Thời gian'] = pd.to_datetime(Nhatkyorder['Thời gian'], format='%d/%m/%Y %H:%M', errors='coerce')
    Nhatkyorder['Thời gian'] = Nhatkyorder['Thời gian'].dt.strftime('%d/%m/%Y %H:%M:00')
    Nhatkyorder['Thời gian'] = pd.to_datetime(Nhatkyorder['Thời gian'], format='%d/%m/%Y %H:%M:%S')
    Nhatkyorder['Ngày']  = Nhatkyorder['Thời gian'].dt.strftime('%d/%m/%Y')
    Nhatkyorder['Năm']   = Nhatkyorder['Thời gian'].dt.year
    Nhatkyorder['Tháng'] = Nhatkyorder['Thời gian'].dt.strftime('%m-%Y')
    iso = Nhatkyorder['Thời gian'].dt.isocalendar()
    Nhatkyorder['Tuần'] = iso.week.astype(str).str.zfill(2) + '-' + iso.year.astype(str)

    # FIFO khấu trừ
    Nhatkyorder['Số lượng'] = pd.to_numeric(Nhatkyorder['Số lượng'], errors='coerce').fillna(0)
    pos_df = Nhatkyorder[Nhatkyorder['Số lượng'] > 0].copy()
    neg_df = Nhatkyorder[Nhatkyorder['Số lượng'] < 0].copy()
    for _, neg_row in neg_df.iterrows():
        amount_to_deduct = abs(neg_row['Số lượng'])
        bill_id = neg_row['Mã hóa đơn']
        item_name = neg_row['Tên hàng']
        mask2 = (pos_df['Mã hóa đơn'] == bill_id) & (pos_df['Tên hàng'] == item_name)
        for idx in reversed(pos_df[mask2].index.tolist()):
            if amount_to_deduct <= 0: break
            current_val = pos_df.at[idx, 'Số lượng']
            if current_val <= amount_to_deduct:
                amount_to_deduct -= current_val
                pos_df.at[idx, 'Số lượng'] = 0
            else:
                pos_df.at[idx, 'Số lượng'] = current_val - amount_to_deduct
                amount_to_deduct = 0
    Nhatkyorder = pos_df[pos_df['Số lượng'] > 0]
    Nhatkyorder = Nhatkyorder.sort_values(by=['Mã hóa đơn','Tên hàng','Thời gian'])
    Nhatkyorder['Mark'] = Nhatkyorder.groupby(['Mã hóa đơn','Tên hàng','Số lượng']).cumcount() + 1
    Nhatkyorder = Nhatkyorder[~Nhatkyorder['Tên hàng'].str.contains('Không lấy Topping', na=False)]

    return Nhatkyorder

# ============================================================
# IV. CLEAN DATA - HỦY MÓN
# ============================================================

def process_huymon(Huymon):
    log.info("🔄 Xử lý Hủy món...")

    def parse_sale_detail(x):
        try:
            if not isinstance(x, str): return []
            data = ast.literal_eval(x)
            return data if isinstance(data, list) else []
        except:
            return []

    Huymon['sale_detail_parsed'] = Huymon['sale_detail'].apply(parse_sale_detail)
    df_huymon = Huymon.explode('sale_detail_parsed')
    df_huymon['item_name'] = df_huymon['sale_detail_parsed'].apply(lambda x: x.get('item_name') if isinstance(x, dict) else None)
    df_huymon['quantity']  = df_huymon['sale_detail_parsed'].apply(lambda x: x.get('quantity') if isinstance(x, dict) else None)
    df_huymon['unit_id']   = df_huymon['sale_detail_parsed'].apply(lambda x: x.get('unit_id') if isinstance(x, dict) else None)
    df_huymon['note']      = df_huymon['sale_detail_parsed'].apply(lambda x: x.get('note') if isinstance(x, dict) else None)

    df_huymon = df_huymon.rename(columns={
        'tran_id':'Mã hóa đơn','tran_no':'Số hóa đơn','updated_at':'Thời gian sửa',
        'item_name':'Tên món','unit_id':'Đơn vị','employee_name':'Người sửa',
        'quantity':'Số lượng','table_name':'Tên bàn','note':'Lý do','_store_uid':'Mã cửa hàng'
    })
    df_huymon = df_huymon.drop(
        columns=['total_tran_id','log_type','log_date','state_action','sale_detail','_date','sale_detail_parsed'],
        errors='ignore'
    )

    df_huymon['Thời gian sửa'] = pd.to_datetime(df_huymon['Thời gian sửa'], format="%d/%m/%Y %H:%M", errors="coerce")
    df_huymon['Ngày']  = df_huymon['Thời gian sửa'].dt.strftime('%d/%m/%Y')
    df_huymon['Năm']   = df_huymon['Thời gian sửa'].dt.year
    df_huymon['Tháng'] = df_huymon['Thời gian sửa'].dt.strftime('%m-%Y')
    iso = df_huymon['Thời gian sửa'].dt.isocalendar()
    df_huymon['Tuần'] = iso.week.astype(str).str.zfill(2) + '-' + iso.year.astype(str)
    df_huymon = df_huymon[df_huymon['Số lượng'] <= 0]
    df_huymon = df_huymon.rename(columns={
        'Người sửa':'Nhân viên','Thời gian sửa':'Thời gian','Tên món':'Món'
    })

    ma_cua_hang = ['CVGTN','CVGHX','TN','THU','CHV']
    ten_cua_hang = ["Chạng Vạng Trần Não","Chạng Vạng Hàng Xanh","Trăng Non","Thương","Chênh Vênh"]
    conditions = [df_huymon['Số hóa đơn'].astype(str).str.contains(ma, na=False) for ma in ma_cua_hang]
    conditions.append(df_huymon['Số hóa đơn'].isna() | df_huymon['Số hóa đơn'].astype(str).str.contains('RC', na=False))
    df_huymon['Cửa hàng'] = np.select(conditions, ten_cua_hang + ["Ráng Chiều"], default='Khác')
    df_huymon['Số lượng'] = df_huymon['Số lượng'].abs()
    df_huymon = df_huymon[['Ngày','Tuần','Tháng','Năm','Cửa hàng','Số hóa đơn','Nhân viên','Món','Số lượng','Lý do','Thời gian']]

    return df_huymon

# ============================================================
# V. TÍNH KPI
# ============================================================
 
def compute_kpi(Hoadontheothoigian, Nhatkyorder, Huymon):
    log.info("Tinh KPI...")
 
    Hoadontheothoigian['Số lượng'] = pd.to_numeric(Hoadontheothoigian['Số lượng'], errors='coerce').fillna(0).astype(int)
    Nhatkyorder['Mark'] = pd.to_numeric(Nhatkyorder['Mark'], errors='coerce').fillna(0).astype(int)
    Hoadontheothoigian['Mark'] = pd.to_numeric(Hoadontheothoigian['Mark'], errors='coerce').fillna(0).astype(int)
 
    Databanhang_merged = pd.merge(
        Nhatkyorder[['Mã hóa đơn','Thời gian','Tên hàng','Mark','Nhân viên','Loại log','Số lượng','Đơn giá','Ngày','Tuần','Tháng','Năm','Thành tiền']],
        Hoadontheothoigian[['Mã hóa đơn','Cửa hàng','Tên hàng','Số lượng','Mark','Số hóa đơn','Bàn','Số khách','Loại khách hàng','SĐT','Tên Khách','Tổng hóa đơn','Giờ vào','Giờ ra','Giảm giá','Loại thành viên']],
        left_on=['Mã hóa đơn','Tên hàng','Số lượng','Mark'],
        right_on=['Mã hóa đơn','Tên hàng','Số lượng','Mark'],
        how='left'
    )
 
    cols = ['Cửa hàng','Số hóa đơn','Bàn','Số khách','Loại khách hàng','SĐT','Tổng hóa đơn']
 
    # Buoc 1: fill trong tung hoa don
    Databanhang_merged[cols] = (
        Databanhang_merged.groupby('Mã hóa đơn')[cols]
        .transform(lambda x: x.ffill().bfill())
        .infer_objects(copy=False)
    )
 
    # Buoc 2: tao lookup chuan tu Hoadontheothoigian
    lookup = (
        Hoadontheothoigian.groupby('Mã hóa đơn').agg({
            'Cửa hàng':'first','Số hóa đơn':'first','Bàn':'first',
            'Số khách':'max','Loại khách hàng':'first','SĐT':'first','Tổng hóa đơn':'max'
        })
    )
 
    # Buoc 3: fill NaN tu lookup
    for col in cols:
        Databanhang_merged[col] = Databanhang_merged[col].fillna(
            Databanhang_merged['Mã hóa đơn'].map(lookup[col])
        )
 
    # Buoc 4: xoa dong khong co So hoa don
    Databanhang_merged.dropna(subset=['Số hóa đơn'], inplace=True)
 
    # Databanhang
    Databanhang = Databanhang_merged.reindex(columns=[
        'Ngày','Tuần','Tháng','Năm','Cửa hàng','Số hóa đơn','Loại khách hàng',
        'Tên hàng','Đơn giá','Nhân viên','Bàn','Loại log','Số lượng','Số khách',
        'Thời gian','Thành tiền','SĐT'
    ])
    Databanhang = Databanhang.sort_values(by=['Thời gian'], ascending=True)
    Databanhang['Số lượng'] = pd.to_numeric(Databanhang['Số lượng'], errors='coerce').astype(int)
    Databanhang = Databanhang.sort_values(by=['Số hóa đơn','Tên hàng','Thời gian'])
 
    # DataKHTT_Cuahang
    
    Hoadontheothoigian['Ngày'] = pd.to_datetime(Hoadontheothoigian['Ngày'], errors='coerce')
    Hoadontheothoigian['Năm']   = Hoadontheothoigian['Ngày'].dt.year
    Hoadontheothoigian['Tháng'] = Hoadontheothoigian['Ngày'].dt.strftime('%m-%Y')
    iso = Hoadontheothoigian['Ngày'].dt.isocalendar()
    Hoadontheothoigian['Tuần'] = iso.week.astype(str).str.zfill(2) + '-' + iso.year.astype(str)
    Hoadontheothoigian['Ngày'] = pd.to_datetime(Hoadontheothoigian['Ngày']).dt.strftime('%d/%m/%Y')
    Hoadontheothoigian['Thời gian'] = pd.to_datetime(Hoadontheothoigian['Ngày'].astype(str) + ' ' + Hoadontheothoigian['Giờ ra'].astype(str), errors='coerce')

    DataKHTT_Cuahang_first = Hoadontheothoigian.reindex(columns=[
        'Ngày','Tuần','Tháng','Năm','Số hóa đơn','Thời gian','Cửa hàng',
        'Tổng hóa đơn','Giờ vào','Giờ ra','Bàn','Số khách','Giảm giá',
        'Loại thành viên','Tên Khách','SĐT','Loại khách hàng'
    ])
    DataKHTT_Cuahang_first = DataKHTT_Cuahang_first.sort_values(by=['Thời gian'], ascending=True)
    DataKHTT_Cuahang_first = DataKHTT_Cuahang_first.sort_values(by=['Số hóa đơn','Thời gian'])

    DataKHTT_Cuahang = DataKHTT_Cuahang_first.drop_duplicates(subset=['Số hóa đơn'], keep='first')
    # Mapping: tiền tố trong "Số hóa đơn" -> tên "Cửa hàng"
    mapping = {
        "THU"   : "Thương",
        "CVGTN" : "Chạng Vạng Trần Não",
        "CVGHX" : "Chạng Vạng Hàng Xanh",
        "CHV"   : "Chênh Vênh",
        "RC"    : "Ráng Chiều",
    }

    # Điều kiện hàng có cột "Cửa hàng" trống/rỗng/0
    mask_empty = (
        DataKHTT_Cuahang["Cửa hàng"].isna() |
        (DataKHTT_Cuahang["Cửa hàng"] == 0) |
        (DataKHTT_Cuahang["Cửa hàng"].astype(str).str.strip() == ""))

    # Với mỗi prefix -> gán tên cửa hàng tương ứng
    for prefix, ten_cua_hang in mapping.items():
        mask = mask_empty & DataKHTT_Cuahang["Số hóa đơn"].astype(str).str.startswith(prefix)
        DataKHTT_Cuahang.loc[mask, "Cửa hàng"] = ten_cua_hang

    DataKHTT_Cuahang["Thời gian"] = pd.to_datetime(DataKHTT_Cuahang["Thời gian"], dayfirst=True)

    # Nếu giờ == 0 thì lùi cột Ngày 1 ngày
    mask_gio00 = DataKHTT_Cuahang["Thời gian"].dt.hour == 0

    DataKHTT_Cuahang.loc[mask_gio00, "Ngày"] = (
        DataKHTT_Cuahang.loc[mask_gio00, "Thời gian"] - pd.Timedelta(days=1)
    ).dt.strftime("%d/%m/%Y")

 
    # group_by_ban_hang
    group_by_ban_hang = (
        Databanhang.groupby(
            ['Ngày','Tuần','Tháng','Năm','Cửa hàng','Số hóa đơn','Số khách','Loại khách hàng','Nhân viên','SĐT'],
            as_index=False
        )['Thành tiền'].sum()
        .sort_values(['Ngày','Số hóa đơn','Nhân viên']).reset_index(drop=True)
    )
 
    # groupby_huy_mon
    groupby_huy_mon = (
        Huymon.groupby(['Ngày','Tuần','Tháng','Năm','Nhân viên','Số hóa đơn','Cửa hàng'], as_index=False)
        .agg(So_lan_huy_mon=('Nhân viên','count'), So_luong_huy=('Số lượng','sum'))
        .rename(columns={'So_lan_huy_mon':'Số lần hủy món','So_luong_huy':'Số lượng hủy món'})
    )
 
    group_by = pd.merge(
        group_by_ban_hang,
        groupby_huy_mon[['Ngày','Nhân viên','Số hóa đơn','Số lần hủy món','Số lượng hủy món']],
        on=['Ngày','Nhân viên','Số hóa đơn'], how='left'
    )
    group_by['Số lần hủy món'] = group_by['Số lần hủy món'].fillna(0)
    group_by['Số lượng hủy món'] = group_by['Số lượng hủy món'].fillna(0)
 
    tong_don = (
        group_by.groupby(['Ngày','Tuần','Tháng','Năm','Cửa hàng','Nhân viên'], as_index=False)['Số hóa đơn']
        .count().rename(columns={'Số hóa đơn':'Tổng số dòng'})
    )
    don_co_tt = (
        group_by[group_by['Loại khách hàng'].isin(["Khách hàng mới","Khách quay lại"])]
        .groupby(['Ngày','Nhân viên'], as_index=False)['Số hóa đơn']
        .count().rename(columns={'Số hóa đơn':'Số dòng có thông tin KH'})
    )
    ty_le = pd.merge(tong_don, don_co_tt, on=['Ngày','Nhân viên'], how='left').fillna(0)
    ty_le['% xin thông tin KH'] = (ty_le['Số dòng có thông tin KH'] / ty_le['Tổng số dòng']).round(2)
    ty_le = ty_le.sort_values(['Ngày','% xin thông tin KH'], ascending=[True,False]).reset_index(drop=True)
 
    result = (
        group_by.groupby(['Ngày','Tuần','Tháng','Năm','Cửa hàng','Nhân viên'], as_index=False)
        .agg({
            'Thành tiền':'sum','Số khách':'sum',
            'Số lần hủy món':'sum','Số lượng hủy món':'sum','Số hóa đơn':'count'
        })
        .rename(columns={
            'Thành tiền':'Tổng giá trị order','Số khách':'Tổng số khách',
            'Số lần hủy món':'Tổng số lần hủy món','Số lượng hủy món':'Tổng số lượng hủy món',
            'Số hóa đơn':'Tổng hoá đơn theo nhân viên'
        })
    )
    result['Doanh thu trên KH'] = (result['Tổng giá trị order'] / result['Tổng số khách']).round(2)
    result = result.sort_values(['Ngày','Tổng giá trị order'], ascending=[True,False]).reset_index(drop=True)
 
    doanh_thu_va_hoa_don_theo_ngay = (
        group_by.groupby(['Ngày','Cửa hàng'], as_index=False)
        .agg(
            tong_gia_tri_order_trong_ngay=('Thành tiền','sum'),
            tong_hoa_don_trong_ngay=('Số hóa đơn','nunique')
        )
        .rename(columns={
            'tong_gia_tri_order_trong_ngay':'Tổng giá trị order trong ngày',
            'tong_hoa_don_trong_ngay':'Tổng hóa đơn trong ngày'
        })
    )
 
    Tonghop_kpi = pd.merge(result, ty_le[['Ngày','Nhân viên','% xin thông tin KH']], on=['Ngày','Nhân viên'], how='left')
    Tonghop_kpi = pd.merge(Tonghop_kpi, doanh_thu_va_hoa_don_theo_ngay, on=['Ngày','Cửa hàng'], how='left')
    Tonghop_kpi = Tonghop_kpi.reindex(columns=[
        'Ngày','Tuần','Tháng','Năm','Cửa hàng','Nhân viên','Tổng giá trị order',
        'Tổng số lần hủy món','Tổng số lượng hủy món','Tổng số khách','Doanh thu trên KH',
        '% xin thông tin KH','Tổng giá trị order trong ngày','Tổng hoá đơn theo nhân viên','Tổng hóa đơn trong ngày'
    ])
 
    Databanhang = Databanhang.rename(columns={
        'Số hóa đơn':'Số hoá đơn','Tên hàng':'Món',
        'Số lượng':'Số lượng order','Thành tiền':'Giá trị order'
    })

    Huymon = Huymon.rename(columns={
        'Số hóa đơn':'Số hoá đơn'
    })

    group_by = group_by.rename(columns={
        'Số hóa đơn':'Số hoá đơn','Thành tiền':'Giá trị order'
    })

    Tonghop_kpi = Tonghop_kpi.rename(columns={
        'Tổng hóa đơn trong ngày':'Tổng hoá đơn trong ngày'
    })


    DataKHTT_Cuahang = DataKHTT_Cuahang.rename(columns={
        'Số hóa đơn':'Số hoá đơn','Loại khách hàng':'Customer Type'
    })

    return Databanhang, Huymon, group_by, Tonghop_kpi, DataKHTT_Cuahang
 

# ============================================================
# VI. ĐẨY LÊN GOOGLE SHEETS
# ============================================================

def push_to_sheets(Databanhang, Huymon, group_by, Tonghop_kpi, DataKHTT_Cuahang):
    log.info("📤 Đẩy data lên Google Sheets...")
    gc = connect_gsheet()
    spreadsheet = gc.open_by_url(SHEET_URL)

    append_to_sheet(gc, spreadsheet, 0, Databanhang)
    append_to_sheet(gc, spreadsheet, 1, Huymon)
    append_to_sheet(gc, spreadsheet, 2, group_by)
    append_to_sheet(gc, spreadsheet, 3, Tonghop_kpi)
    append_to_sheet(gc, spreadsheet, 4, DataKHTT_Cuahang)

    log.info("🎉 Đẩy lên Google Sheets thành công!")

# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    log.info("🚀 Bắt đầu phân tích data iPos...")
    try:
        # Đọc data
        Hoadontheothoigian, Chitiethoadon, Nhatkyorder_raw, Huymon_raw = load_data()

        # Xử lý
        Hoadontheothoigian = process_hoadon(Hoadontheothoigian, Chitiethoadon)
        Nhatkyorder        = process_nhatky(Nhatkyorder_raw)
        Huymon             = process_huymon(Huymon_raw)

        # Tính KPI
        Databanhang, Huymon, group_by, Tonghop_kpi, DataKHTT_Cuahang = compute_kpi(
            Hoadontheothoigian, Nhatkyorder, Huymon
        )

        # Đẩy lên Sheets
        push_to_sheets(Databanhang, Huymon, group_by, Tonghop_kpi, DataKHTT_Cuahang)

        log.info("✅ Hoàn tất!")
        print('{"status": "success"}')

    except Exception as e:
        log.error(f"❌ Lỗi: {e}", exc_info=True)
        print(f'{{"status": "error", "message": "{str(e)}"}}')
        sys.exit(1)
