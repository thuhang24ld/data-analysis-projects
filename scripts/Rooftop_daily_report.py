# -*- coding: utf-8 -*-


import os
import ast
import warnings
import pandas as pd
from datetime import datetime, timedelta
 
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe, set_with_dataframe
 
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ── CẤU HÌNH ──────────────────────────────────────────────────────────────────

CREDENTIALS_PATH = r"CREDENTIALS_FILE"  # <-- đường dẫn file JSON service account

FOLDER = r"D:\Rooftop_IPos_data\Hoa_don"

# Google Sheet URL
SHEET_URL = "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit#gid=0"  # <-- URL Google Sheets (thay SPREADSHEET_ID)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

STORE_MAP = {
    "XXXX": "RC",
    "XXXX": "THU",
    "XXXX": "CVGTN",
    "XXXX": "CVGHX",
    "XXXX": "CVH",
}

# ── XÁC THỰC GOOGLE (Service Account) ────────────────────────────────────────
 
def get_gspread_client() -> gspread.Client:
    creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
    return gspread.authorize(creds)
 
# ── ĐỌC FILE HÓA ĐƠN ─────────────────────────────────────────────────────────
 
def load_invoice_file(folder: str) -> pd.DataFrame:
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    filename  = f"hoa_don_{yesterday}.xlsx"
    filepath  = os.path.join(folder, filename)
 
    if os.path.exists(filepath):
        df = pd.read_excel(filepath)
        print(f"Đọc được {len(df)} rows từ {filename}")
    else:
        files = sorted([f for f in os.listdir(folder) if f.endswith(".xlsx")])
        if files:
            filepath = os.path.join(folder, files[-1])
            df = pd.read_excel(filepath)
            print(f"Không tìm thấy file hôm qua, dùng file mới nhất: {files[-1]} ({len(df)} rows)")
        else:
            raise FileNotFoundError("Không có file nào trong folder!")
 
    return df[[
        "_store_uid", "tran_id", "origin_tran_id", "tran_no",
        "created_at", "start_hour", "start_minute", "end_hour", "end_minute",
        "table_name", "amount_discount_detail", "extra_data", "total_amount"
    ]]
 
# ── XỬ LÝ DỮ LIỆU ────────────────────────────────────────────────────────────
 
def extract_fields(x):
    try:
        data = ast.literal_eval(x)
        return pd.Series({
            "peo_count":            data.get("peo_count"),
            "customer_phone":       data.get("customer_phone"),
            "customer_name":        data.get("customer_name"),
            "Membership_Type_Name": data.get("Membership_Type_Name"),
        })
    except Exception:
        return pd.Series([None, None, None, None])
 
 
def process(df: pd.DataFrame) -> pd.DataFrame:
    df[["peo_count", "customer_phone", "customer_name", "Membership_Type_Name"]] = \
        df["extra_data"].apply(extract_fields)
    mask = df["end_hour"] == 00
    df.loc[mask, "created_at"] = (df.loc[mask, "created_at"] - pd.Timedelta(days=1))
 
    df["Giờ vào"] = (df["start_hour"].astype(str).str.zfill(2) + ":" +
                     df["start_minute"].astype(str).str.zfill(2))
    df["Giờ ra"]  = (df["end_hour"].astype(str).str.zfill(2) + ":" +
                     df["end_minute"].astype(str).str.zfill(2))
 
    df = df.rename(columns={
        "start_hour":             "Hour_in",
        "end_hour":               "Hour_out",     
        "_store_uid":             "Mã cửa hàng",
        "tran_id":                "Mã hóa đơn",
        "origin_tran_id":         "Mã hóa đơn gốc",
        "tran_no":                "Số hóa đơn",
        "created_at":             "Ngày",
        "table_name":             "Bàn",
        "amount_discount_detail": "Giảm giá",
        "total_amount":           "Net Revenue",
        "customer_phone":         "SĐT",
        "customer_name":          "Tên",
        "peo_count":              "Số khách",
        "Membership_Type_Name":   "Loại thành viên",
    })
    df = df.drop(columns=["extra_data", "start_minute", "end_minute"])
    
    df["Month-Year"] = df["Ngày"].dt.month.astype(str) + "-" + df["Ngày"].dt.year.astype(str)
    iso_calendar = df["Ngày"].dt.isocalendar()
    df["Week-Year"] = iso_calendar["week"].astype(str) + "-" + iso_calendar["year"].astype(str)
 
    df["Cửa hàng"] = df["Mã cửa hàng"].map(STORE_MAP).fillna("Chi nhánh là")
    print("Thống kê số đơn theo từng quán:")
    print(df["Cửa hàng"].value_counts(), "\n")
 
    df[["Bàn", "Khu vực"]] = df["Bàn"].str.split("-", n=1, expand=True)
    df["Khu vực"] = df["Khu vực"].str.strip()
 
    df["Ngày"] = (pd.to_datetime(df["Ngày"], format="%d/%m/%Y", errors="coerce")
                    .dt.strftime("%d/%m/%Y"))
 
    t_in  = pd.to_datetime(df["Giờ vào"], format="%H:%M", errors="coerce")
    t_out = pd.to_datetime(df["Giờ ra"],  format="%H:%M", errors="coerce")
    t_out = t_out.where(t_out >= t_in, t_out + pd.Timedelta(days=1))
    df["Thời gian sử dụng (giờ)"] = ((t_out - t_in).dt.total_seconds() / 3600).round(2)
 
    df["Gross Revenue"] = df["Net Revenue"] + df["Giảm giá"]
 
    cols = [
        "Month-Year","Week-Year","Hour_in","Hour_out","Cửa hàng", "Số hóa đơn", "Ngày", "Giờ vào", "Giờ ra",
        "Bàn", "Khu vực", "Số khách", "Gross Revenue", "Giảm giá",
        "Net Revenue", "Loại thành viên", "Tên", "SĐT",
        "Thời gian sử dụng (giờ)"
    ]
    return df[cols].sort_values(["Ngày", "Giờ vào"]).reset_index(drop=True)
 
# ── ĐẨY LÊN GOOGLE SHEET ─────────────────────────────────────────────────────
 
def push_to_sheet(gc: gspread.Client, invoice_info: pd.DataFrame):
    spreadsheet = gc.open_by_url(SHEET_URL)
    worksheet   = spreadsheet.worksheet("hoa don ipos")
 
    old_data = get_as_dataframe(worksheet, evaluate_formulas=True).dropna(how="all")
    final    = pd.concat([old_data, invoice_info], ignore_index=True)
    set_with_dataframe(worksheet, final)
    print(f"Da day {len(invoice_info)} dong moi len Google Sheet (tong: {len(final)} dong)")
 
# ── MAIN ──────────────────────────────────────────────────────────────────────
 
def main():
    print("=== Rooftop Daily Report ===\n")
 
    raw          = load_invoice_file(FOLDER)
    invoice_info = process(raw)
    print(f"Số hóa đơn xu ly: {len(invoice_info)}\n")
 
    print("Đang xác thực Google...")
    gc = get_gspread_client()
    push_to_sheet(gc, invoice_info)
 
    print("\nHoàn thành!")
 
 
if __name__ == "__main__":
    main()
 