import pdfplumber
import pandas as pd
from datetime import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os

# Đường dẫn đến file PDF
file_path = "D:\\KFSP\\automatic\\iRMiA_02082024 viet.pdf"

with pdfplumber.open(file_path) as pdf:
    all_data = []
    for page in pdf.pages:
        tables = page.extract_tables()
        for table in tables:
            all_data.extend(table)

df = pd.DataFrame(all_data[1:], columns=all_data[0])

new_data = []
stop_adding = False

for index, row in df.iterrows():
    if stop_adding:
        break
    if isinstance(row.iloc[0], str) and 'SÀN ĐẠI CHÚNG CHƯA NIÊM YẾT' in row.iloc[0]:
        stop_adding = True
        break
    if isinstance(row.iloc[0], str) and '\n' in row.iloc[0]:
        rows = row.iloc[0].split('\n')
        for r in rows:
            new_row = r.split()
            new_data.append(new_row)
    else:
        new_data.append(row.tolist())

df_cleaned = pd.DataFrame(new_data, columns=df.columns)
df_cleaned = df_cleaned[~df_cleaned['STT'].isin(['STT', 'SÀN'])]
df_cleaned = df_cleaned[df_cleaned['Mã CK'] != '2']

df_final = df_cleaned.iloc[:, [1, 4, 5, 6]].copy()
df_final.columns = ['MA_CK', 'SLCP_SOHUU', 'PHAN_TRAM_SO_HUU', 'ROOM_CON_LAI']

# Chuyển đổi các giá trị trong SLCP_SOHUU và ROOM_CON_LAI về dạng float
df_final['SLCP_SOHUU'] = df_final['SLCP_SOHUU'].str.replace('.', '').str.replace(',', '.').astype(float)
df_final['ROOM_CON_LAI'] = df_final['ROOM_CON_LAI'].str.replace('.', '').str.replace(',', '.').astype(float)

# Chuyển đổi giá trị % sang giá trị tuyệt đối và đảm bảo không làm tròn số
df_final['PHAN_TRAM_SO_HUU'] = df_final['PHAN_TRAM_SO_HUU'].str.replace('%', '').astype(float) / 100

# Định dạng số theo kiểu chuẩn của Mỹ với ngăn cách phần nghìn bằng dấu phẩy và phần thập phân bằng dấu chấm

df_final['PHAN_TRAM_SO_HUU'] = df_final['PHAN_TRAM_SO_HUU'].apply(lambda x: f"{x:,.5f}")

today = datetime.now().strftime('%m/%d/%Y')
df_final.insert(0, 'NGAY', today)

file_date = datetime.now().strftime('%Y-%m-%d')
folder_date = datetime.now().strftime('%Y-%m')
output_path = f'D:\\KFSP\\automatic\\{file_date}.xlsx'

# Ghi file trực tiếp vào đường dẫn hiện tại
df_final.to_excel(output_path, index=False)

print(f"File saved successfully to {output_path}")

# Thiết lập xác thực Google Drive và lưu token vào file credentials.json
gauth = GoogleAuth()
gauth.LoadCredentialsFile("credentials.json")

if not gauth.credentials or gauth.credentials.invalid:
    gauth.LocalWebserverAuth()
    gauth.SaveCredentialsFile("credentials.json")

drive = GoogleDrive(gauth)

# Tìm kiếm thư mục đích
folder_name = 'VSD - NUOC NGOAI'
file_list = drive.ListFile({'q': f"title='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"}).GetList()

if len(file_list) == 0:
    raise FileNotFoundError(f"Folder '{folder_name}' not found on Google Drive")

folder_id = file_list[0]['id']

# Tìm hoặc tạo thư mục con theo định dạng YYYY-MM
subfolder_name = folder_date
subfolder_list = drive.ListFile({'q': f"title='{subfolder_name}' and mimeType='application/vnd.google-apps.folder' and '{folder_id}' in parents and trashed=false"}).GetList()

if len(subfolder_list) == 0:
    subfolder_metadata = {
        'title': subfolder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [{'id': folder_id}]
    }
    subfolder = drive.CreateFile(subfolder_metadata)
    subfolder.Upload()
    subfolder_id = subfolder['id']
else:
    subfolder_id = subfolder_list[0]['id']

# Tải file lên Google Drive vào thư mục con
file_drive = drive.CreateFile({
    'title': f'{file_date}.xlsx',
    'parents': [{'id': subfolder_id}]
})
file_drive.SetContentFile(output_path)
file_drive.Upload()

print(f"File uploaded successfully to folder '{folder_name}/{subfolder_name}' with title '{file_date}.xlsx'")
