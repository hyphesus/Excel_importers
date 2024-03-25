import os
import datetime
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from sqlalchemy import create_engine

Tk().withdraw()  

# Kullanıcıya bir dosya seçtiren diyalog penceresini aç
file_path = askopenfilename(title='Excel Dosyası Seçin', filetypes=[('Excel files', '*.xlsx')])

if not file_path:
    print("Dosya seçilmedi.")
    exit()

print("Seçilen dosya: ", file_path)

# MySQL veritabanı bağlantı bilgileri
DB_USERNAME = 'your_username'
DB_PASSWORD = 'your_password'
DB_HOST = 'localhost'
DB_PORT = '3306'
DB_NAME = 'your_database_name'

# SQLAlchemy engine oluştur
engine = create_engine(f'mysql+pymysql://{DB_USERNAME}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}')

start_time = datetime.datetime.now()
print("Start time: ", start_time)
num = 0

# Seçilen dosyayı işle
try:
    xls = pd.ExcelFile(file_path, engine='openpyxl')
    for sheet_name in xls.sheet_names:
        data = pd.read_excel(xls, sheet_name=sheet_name)
        # Verileri MySQL tablosuna ekle
        data.to_sql(name=sheet_name, con=engine, if_exists='append', index=False)
        num += 1
        print(f"{os.path.basename(file_path)} dosyasının {sheet_name} sekmesi MySQL veritabanına başarıyla işlendi.")
        
except Exception as e:
    print(f"{os.path.basename(file_path)} dosyası işlenirken hata oluştu: {e}")

end_time = datetime.datetime.now()
print("End time: ", end_time)
print("Total time: ", end_time - start_time)
print(f"Toplam {num} sekme veritabanına eklendi.")
