import os
import datetime
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from pymongo import MongoClient

# MongoDB bağlantı bilgileri
MONGO_URI = "mongodb://localhost:27017"  # MongoDb bağlantı adresi (URI) 
client = MongoClient(MONGO_URI)
db = client.get_database("reports")  # Veritabanı adı
collection = db["reports"]  # Koleksiyon adı

Tk().withdraw()  # Tkinter root penceresini gizle

# Kullanıcıya bir dosya seçtiren diyalog penceresini aç
file_path = askopenfilename(title='Excel Dosyası Seçin', filetypes=[('Excel files', '*.xlsx')])

if not file_path:
    print("Dosya seçilmedi.")
    exit()

print("Seçilen dosya: ", file_path)

start_time = datetime.datetime.now()
print("Start time: ", start_time)
num = 0

# Seçilen dosyayı işle
try:
    xls = pd.ExcelFile(file_path, engine='openpyxl')
    for sheet_name in xls.sheet_names:
        data = pd.read_excel(xls, sheet_name=sheet_name)
        records = data.to_dict('records')  # Veriyi MongoDB için uygun bir listeye dönüştür
        collection.insert_many(records)  # MongoDB koleksiyonuna verileri ekle
        num += 1
        print(f"{os.path.basename(file_path)} dosyasının {sheet_name} sekmesi başarıyla işlendi.")
        
except Exception as e:
    print(f"{os.path.basename(file_path)} dosyası işlenirken hata oluştu: {e}")

end_time = datetime.datetime.now()
print("End time: ", end_time)
print("Total time: ", end_time - start_time)
print(f"Toplam {num} sekme veritabanına eklendi.")
