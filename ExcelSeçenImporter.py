import os
import datetime
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename  # Dosya seçmek için
import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore

# Firebase projesi için özel anahtar
cred = credentials.Certificate('/Users/berhansaydam/projectsFolder/JAVA/firebase-adminsdk.json')
firebase_admin.initialize_app(cred)
db = firestore.client()

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
        for index, row in data.iterrows():
            record = {str(k): v for k, v in row.to_dict().items()}
            db.collection('Reports').add(record)
        num += 1
        print(f"{os.path.basename(file_path)} dosyasının {sheet_name} sekmesi başarıyla işlendi.")
        
except Exception as e:
    print(f"{os.path.basename(file_path)} dosyası işlenirken hata oluştu: {e}")

end_time = datetime.datetime.now()
print("End time: ", end_time)
print("Total time: ", end_time - start_time)
print(f"Toplam {num} sekme veritabanına eklendi.")
