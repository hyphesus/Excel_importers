import os
import datetime
import openpyxl
from tkinter import Tk
from tkinter.filedialog import askdirectory
import pandas as pd
from sqlalchemy import create_engine
import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore

# Firebase projesi için özel anahtar 
cred = credentials.Certificate('/Users/berhansaydam/projectsFolder/JAVA/firebase-adminsdk.json')
# Firebase uygulamasını başlat
firebase_admin.initialize_app(cred)
# Firestore veritabanı nesnesi oluştur
db = firestore.client()

# Tkinter penceresini başlatma ve gizleme
Tk().withdraw() 
# Kullanıcıya bir klasör seçtiren diyalog penceresini aç
path = askdirectory(title='Excel Dosyalarının Bulunduğu Klasörü Seçin')
files = os.listdir(path)


if path:  # Eğer bir klasör seçilirse
    print("Seçilen klasör: ", path)
else:
    print("Klasör seçilmedi.")

# Başlangıç zamanını al
start_time = datetime.datetime.now()
print("Start time: ", start_time)
num = 0

for file_name in files:
    file_path = os.path.join(path, file_name)
    # Dosya uzantısını kontrol edin
    if file_path.endswith('.xlsx'):
        engine='openpyxl'
        # Excel dosyasını açın ve tüm sekmeleri alın
        xls = pd.ExcelFile(file_path, engine=engine)
        for sheet_name in xls.sheet_names:
            try:
                # Her sekmedeki verileri okuyun
                data = pd.read_excel(xls, sheet_name=sheet_name)
                # Burada verileri işleyin...
                   # Excel dosyasındaki her satır için
                for index, row in data.iterrows():
                    record = row.to_dict()
                    # Anahtar türlerini standart hale getir
                    record = {str(k): v for k, v in record.items()}
                    # Firestore'a ekle
                    db.collection('Reports').add(record)
                num += 1
                # Firestore koleksiyonuna eklemek için uygun kod buraya gelecek
                print(f"{file_name} dosyasının {sheet_name} sekmesi başarıyla işlendi.")
                print(num, " ", file_name, " dosyası veritabanına eklendi.")
            except Exception as e:
                print(f"{file_name} dosyasının {sheet_name} sekmesi işlenirken hata oluştu: {e}")
    else:
        raise ValueError("Desteklenmeyen dosya formatı: {}".format(os.path.splitext(file_name)[1]))

    data = pd.read_excel(os.path.join(path,file_name), header=0, engine=engine)


    
# Bitiş zamanını al
end_time = datetime.datetime.now()

print("End time: ", end_time)
print("Total time: ", end_time - start_time)
print("Toplam ", num, " dosya veritabanına eklendi.")

