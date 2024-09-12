import pandas as pd
import openpyxl
import re

# Excel dosyasının adını belirtin (*** yerine dosya adı gelmeli)
excel_file = '***.xlsx'

# Excel dosyasını yükleyin
wb = openpyxl.load_workbook(excel_file, data_only=True)

# Belirli bir sayfa yüklenir (*** yerine sayfa adı gelmeli)
ws = wb['***']

# Bu fonksiyon, hücredeki bağlantıdan hyperlink bilgilerini ayıklar.
# Link 'http' ile başlıyorsa, içindeki bilgileri çıkarır ve geri döndürür.
def extract_info(hyperlink):
    if hyperlink and hyperlink.startswith("http"):
        match = re.search(r'q=(-?\d+\.\d+),(-?\d+\.\d+)', hyperlink)
        if match:
            return float(match.group(1)), float(match.group(2))
    return None, None

# 'D' sütunundaki her hücreyi kontrol ederiz ve hücrede bir bağlantı (hyperlink) varsa
# ondan hyperlink bilgilerini çekeriz
info_1 = []
info_2 = []

for cell in ws['D']:
    if cell.hyperlink:
        data1, data2 = extract_info(cell.hyperlink.target)
        info_1.append(data1)
        info_2.append(data2)
    else:
        info_1.append(None) # Eğer link yoksa, None değeri ekler
        info_2.append(None)

# Başka bir sayfadan verileri yükleriz (*** yerine sayfa adı gelmeli)
df2 = pd.read_excel(excel_file, sheet_name='***')

# Çekilen bilgileri DataFrame'e ekliyoruz. Başlık satırı olduğu için [1:] ile atlıyoruz.
df2['Info1'] = info_1[1:]  
df2['Info2'] = info_2[1:] 

# Diğer sayfalardan verileri yükleriz (*** yerine sayfa adı gelmeli)
df1 = pd.read_excel(excel_file, sheet_name='***')
df3 = pd.read_excel(excel_file, sheet_name='***')

# Timestamp bilgisini datetime formatına çeviriyoruz
df1['***'] = pd.to_datetime(df1['***'])
df2['***'] = pd.to_datetime(df2['***'])
df3['***'] = pd.to_datetime(df3['***'])

# Her bir DataFrame için timestamp'i index olarak ayarlıyoruz
time_df1 = df1.set_index('***')
time_df2 = df2.set_index('***')
time_df3 = df3.set_index('***')

# örn. 1 saniyelik periyotlarla verileri özetliyoruz. Burada bazı sütunlar toplama veya ortalama
# işlemi ile özetleniyor (*** yerine sütun adları gelmeli).
df1_sampled = time_df1.resample('1S').agg({
    '***': 'sum',  # Bu sütundaki değerler toplanır
    '***': 'mean', # Bu sütundaki değerlerin ortalaması alınır
    '***': 'sum',  # Bu sütundaki değerler toplanır
    '***': 'sum',  # Bu sütundaki değerler toplanır
    '***': 'sum'   # Bu sütundaki değerler toplanır
})

# Çekilen bilgileri örn. 1 dakikalık periyotlarla ortalıyoruz 
generic_data = time_df2[['Info1', 'Info2']].resample('1T').mean() 

# '***' bilgisini örn. 1 dakikalık periyotlarla ortalıyoruz
df2_sampled = time_df2[['***']].resample('1T').mean()  

# '***' bilgilerini örn. 1 saniyelik periyotlarla ortalıyoruz (*** sütun adları gelmeli)
df3_sampled = time_df3[['***', '***', '***']].resample('1S').mean() 

# df1_sampled verisini ve df2_sampled verisini zaman indexine göre birleştiriyoruz
merged_df = pd.merge(df1_sampled, df2_sampled, left_index=True, right_index=True, how='outer')

# generic_data bilgisini birleştiriyoruz
merged_df = pd.merge(merged_df, generic_data, left_index=True, right_index=True, how='outer')

# Son olarak üçüncü tabloyu (df3_sampled) da zaman indexine göre birleştiriyoruz
final_merged_df = pd.merge(merged_df, df3_sampled, left_index=True, right_index=True, how='outer')

# Boş olan hücreler varsa, bir önceki geçerli değerle dolduruyoruz (forward fill)
final_merged_df.fillna(method='ffill', inplace=True)

# Zaman bilgisini 'Yıl-Ay-Gün Saat:Dakika:Saniye' formatına çeviriyoruz
final_merged_df.index = final_merged_df.index.strftime('%Y-%m-%d %H:%M:%S')

# Sonuçları Excel dosyasına kaydediyoruz, her satırın zamanını 'Timestamp' başlığı altında yazıyoruz
final_merged_df.to_excel('sonuc.xlsx', index_label='Timestamp')
