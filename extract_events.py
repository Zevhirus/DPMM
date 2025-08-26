import pandas as pd
from datetime import datetime, timedelta

# Ganti dengan lokasi file di laptopmu
file_path = "KALENDER ORMAWA TEMPLATE.xlsx"

# Load file Excel
xls = pd.ExcelFile(file_path)
august_df = xls.parse("August")

# Fungsi konversi tanggal dari serial Excel
def excel_date_to_datetime(excel_serial):
    return datetime(1899, 12, 30) + timedelta(days=int(excel_serial))

events = []
current_date = None

# Loop semua kolom (hari dalam bulan)
for col in august_df.columns:
    for val in august_df[col]:
        if isinstance(val, (int, float)) and not pd.isna(val):
            # Simpan tanggal saat ini
            current_date = excel_date_to_datetime(val).strftime("%Y-%m-%d")
        elif isinstance(val, str) and val.strip() != "" and not val.startswith("Unnamed"):
            if current_date:
                # Pisahkan nama kegiatan dan organisasi
                parts = val.strip().rsplit(" ", 1)
                if len(parts) == 2:
                    title, org = parts
                else:
                    title, org = val.strip(), ""
                events.append({
                    "date": current_date,
                    "title": title.strip(),
                    "where": org.strip()
                })

# Cetak hasil dalam format const events
print("const events = [")
for e in events:
    print(f"  {{ date: '{e['date']}', title: '{e['title']}', where: '{e['where']}' }},")
print("];")
