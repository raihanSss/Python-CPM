import csv
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

# Menambahkan header kolom di file Excel
ws.append(["Kode", "Durasi", "Volume", "Produktivitashari(%)", "Produktivitasjam(%)", "ProduktivitasHarianPercepatan", "CrashDuration", "Selisih"])

with open('input_crashing1.csv') as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=';')
    next(csv_reader)  

    durasi_data = []  
    volume_data = []  

    for row in csv_reader:
        if len(row) < 3:
            continue  

        kode = row[0]
        durasi = float(row[1].replace(",", "."))
        volume = float(row[2].replace(",", "."))

        durasi_data.append(durasi)  # simpen data durasi
        volume_data.append(volume)  # simpen data volume

        # rumus
        produktivitashari = volume / durasi
        produktivitasjam = produktivitashari / 8
        produktivitas_harian_percepatan = produktivitashari + produktivitasjam
        crash_duration = volume / produktivitas_harian_percepatan
        selisih = abs(volume - produktivitas_harian_percepatan)

        ws.append([kode, durasi, volume, produktivitashari, produktivitasjam, produktivitas_harian_percepatan, crash_duration, selisih])

    # Mengecek apakah terdapat 3 data durasi dan volume
    if len(durasi_data) != 3 or len(volume_data) != 3:
        print("Jumlah data durasi atau volume tidak sesuai. Pastikan ada 3 data durasi dan volume.")
        exit()

    # memakai data durasi dan volume yang ada
    durasi1 = durasi_data[0]
    durasi2 = durasi_data[1]
    durasi3 = durasi_data[2]

    volume1 = volume_data[0]
    volume2 = volume_data[1]
    volume3 = volume_data[2]


wb.save('crashing1.xlsx')
