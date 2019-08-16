import csv
import glob
from openpyxl import Workbook

print(' 0. Starte Programm.')

###########

print(' 1. Versuche Textdatei im selben Ordner zu finden...')

files = glob.glob('./*.txt')
if len(files) == 0:
    print('FEHLER: Keine Textdatei im gleichen Ordner gefunden!')
    raise SystemExit
elif len(files) > 1:
    print('FEHLER: Mehr als eine Textdatei im gleichen Ordner gefunden!')
    raise SystemExit
filename = files[0]

print('    Datei gefunden:', filename)

###########

print(' 2. Lese Text-Datei ein:', filename)

data = []
with open(filename, 'r', encoding='cp1252') as csvfile:
    csv_reader = csv.reader(csvfile, delimiter=';', quotechar='"')
    for row in csv_reader:
        data.append(row)

print('    Erfolgreich', len(data), 'Zeilen eingelesen')

###########

print(' 4. Entferne Leerzeichen...')

for i in range(len(data)):
    data[i] = [col.strip() for col in data[i]]

print('    Leerzeichen erfolgreich entfernt.')

###########

print(' 5. Loesche Spalten A, B, C, L, M...')

for row in data:
    del row[11:13]
    del row[0:3]

print('    Spalten erfolgreich gelöscht.')

###########

print(' 6. Verschiebe Spalte A nach C...')

for i in range(len(data)):
    row = data[i]
    data[i] = [row[1], row[2], row[0], row[3], row[4], row[5], row[6], row[7], row[8]]

print('    Spalten erfolgreich verschoben.')

###########

print(' 7. Entferne Duplikate, basierend auf Spalte A')

uniques = []
old_count = len(data)
new_data = []
for row in data:
    if row[0] not in uniques:
        uniques.append(row[0])
        new_data.append(row)
data = new_data
    
print('    Insgesamt', old_count - len(data), 'Duplikate entfernt.')

###########

print(' 8. Formatiere Spalte C auf 7-stellige Nummern (führende Nullen)')

count = 0
for row in data[1:]:
    if len(row[2]) < 7:
        count += 1
        row[2] = row[2].zfill(7)
    
print('    Insgesamt', count, 'führende Nullen eingefügt.')

###########

print(' 9. Formatiere Spalten E & F auf Preisformat')

count = 0
for row in data[1:]:
    cell_e = row[4]
    cell_e = cell_e[:-2]    
    if len(cell_e) > 6:
        count += 1
        cell_e = cell_e[:-6] + "'" + cell_e[-6:]
    elif cell_e[0] == '.':
        cell_e = '0' + cell_e
    row[4] = cell_e

    cell_f = row[5]
    cell_f = cell_f[:-2]
    if len(cell_f) > 6:
        count += 1
        cell_f = cell_f[:-6] + "'" + cell_f[-6:]
    elif cell_f[0] == '.':
        cell_f = '0' + cell_f
    row[5] = cell_f

print('    Zellen erfolgreich formatiert,', count, 'Tausender-Trennzeichen eingefügt.')

###########

print('10. Lösche Zeilen mit Präfix 9 in Spalte C')

data = [row for row in data if row[2][0] != '9']

print('    Erfolgreich gefiltert, es bleiben', len(data), 'Zeilen.')

###########

print('11. Lösche Wert in Spalte F, wenn dieser 0.00 beträgt...')

count = 0
for row in data[1:]:
    if row[5] == '0.00':
        count += 1
        row[5] = ''

print('   ', count, 'Werte erfolgreich gelöscht.')

###########

print('12. Füge Stern hinzu, wenn Spalte I nicht in blacklist...')

count = 0
blacklist = ['FOO', 'BAR', 'TEST', '']
print('    Aktuelle blacklist:', blacklist)

for row in data[1:]:
    if row[8] not in blacklist:
        row[0] = row[0] + ' *'
        row[1] = row[1] + ' *'
        count += 1

print('    Stern bei', count, 'Zeilen hinzugefügt.')

###########

print('13. Spalte G einstellig formatieren...')

for row in data[1:]:
    row[6] = row[6][-1]

print('    Spalte G erfolgreich formatiert.')

###########

print('14. Sortiere nach erster Spalte alphabetisch...')

header = data[0]
del data[0]
data.sort(key=lambda x: x[0].replace('Ö', 'O').replace('Ü', 'U').replace('Ä', 'A'))
data.insert(0, header)

print('    Daten alphabetisch sortiert.')

###########

print('15. IN: Speichere in.xlsx...')

wb = Workbook()
ws = wb.active
count = 0
for row in data:
    if row[7] != '7':
        ws.append(row[:7])
        count += 1
wb.save("in.xlsx")

print('    IN: Datei in.xlsx erfolgreich gespeichert mit', count, 'Zeilen.')

###########

print('16. OUT: Speichere out.xlsx...')

wb = Workbook()
ws = wb.active
count = 0
for row in data:
    if count == 0 or row[7] == '7':
        ws.append([row[0], row[1], row[2], row[3], row[5], row[6]])
        count += 1
wb.save("out.xlsx")

print('    OUT: Datei out.xlsx erfolgreich gespeichert mit', count, 'Zeilen.')

input('Prozess beendet. Beliebige Taste drücken zum fortfahren...')
