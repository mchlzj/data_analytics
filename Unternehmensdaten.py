from datetime import date, timedelta
from xmlrpc.client import DateTime
from openpyxl import Workbook, load_workbook
import random
import datetime
import math

wb = load_workbook('Unternehmensdaten.xlsx')

beginning_date = datetime.date(2021,1,1)
end_date = datetime.date(2022,1,1)
row_mitarbeiter = 2
id_mitarbeiter = 1

ABTEILUNG_MONTAGE = 'montage'
ABTEILUNG_EINZELTEIL_A = 'einzelteil_a'
ABTEILUNG_EINZELTEIL_B = 'einzelteil_b'

ANWESEND = 'anwesend'
ABWESEND = 'abwesend'
ws_ma = wb['Anwesenheit']
ws_ma['A1'].value = "id"
ws_ma['B1'].value = "datum"
ws_ma['C1'].value = 'mitarbeiter_id'
ws_ma['D1'].value = "abteilung"
ws_ma['E1'].value = "anwesenheit"
mitarbeiter_montage = 10
mitarbeiter_eintelteil_a = 4
mitarbeiter_einzelteil_b = 4

TYP_EINZELTEIL_A = "einzelteil_a"
TYP_EINZELTEIL_B = "einzelteil_b"

ws_et = wb['Einzelteile']
ws_et['A1'].value = "id"
ws_et['B1'].value = "datum"
ws_et['C1'].value = "typ"
ws_et['D1'].value = "fertigungszeit"
ws_et['E1'].value = "qualitaetsmessung"
ws_et['F1'].value = "maschine_id"
menge_einzelteile_a = random.randrange(130,150,1)
fertigungszeit_einzelteil_a = random.randrange(500,700,1)
maschinen_id_einzelteil_a = 1
qualitaetsmessung_einzelteil_a = random.randrange(90,100,1)

ws_mt = wb['Montage']
ws_mt['A1'].value = "id"
ws_mt['B1'].value = "datum"
ws_mt['C1'].value = "montagezeit"

menge_montage = random.randrange(57,83,1)
row_montage = 2
id_montage = 1

ws_ms = wb['Maschinen']
ws_ms['A1'].value = "id"
ws_ms['B1'].value = "datum"
ws_ms['C1'].value = "maschinen_id"
ws_ms['D1'].value = 'typ'
ws_ms['E1'].value = 'datum_letzte_wartung'

row_maschine = 2
id_maschine = 1
TYP_MASCHINE_A = "A"
TYP_MASCHINE_B = "B"

beginning_date = datetime.date(2021,1,1)
row_einzelteil = 2
id_einzelteil = 1

regression0 = 0.00
regression05 = 0.015
regression1 = 0.03
regression15 = 0.045
regression2 = 0.06
regression25 = 0.075
regression3 = 0.09
regression35 = 0.105
regression4 = 0.12
regression45 = 0.135
regression5 = 0.15
regression6 = 0.18

while (beginning_date < end_date):

    if(beginning_date.weekday() == 5 or beginning_date.weekday() == 6):
        print(beginning_date)
        beginning_date += datetime.timedelta(days=1)
    else:
        mitarbeiter_id = 1
        for mitarbeiter in range(mitarbeiter_montage):
            ws_ma['B' + str(row_mitarbeiter)].value = beginning_date
            ws_ma['A' + str(row_mitarbeiter)].value = id_mitarbeiter
            ws_ma['C' + str(row_mitarbeiter)].value = mitarbeiter_id
            ws_ma['D' + str(row_mitarbeiter)].value = ABTEILUNG_MONTAGE

            if mitarbeiter_id % 5 == 0:
                my_status = [ANWESEND, ABWESEND]
                ws_ma['E' + str(row_mitarbeiter)].value = random.choice(my_status)
            else:
                ws_ma['E' + str(row_mitarbeiter)].value = ANWESEND

            mitarbeiter_id += 1
            id_mitarbeiter += 1
            row_mitarbeiter += 1
        
        mitarbeiter_id = 11
        for mitarbeiter in range(mitarbeiter_eintelteil_a):
            ws_ma['B' + str(row_mitarbeiter)].value = beginning_date
            ws_ma['A' + str(row_mitarbeiter)].value = id_mitarbeiter
            ws_ma['C' + str(row_mitarbeiter)].value = mitarbeiter_id
            ws_ma['D' + str(row_mitarbeiter)].value = ABTEILUNG_EINZELTEIL_A

            if mitarbeiter_id % 4 == 0:
                my_status = [ANWESEND, ABWESEND]
                ws_ma['E' + str(row_mitarbeiter)].value = random.choice(my_status)
            else:
                ws_ma['E' + str(row_mitarbeiter)].value = ANWESEND

            mitarbeiter_id += 1
            id_mitarbeiter += 1
            row_mitarbeiter += 1

        mitarbeiter_id = 15
        for mitarbeiter in range(mitarbeiter_einzelteil_b):
            ws_ma['B' + str(row_mitarbeiter)].value = beginning_date
            ws_ma['A' + str(row_mitarbeiter)].value = id_mitarbeiter
            ws_ma['C' + str(row_mitarbeiter)].value = mitarbeiter_id
            ws_ma['D' + str(row_mitarbeiter)].value = ABTEILUNG_EINZELTEIL_B

            if mitarbeiter_id % 4 == 0:
                my_status = [ANWESEND, ABWESEND]
                ws_ma['E' + str(row_mitarbeiter)].value = random.choice(my_status)
            else:
                ws_ma['E' + str(row_mitarbeiter)].value = ANWESEND

            mitarbeiter_id += 1
            id_mitarbeiter += 1
            row_mitarbeiter += 1

        if (beginning_date >= datetime.date(2021,1,1) and beginning_date < datetime.date(2021,1,15)) or (beginning_date >= datetime.date(2021,4,1) and beginning_date < datetime.date(2021,4,15)) or (beginning_date >= datetime.date(2021,7,1) and beginning_date < datetime.date(2021,7,15)) or (beginning_date >= datetime.date(2021,10,1) and beginning_date < datetime.date(2021,10,15)):
            menge_einzelteile_a = random.randrange(math.ceil(130 - 130*regression0),math.ceil(150 - 150*regression1),1)
            menge_montage = random.randrange(math.ceil(57 -57*regression1), math.ceil(83 -83*regression1),1)
            for menge in range(menge_einzelteile_a):
                fertigungszeit_einzelteil_a = random.randrange(math.ceil(500 + 500*regression0), math.ceil(700 + 700*regression1))
                qualitaetsmessung_einzelteil_a = random.randrange(math.ceil(90 - 90*regression0), math.ceil(100 - 100*regression0))
                ws_et['A' + str(row_einzelteil)].value = id_einzelteil
                ws_et['B' + str(row_einzelteil)].value = beginning_date
                ws_et['C' + str(row_einzelteil)].value = TYP_EINZELTEIL_A
                ws_et['D' + str(row_einzelteil)].value = fertigungszeit_einzelteil_a
                ws_et['E' + str(row_einzelteil)].value = qualitaetsmessung_einzelteil_a
                ws_et['F' + str(row_einzelteil)].value = maschinen_id_einzelteil_a
                row_einzelteil+=1
                id_einzelteil+=1
            
            for menge in range(menge_montage):
                montagezeit = random.randrange(5000,7000,1)
                ws_mt['A' + str(row_montage)].value = id_montage
                ws_mt['B' + str(row_montage)].value = beginning_date
                ws_mt['C' + str(row_montage)].value = montagezeit
                row_montage += 1
                id_montage += 1
            
        elif (beginning_date >= datetime.date(2021,1,15) and beginning_date < datetime.date(2021,2,1)) or (beginning_date >= datetime.date(2021,4,15) and beginning_date < datetime.date(2021,5,1)) or (beginning_date >= datetime.date(2021,7,15) and beginning_date < datetime.date(2021,8,1)) or (beginning_date >= datetime.date(2021,10,15) and beginning_date < datetime.date(2021,11,1)):
            menge_einzelteile_a = random.randrange(math.ceil(130 - 130*regression1),math.ceil(150 - 150*regression2),1)
            menge_montage = random.randrange(math.ceil(57 -57*regression2), math.ceil(83 -83*regression2),1)
            for menge in range(menge_einzelteile_a):
                fertigungszeit_einzelteil_a = random.randrange(math.ceil(500 + 500*regression0), math.ceil(700 + 700*regression2))
                qualitaetsmessung_einzelteil_a = random.randrange(math.ceil(90 - 90*regression05), math.ceil(100 - 100*regression0))
                ws_et['A' + str(row_einzelteil)].value = id_einzelteil
                ws_et['B' + str(row_einzelteil)].value = beginning_date
                ws_et['C' + str(row_einzelteil)].value = TYP_EINZELTEIL_A
                ws_et['D' + str(row_einzelteil)].value = fertigungszeit_einzelteil_a
                ws_et['E' + str(row_einzelteil)].value = qualitaetsmessung_einzelteil_a
                ws_et['F' + str(row_einzelteil)].value = maschinen_id_einzelteil_a
                row_einzelteil+=1
                id_einzelteil+=1

            for menge in range(menge_montage):
                montagezeit = random.randrange(5000,7000,1)
                ws_mt['A' + str(row_montage)].value = id_montage
                ws_mt['B' + str(row_montage)].value = beginning_date
                ws_mt['C' + str(row_montage)].value = montagezeit
                row_montage += 1
                id_montage += 1

        elif (beginning_date >= datetime.date(2021,2,1) and beginning_date < datetime.date(2021,2,15)) or (beginning_date >= datetime.date(2021,5,1) and beginning_date < datetime.date(2021,5,15)) or (beginning_date >= datetime.date(2021,8,1) and beginning_date < datetime.date(2021,8,15)) or (beginning_date >= datetime.date(2021,11,1) and beginning_date < datetime.date(2021,11,15)):
            menge_einzelteile_a = random.randrange(math.ceil(130 - 130*regression2),math.ceil(150 - 150*regression3),1)
            menge_montage = random.randrange(math.ceil(57 -57*regression3), math.ceil(83 -83*regression3),1)
            for menge in range(menge_einzelteile_a):
                fertigungszeit_einzelteil_a = random.randrange(math.ceil(500 + 500*regression1), math.ceil(700 + 700*regression2))
                qualitaetsmessung_einzelteil_a = random.randrange(math.ceil(90 - 90*regression1), math.ceil(100 - 100*regression0))
                ws_et['A' + str(row_einzelteil)].value = id_einzelteil
                ws_et['B' + str(row_einzelteil)].value = beginning_date
                ws_et['C' + str(row_einzelteil)].value = TYP_EINZELTEIL_A
                ws_et['D' + str(row_einzelteil)].value = fertigungszeit_einzelteil_a
                ws_et['E' + str(row_einzelteil)].value = qualitaetsmessung_einzelteil_a
                ws_et['F' + str(row_einzelteil)].value = maschinen_id_einzelteil_a
                row_einzelteil+=1
                id_einzelteil+=1

            for menge in range(menge_montage):
                montagezeit = random.randrange(5000,7000,1)
                ws_mt['A' + str(row_montage)].value = id_montage
                ws_mt['B' + str(row_montage)].value = beginning_date
                ws_mt['C' + str(row_montage)].value = montagezeit
                row_montage += 1
                id_montage += 1

        elif (beginning_date >= datetime.date(2021,2,15) and beginning_date < datetime.date(2021,3,1)) or (beginning_date >= datetime.date(2021,5,15) and beginning_date < datetime.date(2021,6,1)) or (beginning_date >= datetime.date(2021,8,15) and beginning_date < datetime.date(2021,9,1)) or (beginning_date >= datetime.date(2021,11,15) and beginning_date < datetime.date(2021,12,1)):
            menge_einzelteile_a = random.randrange(math.ceil(130 - 130*regression3),math.ceil(150 - 150*regression4),1)
            menge_montage = random.randrange(math.ceil(57 -57*regression4), math.ceil(83 -83*regression4),1)
            for menge in range(menge_einzelteile_a):
                fertigungszeit_einzelteil_a = random.randrange(math.ceil(500 + 500*regression1), math.ceil(700 + 700*regression3))
                qualitaetsmessung_einzelteil_a = random.randrange(math.ceil(90 - 90*regression1), math.ceil(100 - 100*regression0))
                ws_et['A' + str(row_einzelteil)].value = id_einzelteil
                ws_et['B' + str(row_einzelteil)].value = beginning_date
                ws_et['C' + str(row_einzelteil)].value = TYP_EINZELTEIL_A
                ws_et['D' + str(row_einzelteil)].value = fertigungszeit_einzelteil_a
                ws_et['E' + str(row_einzelteil)].value = qualitaetsmessung_einzelteil_a
                ws_et['F' + str(row_einzelteil)].value = maschinen_id_einzelteil_a
                row_einzelteil+=1
                id_einzelteil+=1

            for menge in range(menge_montage):
                montagezeit = random.randrange(5000,7000,1)
                ws_mt['A' + str(row_montage)].value = id_montage
                ws_mt['B' + str(row_montage)].value = beginning_date
                ws_mt['C' + str(row_montage)].value = montagezeit
                row_montage += 1
                id_montage += 1

        elif (beginning_date >= datetime.date(2021,3,1) and beginning_date < datetime.date(2021,3,15)) or (beginning_date >= datetime.date(2021,6,1) and beginning_date < datetime.date(2021,6,15)) or (beginning_date >= datetime.date(2021,9,1) and beginning_date < datetime.date(2021,9,15)) or (beginning_date >= datetime.date(2021,12,1) and beginning_date < datetime.date(2021,12,15)):
            menge_einzelteile_a = random.randrange(math.ceil(130 - 130*regression4),math.ceil(150 - 150*regression5),1)
            menge_montage = random.randrange(math.ceil(57 - 57*regression5), math.ceil(83 - 83*regression5),1)
            for menge in range(menge_einzelteile_a):
                fertigungszeit_einzelteil_a = random.randrange(math.ceil(500 + 500*regression1), math.ceil(700 + 700*regression4))
                qualitaetsmessung_einzelteil_a = random.randrange(math.ceil(90 - 90*regression15), math.ceil(100 - 100*regression0))
                ws_et['A' + str(row_einzelteil)].value = id_einzelteil
                ws_et['B' + str(row_einzelteil)].value = beginning_date
                ws_et['C' + str(row_einzelteil)].value = TYP_EINZELTEIL_A
                ws_et['D' + str(row_einzelteil)].value = fertigungszeit_einzelteil_a
                ws_et['E' + str(row_einzelteil)].value = qualitaetsmessung_einzelteil_a
                ws_et['F' + str(row_einzelteil)].value = maschinen_id_einzelteil_a
                row_einzelteil+=1
                id_einzelteil+=1

            for menge in range(menge_montage):
                montagezeit = random.randrange(5000,7000,1)
                ws_mt['A' + str(row_montage)].value = id_montage
                ws_mt['B' + str(row_montage)].value = beginning_date
                ws_mt['C' + str(row_montage)].value = montagezeit
                row_montage += 1
                id_montage += 1

        elif (beginning_date >= datetime.date(2021,3,15) and beginning_date < datetime.date(2021,4,1)) or (beginning_date >= datetime.date(2021,6,15) and beginning_date < datetime.date(2021,7,1)) or (beginning_date >= datetime.date(2021,9,15) and beginning_date < datetime.date(2021,10,1)) or (beginning_date >= datetime.date(2021,12,15) and beginning_date <= datetime.date(2021,12,31)):
        # else: 
            menge_einzelteile_a = random.randrange(math.ceil(130 - 130*regression5),math.ceil(150 - 150*regression6),1)
            menge_montage = random.randrange(math.ceil(57 -57*regression6), math.ceil(83 -83*regression6),1)
            for menge in range(menge_einzelteile_a):
                fertigungszeit_einzelteil_a = random.randrange(math.ceil(500 + 500*regression2), math.ceil(700 + 700*regression4))
                qualitaetsmessung_einzelteil_a = random.randrange(math.ceil(90 - 90*regression15), math.ceil(100 - 100*regression0))
                ws_et['A' + str(row_einzelteil)].value = id_einzelteil
                ws_et['B' + str(row_einzelteil)].value = beginning_date
                ws_et['C' + str(row_einzelteil)].value = TYP_EINZELTEIL_A
                ws_et['D' + str(row_einzelteil)].value = fertigungszeit_einzelteil_a
                ws_et['E' + str(row_einzelteil)].value = qualitaetsmessung_einzelteil_a
                ws_et['F' + str(row_einzelteil)].value = maschinen_id_einzelteil_a
                row_einzelteil+=1
                id_einzelteil+=1

            for menge in range(menge_montage):
                montagezeit = random.randrange(5000,7000,1)
                ws_mt['A' + str(row_montage)].value = id_montage
                ws_mt['B' + str(row_montage)].value = beginning_date
                ws_mt['C' + str(row_montage)].value = montagezeit
                row_montage += 1
                id_montage += 1

        menge_einzelteile_b = random.randrange(130,150,1)
        for menge in range(menge_einzelteile_b):
            fertigungszeit_einzelteil_b = random.randrange(500,700,1)
            maschinen_id_einzelteil_b = 2
            qualitaetsmessung_einzelteil_b = random.randrange(90, 100, 1)
            ws_et['A' + str(row_einzelteil)].value = id_einzelteil
            ws_et['B' + str(row_einzelteil)].value = beginning_date
            ws_et['C' + str(row_einzelteil)].value = TYP_EINZELTEIL_B
            ws_et['D' + str(row_einzelteil)].value = fertigungszeit_einzelteil_b
            ws_et['E' + str(row_einzelteil)].value = qualitaetsmessung_einzelteil_b
            ws_et['F' + str(row_einzelteil)].value = maschinen_id_einzelteil_b
            row_einzelteil+=1
            id_einzelteil+=1

        if beginning_date < datetime.date(2021,4,1):
            ws_ms['A' + str(row_maschine)].value = id_maschine
            ws_ms['B' + str(row_maschine)].value = beginning_date
            ws_ms['C' + str(row_maschine)].value = 1
            ws_ms['D' + str(row_maschine)].value = TYP_MASCHINE_A
            ws_ms['E' + str(row_maschine)].value = datetime.date(2021,1,1)
            row_maschine += 1
            id_maschine += 1

            ws_ms['A' + str(row_maschine)].value = id_maschine
            ws_ms['B' + str(row_maschine)].value = beginning_date
            ws_ms['C' + str(row_maschine)].value = 2
            ws_ms['D' + str(row_maschine)].value = TYP_MASCHINE_B
            ws_ms['E' + str(row_maschine)].value = datetime.date(2021,1,1)
            row_maschine += 1
            id_maschine += 1

        elif beginning_date < datetime.date(2021,7,1):
            ws_ms['A' + str(row_maschine)].value = id_maschine
            ws_ms['B' + str(row_maschine)].value = beginning_date
            ws_ms['C' + str(row_maschine)].value = 1
            ws_ms['D' + str(row_maschine)].value = TYP_MASCHINE_A
            ws_ms['E' + str(row_maschine)].value = datetime.date(2021,4,1)

            row_maschine += 1
            id_maschine += 1

            ws_ms['A' + str(row_maschine)].value = id_maschine
            ws_ms['B' + str(row_maschine)].value = beginning_date
            ws_ms['C' + str(row_maschine)].value = 2
            ws_ms['D' + str(row_maschine)].value = TYP_MASCHINE_B
            ws_ms['E' + str(row_maschine)].value = datetime.date(2021,4,1)
            row_maschine += 1
            id_maschine += 1

        elif beginning_date < datetime.date(2021,10,1):
            ws_ms['A' + str(row_maschine)].value = id_maschine
            ws_ms['B' + str(row_maschine)].value = beginning_date
            ws_ms['C' + str(row_maschine)].value = 1
            ws_ms['D' + str(row_maschine)].value = TYP_MASCHINE_A
            ws_ms['E' + str(row_maschine)].value = datetime.date(2021,7,1)

            row_maschine += 1
            id_maschine += 1     

            ws_ms['A' + str(row_maschine)].value = id_maschine
            ws_ms['B' + str(row_maschine)].value = beginning_date
            ws_ms['C' + str(row_maschine)].value = 2
            ws_ms['D' + str(row_maschine)].value = TYP_MASCHINE_B
            ws_ms['E' + str(row_maschine)].value = datetime.date(2021,7,1)
            row_maschine += 1
            id_maschine += 1 

        elif beginning_date < datetime.date(2022,1,1):
            ws_ms['A' + str(row_maschine)].value = id_maschine
            ws_ms['B' + str(row_maschine)].value = beginning_date
            ws_ms['C' + str(row_maschine)].value = 1
            ws_ms['D' + str(row_maschine)].value = TYP_MASCHINE_A
            ws_ms['E' + str(row_maschine)].value = datetime.date(2021,10,1)
            row_maschine += 1
            id_maschine += 1 

            ws_ms['A' + str(row_maschine)].value = id_maschine
            ws_ms['B' + str(row_maschine)].value = beginning_date
            ws_ms['C' + str(row_maschine)].value = 2
            ws_ms['D' + str(row_maschine)].value = TYP_MASCHINE_B
            ws_ms['E' + str(row_maschine)].value = datetime.date(2021,10,1)
            row_maschine += 1
            id_maschine += 1  

        beginning_date += datetime.timedelta(days=1)
wb.save('Unternehmensdaten.xlsx')

