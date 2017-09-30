# Tabel incarcari verticale
from openpyxl import Workbook
wb = Workbook() # current workbook
ws = wb.active  # default worksheet
#==========================================#==========================================
# STAGE ONE
# CREATE THE TABLE AND FORMAT IT
#==========================================#==========================================
#A1:X1 (CAP TABEL)
ws.merge_cells("A1:X1")
#A2:A3 diafrag
ws.merge_cells("A2:A3")
#B2:B3 SPALETI
ws.merge_cells("B2:B3")



measure_units = {'m1': 'm', 'm2': 'm^2'}
th_text = {'A1': 'ÎNCĂRCĂRI VERTICALE PE ȘPALEȚI', \
           'A2':'Diafrag',\
           'B2':'',\
           'C2':'',\
           'D2':'',\
           'E2':'',\
           'F2':'',\
           'G2':'',\
           'H2':'',\
           'I2':'',\
           'J2':'',\
           'K2':'',\
           'L2':'',\
           'M2':'',\
           'N2':'',\
           'O2':'',\
           'P2':'',\
           'Q2':'',\
           'R2':'',\
           'S2':'',\
           'T2':'',\
           'U2':'',\
           }

ws['A1'] = th_text['A1']

wb.save("ivps.xlsx")