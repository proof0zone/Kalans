# Tabel incarcari verticale
from openpyxl import Workbook
wb = Workbook() # current workbook
ws = wb.active  # default worksheet

#A1:X1 (CAP TABEL)
ws.merge_cells("A1:X1")
ct = ws["A1"]
measure_units = {'m1': 'm', 'm2': 'm^2'}
th_text = {'A1': 'ÎNCĂRCĂRI VERTICALE PE ȘPALEȚI',}

wb.save("ivps.xlsx")