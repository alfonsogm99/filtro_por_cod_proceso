import win32com.client as win32
#from tkinter import filedialog

# ---------------- config ----------------
RUTA_ARCHIVO = r'E:\pruebas.xlsx'
#RUTA_ARCHIVO = filedialog.askopenfilename()
HOJA_ORIGEN = 'xxx3'
HOJA_DESTINO = 'bbdd_ins'
HOJA_DEBUG = 'debug'
COLUMNA_FILTRO = 'cod_operacion'
PREF_VALIDOS = ('01', "'01", '06', "'06")
PREF_EXCLUIDOS = ('0150', "'0150", '0151', "'0151", '069', "'069")

#
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False 
wb = excel.Workbooks.Open(RUTA_ARCHIVO)

hoja_origen = wb.Sheets(HOJA_ORIGEN)
hoja_bbdd = wb.Sheets(HOJA_DESTINO)
hoja_debug = wb.Sheets(HOJA_DEBUG)

#
last_col = hoja_origen.UsedRange.Columns.Count
headers = [hoja_origen.Cells(1, col).Value for col in range(1, last_col + 1)]

#
try:
    col_filtro_idx = headers.index(COLUMNA_FILTRO) + 1
except ValueError:
    raise Exception(f"No se encontró la columna '{COLUMNA_FILTRO}' en la hoja {HOJA_ORIGEN}.")

#
hoja_bbdd.Rows(1).Copy(hoja_bbdd.Rows(1))
hoja_debug.Rows(1).Copy(hoja_debug.Rows(1))

#
last_row = hoja_origen.UsedRange.Rows.Count
bbdd_row = 2
debug_row = 2

for row in range(2, last_row + 1):
    valor = str(hoja_origen.Cells(row, col_filtro_idx).Text).strip()

    if any(valor.startswith(pref) for pref in PREF_VALIDOS) and not any(valor.startswith(ban) for ban in PREF_EXCLUIDOS):
        hoja_origen.Rows(row).Copy(hoja_bbdd.Rows(bbdd_row))
        bbdd_row += 1
    else:
        hoja_origen.Rows(row).Copy(hoja_debug.Rows(debug_row))
        debug_row += 1

#
wb.Save()
wb.Close(SaveChanges=True)
excel.Quit()

print(f"i: Datos separados en las hojas: '{HOJA_DESTINO}' y '{HOJA_DEBUG}'.")
