from datetime import datetime, date, timedelta
import win32com.client as win32
from tkinter import filedialog
import subprocess
import calendar
import time
import os



# logo
print("""  		    _       _____   ____   _   _   _____    ____   	  _____  __  __
            /\     | |     |  ___| / __ \ | \ | | / ____|  / __ \ 	 / ____||  \/  |
           /  \    | |     | |_	  | |  | ||  \| || (___   | |  | |	| |  __ | |\/| |
          / /\ \   | |     |  _|  | |  | || . \`| \___ \  | |  | |	| | |_ || |  | |
         / ____ \  | |___  | |	  | |__| || |\  | ____) | | |__| |	| |__| || |  | |
        /_/    \_\ |_____| |_|     \____/ |_| \_||_____/   \____/ 	 \_____||_|  |_|

        ----------------------------------------------------------------------------
        
        """)





print("\n"+"// Código creado por Alfonso Gómez Medel @ 2025 //"+"\n")

#
filtro = subprocess.Popen(["python", "filtro.py"])
filtro.wait()

#
today=calendar.day_name[date.today().weekday()]
hoy=date.today().strftime("%d-%m-%Y")

anteayer = (date.today() - timedelta(days=2)).strftime("%d-%m-%Y")
ayer = datetime.now() - timedelta(days=1)
ayer = ayer.strftime("%d-%m-%Y")


month = date.today().strftime("%m")
month = int(month)

ayermonth = datetime.now() - timedelta(days=2)
ayermonth = ayermonth.strftime("%m")
ayermonth = int(ayermonth)

anio = date.today().strftime("%Y")
anio = int(anio)

#meses (dir)
mesesDic = {
    "1":'Enero',
    "2":'Febrero',
    "3":'Marzo',
    "4":'Abril',
    "5":'Mayo',
    "6":'Junio',
    "7":'Julio',
    "8":'Agosto',
    "9":'Septiembre',
    "10":'Octubre',
    "11":'Noviembre',
    "12":'Diciembre'
}

nmonth=mesesDic[str(ayermonth)[:3]]

#meses (excel)
mesesDic2 = {
    "1":'ENERO',
    "2":'FEBRERO',
    "3":'MARZO',
    "4":'ABRIL',
    "5":'MAYO',
    "6":'JUNIO',
    "7":'JULIO',
    "8":'AGOSTO',
    "9":'SEPTIEMBRE',
    "10":'OCTUBRE',
    "11":'NOVIEMBRE',
    "12":'DICIEMBRE'
}

Emonth=mesesDic2[str(ayermonth)][:3]



#archivo_excel = filedialog.askopenfilename()
archivo_excel = r'E:\pruebas.xlsx'

if os.path.exists(archivo_excel):
    print("")
else:
    print(f"Atención! El archivo '{archivo_excel}' no existe.")
    time.sleep(5)
    exit(1)
    
excel = win32.Dispatch('Excel.Application')
excel.Visible = False #hidden


workbook = excel.Workbooks.Open(archivo_excel)

clear1 = workbook.Sheets("-")
#clear1.Visible = True

empresa01_op = workbook.Sheets("empresa01_op")
#empresa01_op.Visible = True

empresa02_op = workbook.Sheets("empresa02_op")
#empresa02_op.Visible = True

empresa03_op = workbook.Sheets("empresa03_op")
#empresa03_op.Visible = True

empresa04_op = workbook.Sheets("empresa04_op")
#empresa04_op.Visible = True

empresa05_op = workbook.Sheets("empresa05_op")
#empresa05_op.Visible = True

clear2 = workbook.Sheets("--")
#clear2.Visible = True

xxx3 = workbook.Sheets("xxx3")
#xxx3.Visible = True

filtro = workbook.Sheets("filtro")
#filtro.Visible = True

clear3 = workbook.Sheets("---")
#clear3.Visible = True

empresa01 = workbook.Sheets("empresa01")
#empresa01.Visible = True

empresa02 = workbook.Sheets("empresa02")
#empresa01.Visible = True

empresa03 = workbook.Sheets("empresa03")
#empresa03.Visible = True

empresa04 = workbook.Sheets("empresa04")
#empresa04.Visible = True

empresa05 = workbook.Sheets("empresa05")
#empresa05.Visible = True



# update -> tablas dinámicas
tabla_inspecymant = filtro.PivotTables("TablaDinámica2")
tabla_inspecymant.RefreshTable()

tabla_total = filtro.PivotTables("TablaDinámica1")
tabla_total.RefreshTable()

time.sleep(4)


#
workbook.Save()
workbook.Close()
excel.Quit()

print("\n"+"i: Actualización de datos realizada correctamente.")
time.sleep(1)

#
print("\n"+"Proceso completado con éxito. ✅")
time.sleep(5)
