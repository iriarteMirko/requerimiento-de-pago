from tkinter import messagebox, filedialog
from ..database.conexion import conexionSQLite


def seleccionar_dacxanalista():
    archivo_excel = filedialog.askopenfilename(
        initialdir="/",
        title="Seleccionar archivo DACxANALISTA",
        filetypes=(("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
    )
    dacxanalista_path = archivo_excel
    
    query = ("""UPDATE RUTAS
                SET DACXANALISTA == '""" + dacxanalista_path + """'
                WHERE ID == 0""")
    conexion = conexionSQLite()
    try:
        cursor = conexion.cursor()
        cursor.execute(query)
        conexion.commit()
    except Exception as ex:
        messagebox.showerror("Error", str(ex))
    finally:
        cursor.close()
        conexion.close

def seleccionar_dac_cdr():
    archivo_excel = filedialog.askopenfilename(
        initialdir="/",
        title="Seleccionar archivo Base DAC y CDR",
        filetypes=(("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
    )
    dac_cdr_path = archivo_excel
    
    query = ("""UPDATE RUTAS
                SET BASE_DAC_CDR == '""" + dac_cdr_path + """'
                WHERE ID == 0""")
    conexion = conexionSQLite()
    try:
        cursor = conexion.cursor()
        cursor.execute(query)
        conexion.commit()
    except Exception as ex:
        messagebox.showerror("Error", str(ex))
    finally:
        cursor.close()
        conexion.close