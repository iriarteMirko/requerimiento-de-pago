import warnings
import time

from customtkinter import *
from tkinter import messagebox, filedialog
from datetime import datetime
from threading import Thread

from src.database.conexion import *
from src.models.validar_data import validar_cuentas, validar_analistas
from src.models.generar_dataframes import generar_dataframes
from src.models.generar_doc import generar_doc
from src.models.generar_excel import generar_excel


warnings.filterwarnings("ignore")


class Cartas():
    def __init__(self):
        self.base = resource_path("BASE.xlsx")
        self.modelo_1 = resource_path("./modelos/MODELO_1.docx")
        self.modelo_2 = resource_path("./modelos/MODELO_2.docx")
        self.unidades = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"]
        self.diez_a_diecinueve = ["diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve"]
        self.veintiuno_a_veintinueve = ["", "veintiuno", "veintidos", "veintitres", "veinticuatro", "veinticinco", "veintiseis", "veintisiete", "veintiocho", "veintinueve"]
        self.decenas = ["", "", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"]
        self.centenas = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"]
        self.analistas_validados = ["WALTER LOPEZ", "YOLANDA OLIVA", "JUAN CARLOS HUATAY", "RAQUEL CAYETANO", "JOSE LUIS VALVERDE", "DIEGO RODRIGUEZ"]
        self.correos_analistas = {
            "Walter Lopez" : "wlopez@claro.com.pe",
            "Yolanda Oliva" : "yolanda.oliva@claro.com.pe",
            "Juan Carlos Huatay" : "juan.huatay@claro.com.pe",
            "Raquel Cayetano" :"rcayetano@claro.com.pe",
            "Jose Luis Valverde" : "jvalverde@claro.com.pe",
            "Diego Rodriguez": "diego.rodriguez@claro.com.pe"}
        meses = {
            "01": "enero",
            "02": "febrero",
            "03": "marzo",
            "04": "abril",
            "05": "mayo",
            "06": "junio",
            "07": "julio",
            "08": "agosto",
            "09": "septiembre",
            "10": "octubre",
            "11": "noviembre",
            "12": "diciembre"}
        hoy = datetime.today()
        dia = hoy.strftime("%d")
        mes = hoy.strftime("%m")
        año = hoy.strftime("%Y")
        nombre_mes = meses.get(mes)
        self.fecha_hoy = f"{dia} de {nombre_mes} de {año}"
        print(f"Fecha hoy: {self.fecha_hoy}\n")
    
    def deshabilitar_botones(self):
        self.boton_ejecutar.configure(state="disabled")
        self.boton_dacx.configure(state="disabled")
        self.boton_dac_cdr.configure(state="disabled")
    
    def habilitar_botones(self):
        self.boton_ejecutar.configure(state="normal")
        self.boton_dacx.configure(state="normal")
        self.boton_dac_cdr.configure(state="normal")
    
    def verificar_thread(self, thread):
        if thread.is_alive():
            self.app.after(1000, self.verificar_thread, thread)
        else:
            self.habilitar_botones()
    
    def iniciar_tarea(self):
        self.deshabilitar_botones()
        thread = Thread(target=self.ejecutar)
        thread.start()
        self.app.after(1000, self.verificar_thread, thread)
    
    def generar_cartas_requerimiento_pago(self):
        dataframes = generar_dataframes(self.base, self.ruta_dacxa, self.ruta_dac_cdr)
        self.df_base = dataframes[0]
        self.df_cruce = dataframes[1]
        
        print(f"Registros Base: [{self.df_base.shape[0]}]\n")
        cuentas_base = self.df_base["Cuenta"].drop_duplicates().to_list()
        cuentas_cruce = self.df_cruce["Deudor"].to_list()
        analistas = self.df_cruce["ANALISTA_ACT"].drop_duplicates().to_list()
        
        cuentas = validar_cuentas(cuentas_base, cuentas_cruce)
        validar_analistas(analistas, self.analistas_validados)
        
        for cuenta in cuentas:
            self.df_cuenta = self.df_base[self.df_base["Cuenta"] == cuenta]
            if (self.df_cuenta["Demora"] >= 0).all():
                self.generar_cartas_sin_deudaxvencer(cuenta)
            else:
                self.generar_cartas_con_deudaxvencer(cuenta)
    
    def generar_cartas_sin_deudaxvencer(self, cuenta):
        razon_social = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["NOMBRE DAC"].iloc[0].upper()
        direccion_legal = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["DIRECCIÓN LEGAL"].iloc[0]
        distrito = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["DISTRITO"].iloc[0]
        provincia = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["PROVINCIA"].iloc[0]
        departamento = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["DPTO."].iloc[0]
        dias_demora = self.df_cuenta["Demora"].iloc[0]
        
        deuda_vencida = round(self.df_cuenta["Importe"].sum(),2)
        parte_entera_deuda_vencida, parte_decimal_deuda_vencida = self.separar_entero_decimal(deuda_vencida)
        deuda_vencida_soles = f"S/ {parte_entera_deuda_vencida}.{parte_decimal_deuda_vencida}"
        parte_entera_deuda_vencida_a_texto = self.numero_entero_a_texto(int(parte_entera_deuda_vencida))
        deuda_vencida_texto = f"({parte_entera_deuda_vencida_a_texto} con {parte_decimal_deuda_vencida}/100 soles)"
        
        analista_mayuscula = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["ANALISTA_ACT"].iloc[0]
        analista = " ".join([palabra.capitalize() for palabra in analista_mayuscula.lower().split(" ")])
        correo_analista = self.correos_analistas.get(analista)
        
        dias_demora_2 = dias_demora
        razon_social_2 = razon_social
        
        ruta_doc = resource_path("./results/"+razon_social+".docx")
        self.df_cuenta.to_excel(resource_path("./results/"+razon_social+".xlsx"), index=False) # Sin deudas por vencer
        
        replacements = {
            "[fecha_hoy]": {"value": str(self.fecha_hoy), "font_size": 11},
            "[razon_social]": {"value": str(razon_social), "font_size": 11, "bold": True},
            "[direccion_legal]": {"value": str(direccion_legal), "font_size": 11},
            "[distrito]": {"value": str(distrito), "font_size": 11},
            "[provincia]": {"value": str(provincia), "font_size": 11},
            "[departamento]": {"value": str(departamento), "font_size": 11},
            "[dias_demora]": {"value": str(dias_demora), "font_size": 11},
            "[deuda_vencida_soles]": {"value": str(deuda_vencida_soles), "font_size": 11},
            "[deuda_vencida_texto]": {"value": str(deuda_vencida_texto), "font_size": 11},
            "[analista]": {"value": str(analista), "font_size": 11},
            "[correo_analista]": {"value": str(correo_analista), "font_size": 11},
            "[dias_demora_2]": {"value": str(dias_demora_2), "font_size": 8},
            "[razon_social_2]": {"value": str(razon_social_2), "font_size": 8, "bold": True},
        }
        
        generar_doc(self.modelo_2, replacements, ruta_doc)
        generar_excel(razon_social)

    def generar_cartas_con_deudaxvencer(self, cuenta):
        razon_social = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["NOMBRE DAC"].iloc[0].upper()
        direccion_legal = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["DIRECCIÓN LEGAL"].iloc[0]
        distrito = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["DISTRITO"].iloc[0]
        provincia = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["PROVINCIA"].iloc[0]
        departamento = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["DPTO."].iloc[0]
        dias_demora = self.df_cuenta["Demora"].iloc[0]
        
        deuda_vencida = round(self.df_cuenta[self.df_cuenta["Demora"] >= 0]["Importe"].sum(),2)
        parte_entera_deuda_vencida, parte_decimal_deuda_vencida = self.separar_entero_decimal(deuda_vencida)
        deuda_vencida_soles = f"S/ {parte_entera_deuda_vencida}.{parte_decimal_deuda_vencida}"
        parte_entera_deuda_vencida_a_texto = self.numero_entero_a_texto(int(parte_entera_deuda_vencida))
        deuda_vencida_texto = f"({parte_entera_deuda_vencida_a_texto} con {parte_decimal_deuda_vencida}/100 soles)"
        
        deuda_por_vencer = round(self.df_cuenta[self.df_cuenta["Demora"] < 0]["Importe"].sum(),2)
        parte_entera_deuda_por_vencer, parte_decimal_deuda_por_vencer = self.separar_entero_decimal(deuda_por_vencer)
        deuda_por_vencer_soles = f"S/ {parte_entera_deuda_por_vencer}.{parte_decimal_deuda_por_vencer}"
        parte_entera_deuda_por_vencer_a_texto = self.numero_entero_a_texto(int(parte_entera_deuda_por_vencer))
        deuda_por_vencer_texto = f"({parte_entera_deuda_por_vencer_a_texto} con {parte_decimal_deuda_por_vencer}/100 soles)"
        
        analista_mayuscula = self.df_cruce[self.df_cruce["Deudor"]==cuenta]["ANALISTA_ACT"].iloc[0]
        analista = " ".join([palabra.capitalize() for palabra in analista_mayuscula.lower().split(" ")])
        correo_analista = self.correos_analistas.get(analista)
        
        dias_demora_2 = dias_demora
        razon_social_2 = razon_social
        
        ruta_doc = resource_path("./results/"+razon_social+".docx")
        self.df_cuenta.to_excel(resource_path("./results/"+razon_social+".xlsx"), index=False) # Con deudas por vencer
        
        replacements = {
            "[fecha_hoy]": {"value": str(self.fecha_hoy), "font_size": 11},
            "[razon_social]": {"value": str(razon_social), "font_size": 11, "bold": True},
            "[direccion_legal]": {"value": str(direccion_legal), "font_size": 11},
            "[distrito]": {"value": str(distrito), "font_size": 11},
            "[provincia]": {"value": str(provincia), "font_size": 11},
            "[departamento]": {"value": str(departamento), "font_size": 11},
            "[dias_demora]": {"value": str(dias_demora), "font_size": 11},
            "[deuda_vencida_soles]": {"value": str(deuda_vencida_soles), "font_size": 11},
            "[deuda_vencida_texto]": {"value": str(deuda_vencida_texto), "font_size": 11},
            "[deuda_por_vencer_soles]": {"value": str(deuda_por_vencer_soles), "font_size": 11},
            "[deuda_por_vencer_texto]": {"value": str(deuda_por_vencer_texto), "font_size": 11},
            "[analista]": {"value": str(analista), "font_size": 11},
            "[correo_analista]": {"value": str(correo_analista), "font_size": 11},
            "[dias_demora_2]": {"value": str(dias_demora_2), "font_size": 8},
            "[razon_social_2]": {"value": str(razon_social_2), "font_size": 8, "bold": True},
        }
        
        generar_doc(self.modelo_1, replacements, ruta_doc)
        generar_excel(razon_social)

    def separar_entero_decimal(self, numero):
        numero_str = str(numero)
        if "." not in numero_str:
            parte_entera = numero_str
            parte_decimal = "00"
        else:
            parte_entera, parte_decimal = numero_str.split(".")
        
        if len(parte_decimal) > 2:
            parte_decimal = parte_decimal[:2]
        elif len(parte_decimal) < 2:
            parte_decimal = parte_decimal.ljust(2, "0")
        
        return parte_entera, parte_decimal

    def numero_entero_a_texto(self, num):
        if num == 0:
            return "cero"
        grupos = []
        while num > 0:
            grupos.append(num % 1000)
            num //= 1000
        textos = [self.convertir_grupo_a_texto(grupo) for grupo in grupos]
        if len(textos) > 1:
            textos[1] += " mil"
        if len(textos) > 2:
            textos[2] = "un millón" if textos[2] == "uno" else textos[2] + " millones"
        return " ".join(textos[::-1]).strip()

    def convertir_grupo_a_texto(self, num):
        texto = ""
        if num >= 100:
            texto += self.centenas[num // 100]
            num %= 100
        if num >= 20:
            texto += " " + self.decenas[num // 10]
            num %= 10
        elif num >= 10:
            texto += " " + self.diez_a_diecinueve[num - 10]
            num = 0
        if num > 0:
            texto += " y " + self.unidades[num]
        return texto.strip()
    
    def seleccionar_dacxanalista(self):
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
    
    def seleccionar_base_dac_cdr(self):
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
    
    def ejecutar(self):
        self.progressbar.start()
        start = time.time()
        query = """SELECT * FROM RUTAS WHERE ID == 0"""
        try:
            datos = ejecutar_query(query)
            self.ruta_dacxa = datos[0][1]
            self.ruta_dac_cdr = datos[0][2]
            if self.ruta_dacxa is None or self.ruta_dac_cdr is None:
                messagebox.showerror("Error", "Por favor, configure las rutas de los archivos.")
            elif not os.path.exists(self.ruta_dacxa):
                messagebox.showerror("Error", "No se encontraró el archivo DACxANALISTA en la ruta especificada.")
            elif not os.path.exists(self.ruta_dac_cdr):
                messagebox.showerror("Error", "No se encontraró el archivo DAC y CDR en la ruta especificada.")
            else:
                self.generar_cartas_requerimiento_pago()
        except Exception as ex:
            messagebox.showerror("Error", "Detalle:\n" + str(ex))
        finally:
            end = time.time()
            self.progressbar.stop()
            tiempo_promedio = end - start
            print(f"Tiempo ejecución: {tiempo_promedio} segundos.")
    
    def crear_app(self):
        self.app = CTk()
        self.app.title("Generador de Cartas")
        icon_path = resource_path("./src/images/icono.ico")
        if os.path.isfile(icon_path):
            self.app.iconbitmap(icon_path)
        else:
            messagebox.showwarning("ADVERTENCIA", "No se encontró el archivo 'icono.ico' en la ruta: " + icon_path)
        self.app.resizable(False, False)
        set_appearance_mode("dark")
        
        main_frame = CTkFrame(self.app)
        main_frame.pack_propagate("True")
        main_frame.pack(fill="both", expand=True)
        
        frame_base = CTkFrame(main_frame)
        frame_base.grid(row=0, column=0, padx=(20, 10), pady=(20, 0), sticky="nsew")
        
        ruta_dacxa = CTkLabel(frame_base, text="Ruta DACxAnalista", font=("Calibri",15))
        ruta_dacxa.pack(padx=(20, 20), pady=(5, 0), fill="both", expand=True, anchor="center", side="top")
        self.boton_dacx = CTkButton(frame_base, text="Seleccionar", font=("Calibri",15), text_color="white",
                                fg_color="transparent", border_color="#d11515", border_width=2, hover_color="#d11515", 
                                width=25, corner_radius=25, command=lambda: self.seleccionar_dacxanalista())
        self.boton_dacx.pack(padx=(20, 20), pady=(0, 15), fill="both", anchor="center", side="bottom")
        
        frame_dacx = CTkFrame(main_frame)
        frame_dacx.grid(row=0, column=1, padx=(10, 20), pady=(20, 0), sticky="nsew")
        
        ruta_daccdr = CTkLabel(frame_dacx, text="Ruta DAC y CDR", font=("Calibri",15))
        ruta_daccdr.pack(padx=(20, 20), pady=(5, 0), fill="both", expand=True, anchor="center", side="top")
        self.boton_dac_cdr = CTkButton(frame_dacx, text="Seleccionar", font=("Calibri",15), text_color="white",
                                fg_color="transparent", border_color="#d11515", border_width=2, hover_color="#d11515", 
                                width=25, corner_radius=25, command=lambda: self.seleccionar_base_dac_cdr())
        self.boton_dac_cdr.pack(padx=(20, 20), pady=(0, 15), fill="both", anchor="center", side="bottom")
        
        self.boton_ejecutar = CTkButton(main_frame, text="GENERAR CARTAS", text_color="black", font=("Calibri",20,"bold"), 
                                    border_color="black", border_width=3, fg_color="gray", 
                                    hover_color="red", command=lambda: self.iniciar_tarea())
        self.boton_ejecutar.grid(row=1, column=0, columnspan=2, ipady=20, padx=(20, 20), pady=(20, 0), sticky="nsew")
        
        self.cuadro = CTkTextbox(main_frame, font=("Calibri",15), height=50, border_color="black", border_width=2)
        self.cuadro.grid(row=2, column=0, columnspan=2, padx=(20, 20), pady=(20, 0), sticky="nsew")
        self.cuadro.configure(state="disabled")
        
        self.progressbar = CTkProgressBar(main_frame, mode="indeterminate", orientation="horizontal", 
                                        progress_color="#d11515", height=10, border_width=0)
        self.progressbar.grid(row=3, column=0, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="nsew")
        
        self.app.mainloop()



def main():
    app = Cartas()
    app.crear_app()


if __name__ == "__main__":
    main()