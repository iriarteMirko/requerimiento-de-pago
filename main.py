from docx import Document
from datetime import datetime
from docx.shared import Pt
import pandas as pd
import warnings
import time

warnings.filterwarnings("ignore")

def main():
    def generar_cartas_requerimiento_pago(df_base, df_cruce):
        ########## BASE ##########
        columnas_deseadas_base = [
            "Cuenta", "Nº documento", "Referencia", "Fecha de documento", "Clase de documento", 
            "Demora tras vencimiento neto", "Moneda del documento", "Importe en moneda local"]
        nuevas_columnas_base = ["Fecha de doc.", "CL", "Demora", "Moneda", "Importe"]
        df_base = df_base[columnas_deseadas_base]
        df_base = df_base[df_base["Cuenta"].notna()]

        nombres_columnas = dict(zip(columnas_deseadas_base[3:], nuevas_columnas_base))
        df_base = df_base.rename(columns=nombres_columnas)

        df_base["Cuenta"] = df_base["Cuenta"].astype("Int64").astype("str")
        df_base["Nº documento"] = df_base["Nº documento"].astype("Int64").astype("str")
        df_base["Demora"] = df_base["Demora"].astype("Int64")
        df_base.sort_values(by=["Cuenta","Demora"], ascending=[True, False], inplace=True)
        print("Base: ",df_base.shape)

        cuentas = df_base["Cuenta"].drop_duplicates().to_list()
        print("Deudores: ",cuentas,"\n")

        ########## CRUCE ##########
        df_cruce.drop(columns=["DEUDOR"], inplace=True)
        df_cruce = df_cruce[df_cruce["ANALISTA_ACT"].notna()]
        columnas_deseadas_cruce = [
            "Deudor", "NOMBRE DAC", "DIRECCIÓN LEGAL", "DISTRITO", "PROVINCIA", "DPTO.", "ANALISTA_ACT"]
        analistas_no_deseados = ["REGION NORTE", "REGION SUR", "SIN INFORMACION"]
        df_cruce = df_cruce[columnas_deseadas_cruce]
        df_cruce = df_cruce[df_cruce["Deudor"].notna()]
        df_cruce["Deudor"] = df_cruce["Deudor"].astype("Int64").astype("str")
        df_cruce = df_cruce.loc[~df_cruce["ANALISTA_ACT"].isin(analistas_no_deseados)]

        analistas = df_cruce["ANALISTA_ACT"].drop_duplicates().to_list()
        analistas_no_validados = []
        for analista in analistas:
            if analista in analistas_validados:
                pass
            else:
                analistas_no_validados.append(analista)
        if len(analistas_no_validados) == 0:
            print("Analistas validados\n")
        else:
            print("Analistas no validados: ",analistas_no_validados,"\n")
        
        cuentas_encontradas = encontrar_cuentas(df_cruce, cuentas)

        ########## CRUCE BASE ##########
        for cuenta in cuentas_encontradas:
            df_cuenta = df_base[df_base["Cuenta"] == cuenta]
            if (df_cuenta["Demora"] >= 0).all():
                generar_cartas_sin_deudaxvencer(df_cuenta, df_cruce, cuenta)
            else:
                generar_cartas_con_deudaxvencer(df_cuenta, df_cruce, cuenta)

    def encontrar_cuentas(df_cruce, cuentas):
        cuentas_no_encontradas = []
        cuentas_copia = cuentas.copy()

        for cuenta in cuentas_copia:
            if cuenta not in df_cruce["Deudor"].to_list():
                cuentas.remove(cuenta)
                cuentas_no_encontradas.append(cuenta)

        if len(cuentas_no_encontradas) == 0:
            print("Deudores validados.\n")
        else:
            print("Deudores no encontrados: ",cuentas_no_encontradas,"\n")
        
        return cuentas

    def generar_cartas_sin_deudaxvencer(df_cuenta, df_cruce, cuenta):
        razon_social = df_cruce[df_cruce["Deudor"]==cuenta]["NOMBRE DAC"].iloc[0].upper()
        direccion_legal = df_cruce[df_cruce["Deudor"]==cuenta]["DIRECCIÓN LEGAL"].iloc[0]
        distrito = df_cruce[df_cruce["Deudor"]==cuenta]["DISTRITO"].iloc[0]
        provincia = df_cruce[df_cruce["Deudor"]==cuenta]["PROVINCIA"].iloc[0]
        departamento = df_cruce[df_cruce["Deudor"]==cuenta]["DPTO."].iloc[0]
        dias_demora = df_cuenta["Demora"].iloc[0]

        deuda_vencida = round(df_cuenta["Importe"].sum(),2)
        parte_entera_deuda_vencida, parte_decimal_deuda_vencida = separar_entero_decimal(deuda_vencida)
        deuda_vencida_soles = f"S/ {parte_entera_deuda_vencida}.{parte_decimal_deuda_vencida}"
        parte_entera_deuda_vencida_a_texto = convertir_entero_a_texto(int(parte_entera_deuda_vencida))
        deuda_vencida_texto = f"({parte_entera_deuda_vencida_a_texto} con {parte_decimal_deuda_vencida}/100 soles)"
        
        analista_mayuscula = df_cruce[df_cruce["Deudor"]==cuenta]["ANALISTA_ACT"].iloc[0]
        analista = " ".join([palabra.capitalize() for palabra in analista_mayuscula.lower().split(" ")])
        correo_analista = correos_analistas.get(analista)
        
        dias_demora_2 = dias_demora
        razon_social_2 = razon_social
        
        df_cuenta.to_excel("./FINAL/"+razon_social+".xlsx", index=False) # Sin deudas por vencer
        
        #word_file = "MODELO_2.docx"
        doc = Document(modelo_2)
        replacements = {
            "[fecha_hoy]": {"value": str(fecha_hoy), "font_size": 11},
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
        
        for paragraph in doc.paragraphs:
            for key, attributes in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, attributes["value"])
                    run = paragraph.runs[0]
                    run.font.name = 'Arial'
                    run.font.size = Pt(attributes["font_size"])
                    run.bold = attributes.get("bold", False)
        
        guardar_documentos(doc, razon_social)

    def generar_cartas_con_deudaxvencer(df_cuenta, df_cruce, cuenta):
        razon_social = df_cruce[df_cruce["Deudor"]==cuenta]["NOMBRE DAC"].iloc[0].upper()
        direccion_legal = df_cruce[df_cruce["Deudor"]==cuenta]["DIRECCIÓN LEGAL"].iloc[0]
        distrito = df_cruce[df_cruce["Deudor"]==cuenta]["DISTRITO"].iloc[0]
        provincia = df_cruce[df_cruce["Deudor"]==cuenta]["PROVINCIA"].iloc[0]
        departamento = df_cruce[df_cruce["Deudor"]==cuenta]["DPTO."].iloc[0]
        dias_demora = df_cuenta["Demora"].iloc[0]
        
        deuda_vencida = round(df_cuenta[df_cuenta["Demora"] >= 0]["Importe"].sum(),2)
        parte_entera_deuda_vencida, parte_decimal_deuda_vencida = separar_entero_decimal(deuda_vencida)
        deuda_vencida_soles = f"S/ {parte_entera_deuda_vencida}.{parte_decimal_deuda_vencida}"
        parte_entera_deuda_vencida_a_texto = convertir_entero_a_texto(int(parte_entera_deuda_vencida))
        deuda_vencida_texto = f"({parte_entera_deuda_vencida_a_texto} con {parte_decimal_deuda_vencida}/100 soles)"
        
        deuda_por_vencer = round(df_cuenta[df_cuenta["Demora"] < 0]["Importe"].sum(),2)
        parte_entera_deuda_por_vencer, parte_decimal_deuda_por_vencer = separar_entero_decimal(deuda_por_vencer)
        deuda_por_vencer_soles = f"S/ {parte_entera_deuda_por_vencer}.{parte_decimal_deuda_por_vencer}"
        parte_entera_deuda_por_vencer_a_texto = convertir_entero_a_texto(int(parte_entera_deuda_por_vencer))
        deuda_por_vencer_texto = f"({parte_entera_deuda_por_vencer_a_texto} con {parte_decimal_deuda_por_vencer}/100 soles)"
        
        analista_mayuscula = df_cruce[df_cruce["Deudor"]==cuenta]["ANALISTA_ACT"].iloc[0]
        analista = " ".join([palabra.capitalize() for palabra in analista_mayuscula.lower().split(" ")])
        correo_analista = correos_analistas.get(analista)
        
        dias_demora_2 = dias_demora
        razon_social_2 = razon_social
        
        df_cuenta.to_excel("./FINAL/"+razon_social+".xlsx", index=False) # Con deudas por vencer
        
        #word_file = "MODELO_1.docx"
        doc = Document(modelo_1)
        replacements = {
            "[fecha_hoy]": {"value": str(fecha_hoy), "font_size": 11},
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
        
        for paragraph in doc.paragraphs:
            for key, attributes in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, attributes["value"])
                    run = paragraph.runs[0]
                    run.font.name = 'Arial'
                    run.font.size = Pt(attributes["font_size"])
                    run.bold = attributes.get("bold", False)
        
        guardar_documentos(doc, razon_social)

    def generar_fecha(hoy, meses):
        dia = hoy.strftime("%d")
        mes = hoy.strftime("%m")
        año = hoy.strftime("%Y")
        nombre_mes = meses.get(mes)
        return dia, nombre_mes, año

    def guardar_documentos(doc, nombre_doc):
        doc_final = "./FINAL/"+nombre_doc+".docx"
        doc.save(doc_final)

    def separar_entero_decimal(numero):
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

    def convertir_entero_a_texto(num):
        if num == 0:
            return "cero"
        texto = ""
        if num >= 1000:
            texto += miles[int(str(num)[0])]
            num %= 1000
        if num >= 100:
            texto += " " + centenas[int(str(num)[0])]
            num %= 100
        if num >= 10:
            if num < 30 and num > 20:
                texto += " " + veintiuno_a_veintinueve[num-20]
                return texto
            if num < 20 and num >10:
                texto += " " + diez_a_diecinueve[num-10]
                return texto
            texto += " " + decenas[int(str(num)[0])]
            num %= 10
        if num > 0:
            texto += " y " + unidades[num]
        return texto.strip()
    
    global analistas_validados, modelo_1, modelo_2
    
    unidades = [
        "", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"]
    diez_a_diecinueve = [
        "diez", "once", "doce", "trece", "catorce", "quince", 
        "dieciséis", "diecisiete", "dieciocho", "diecinueve"]
    veintiuno_a_veintinueve = [
        "", "veintiuno", "veintidos", "veintitres", "veinticuatro", "veinticinco", 
        "veintiseis", "veintisiete", "veintiocho", "veintinueve"]
    decenas = [
        "", "", "veinte", "treinta", "cuarenta", "cincuenta", 
        "sesenta", "setenta", "ochenta", "noventa"]
    centenas = [
        "", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", 
        "seiscientos", "setecientos", "ochocientos", "novecientos"]
    miles = [
        "", "mil", "dos mil", "tres mil", "cuatro mil", "cinco mil", 
        "seis mil", "siete mil", "ocho mil", "nueve mil"]

    analistas_validados = [
        "WALTER LOPEZ", "YOLANDA OLIVA", "JUAN CARLOS HUATAY", 
        "RAQUEL CAYETANO", "JOSE LUIS VALVERDE", "DIEGO RODRIGUEZ"]
    correos_analistas = {
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
        "12": "diciembre"
    }
    
    hoy = datetime.today()
    dia, mes, año = generar_fecha(hoy, meses)
    fecha_hoy = f"{dia} de {mes} de {año}"
    print("Fecha hoy: ", fecha_hoy,"\n")

    ########## RUTAS ##########
    base = "BASE.xlsx"
    dac_cdr = "./FUENTES/BASE DAC Y CDR ac.xlsx" #"Z:/Base Datos Contratos/base actualizada DAC Y CDR/"
    dac_x_analista = "./FUENTES/Nuevo_DACxANALISTA.xlsx" #"Z:/JEFATURA CCD/"
    modelo_1 = "MODELO_1.docx"
    modelo_2 = "MODELO_2.docx"
    
    ########## DATAFRAMES ##########
    df_base = pd.read_excel(base, sheet_name="BASE")
    df_dac_cdr = pd.read_excel(dac_cdr, sheet_name=" CONTRATOS DAC-DACES")
    df_dac_x_analista = pd.read_excel(dac_x_analista, sheet_name="Base_NUEVA")
    df_cruce = pd.merge(df_dac_cdr, df_dac_x_analista, left_on="Deudor", right_on="DEUDOR", how="left")
    
    generar_cartas_requerimiento_pago(df_base, df_cruce)

if __name__ == "__main__":
    start = time.time()
    main()
    end = time.time()
    tiempo_promedio = end - start
    print(f"Tiempo ejecución: {tiempo_promedio} segundos.")