from datetime import datetime
from .validar_data import validar_cuentas, validar_analistas
from .generar_dataframes import generar_dataframes
from .generar_doc import generar_doc
from .generar_excel import generar_excel
from .numeros import formato_numero
from ..routes.rutas import *
from ..utils.variables import meses, correos_analistas


hoy = datetime.today()
dia = hoy.strftime("%d")
mes = hoy.strftime("%m")
año = hoy.strftime("%Y")
nombre_mes = meses.get(mes)
fecha_hoy = f"{dia} de {nombre_mes} de {año}"

def generar_cartas(ruta_dacxa, ruta_dac_cdr, cuadro):
    dataframes = generar_dataframes(base, ruta_dacxa, ruta_dac_cdr)
    df_base = dataframes[0]
    df_cruce = dataframes[1]
    
    cuadro.insert("end", f"Registros: {df_base.shape[0]}\n")
    
    cuentas_base = df_base["Cuenta"].drop_duplicates().to_list()
    cuentas_cruce = df_cruce["Deudor"].to_list()
    analistas = df_cruce["ANALISTA_ACT"].drop_duplicates().to_list()
    
    cuadro.insert("end", f"Deudores: {cuentas_base}\n")
    cuentas, mensaje_cuentas = validar_cuentas(cuentas_base, cuentas_cruce)
    cuadro.insert("end", mensaje_cuentas)
    
    mensaje_analistas = validar_analistas(analistas)
    cuadro.insert("end", mensaje_analistas)
    
    for cuenta in cuentas:
        df_cuenta = df_base[df_base["Cuenta"] == cuenta]
        if (df_cuenta["Demora"] >= 0).all():
            generar_cartas_sin_deudaxvencer(cuenta, df_cuenta, df_cruce)
        else:
            generar_cartas_con_deudaxvencer(cuenta, df_cuenta, df_cruce)

def generar_cartas_sin_deudaxvencer(cuenta, df_cuenta, df_cruce):
    razon_social = df_cruce[df_cruce["Deudor"]==cuenta]["NOMBRE DAC"].iloc[0].upper()
    razon_social_2 = razon_social
    direccion_legal = df_cruce[df_cruce["Deudor"]==cuenta]["DIRECCIÓN LEGAL"].iloc[0]
    distrito = df_cruce[df_cruce["Deudor"]==cuenta]["DISTRITO"].iloc[0]
    provincia = df_cruce[df_cruce["Deudor"]==cuenta]["PROVINCIA"].iloc[0]
    departamento = df_cruce[df_cruce["Deudor"]==cuenta]["DPTO."].iloc[0]
    dias_demora = df_cuenta["Demora"].iloc[0]
    dias_demora_2 = dias_demora
    
    deuda_vencida = round(df_cuenta["Importe"].sum(),2)
    deuda_vencida_soles, deuda_vencida_texto = formato_numero(deuda_vencida)
    
    analista_mayuscula = df_cruce[df_cruce["Deudor"]==cuenta]["ANALISTA_ACT"].iloc[0]
    analista = " ".join([palabra.capitalize() for palabra in analista_mayuscula.lower().split(" ")])
    correo_analista = correos_analistas.get(analista)
    
    ruta_doc = resource_path("./results/"+razon_social+".docx")
    df_cuenta.to_excel(resource_path("./results/"+razon_social+".xlsx"), index=False) # Sin deudas por vencer
    
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
    
    generar_doc(modelo_2, replacements, ruta_doc)
    generar_excel(razon_social)

def generar_cartas_con_deudaxvencer(cuenta, df_cuenta, df_cruce):
    razon_social = df_cruce[df_cruce["Deudor"]==cuenta]["NOMBRE DAC"].iloc[0].upper()
    razon_social_2 = razon_social
    direccion_legal = df_cruce[df_cruce["Deudor"]==cuenta]["DIRECCIÓN LEGAL"].iloc[0]
    distrito = df_cruce[df_cruce["Deudor"]==cuenta]["DISTRITO"].iloc[0]
    provincia = df_cruce[df_cruce["Deudor"]==cuenta]["PROVINCIA"].iloc[0]
    departamento = df_cruce[df_cruce["Deudor"]==cuenta]["DPTO."].iloc[0]
    dias_demora = df_cuenta["Demora"].iloc[0]
    dias_demora_2 = dias_demora
    
    deuda_vencida = round(df_cuenta[df_cuenta["Demora"] >= 0]["Importe"].sum(),2)
    deuda_vencida_soles, deuda_vencida_texto = formato_numero(deuda_vencida)
    
    deuda_por_vencer = round(df_cuenta[df_cuenta["Demora"] < 0]["Importe"].sum(),2)
    deuda_por_vencer_soles, deuda_por_vencer_texto = formato_numero(deuda_por_vencer)
    
    analista_mayuscula = df_cruce[df_cruce["Deudor"]==cuenta]["ANALISTA_ACT"].iloc[0]
    analista = " ".join([palabra.capitalize() for palabra in analista_mayuscula.lower().split(" ")])
    correo_analista = correos_analistas.get(analista)
    
    ruta_doc = resource_path("./results/"+razon_social+".docx")
    df_cuenta.to_excel(resource_path("./results/"+razon_social+".xlsx"), index=False) # Con deudas por vencer
    
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
    
    generar_doc(modelo_1, replacements, ruta_doc)
    generar_excel(razon_social)