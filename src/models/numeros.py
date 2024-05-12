from ..utils.variables import *


def separar_entero_decimal(num):
        num_str = str(num)
        if "." not in num_str:
            parte_entera = num_str
            parte_decimal = "00"
        else:
            parte_entera, parte_decimal = num_str.split(".")
        
        if len(parte_decimal) > 2:
            parte_decimal = parte_decimal[:2]
        elif len(parte_decimal) < 2:
            parte_decimal = parte_decimal.ljust(2, "0")
        
        return parte_entera, parte_decimal

def numero_entero_a_texto(num):
    if num == 0:
        return "cero"
    grupos = []
    while num > 0:
        grupos.append(num % 1000)
        num //= 1000
    textos = [convertir_grupo_a_texto(grupo) for grupo in grupos]
    if len(textos) > 1:
        textos[1] += " mil"
    if len(textos) > 2:
        textos[2] = "un millón" if textos[2] == "uno" else textos[2] + " millones"
    return " ".join(textos[::-1]).strip()

def convertir_grupo_a_texto(num):
    texto = ""
    if num >= 100:
        texto += centenas[num // 100]
        num %= 100
    if num >= 20:
        texto += " " + decenas[num // 10]
        num %= 10
    elif num >= 10:
        texto += " " + diez_a_diecinueve[num - 10]
        num = 0
    if num > 0:
        texto += " y " + unidades[num]
    return texto.strip()