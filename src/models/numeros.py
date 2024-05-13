from ..utils.variables import unidades, decenas, centenas, veinte_a_veintinueve, diez_a_diecinueve


def separar_entero_decimal(num):
    num_str = str(num)
    if "." not in num_str:
        entero = num_str
        decimal = "00"
    else:
        entero, decimal = num_str.split(".")
    
    if len(decimal) > 2:
        decimal = decimal[:2]
    elif len(decimal) < 2:
        decimal = decimal.ljust(2, "0")
    
    return entero, decimal

def convertir_grupo_a_texto(num):
    texto = ""
    if len(str(num)) == 1:
        texto = unidades[num]
    else:
        if num > 100:
            texto += centenas[num // 100]
            num %= 100
        
        if num >= 30:
            texto += " " + decenas[num // 10]
            num %= 10
        elif num >=20:
            texto += " " + veinte_a_veintinueve[num - 20]
            num = 0
        elif num >= 10:
            texto += " " + diez_a_diecinueve[num - 10]
            num = 0
        
        if num >= 1 and " " not in texto:
            texto += " " + unidades[num]
        elif num >= 1:
            texto += " y " + unidades[num]
        else:
            if num > 1:
                texto += " y " + unidades[num]
            elif num == 1:
                texto += " un"
            elif num == 0:
                pass
    return texto.strip()

def numero_entero_a_texto(num):
    if num == 0:
        return "cero"
    elif num == 100:
        return "cien"
    elif num == 1000:
        return "mil"
    grupos = []
    while num > 0:
        grupos.append(num % 1000)
        num //= 1000
    textos = [convertir_grupo_a_texto(grupo) for grupo in grupos]
    
    if len(textos) > 1:
        if textos[1] == "":
            pass
        elif textos[1] == "uno":
            textos[1] = "mil"
        else:
            textos[1] += " mil"
        if "uno mil" in textos[1]:
            textos[1] = textos[1].replace("uno mil", "un mil")
    if len(textos) > 2:
        textos[2] = "un millón" if textos[2] == "uno" else textos[2] + " millones"
    return " ".join(textos[::-1]).strip().replace("  ", " ")

def formato_miles(num):
    num_str = str(num)
    if len(num_str) <= 3:
        return num_str
    miles = num_str[:-3]
    cientos = num_str[-3:]
    miles_con_comas = ""
    for i, digito in enumerate(miles[::-1]):
        miles_con_comas += digito
        if (i + 1) % 3 == 0 and (i + 1) != len(miles):
            miles_con_comas += ","
    return miles_con_comas[::-1] + cientos

def formato_numero(num):
    entero, decimal = separar_entero_decimal(num)
    deuda_soles = f"S/ {formato_miles(entero)}.{decimal}"
    
    entero_texto = numero_entero_a_texto(int(entero))
    deuda_texto = f"({entero_texto} con {decimal}/100 soles)"
    
    return deuda_soles, deuda_texto