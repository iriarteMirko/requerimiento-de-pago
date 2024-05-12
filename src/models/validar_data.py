from ..utils.variables import analistas_validados


def validar_cuentas(cuentas_base, cuentas_cruce):
        cuentas_no_encontradas = []
        cuentas_copia = cuentas_base.copy()
        
        for cuenta in cuentas_copia:
            if cuenta not in cuentas_cruce:
                cuentas_base.remove(cuenta)
                cuentas_no_encontradas.append(cuenta)
        
        if len(cuentas_no_encontradas) == 0:
            mensaje = "Deudores OK\n"
        else:
            mensaje = f"Deudores no encontrados: {cuentas_no_encontradas}\n"
        
        return cuentas_base, mensaje
    
def validar_analistas(analistas):
    analistas_no_validados = []
    for analista in analistas:
        if analista not in analistas_validados:
            analistas_no_validados.append(analista)
    
    if len(analistas_no_validados) == 0:
        mensaje = "Analistas OK\n"
    else:
        mensaje = f"Analistas no validados: {analistas_no_validados}\n"
    
    return mensaje