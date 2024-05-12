def validar_cuentas(cuentas_base, cuentas_cruce):
        print(f"Deudores: {cuentas_base}\n")
        cuentas_no_encontradas = []
        cuentas_copia = cuentas_base.copy()
        
        for cuenta in cuentas_copia:
            if cuenta not in cuentas_cruce:
                cuentas_base.remove(cuenta)
                cuentas_no_encontradas.append(cuenta)
        
        if len(cuentas_no_encontradas) == 0:
            print("Deudores OK.\n")
        else:
            print(f"Deudores no encontrados: {cuentas_no_encontradas}\n")
        
        return cuentas_base
    
def validar_analistas(analistas, analistas_validados):
    analistas_no_validados = []
    for analista in analistas:
        if analista not in analistas_validados:
            analistas_no_validados.append(analista)
    
    if len(analistas_no_validados) == 0:
        print("Analistas OK\n")
    else:
        print("Analistas no validados: ",analistas_no_validados,"\n")