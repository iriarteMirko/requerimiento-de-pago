import pandas as pd


def generar_dataframes(base, ruta_dacxa, ruta_dac_cdr):
        df_dac_cdr = pd.read_excel(ruta_dac_cdr, sheet_name=" CONTRATOS DAC-DACES")
        df_dacxanalista = pd.read_excel(ruta_dacxa, sheet_name="Base_NUEVA")
        # BASE #
        df_base = pd.read_excel(base, sheet_name="BASE")
        columnas_deseadas_base = ["Cuenta", "Nº documento", "Referencia", "Fecha de documento", "Clase de documento", "Demora tras vencimiento neto", "Moneda del documento", "Importe en moneda local"]
        nuevas_columnas_base = ["Fecha de doc.", "CL", "Demora", "Moneda", "Importe"]
        df_base = df_base[columnas_deseadas_base]
        df_base = df_base[df_base["Cuenta"].notna()]
        nombres_columnas = dict(zip(columnas_deseadas_base[3:], nuevas_columnas_base))
        df_base = df_base.rename(columns=nombres_columnas)
        df_base["Cuenta"] = df_base["Cuenta"].astype("Int64").astype("str")
        df_base["Nº documento"] = df_base["Nº documento"].astype("Int64").astype("str")
        df_base["Demora"] = df_base["Demora"].astype("Int64")
        df_base.sort_values(by=["Cuenta","Demora"], ascending=[True, False], inplace=True)
        df_base = df_base.reset_index(drop=True)
        # CRUCE #
        df_cruce = pd.merge(df_dac_cdr, df_dacxanalista, left_on="Deudor", right_on="DEUDOR", how="left")
        df_cruce.drop(columns=["DEUDOR"], inplace=True)
        df_cruce = df_cruce[df_cruce["ANALISTA_ACT"].notna()]
        columnas_deseadas_cruce = ["Deudor", "NOMBRE DAC", "DIRECCIÓN LEGAL", "DISTRITO", "PROVINCIA", "DPTO.", "ANALISTA_ACT"]
        analistas_no_deseados = ["REGION NORTE", "REGION SUR", "SIN INFORMACION"]
        df_cruce = df_cruce[columnas_deseadas_cruce]
        df_cruce = df_cruce[df_cruce["Deudor"].notna()]
        df_cruce["Deudor"] = df_cruce["Deudor"].astype("Int64").astype("str")
        df_cruce = df_cruce.loc[~df_cruce["ANALISTA_ACT"].isin(analistas_no_deseados)]
        df_cruce = df_cruce.reset_index(drop=True)
        
        return df_base, df_cruce