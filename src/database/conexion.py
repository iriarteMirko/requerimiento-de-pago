from src.utils.resource_path import resource_path
import sqlite3 as sql

def conexionSQLite():
    try:
        conexion = sql.connect(resource_path("../database/db.db"))
        return conexion
    except sql.Error as ex:
        error = "Error al conectar a la base de datos SQLite:" + str(ex)
        return error

def ejecutar_query(query):
    conexion = conexionSQLite()
    try:
        cursor = conexion.cursor()
        cursor.execute(query)
        resultados = cursor.fetchall()
        return resultados
    except sql.Error as ex:
        error = "Error al ejecutar la consulta:" + str(ex)
        return error
    finally:
        cursor.close()
        conexion.close