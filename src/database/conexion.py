from src.utils.resource_path import resource_path
import sqlite3 as sql

def conexionSQLite():
    try:
        conexion = sql.connect(resource_path("./src/database/db.db"))
        return conexion
    except sql.Error as ex:
        error = "Error al conectar a la base de datos SQLite:" + str(ex)
        return error

def ejecutar_query(query):
    try:
        conexion = conexionSQLite()
        cursor = conexion.cursor()
        cursor.execute(query)
        resultados = cursor.fetchall()
        return resultados
    except sql.Error as ex:
        error = "Error al ejecutar la consulta:" + str(ex)
        return error
    finally:
        if cursor is not None:
            cursor.close()
        if conexion is not None:
            conexion.close()
