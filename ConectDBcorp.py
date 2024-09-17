import pyodbc
from xml.dom import minidom

with open("V:\Informática\EmailsPython\emails_parametros.xml", "r", encoding="utf-8") as xmlFile:
    config = minidom.parse(xmlFile)

config_servidor = ''
config_database = ''
config_usuario = ''
config_senha = ''

for configs in config.getElementsByTagName("BancoDados"):
    config_servidor = configs.getAttribute("Servidor")
    config_database = configs.getAttribute("Database")
    config_usuario = configs.getAttribute("Usuario")
    config_senha = configs.getAttribute("Senha")

def ConectaSQL():
    conn_str = (
        "DRIVER={SQL Server};"
        f"SERVER={config_servidor};"
        f"DATABASE={config_database};"
        f"UID={config_usuario};"
        f"PWD={config_senha};"
    )

    try:
        conn = pyodbc.connect(conn_str)
        print("Conexão Ok")
    except pyodbc.Error as e:
        print(f"Erro de conexão: {e}")

    return conn.cursor()

def arrayAdd(array, search_valor, valorAdd):
    indice = None

    for x in range(len(array)):
        if array[x][0] == search_valor:
            indice = x

    if indice == None:
        array.append([search_valor, valorAdd])
    else:
        array[indice][1] += valorAdd

    return array