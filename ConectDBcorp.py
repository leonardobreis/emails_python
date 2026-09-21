import pyodbc
from xml.dom import minidom

with open("V:\\Informática\\EmailsPython\\emails_parametros.xml", "r", encoding="utf-8") as xmlFile:
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

config_servidor_bi = ''
config_database_bi = ''
config_usuario_bi = ''
config_senha_bi = ''

for configs in config.getElementsByTagName("BancoDadosAirzapBI"):
    config_servidor_bi = configs.getAttribute("Servidor")
    config_database_bi = configs.getAttribute("Database")
    config_usuario_bi = configs.getAttribute("Usuario")
    config_senha_bi = configs.getAttribute("Senha")


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
        print("Conexão DBCORP Ok")
    except pyodbc.Error as e:
        print(f"Erro de conexão DBCORP: {e}")
        raise e

    return conn.cursor()


def ConectaAirzapBI():
    conn_str = (
        "DRIVER={SQL Server};"
        f"SERVER={config_servidor_bi};"
        f"DATABASE={config_database_bi};"
        f"UID={config_usuario_bi};"
        f"PWD={config_senha_bi};"
    )

    try:
        conn = pyodbc.connect(conn_str)
        print("Conexão AirzapMyBIGateway Ok")
        return conn  # Retorna o objeto de conexão completo (necessário para o commit/close do ETL)
    except pyodbc.Error as e:
        print(f"Erro de conexão AirzapMyBIGateway: {e}")
        raise e


def arrayAdd(array, search_valor, valor1Add, valor2Add=0):
    indice = None

    for x in range(len(array)):
        if array[x][0] == search_valor:
            indice = x

    if indice == None:
        array.append([search_valor, valor1Add, valor2Add])
    else:
        array[indice][1] += valor1Add
        array[indice][2] += valor2Add

    return array