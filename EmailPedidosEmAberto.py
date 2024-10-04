import ConectDBcorp
import EnviarEmail
from xml.dom import minidom
import sys
from datetime import datetime

with open("V:\Informática\EmailsPython\SQLQuery\SQLQuery - PedidosEmAberto.sql", "r") as arquivo:
    SQLQuery = arquivo.read()

with open("V:\Informática\EmailsPython\emails_parametros.xml", "r", encoding="utf-8") as xmlFile:
    config = minidom.parse(xmlFile)

parametros = sys.argv

config_inTeste: int = 1
config_inHomolog: int = 1
config_emailsPara: str  = ""
config_emailsPara_CD: str  = ""
config_emailsPara_AT: str  = ""
config_emailsPara_Thiago: str  = ""
config_emailsPara_Ricardo: str  = ""
config_emailsPara_Outros: str  = ""
config_emailsCc: str = ""
config_emailsTeste: str = ""
config_assunto: str = ""
parametro1: str = ""

for configs in config.getElementsByTagName("EmailPedidosEmAberto"):
    config_inTeste = int(configs.getAttribute("inTeste"))
    config_inHomolog = int(configs.getAttribute("inHomolog"))
    config_emailsPara_CD = configs.getAttribute("EmailsPara_CD")
    config_emailsPara_AT = configs.getAttribute("EmailsPara_AT")
    config_emailsPara_Thiago = configs.getAttribute("EmailsPara_Thiago")
    config_emailsPara_Ricardo = configs.getAttribute("EmailsPara_Ricardo")
    config_emailsPara_Outros = configs.getAttribute("EmailsPara_Outros")
    config_emailsCc = configs.getAttribute("EmailCc")
    config_emailsTeste = configs.getAttribute("EmailTeste")
    config_assunto = configs.getAttribute("Assunto")
    config_tituloTabela1 = configs.getAttribute("TituloTabela1")
    config_tituloTabela2 = configs.getAttribute("TituloTabela2")

if len(parametros) > 1:
    parametro1 = parametros[1]
else:
    parametro1 = 'AT'

if parametro1 == 'CD':
    config_emailsPara = config_emailsPara_CD
elif parametro1 == 'AT':
    config_emailsPara = config_emailsPara_AT
elif parametro1 == 'RICARDO':
    config_emailsPara = config_emailsPara_Ricardo
elif parametro1 == 'THIAGO':
    config_emailsPara = config_emailsPara_Thiago
elif parametro1 == 'AED':
    config_emailsPara = config_emailsPara_Outros
else:
    config_emailsPara = config_emailsTeste

config_assunto = f'{config_assunto}: {parametro1}'

SQLQuery = SQLQuery.replace('@Division', f"'{parametro1}'")

cursor = ConectDBcorp.ConectaSQL()
cursor.execute(SQLQuery)

rows = cursor.fetchall()

email_head = """<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>"""

email_titulo_colunas_tabela1 = """<tr>
                                    <th>Prev. Fatur.</th>
                                    <th>Nº Pedido</th>
                                    <th>Cliente</th>
                                    <th>Status Pedido</th>
                                    <th>Status Separação</th>
                                    <th>% Estoque*</th>                                    
                                    <th>Vendedor</th>
                                    <th>Pagamento</th>
                                    <th>Valor</th>
                                  </tr>"""
total_colunas_tabela1 = 9

email_titulo_colunas_tabela2 = """<tr>
                                    <th>Prev. Fatur.</th>
                                    <th>Nº Pedido</th>
                                    <th>Cliente</th>
                                    <th>Status Pedido</th>
                                    <th>Status Separação</th>
                                    <th>% Estoque*</th>
                                    <th>Vendedor</th>
                                    <th>Natureza</th>
                                    <th>Valor</th>
                                  </tr>"""
total_colunas_tabela2 = 9

titulo_tabela1 = f"<tr><th colspan='{total_colunas_tabela1}'>{config_tituloTabela1}: {parametro1}</th></tr>"
email_body_tabela1 = ''
total_tabela1 = 0
meses_tabela1 = []
grupos_tabela1 = []

titulo_tabela2 = f"<tr><th colspan='{total_colunas_tabela2}'>{config_tituloTabela2}: {parametro1}</th></tr>"
email_body_tabela2 = ''
total_tabela2 = 0
meses_tabela2 = []
grupos_tabela2 = []

for x in rows:
    if x.ReceitaBruta == 1: #Pedidos de Venda
        total_tabela1 += float(x.Valor)
        if x.DataPrevFaturamento.date() < datetime.now().date():
            email_body_tabela1 += f"<tr style='color:red'>"
        elif x.DataPrevFaturamento.date() == datetime.now().date():
            email_body_tabela1 += f"<tr style='color:blue'>"
        else:
            email_body_tabela1 += f"<tr style='color:black'>"

        email_body_tabela1 += (f"<td>{x.PrevFaturamento}</td>"
                               f"<td>{x.Pedido}</td>"
                               f"<td>{x.Cliente}</td>"
                               f"<td>{x.StatusPedido}</td>"
                               f"<td>{x.StatusSeparacao}</td>"
                               f"<td>{x.Estoque}</td>"
                               f"<td>{x.Representante}</td>"
                               f"<td>{x.Pagamento}</td>"
                               f"<td align ='right'>{float(x.Valor):_.2f}</td>".replace('.',',').replace('_','.')+
                               f"</tr>")
        meses_tabela1 = ConectDBcorp.arrayAdd(meses_tabela1, x.PrevFaturamento[-7:], float(x.Valor))
        if x.Cliente[:1] == 'G':
            grupos_tabela1 = ConectDBcorp.arrayAdd(grupos_tabela1, x.Cliente, float(x.Valor))
        else:
            grupos_tabela1 = ConectDBcorp.arrayAdd(grupos_tabela1, "Outros", float(x.Valor))

    if x.ReceitaBruta == 0: #Outras Naturezas
        total_tabela2 += float(x.Valor)
        if x.DataPrevFaturamento.date() < datetime.now().date():
            email_body_tabela2 += f"<tr style='color:red'>"
        elif x.DataPrevFaturamento.date() == datetime.now().date():
            email_body_tabela2 += f"<tr style='color:blue'>"
        else:
            email_body_tabela2 += f"<tr style='color:black'>"

        email_body_tabela2 += (f"<td>{x.PrevFaturamento}</td>"
                               f"<td>{x.Pedido}</td>"
                               f"<td>{x.Cliente}</td>"
                               f"<td>{x.StatusPedido}</td>"
                               f"<td>{x.StatusSeparacao}</td>"
                               f"<td>{x.Estoque}</td>"
                               f"<td>{x.Representante}</td>"
                               f"<td>{x.Natur_Descr}</td>"
                               f"<td align ='right'>{float(x.Valor):_.2f}</td>".replace('.',',').replace('_','.')+
                               f"</tr>")
        meses_tabela2 = ConectDBcorp.arrayAdd(meses_tabela2, x.PrevFaturamento[-7:], float(x.Valor))
        if x.Cliente[:1] == 'G':
            grupos_tabela2 = ConectDBcorp.arrayAdd(grupos_tabela2, x.Cliente, float(x.Valor))
        else:
            grupos_tabela2 = ConectDBcorp.arrayAdd(grupos_tabela2, "Outros", float(x.Valor))

#Totais Tabela1
email_body_tabela1 += f"<tr><th colspan='{total_colunas_tabela1}'>Total por Mês</th></tr>"
for x in range(len(meses_tabela1)):
    email_body_tabela1 += f"<tr><td colspan='{total_colunas_tabela1-1}'>{meses_tabela1[x][0]}</td><td align ='right'>{meses_tabela1[x][1]:_.2f}</td></tr>".replace('.', ',').replace('_', '.')

email_body_tabela1 += f"<tr><th colspan='{total_colunas_tabela1}'>Total por Grupo Economico</th></tr>"
grupos_tabela1 = sorted(grupos_tabela1, key=lambda x: x[0])
for x in range(len(grupos_tabela1)):
    email_body_tabela1 += f"<tr><td colspan='{total_colunas_tabela1-1}'>{grupos_tabela1[x][0]}</td><td align ='right'>{grupos_tabela1[x][1]:_.2f}</td></tr>".replace('.', ',').replace('_', '.')

email_body_tabela1 += f"<tr><th colspan='{total_colunas_tabela1}'>Total Geral</th></tr>"
email_body_tabela1 += f"<tr><td colspan='{total_colunas_tabela1-1}'>Total</td><td align ='right'>{total_tabela1:_.2f}</td></tr>".replace('.', ',').replace('_', '.')


#Totais Tabela2
email_body_tabela2 += f"<tr><th colspan='{total_colunas_tabela2}'>Total por Mês</th></tr>"
for x in range(len(meses_tabela2)):
    email_body_tabela2 += f"<tr><td colspan='{total_colunas_tabela2-1}'>{meses_tabela2[x][0]}</td><td align ='right'>{meses_tabela2[x][1]:_.2f}</td></tr>".replace('.', ',').replace('_', '.')

email_body_tabela2 += f"<tr><th colspan='{total_colunas_tabela2}'>Total por Grupo Economico</th></tr>"
grupos_tabela2 = sorted(grupos_tabela2, key=lambda x: x[0])
for x in range(len(grupos_tabela2)):
    email_body_tabela2 += f"<tr><td colspan='{total_colunas_tabela2-1}'>{grupos_tabela2[x][0]}</td><td align ='right'>{grupos_tabela2[x][1]:_.2f}</td></tr>".replace(
        '.', ',').replace('_', '.')

email_body_tabela2 += f"<tr><th colspan='{total_colunas_tabela2}'>Total Geral</th></tr>"
email_body_tabela2 += f"<tr><td colspan='{total_colunas_tabela2-1}'>Total</td><td align ='right'>{total_tabela2:_.2f}</td></tr>".replace('.', ',').replace('_', '.')


corpo_footer = ("<br/><br/>"
                "* % Estoque considera o Saldo a separar X Estoque atual, não considera o consumo do item por outros pedidos, OPs etc."
                "<body></html>")

corpo_email = email_head

if total_tabela1 > 0:
    corpo_email += "<table border='1' style='color:black'>" + titulo_tabela1 + email_titulo_colunas_tabela1 + email_body_tabela1 + "</table><br/><br/>"

if total_tabela2 > 0:
    corpo_email += "<table border='1' style='color:black'>" + titulo_tabela2 + email_titulo_colunas_tabela2 + email_body_tabela2 + "</table>"

corpo_email += corpo_footer

if config_inTeste == 1:
    emails_Para = config_emailsTeste
    emails_Cc = ''
else:
    emails_Para = config_emailsPara
    emails_Cc = config_emailsCc

if config_inHomolog == 1 and config_inTeste == 0:
    emails_BCc = config_emailsTeste
else:
    emails_BCc = ''

EnviarEmail.Enviar(config_assunto, corpo_email, emails_Para, emails_Cc, emails_BCc)