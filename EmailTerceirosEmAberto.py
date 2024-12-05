import ConectDBcorp
import EnviarEmail
from xml.dom import minidom
import sys
from datetime import datetime, timedelta

with open("V:\Informática\EmailsPython\SQLQuery\SQLQuery - TerceirosEmAberto.sql", "r") as arquivo:
    SQLQuery = arquivo.read()

with open("V:\Informática\EmailsPython\emails_parametros.xml", "r", encoding="utf-8") as xmlFile:
    config = minidom.parse(xmlFile)

parametros = sys.argv

config_inTeste: int = 1
config_inHomolog: int = 1
config_emailsPara: str  = ""
config_emailsPara_CD: str  = ""
config_emailsPara_AED: str  = ""
config_emailsPara_COMPRAS: str  = ""
config_emailsCc: str = ""
config_emailsTeste: str = ""
config_assunto: str = ""
parametro1: str = ""

for configs in config.getElementsByTagName("EmailTerceirosEmAberto"):
    config_inTeste = int(configs.getAttribute("inTeste"))
    config_inHomolog = int(configs.getAttribute("inHomolog"))
    config_emailsPara_CD = configs.getAttribute("EmailsPara_CD")
    config_emailsPara_AED = configs.getAttribute("EmailsPara_AED")
    config_emailsPara_COMPRAS = configs.getAttribute("EmailsPara_COMPRAS")
    config_emailsCc = configs.getAttribute("EmailCc")
    config_emailsTeste = configs.getAttribute("EmailTeste")
    config_assunto = configs.getAttribute("Assunto")
    config_tituloTabela1 = configs.getAttribute("TituloTabela1")

if len(parametros) > 1:
    parametro1 = parametros[1]
else:
    parametro1 = 'CD'

if parametro1 == 'CD':
    config_emailsPara = config_emailsPara_CD
elif parametro1 == 'AED':
    config_emailsPara = config_emailsPara_AED
elif parametro1 == 'COMPRAS':
    config_emailsPara = config_emailsPara_COMPRAS
else:
    config_emailsPara = config_emailsTeste

config_assunto = f'{config_assunto}: {parametro1}'

SQLQuery = SQLQuery.replace('@Division', f"'{parametro1}'")

cursor = ConectDBcorp.ConectaSQL()
cursor.execute(SQLQuery)

rows = cursor.fetchall()

email_head = """<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>"""

email_titulo_colunas_tabela1 = """<tr>
                                    <th>NF</th>
                                    <th>Emissão</th>
                                    <th>Destinatário</th>
                                    <th>Natureza</th>
                                    <th>Itens</th>
                                    <th>Quantidade</th>
                                  </tr>"""
total_colunas_tabela1 = 6

titulo_tabela1 = f"<tr><th colspan='{total_colunas_tabela1}'>{config_tituloTabela1}: {parametro1}</th></tr>"
email_body_tabela1 = ''
total_tabela1 = 0

for x in rows:
    total_tabela1 += float(x.SaldoConsiderar)
    if x.DataEmissao.date() < (datetime.now().date() - timedelta(days=60)):
        email_body_tabela1 += f"<tr style='color:red'>"
    elif x.DataEmissao.date() < (datetime.now().date() - timedelta(days=30)):
        email_body_tabela1 += f"<tr style='color:blue'>"
    else:
        email_body_tabela1 += f"<tr style='color:black'>"

    email_body_tabela1 += (f"<td>{x.NF}</td>"
                           f"<td>{x.Emissao}</td>"
                           f"<td>{x.DestinatárioGr}</td>"
                           f"<td>{x.Natureza}</td>"
                           f"<td>{x.Itens}</td>"
                           f"<td align ='right'>{float(x.SaldoConsiderar):_.2f}</td>".replace('.',',').replace('_','.')+
                           f"</tr>")

email_body_tabela1 += f"<tr><td colspan='{total_colunas_tabela1-1}'>Total</td><td align ='right'>{total_tabela1:_.2f}</td></tr>".replace('.', ',').replace('_', '.')

corpo_footer = ("<br/>"
                "<table border='1' style='color:black'>"
                "<tr><th style='color:black'>Legenda:</th></tr>"
                "<tr><td style='color:red'>NF emitida a mais que 60 dias</td></tr>"
                "<tr><td style='color:blue'>NF emitida de 30 a 60 dias</td></tr>"
                "<tr><td style='color:black'>NF emitida até 30 dias</td></tr>"
                "</table><body></html>")

corpo_email = email_head

if total_tabela1 > 0:
    corpo_email += "<table border='1' style='color:black'>" + titulo_tabela1 + email_titulo_colunas_tabela1 + email_body_tabela1 + "</table><br/><br/>"

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