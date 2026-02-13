import ConectDBcorp
import EnviarEmail
from xml.dom import minidom
import sys
from datetime import datetime, timedelta

with open("V:\Informática\EmailsPython\SQLQuery\SQLQuery - FaltaComprar.sql", "r") as arquivo:
    SQLQuery = arquivo.read()

with open("V:\Informática\EmailsPython\emails_parametros.xml", "r", encoding="utf-8") as xmlFile:
    config = minidom.parse(xmlFile)

parametros = sys.argv

config_inTeste: int = 1
config_inHomolog: int = 1
config_emailsPara: str  = ""
config_emailsCc: str = ""
config_emailsTeste: str = ""
config_assunto: str = ""

for configs in config.getElementsByTagName("EmailFaltaComprar"):
    config_inTeste = int(configs.getAttribute("inTeste"))
    config_inHomolog = int(configs.getAttribute("inHomolog"))
    config_emailsPara = configs.getAttribute("EmailsPara")
    config_emailsCc = configs.getAttribute("EmailCc")
    config_emailsTeste = configs.getAttribute("EmailTeste")
    config_assunto = configs.getAttribute("Assunto")
    config_tituloTabela1 = configs.getAttribute("TituloTabela1")

config_assunto = f'{config_assunto}'

cursor = ConectDBcorp.ConectaSQL()
cursor.execute(SQLQuery)

rows = cursor.fetchall()

email_head = """<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>"""

email_titulo_colunas_tabela1 = """<tr>
                                    <th>OP Pai</th>
                                    <th>Status</th>
                                    <th colspan=2>Item</th>
                                    <th>Grupo Estoque</th>
                                    <th>Qtd</th>
                                    <th>Falta</th>
                                  </tr>"""
total_colunas_tabela1 = 7

titulo_tabela1 = f"<tr><th colspan='{total_colunas_tabela1}'>{config_tituloTabela1}</th></tr>"
email_body_tabela1 = ''
total_tabela1 = 0
op_pai = 0

for x in rows:
    if op_pai == 0 or op_pai != x.OrdProdPaiId:
        email_body_tabela1 += (f"<tr style='color:black'>"
                               f"<td>{x.OrdProdPaiId}</td>"
                               f"<td colspan=6>{x.ItemEmpresaIdOPPai} - {x.ItemDescrOPPai}</td>"
                               f"</tr>")

    if x.Status == 'Falta Comprar':
        email_body_tabela1 += f"<tr style='color:red'>"
    else:
        email_body_tabela1 += f"<tr style='color:blue'>"

    email_body_tabela1 += (f"<td align ='right'> -></td>"
                           f"<td>{x.Status}</td>"
                           f"<td>{x.ItemEmpresaId}</td>"
                           f"<td>{x.ItemDescr}</td>"
                           f"<td>{x.GpEstDescr}</td>"
                           f"<td align ='right'>{float(x.QtdNecessaria):_.2f}</td>".replace('.',',').replace('_','.')+
                           f"<td align ='right'>{float(x.Falta):_.2f}</td>".replace('.', ',').replace('_', '.')+
                           f"</tr>")
    op_pai = x.OrdProdPaiId
    total_tabela1 += 1

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