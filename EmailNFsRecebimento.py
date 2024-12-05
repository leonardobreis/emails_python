from typing import List, Any

import requests
from datetime import date, datetime, timedelta
import EnviarEmail
from collections import namedtuple
from xml.dom import minidom

with open("V:\Informática\EmailsPython\emails_parametros.xml", "r", encoding="utf-8") as xmlFile:
    config = minidom.parse(xmlFile)

config_inTeste: int = 1
config_emailsPara: str  = ""
config_emailsTeste: str = ""
config_assunto: str = ""
config_tituloTabela1: str = ""
config_tituloTabela2: str = ""
config_APItoken: str = ""
config_APICNPJ: str = ""
config_URLBase: str = ""

for configs in config.getElementsByTagName("EmailNFsRecebimento"):
    config_inTeste= int(configs.getAttribute("inTeste"))
    config_inHomolog = int(configs.getAttribute("inHomolog"))
    config_emailsPara = configs.getAttribute("EmailsPara")
    config_emailsTeste = configs.getAttribute("EmailTeste")
    config_emailsCc = configs.getAttribute("EmailCc")
    config_assunto = configs.getAttribute("Assunto")
    config_tituloTabela1 = configs.getAttribute("TituloTabela1")

for configs in config.getElementsByTagName("APIBuscaNFE"):
    config_APItoken = configs.getAttribute("APItoken")
    config_APICNPJ = configs.getAttribute("APICNPJ")
    config_URLBase = configs.getAttribute("URLBase")

dt_ini = '2024-01-01'
hoje = date.today()
dt_fin = hoje.strftime('%Y-%m-%d')

url = f"{config_URLBase}?cnpj={config_APICNPJ}&dt_ini={dt_ini}&dt_fin={dt_fin}&token={config_APItoken}"

response = requests.get(url)
data = response.json()
columns = data[0].keys()
Row = namedtuple('Row', columns)

rows = [Row(**item) for item in data]

email_head = """<html><html><body>"""

email_titulo_colunas = """<tr>
                            <th>NF</th>
                            <th>Fornecedor</th>
                            <th>CNPJ</th>
                            <th>Emissão</th>
                            <th>Recebimento Físico</th>
                            <th>Status NFe</th>
                            <th>Valor</th>
                          </tr>"""

total_colunas_tabela1 = 7

email_body = ''

titulo_tabela1 = f"<tr><th colspan='{total_colunas_tabela1}'>{config_tituloTabela1}</th></tr>"
email_body_tabela1 = ''
total_tabela1 = 0
quantidade_atrasado1 = 0
quantidade_normal1 = 0

for x in rows:
    if x.descEvento != 'Ciencia da Operacao':
        continue

    dataEmi = datetime.strptime(x.dhEmi, "%Y-%m-%d")

    if x.recibo_dt == 'NULL':
        recebimento = 'Não Recebido'
    else:
        dataRec = datetime.strptime(x.recibo_dt, "%Y-%m-%d")
        recebimento = f"{str(dataRec.strftime('%d/%m/%Y'))} - {x.recibo_user}"

    cnpj = x.CNPJ
    cnpj_formatado = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

    if dataEmi.date() < (datetime.now().date() - timedelta(days=7)):
        email_body = f"<tr style='color:red'>"
        quantidade_atrasado1 += 1
    else:
        email_body = f"<tr style='color:black'>"
        quantidade_normal1 += 1

    email_body += (f"<td>{x.nNF}</td>"
                  f"<td>{x.xNome}</td>"+
                  f"<td>{cnpj_formatado}</td>"+
                  f"<td>{dataEmi.strftime('%d/%m/%Y')}</td>"+
                  f"<td>{recebimento}</td>"+
                  f"<td>{x.descEvento}</td>"+
                  f"<td align ='right'>{float(x.vNF):_.2f}</td>".replace('.',',').replace('_','.')+
                  f"</tr>")

    total_tabela1 += float(x.vNF)
    email_body_tabela1 += email_body

email_body_tabela1 += f"<tr><th colspan='{total_colunas_tabela1}'>Total Geral</th></tr>"
email_body_tabela1 += (f"<tr><td colspan='{total_colunas_tabela1-3}'>Total</td>"
                       f"<td align ='left'>Qtd: {quantidade_normal1:_.0f}</td>"
                       f"<td align ='left' style='color:red'>Qtd: {quantidade_atrasado1:_.0f}</td>"
                       f"<td align ='right'>{total_tabela1:_.2f}</td></tr>").replace('.', ',').replace('_', '.')


corpo_footer = ("<br/>"
                "<table border='1' style='color:black'>"
                "<tr><th style='color:black'>Legenda:</th></tr>"
                "<tr><td style='color:red'>NF emitida a mais de 7 dias</td></tr>"
                "<tr><td style='color:black'>NF emitida até 7 dias</td></tr>"
                "</table><body></html>")

corpo_email = (email_head +
               "<table border='1' style='color:black'>" + titulo_tabela1 + email_titulo_colunas + email_body_tabela1 + "</table><br/><br/>" +
               corpo_footer)

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