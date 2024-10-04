from typing import List, Any

import ConectDBcorp
import EnviarEmail
from xml.dom import minidom

with open("V:\Informática\EmailsPython\SQLQuery\SQLQuery - Expedicao.sql", "r") as arquivo:
    SQLQuery = arquivo.read()

with open("V:\Informática\EmailsPython\emails_parametros.xml", "r", encoding="utf-8") as xmlFile:
    config = minidom.parse(xmlFile)

config_inTeste: int = 1
config_emailsPara: str  = ""
config_emailsTeste: str = ""
config_assunto: str = ""
config_tituloTabela1: str = ""

for configs in config.getElementsByTagName("EmailExpedicao"):
    config_inTeste= int(configs.getAttribute("inTeste"))
    config_emailsPara = configs.getAttribute("EmailsPara")
    config_emailsTeste = configs.getAttribute("EmailTeste")
    config_assunto = configs.getAttribute("Assunto")
    config_tituloTabela1 = configs.getAttribute("TituloTabela1")

cursor = ConectDBcorp.ConectaSQL()
cursor.execute(SQLQuery)

rows = cursor.fetchall()

email_head = """<html><html><body>"""

email_titulo_colunas = """<tr>
                            <th>NF</th>
                            <th>Emissão</th>
                            <th>Destinatário</th>
                            <th>Nº Coleta</th>
                            <th>Solicitação</th>
                            <th>Transportadora</th>
                            <th>Peso Bruto</th>
                            <th>Qtd Volumes</th>
                          </tr>"""

total_colunas_tabela1 = 8

email_body = ''

titulo_tabela1 = f"<tr><th colspan='{total_colunas_tabela1}'>{config_tituloTabela1}</th></tr>"
total1_tabela1 = 0
total2_tabela1 = 0
divisoes_tabela1 = []

for x in rows:
    email_body += (f"<tr>"
                   f"<td>{x.NFNum}</td>"
                   f"<td>{x.Emissao}</td>"
                   f"<td>{x.DestNFNomeRazaoSocial}</td>"
                   f"<td>{x.NFOrdColId}</td>"
                   f"<td>{x.Solicitacao}</td>"
                   f"<td>{x.Transportadora}</td>"
                   f"<td align ='right'>{float(x.NFPesoBruto):_.2f}</td>".replace('.', ',').replace('_', '.')+
                   f"<td align ='right'>{float(x.NFQtdVolumes):_.2f}</td>".replace('.', ',').replace('_', '.') +
                   f"</tr>")

    total1_tabela1 += float(x.NFPesoBruto)
    total2_tabela1 += float(x.NFQtdVolumes)
    divisoes_tabela1 = ConectDBcorp.arrayAdd(divisoes_tabela1, x.Transportadora, float(x.NFPesoBruto), float(x.NFQtdVolumes))

#Totais Transportadora
email_body += f"<tr><th colspan='{total_colunas_tabela1}'>Total por Transportadora</th></tr>"
divisoes_tabela1 = sorted(divisoes_tabela1, key=lambda x: x[0])
for x in range(len(divisoes_tabela1)):
    email_body += (f"<tr><td colspan='{total_colunas_tabela1-2}'>{divisoes_tabela1[x][0]}</td>"
                   f"<td align ='right'>{divisoes_tabela1[x][1]:_.2f}</td>"
                   f"<td align ='right'>{divisoes_tabela1[x][2]:_.2f}</td></tr>").replace('.', ',').replace('_', '.')

email_body += f"<tr><th colspan='{total_colunas_tabela1}'>Total Geral</th></tr>"
email_body += (f"<tr><td colspan='{total_colunas_tabela1-2}'>Total</td>"
               f"<td align ='right'>{total1_tabela1:_.2f}</td>"
               f"<td align ='right'>{total2_tabela1:_.2f}</td></tr>").replace('.', ',').replace('_', '.')


corpo_footer = "<body></html>"

corpo_email = (email_head +
               "<table border='1' style='color:black'>" + titulo_tabela1 + email_titulo_colunas + email_body + "</table>" +
               corpo_footer)

if config_inTeste == 1:
    emails_Para = config_emailsTeste
else:
    emails_Para = config_emailsPara

EnviarEmail.Enviar(config_assunto, corpo_email, emails_Para)