from typing import List, Any

import ConectDBcorp
import EnviarEmail
from xml.dom import minidom

with open("V:\Informática\EmailsPython\SQLQuery\SQLQuery - TitulosPagarReceberVencidos.sql", "r") as arquivo:
    SQLQuery = arquivo.read()

with open("V:\Informática\EmailsPython\emails_parametros.xml", "r", encoding="utf-8") as xmlFile:
    config = minidom.parse(xmlFile)

config_inTeste: int = 1
config_emailsPara: str  = ""
config_emailsTeste: str = ""
config_assunto: str = ""

for configs in config.getElementsByTagName("EmailContasReceber"):
    config_inTeste= int(configs.getAttribute("inTeste"))
    config_emailsPara = configs.getAttribute("EmailsPara")
    config_emailsTeste = configs.getAttribute("EmailTeste")
    config_assunto = configs.getAttribute("Assunto")
    config_tituloTabela1 = configs.getAttribute("TituloTabela1")
    config_tituloTabela2 = configs.getAttribute("TituloTabela2")

cursor = ConectDBcorp.ConectaSQL()
cursor.execute(SQLQuery)

rows = cursor.fetchall()

email_head = """<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>"""

email_titulo_colunas = """<tr>
                            <th>Divisão</th>
                            <th>Documento</th>
                            <th>Parcela</th>
                            <th>Vencimento</th>
                            <th>Pagamento</th>
                            <th>Dias Atraso</th>
                            <th>Terceiro</th>
                            <th>Valor</th>
                          </tr>"""

email_body = ''

titulo_tabela1 = f"<tr><th colspan='8'>{config_tituloTabela1}</th></tr>"
email_body_tabela1 = ''
total_tabela1 = 0
divisoes_tabela1 = []
grupos_tabela1 = []

titulo_tabela2 = f"<tr><th colspan='8'>{config_tituloTabela2}</th></tr>"
email_body_tabela2 = ''
total_tabela2 = 0
divisoes_tabela2 = []
grupos_tabela2 = []

for x in rows:
    email_body = (f"<tr>"
                  f"<td>{x.Division}</td>"
                  f"<td>{x.Documento}</td>"
                  f"<td>{x.Parcela}</td>"
                  f"<td>{x.Vencimento}</td>"
                  f"<td>{x.Pagamento}</td>"
                  f"<td align ='right'>{x.DiasPagto}</td>"
                  f"<td>{x.Terceiro}</td>"+
                  f"<td align ='right'>{float(x.ValTitulo):_.2f}</td>".replace('.',',').replace('_','.')+
                  f"</tr>")

    if x.Tipo == 'Receber' and x.Status == 'Vencido':
        total_tabela1 += float(x.ValTitulo)
        email_body_tabela1 += email_body
        divisoes_tabela1 = ConectDBcorp.arrayAdd(divisoes_tabela1, x.Division, float(x.ValTitulo))
        if x.Terceiro[:1] == 'G':
            grupos_tabela1 = ConectDBcorp.arrayAdd(grupos_tabela1, x.Terceiro, float(x.ValTitulo))
        else:
            grupos_tabela1 = ConectDBcorp.arrayAdd(grupos_tabela1, "Outros", float(x.ValTitulo))

    if x.Tipo == 'Receber' and x.Status == 'Liquidado':
        total_tabela2 += float(x.ValTitulo)
        email_body_tabela2 += email_body
        divisoes_tabela2 = ConectDBcorp.arrayAdd(divisoes_tabela2, x.Division, float(x.ValTitulo))
        if x.Terceiro[:1] == 'G':
            grupos_tabela2 = ConectDBcorp.arrayAdd(grupos_tabela2, x.Terceiro, float(x.ValTitulo))
        else:
            grupos_tabela2 = ConectDBcorp.arrayAdd(grupos_tabela2, "Outros", float(x.ValTitulo))

#Totais Inadimplencia
email_body_tabela1 += f"<tr><th colspan='8'>Total por Divisão</th></tr>"
divisoes_tabela1 = sorted(divisoes_tabela1, key=lambda x: x[0])
for x in range(len(divisoes_tabela1)):
    email_body_tabela1 += f"<tr><td colspan='7'>{divisoes_tabela1[x][0]}</td><td align ='right'>{divisoes_tabela1[x][1]:_.2f}</td></tr>".replace(
        '.', ',').replace('_', '.')

email_body_tabela1 += f"<tr><th colspan='8'>Total por Grupo Economico</th></tr>"
grupos_tabela1 = sorted(grupos_tabela1, key=lambda x: x[0])
for x in range(len(grupos_tabela1)):
    email_body_tabela1 += f"<tr><td colspan='7'>{grupos_tabela1[x][0]}</td><td align ='right'>{grupos_tabela1[x][1]:_.2f}</td></tr>".replace(
        '.', ',').replace('_', '.')

email_body_tabela1 += f"<tr><th colspan='8'>Total Geral</th></tr>"
email_body_tabela1 += f"<tr><td colspan='7'>Total</td><td align ='right'>{total_tabela1:_.2f}</td></tr>".replace('.', ',').replace('_', '.')


#Totais Atrasos
email_body_tabela2 += f"<tr><th colspan='8'>Total por Divisão</th></tr>"
divisoes_tabela2 = sorted(divisoes_tabela2, key=lambda x: x[0])
for x in range(len(divisoes_tabela2)):
    email_body_tabela2 += f"<tr><td colspan='7'>{divisoes_tabela2[x][0]}</td><td align ='right'>{divisoes_tabela2[x][1]:_.2f}</td></tr>".replace(
        '.', ',').replace('_', '.')

email_body_tabela2 += f"<tr><th colspan='8'>Total por Grupo Economico</th></tr>"
grupos_tabela2 = sorted(grupos_tabela2, key=lambda x: x[0])
for x in range(len(grupos_tabela2)):
    email_body_tabela2 += f"<tr><td colspan='7'>{grupos_tabela2[x][0]}</td><td align ='right'>{grupos_tabela2[x][1]:_.2f}</td></tr>".replace(
        '.', ',').replace('_', '.')

email_body_tabela2 += f"<tr><th colspan='8'>Total Geral</th></tr>"
email_body_tabela2 += f"<tr><td colspan='7'>Total</td><td align ='right'>{total_tabela2:_.2f}</td></tr>".replace('.', ',').replace('_', '.')


corpo_footer = "<body></html>"

corpo_email = (email_head +
               "<table border='1' style='color:red'>" + titulo_tabela1 + email_titulo_colunas + email_body_tabela1 + "</table><br/><br/>" +
               "<table border='1' style='color:blue'>" + titulo_tabela2 + email_titulo_colunas + email_body_tabela2 + "</table>" +
               corpo_footer)

if config_inTeste == 1:
    emails_Para = config_emailsTeste
else:
    emails_Para = config_emailsPara

EnviarEmail.Enviar(config_assunto, corpo_email, emails_Para)