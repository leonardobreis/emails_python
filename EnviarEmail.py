import smtplib
import email.message
from xml.dom import minidom

def Enviar(_Assunto, _Corpo_Email, _Emails_Para, _Emails_CC = '', _Emails_CCO = ''):

    with open("V:\Informática\EmailsPython\emails_parametros.xml", "r", encoding="utf-8") as xmlFile:
        config = minidom.parse(xmlFile)

    config_login: str = ''
    config_senha: str = ''

    for configs in config.getElementsByTagName("ConfigEmail"):
        config_login = configs.getAttribute("Usuario")
        config_senha = configs.getAttribute("Senha")

    msg = email.message.Message()
    msg['Subject'] = _Assunto
    msg['From'] = f"Airzap - Anest Iwata<{config_login}>"

    _Emails_Para = _Emails_Para.split(',')
    msg['To'] = ','.join(_Emails_Para)

    if _Emails_CC != '':
        _Emails_CC = _Emails_CC.split(',')
        msg['Cc'] = ','.join(_Emails_CC)

    if _Emails_CCO != '':
        _Emails_CCO = _Emails_CCO.split(',')
        msg['Bcc'] = ','.join(_Emails_CCO)

    login = config_login
    password = config_senha

    # Lista de destinatários (To + Cc + Bcc)
    destinatarios = _Emails_Para.split(',') if isinstance(_Emails_Para, str) else _Emails_Para
    if _Emails_CC != '':
        destinatarios += _Emails_CC.split(',') if isinstance(_Emails_CC, str) else _Emails_CC
    if _Emails_CCO != '':
        destinatarios += _Emails_CCO.split(',') if isinstance(_Emails_CCO, str) else _Emails_CCO

    msg.add_header('Content-Type', 'text/html')
    msg.set_payload( _Corpo_Email, charset='utf-8')

    send = smtplib.SMTP('smtp.gmail.com: 587')

    print(f"destinatarios: {destinatarios}")
    print(f"msg['To']: {msg['To']}")
    print(f"msg['Cc']: {msg['Cc']}")
    print(f"msg['Bcc']: {msg['Bcc']}")

    try:
        send.starttls()
        send.login(login, password)
    except smtplib.SMTPAuthenticationError as e:
        print(f"Erro de envio: {e}")

    try:
        send.sendmail(msg['From'], destinatarios, msg.as_string().encode('utf-8'))
    except smtplib.SMTPDataError as e:
        print(f"Erro de envio: {e}")

    print('Email enviado')