import smtplib
import email.message

def Enviar(_Assunto, _Corpo_Email, _Emails_Para):
    msg = email.message.Message()
    msg['Subject'] = _Assunto
    msg['From'] = "Airzap - Anest Iwata<contato@airzap.com.br>"

    _Emails_Para = _Emails_Para.split(',')
    msg['To'] = ','.join(_Emails_Para)

    login = "contato@airzap-anestiwata.com.br"
    password = "f50I+ht8"

    msg.add_header('Content-Type', 'text/html')
    msg.set_payload( _Corpo_Email, charset='utf-8')

    send = smtplib.SMTP('smtp.gmail.com: 587')

    try:
        send.starttls()
        send.login(login, password)
    except smtplib.SMTPAuthenticationError as e:
        print(f"Erro de envio: {e}")

    try:
        send.sendmail(msg['From'], _Emails_Para, msg.as_string().encode('utf-8'))
    except smtplib.SMTPDataError as e:
        print(f"Erro de envio: {e}")

    print('Email enviado')