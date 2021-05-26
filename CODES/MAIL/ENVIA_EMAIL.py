import win32com.client as win32


def send_email(path_anexo,mail_to,mail_cc,subject, tipo_envio,*protocolo_ou_boleto):
    
    # CRIA INSTÂNCIA OUTLOOK
    outlook = win32.Dispatch('outlook.application')
    
    # CRIA EMAIL
    mail = outlook.CreateItem(0)
    mail.To = mail_to
    mail.CC = mail_cc
    mail.Subject = subject
    mail.Attachments.Add(path_anexo)

    assinatura_jr = '\n\nAtenciosamente,\nJR & Marto\n(11) 3977-1300\n(11) 93235-4915'

    # DEFINE ASSUNTO E CORPO DO E-MAIL
    if tipo_envio in [1,5]:
        if tipo_envio == 1:
            if 'RAZÃO SOCIAL 3' in subject:
                mail.Body = 'Prezados,\n\nSegue fatura fechada com clientes ativos.\n\nAguardo fatura e boleto com vencimento para dia 15.\n\nEMITIR: COMISSÃO: 10% p/ JR MARTO; 40% p/ NITEROI SELF STORAGE SPE LTDA (pró-labore)' + assinatura_jr
            else:
                mail.Body = 'Prezados,\n\nSegue fatura fechada com clientes ativos.\n\nAguardo fatura e boleto com vencimento para dia 15.' + assinatura_jr
        else:
            mail.Body = 'Prezados,\n\nSegue base de clientes em previsão para o próximo fechamento.' + assinatura_jr

    elif tipo_envio in [2,99]:
        mail.Body = 'Prezados,\n\nSegue base de clientes em previsão. Atualizada até o último protocolo recebido.' + assinatura_jr
        if None not in protocolo_ou_boleto:
            mail.Attachments.Add(protocolo_ou_boleto[0])

    elif tipo_envio in [3,4]:
        mail.Body = 'Prezados,\n\nSegue anexo fechamento do seguro + Boleto.' + assinatura_jr
        mail.Attachments.Add(protocolo_ou_boleto[0])

    # ENVIA E-MAIL
    mail.Display(True)
    # mail.Send()