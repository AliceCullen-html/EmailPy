
# Esse caso funciona se o seu OUTLOOK tem várias contas ou apenas uma =D

# Primeiro use o comando

# Primeiro import o win32com.client
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
conta = ''


for item in outlook.Session.Accounts:
    # coloque o seu email no lugar remetente@email.com
    if item.SmtpAddress == 'remetente@email.com':
        conta = item
        break


mail = outlook.CreateItem(0)


if conta:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, conta))


mail.To = 'test@hotmail.com'
mail.Subject = 'Email vindo do Outlook'
#mail.CC = 'email@gmail.com'
#mail.BCC = 'email@gmail.com'

# se preferir pode mandar um email mais simples usando o mail.Body


mail.HTMLBody = '<h1> Corpo do Email em HTML </h1>'
'<p> Lorem ipsum dolor sit amet, consectetur adipiscing elit. Fusce lacinia, enim ut egestas pulvinar, lorem nisi consectetur leo, eget euismod ante leo ut dui. Nunc at tempor erat, a imperdiet nisi. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Nunc ex nisl, congue a ante sit amet, sagittis euismod purus. Morbi quam ligula, iaculis sed tincidunt eget, vehicula eget erat. Proin neque felis, aliquet sed odio id, eleifend efficitur velit. Nulla id lacus dictum neque bibendum ullamcorper ut eu tortor. Donec mollis risus nunc, eu scelerisque enim suscipit a. </p>'


# Anexos (pode colocar quantos quiser):
attachment = r'C:\Users\vulco\Downloads\Cvnew1.pdf'
mail.Attachments.Add(attachment)


mail.Send()
