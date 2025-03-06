import win32com.client as win32
import os

caminho_arquivos_auxiliares = os.path.abspath("arquivos_auxiliares")
caminho_assinatura = os.path.join(caminho_arquivos_auxiliares, "assinatura.html")
caminho_logo = os.path.join(caminho_arquivos_auxiliares, "logo.png")
message = "Este é O CORPO DE EMAIL com Python"
# Abrindo a assinatura
with open(caminho_assinatura, "r", encoding="utf-8") as file:
    assinatura = file.read()

# Testando a conexão com o Outlook 
outlook = win32.Dispatch("Outlook.Application")

#Criar uma mensagem simples
mail = outlook.CreateItem(0) # 0 é o código para mensagem normal
mail.Subject = "Testando envio de email - TÍTULO"
mail.HTMLBody = message + "<br><br>" + assinatura
mail.To = "pessoal.thiagomartins@gmail.com"

anexo = mail.Attachments.Add(caminho_logo) # Adicionar um anexo
anexo.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo_empresa")
# Enviar a mensagem
mail.Send()
print("Mensagem enviada com sucesso!")
