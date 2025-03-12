import win32com.client as win32
import os

# Caminhos dos arquivos auxiliares
caminho_arquivos_auxiliares = os.path.abspath("arquivos_auxiliares")
caminho_relatorio = os.path.join(caminho_arquivos_auxiliares, "guia_envio_emails.docx")
caminho_guia = os.path.join(caminho_arquivos_auxiliares, "relatorio_envio_emails.xlsx")
caminho_video = os.path.join(caminho_arquivos_auxiliares, "video.mkv")

lista_anexos = [caminho_relatorio, caminho_guia]
limite_envio_anexo = 25 * 1024 * 1024


# calcular tamanho total dos anexos
def calcular_tamanho_total_anexos(anexos):
    tamanho_total = 0
    for anexo in anexos:
        tamanho_total += os.path.getsize(anexo)
    return tamanho_total

tamanho_total_anexos = calcular_tamanho_total_anexos(lista_anexos)

outlook = win32.Dispatch("Outlook.Application")

if  tamanho_total_anexos > limite_envio_anexo:
    print("O tamanho total dos anexos excede o limite de envio do Outlook")
else:
    # Criar uma mensagem simples
    mail = outlook.CreateItem(0)  # 0 é o código para mensagem normal
    mail.Subject = "Testando envio de email - TÍTULO"
    mail.Body = "Este é O CORPO DE EMAIL com Python"
    mail.To = "pessoal.thiagomartins@gmail.com"
    for anexo in lista_anexos:
        mail.Attachments.Add(anexo)  # Adicionar os anexos na mensagem
    # Enviar a mensagem
    mail.Send()
    print("Mensagem enviada com sucesso!")