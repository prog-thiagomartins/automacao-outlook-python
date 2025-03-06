import win32com.client as win32

# Testando a conexão com o Outlook

outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Verificar títulos da caixa de entrada

inbox = namespace.GetDefaultFolder(6)  # 6 é o código da caixa de entrada
messages = inbox.Items

for message in messages:
    print(message.Subject)
    print(message.Body)

