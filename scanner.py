import win32com.client

print("A conectar ao SAP...")
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
session = application.Children(0).Children(0)

print("A escanear os nomes secretos das caixinhas...\n")
painel = session.findById("wnd[0]/usr/sub:ZWMO_015:0005")

for i in range(painel.Children.Count):
    campo = painel.Children(i)
    texto = campo.Text.strip()
    
    # Se o campo tiver algum texto, ele mostra o ID
    if texto != "":
        nome_campo = campo.ID.split("/")[-1]
        print(f"Texto na tela: {texto} ---> NOME SECRETO: {nome_campo}")
        
print("\n--- FIM DO MAPEAMENTO ---")