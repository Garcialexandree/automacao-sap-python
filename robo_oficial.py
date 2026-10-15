import gspread
from oauth2client.service_account import ServiceAccountCredentials
import win32com.client
import time
import os
import json
import pandas as pd

# ==========================================
# CONFIGURAÇÕES E FUNÇÕES AUXILIARES
# ==========================================
NOME_SISTEMA_SAP = "ECC R/3 - 02 - Produção" 

def atualizar_historico_json(ordem, mat, desc, qtd, und, status, func):
    arquivo = "historico_almoxarifado.json"
    historico = []
    if os.path.exists(arquivo):
        with open(arquivo, "r", encoding="utf-8") as f:
            try: historico = json.load(f)
            except: pass
            
    historico.insert(0, {
        "ordem": ordem, "material": mat, "descricao": desc,
        "qtd": qtd, "unidade": und, "funcionario": func, "status": status
    })
    
    with open(arquivo, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, indent=4)

def capturar_erro_barra_sap(session):
    try:
        tipo = session.findById("wnd[0]/sbar").messageType
        texto = session.findById("wnd[0]/sbar").text
        if tipo == "E" or tipo == "A": 
            return texto
    except:
        pass
    return None

def conectar_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except:
        print("-> SAP não encontrado. A iniciar o SAP Logon automaticamente...")
        os.system('start saplogon')
        time.sleep(12) 
        SapGuiAuto = win32com.client.GetObject("SAPGUI")

    application = SapGuiAuto.GetScriptingEngine
    if application.Children.Count == 0:
        print(f"-> A abrir a ligação ao ambiente: {NOME_SISTEMA_SAP}...")
        connection = application.OpenConnection(NOME_SISTEMA_SAP, True)
        time.sleep(5) 
    else:
        connection = application.Children(0)
        
    return connection.Children(0)

# ==========================================
# 0. CARREGAR BASE DE MATERIAIS
# ==========================================
print("A ler base de materiais do Excel...")
dict_materiais = {}
try:
    df_base = pd.read_excel("base_materiais.xlsx", dtype=str)
    col_mat = [c for c in df_base.columns if "Material" in str(c)][0]
    col_umb = [c for c in df_base.columns if "UMB" in str(c) or "Unid" in str(c) or "UM" in str(c)][0]
    dict_materiais = dict(zip(df_base[col_mat].str.strip(), df_base[col_umb].str.strip()))
    print(f"✅ {len(dict_materiais)} materiais memorizados com sucesso!\n")
except Exception as e:
    print(f"⚠️ Aviso: Não consegui ler o 'base_materiais.xlsx'. (Erro: {e})\n")

# ==========================================
# 1. LIGAÇÃO AO GOOGLE SHEETS
# ==========================================
print("A ligar ao Google Sheets...")
try:
    escopos = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credenciais = ServiceAccountCredentials.from_json_keyfile_name("credenciais.json", escopos)
    cliente = gspread.authorize(credenciais)
    arquivo_google = cliente.open("Requisição de Material - Respostas")
    
    planilha_requisicoes = arquivo_google.worksheet("Respostas ao formulário 1")
    planilha_inventario = arquivo_google.worksheet("Inventário")
    print("SHEETS OK")
except Exception as erro:
    print(f"❌ ERRO CRÍTICO NO GOOGLE SHEETS: {erro}")
    input("Pressione ENTER para fechar...")
    exit()

# ==========================================
# 2. LIGAÇÃO INICIAL AO SAP
# ==========================================
print("A ligar ao SAP pela primeira vez...")
try:
    session = conectar_sap()
    print("SAP OK!\n")
except Exception as e:
    print(f"❌ Não foi possível ligar ao SAP inicialmente. (Erro: {e})")

print("O ROBO LIGOU OK")
print("🤖 MODO CONTÍNUO ATIVADO! Pressione Ctrl+C para desligar o robô.")

# ==========================================
# 3. LOOP INFINITO
# ==========================================
try:
    while True:
        try:
            try:
                _ = session.info.User
            except:
                print("\n⚠️ Ligação ao SAP perdida (Fechado). A tentar religar...")
                session = conectar_sap()
                print("✅ Ligação ao SAP recuperada com sucesso!")

            # -------------------------------------------------------------
            # PARTE 1: PROCESSAR REQUISIÇÕES DE MATERIAL (MIGO/IW32)
            # -------------------------------------------------------------
            print("\nA procurar novos pedidos no formulário...")
            linhas_req = planilha_requisicoes.get_all_values() 
            
            for indice, linha in enumerate(linhas_req[1:], start=2): 
                while len(linha) < 10: linha.append("")
                status_robo = linha[9]
                
                if status_robo == "":
                    nome_funcionario = linha[1] 
                    triagem = linha[2]          
                    
                    if "Sim, vou retirar" in triagem:
                        numero_ordem = linha[3]
                        print(f"-> A executar MIGO para Ordem Completa: {numero_ordem}")
                        session.findById("wnd[0]/tbar[0]/okcd").text = "/nmigo"
                        session.findById("wnd[0]").sendVKey(0)
                        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").text = numero_ordem
                        session.findById("wnd[0]").sendVKey(0)
                        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-MTSNR").text = numero_ordem

                        linha_migo = 0
                        while True:
                            try:
                                session.findById(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,{linha_migo}]").selected = True
                                session.findById(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-LGOBE[6,{linha_migo}]").text = "BA04"
                                linha_migo += 1
                            except: break

                        session.findById("wnd[0]/tbar[1]/btn[7]").press() 
                        time.sleep(1)
                        erro_verificacao = capturar_erro_barra_sap(session)
                        if erro_verificacao:
                            planilha_requisicoes.update_cell(indice, 10, f"ERRO: {erro_verificacao}")
                            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                            session.findById("wnd[0]").sendVKey(0)
                            continue 

                        session.findById("wnd[0]/tbar[1]/btn[23]").press() 
                        time.sleep(1)
                        erro_gravacao = capturar_erro_barra_sap(session)
                        if erro_gravacao:
                            planilha_requisicoes.update_cell(indice, 10, f"ERRO: {erro_gravacao}")
                            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                            session.findById("wnd[0]").sendVKey(0)
                            continue
                        
                        planilha_requisicoes.update_cell(indice, 10, "LANÇADO")
                        atualizar_historico_json(numero_ordem, "-", "Baixa Direta na MIGO", "-", "-", "Baixado", nome_funcionario)
                        print("✅ Concluído!\n")
                        
                    elif "Não, vou pedir peças avulsas" in triagem:
                        numero_ordem = linha[4] 
                        codigo_sap = str(linha[5]).strip()   
                        quantidade = linha[7]   
                        unidade_mecanico = str(linha[8]).strip()
                        unidade_final = unidade_mecanico
                        
                        if codigo_sap in dict_materiais:
                            if unidade_mecanico.upper() != dict_materiais[codigo_sap].upper():
                                unidade_final = dict_materiais[codigo_sap]
                        
                        print(f"-> A adicionar Peça {codigo_sap} na Ordem {numero_ordem} (IW32)...")
                        session.findById("wnd[0]/tbar[0]/okcd").text = "/niw32"
                        session.findById("wnd[0]").sendVKey(0)
                        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = numero_ordem
                        session.findById("wnd[0]").sendVKey(0)
                        
                        erro_abrir_ordem = capturar_erro_barra_sap(session)
                        if erro_abrir_ordem:
                            planilha_requisicoes.update_cell(indice, 10, f"ERRO: {erro_abrir_ordem}")
                            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                            session.findById("wnd[0]").sendVKey(0)
                            continue

                        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpMUEB").select()
                        tabela_base = "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpMUEB/ssubSUB_AUFTRAG:SAPLCOMK:3020/tblSAPLCOMKTCTRL_3020"
                        
                        linha_sap = 0
                        inserido_com_sucesso = False
                        desc_sap_texto = "Aguardando SAP"
                        
                        while not inserido_com_sucesso and linha_sap <= 30:
                            try:
                                session.findById(f"{tabela_base}/ctxtRESBD-MATNR[1,{linha_sap}]").text = codigo_sap
                                session.findById(f"{tabela_base}/txtRESBD-MENGE[4,{linha_sap}]").text = quantidade
                                session.findById(f"{tabela_base}/ctxtRESBD-EINHEIT[5,{linha_sap}]").text = unidade_final 
                                session.findById(f"{tabela_base}/ctxtRESBD-POSTP[6,{linha_sap}]").text = "L"
                                session.findById(f"{tabela_base}/ctxtRESBD-WERKS[9,{linha_sap}]").text = "BA01"
                                session.findById(f"{tabela_base}/txtRESBD-VORNR[10,{linha_sap}]").text = "0010"
                                session.findById("wnd[0]").sendVKey(0) 
                                
                                try: desc_sap_texto = session.findById(f"{tabela_base}/txtRESBD-MAKTX[2,{linha_sap}]").text
                                except: 
                                    try: desc_sap_texto = session.findById(f"{tabela_base}/txtRESBD-MAKTX[3,{linha_sap}]").text
                                    except: pass
                                inserido_com_sucesso = True 
                            except: linha_sap += 1
                                
                        session.findById("wnd[0]/tbar[0]/btn[11]").press() 
                        time.sleep(1)
                        erro_iw32 = capturar_erro_barra_sap(session)
                        if erro_iw32:
                            planilha_requisicoes.update_cell(indice, 10, f"ERRO: {erro_iw32}")
                            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                            session.findById("wnd[0]").sendVKey(0)
                            continue

                        print("-> A executar baixa na MIGO...")
                        session.findById("wnd[0]/tbar[0]/okcd").text = "/nmigo"
                        session.findById("wnd[0]").sendVKey(0)
                        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").text = numero_ordem
                        session.findById("wnd[0]").sendVKey(0)
                        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-MTSNR").text = numero_ordem

                        linha_migo = 0
                        while True:
                            try:
                                session.findById(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,{linha_migo}]").selected = True
                                session.findById(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-LGOBE[6,{linha_migo}]").text = "BA04"
                                linha_migo += 1
                            except: break

                        session.findById("wnd[0]/tbar[1]/btn[7]").press() 
                        time.sleep(1)
                        erro_verificacao2 = capturar_erro_barra_sap(session)
                        if erro_verificacao2:
                            planilha_requisicoes.update_cell(indice, 10, f"ERRO: {erro_verificacao2}")
                            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                            session.findById("wnd[0]").sendVKey(0)
                            continue

                        session.findById("wnd[0]/tbar[1]/btn[23]").press() 
                        time.sleep(1)
                        erro_gravacao2 = capturar_erro_barra_sap(session)
                        if erro_gravacao2:
                            planilha_requisicoes.update_cell(indice, 10, f"ERRO: {erro_gravacao2}")
                            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                            session.findById("wnd[0]").sendVKey(0)
                            continue
                        
                        planilha_requisicoes.update_cell(indice, 7, desc_sap_texto.strip()) 
                        planilha_requisicoes.update_cell(indice, 10, "LANÇADO")            
                        atualizar_historico_json(numero_ordem, codigo_sap, desc_sap_texto.strip(), quantidade, unidade_final, "Baixado", nome_funcionario)
                        print(f"✅ Peça {codigo_sap} processada com sucesso!\n")
            
            # -------------------------------------------------------------
            # PARTE 2: PROCESSAR INVENTÁRIO (ZWM063) - MODO SIMPLIFICADO
            # -------------------------------------------------------------
            print("\nA verificar aba Inventário...")
            linhas_inv = planilha_inventario.get_all_values()
            
            iniciou_tabela = False 
            
            for indice_inv, linha_inv in enumerate(linhas_inv[1:], start=2):
                while len(linha_inv) < 3: linha_inv.append("")
                
                material_inv = str(linha_inv[0]).strip()
                qtd_fisica = str(linha_inv[1]).strip()
                status_inv = str(linha_inv[2]).strip()
                
                # Ignora linhas totalmente vazias
                if material_inv == "":
                    continue
                
                # Se ainda não chegou à tabela, procura a palavra "Material"
                if not iniciou_tabela:
                    if material_inv.upper() == "MATERIAL":
                        iniciou_tabela = True 
                    continue 
                
                # DAQUI PARA BAIXO, ELE SÓ LÊ OS MATERIAIS REAIS
                if qtd_fisica != "" and status_inv == "":
                    print(f"-> A lançar Inventário: Material {material_inv} - Qtd: {qtd_fisica}")
                    
                    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzwm063"
                    session.findById("wnd[0]").sendVKey(0)
                    time.sleep(1) 
                    
                    try:
                        session.findById("wnd[0]/usr/btnOPCAO4").press()
                        session.findById("wnd[0]/usr/btnOPCAO1").press()
                    except: pass

                    time.sleep(1) # Dá 1 segundo para a janelinha amarela aparecer
                    
                    # --- NOVO: BLOCO PARA FECHAR A JANELINHA DO "B03" ÀS CEGAS ---
                    try:
                        inseriu_b03 = False
                        for id_campo in ["ctxt[0]", "ctxt[1]", "ctxt[2]", "txt[0]", "txt[1]"]:
                            try:
                                session.findById(f"wnd[1]/usr/{id_campo}").text = "B03"
                                inseriu_b03 = True
                                break 
                            except:
                                pass
                        
                        try:
                            session.findById("wnd[1]/tbar[0]/btn[0]").press() 
                        except:
                            session.findById("wnd[1]").sendVKey(0) 
                            
                        time.sleep(1) 
                    except:
                        pass 
                    # -------------------------------------------------------------

                    # PASSO A: Busca do Material
                    mat_procurado = material_inv.lstrip('0') 
                    achou_mat = False
                    
                    for scroll in range(30): 
                        for linha_mat in range(0, 30, 2): 
                            try:
                                mat_na_tela = session.findById(f"wnd[0]/usr/sub:ZWMO_015:0005/ctxtW_TELA-MATNR[{linha_mat},5]").text
                                mat_limpo = str(mat_na_tela).strip().lstrip('0')
                                
                                if mat_procurado in mat_limpo or mat_limpo in mat_procurado:
                                    session.findById(f"wnd[0]/usr/sub:ZWMO_015:0005/chkW_TELA-OPCAO[{linha_mat},19]").selected = True
                                    achou_mat = True
                                    time.sleep(1)
                                    break
                            except: 
                                pass 
                            
                        if achou_mat:
                            break 
                            
                        try:
                            pos_atual = session.findById("wnd[0]/usr").verticalScrollbar.position
                            session.findById("wnd[0]/usr").verticalScrollbar.position = pos_atual + 3
                            time.sleep(0.5) 
                        except:
                            break

                    if not achou_mat:
                        print(f"❌ Material {material_inv} não encontrado para contagem.")
                        planilha_inventario.update_cell(indice_inv, 3, "ERRO: Não encontrado") 
                        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                        session.findById("wnd[0]").sendVKey(0)
                        continue

                    # PASSO B: Avançar a tela do Local
                    try:
                        session.findById("wnd[0]/usr/sub:ZWMO_015:0006/chkW_TELA-OPCAO[0,20]").selected = True
                        time.sleep(1)
                    except:
                        pass 

                    # PASSO C: Preencher quantidade
                    try:
                        session.findById("wnd[0]/usr/txtTXT_QTD").text = qtd_fisica
                        session.findById("wnd[0]").sendVKey(0) 
                        session.findById("wnd[0]").sendVKey(0) 
                    except: pass
                    
                    planilha_inventario.update_cell(indice_inv, 3, "LANÇADO") 
                    print(f"✅ Quantidade {qtd_fisica} lançada com sucesso!\n")

            print("Ronda concluída. A fazer pausa de 2 minutos...")
            time.sleep(120) 
            
        except Exception as e:
            print(f"Aviso: Houve uma falha de comunicação com o Google/SAP. (Erro: {e})")
            time.sleep(10)

except KeyboardInterrupt:
    print("\nDESLIGAMENTO DO ROBO OK")