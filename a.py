import os
import pandas as pd
import pyautogui as auto
import subprocess
import sys
import win32com.client
import time
import datetime

data = datetime.date.today()
data_formatada = data.strftime('%d.%m.%Y')
print(data_formatada)

exportado = []
erro = []

def get_files(dir=None,return_dir = False):
    list = []
    for arqiuivo in os.listdir(dir):
        if arqiuivo.endswith(".xlsx"):
            dir_arquivo = f"{arqiuivo}"
            if return_dir == True:
                list.append(str(dir+"/"+dir_arquivo))
            else:
                list.append(str(dir_arquivo))
    return (list)


pasta_login = r"C:\Users\igor.gabriel\OneDrive - JSL SA\Área de Trabalho\AUTOMAÇÃO EXPORT SAP\LOGIN"
lista_arquivos_login = os.listdir(pasta_login)

for arquivo in lista_arquivos_login:
    arquivo_encontrado = os.path.join(pasta_login, arquivo)
    login = pd.read_excel(arquivo_encontrado, sheet_name='Planilha1')
    user = login.at[0, 'login']
    senha = login.at[0, 'senha']
    print(f"{user} {senha}") 

def saplogin(): # Função de Login
    global session
    try:
        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe" # Define o caminho para o SAP
        subprocess.Popen(path) # Inicia o SAP
        time.sleep(2) # Espera 5 segundos para que o SAP GUI seja carregado
        SapGuiAuto = win32com.client.GetObject('SAPGUI') # Obtém a a\utomação do SAP GUI
        if not type(SapGuiAuto) == win32com.client.CDispatch: # Verifica se o objeto obtido é do tipo esperado
            return
        application = SapGuiAuto.GetScriptingEngine # Obtém a aplicação SAP GUI
        if not type(application) == win32com.client.CDispatch: # Verifica se o objeto obtido é do tipo esperado        
            SapGuiAuto = None
            return

        connection = application.OpenConnection("# JSL -  ECC - Produção (ECP)", True) # Conexão Qualidade(TESTES)

        if not type(connection) == win32com.client.CDispatch: # Verifica se o objeto obtido é do tipo esperado
            application = None
            SapGuiAuto = None
            return
        session = connection.Children(0) # Obtém a sessão ativa
        if not type(session) == win32com.client.CDispatch: # Verifica se o objeto obtido é do tipo esperado
            connection = None
            application = None
            SapGuiAuto = None
            return
        # Preenche os campos de nome de usuário e senha no SAP GUI
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "30120919" # Preenche a matricula no campo de login
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Abbeyroad123!" # Preenche a senha
        auto.hotkey('enter') # Pressiona enter para completar o login
    except:
        print(sys.exc_info()[0]) # Em caso de exceção, aparece mensagem de erro

    finally:
        # Limpa os objetos e recursos
        connection = None
        application = None
        SapGuiAuto = None
        # Chama a função saplogin para executar o login no SAP GUI

saplogin()
print("ABRINDO O SAP...")

def get_connection(ambiente = "# JSL -  ECC - Produção (ECP)"): # Função que define a conexão no SAP
    SapGuiAuto = win32com.cliente.GetObject('SAPGUI') 
    application = SapGuiAuto.GetScriptingEngine 
    connection = application.OpenConnection(ambiente, True)
    return connection

def get_session(connection):
    session = connection.Children(0)
    return session

def pesquisa():
    migo = session.findById("wnd[0]/tbar[0]/okcd")
    migo.text = "MIGO"
    auto.hotkey('enter')
pesquisa()

lista_arquivos = get_files(dir=r"FILES", return_dir= True)
for file in lista_arquivos:
    print(file)
    try:
        dt = pd.read_excel(file, sheet_name= 'Planilha1')
        for index, linha in dt.iterrows():
            codigo = str(linha['codigo'])
            item = int(linha['item'])
            numero = int(linha['numero'])
            print(f"CODIGO: {codigo}")
            print(f"ITEM: {item}")
            print(f"NUMERO: {numero}")

            time.sleep(3)
            # MUDA PARA ESTORNO
            try:
                auto.hotkey('tab')
                auto.hotkey("Ctrl","a")
                auto.hotkey('delete')
                time.sleep(2)
                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-MAT_DOC").text = f"{codigo}"
                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-MAT_DOC").caretPosition = 0
                auto.hotkey('enter')
                time.sleep(1)

                try:
                    msg = session.findById("wnd[0]/sbar").text
                    print(msg)
                except Exception as e:
                    print(f"CARAIO: {e}")
                
                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").setFocus
                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").key = "A03"

            except Exception as e:
                
                if msg == f"Já foram estornados todos os itens do documento {codigo}":
                    try:
                        print("DIALOGO NOVO")
                        time.sleep(3)
                        auto.hotkey('F5')
                        time.sleep(2)
                        auto.hotkey('tab')
                        auto.hotkey('enter')
                        msg = ""
                        time.sleep(3)
                        index += 5
                        print(codigo)
                        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-MAT_DOC").text = f"{codigo}"
                        auto.hotkey('enter')
                    except Exception as e:
                         print(f"NAO PREENCHEU e")

                elif msg == f"Documento {codigo} não existe no ano 2024" :
                    try:
                        print("ENTRO NO NAO EXISTE NO ANO")
                        auto.hotkey('tab')
                        auto.hotkey("Ctrl","a")
                        auto.hotkey('delete')
                        time.sleep(3)
                        index -= 1
                        msg = ""
                        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-MAT_DOC").text = f"{codigo}"
                        #session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2010/txtGODYNPRO-MAT_DOC").caretPosition = 0
                        auto.hotkey('enter')         
                    except Exception as e:
                        print(e)           
                else:
                     pass
                print(f"PREENCHIMENTO CODIGO {e}")
                line = True
                quanta_rolada = 0
                cont_linha = 0
                i = 0
                while line == True:
                    if cont_linha == 12:
                        cont_linha = 0
                        quanta_rolada += 12
                        rolada = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM").verticalScrollbar.position = quanta_rolada
                        linha = session.findById(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-EBELP[29,{cont_linha}]").text
                        linha = int(linha)
                        i+=1
                    else:
                        linha = session.findById(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-EBELP[29,{cont_linha}]").text
                        linha = int(linha)
                        i+=1
                    cont_linha+=1
                    print(linha)
                    print(i)

                    if linha == item:
                        print("ACHO")
                        seleciona = session.findById(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,{cont_linha-1}]").selected = True
                        line = False
                    else:
                        pass

                time.sleep(2)
                data = auto.locateCenterOnScreen('PRINT\data.png', confidence=0.9)
                auto.click(x=180, y=261)
                auto.hotkey('ctrl','a')
                auto.hotkey('delete')
                auto.typewrite(data_formatada)
                time.sleep(3)
                auto.hotkey('ctrl','s')

                msg = session.findById("wnd[0]/sbar").text
                print(f"\nMENSAGEM: {msg}\n")

                if msg == "O item foi marcado com OK": 
                    time.sleep(3)
                    auto.hotkey('F5')
                    time.sleep(1)
                    auto.hotkey('tab')
                    auto.hotkey('enter')
                    erro.append(codigo)
                else:
                    exportado.append(codigo)
                    pass
            else:
                pass
    except Exception as e:
        print(f"ERRO ARQUIVO: {file} | {e}")


dt_exportado = pd.DataFrame(exportado)
dt_exportado.to_excel(r"C:\Users\igor.gabriel\OneDrive - JSL SA\Área de Trabalho\AUTOMAÇÃO EXPORT SAP\EXPORTADO.xlsx", index = False)
dt_erro = pd.DataFrame(erro)
dt_erro.to_excel(r"C:\Users\igor.gabriel\OneDrive - JSL SA\Área de Trabalho\AUTOMAÇÃO EXPORT SAP\ERRO.xlsx", index = False)
    