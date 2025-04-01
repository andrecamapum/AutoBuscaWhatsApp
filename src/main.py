# Autor: André Camapum Carvalho de Freitas
# Data: 03/03/2025
# Versão: 1.0
# Descrição: Script para automatizar a busca de listas de termos no WhatsApp Desktop.
#            O script é dividido em seções para facilitar a compreensão e a manutenção.
#            O script é compatível com Python 3.7+.
#            O script requer a instalação de bibliotecas adicionais, como pyautogui, pytesseract, opencv-python, pynput e pyperclip.
#            O script foi projetado para ser executado em um ambiente macOS. Funcionalidades em Windows com construção prevista.
#            O script foi projetado para ser executado em um ambiente de desenvolvimento local.


# ------------------------------------------------
# Modo de Depuração via JSON com Relatórios de Log
# ------------------------------------------------

# Importação de Bibliotecas
import json
import os
from datetime import datetime

# Criando o dicionário com todas as variáveis de depuração de bloco iniciando como True
registros_dep = {
    "Reg_Ver_WhatsApp_Exec": True,
    "Reg_Abrir_WhatsApp": True,
    "Reg_Reabrir_WhatsApp": True,
    "Reg_Id_Maior_Monitor": True,
    "Reg_Posicionar_WhatsApp": True,
    "Reg_Maximizar_WhatsApp": True,
    "Reg_Pesquisa_Global_WhatsApp": True,
    "Reg_Abrir_Img_Ilustrativa": True,
    "Reg_Mesa_Virtual_Frente": True,
    "Reg_Capturar_Painel_Pesquisa": True,
    "Reg_Pesquisa_Termo_Inicial": True,
    "Reg_Calibrar_Rolagem": True,
    "Reg_Capturas_OCR_Tratamento": True,
    "Reg_Conclusao_Programa": True
}

# Configurações de Depuração via JSON

debug_mode = True

if not debug_mode:
    print("🔹 Modo de depuração desativado.")
    for key in registros_dep:
        registros_dep[key] = False  # Define todas as chaves como False
else:
    print("🔹 Modo de depuração ativado.")
    Reg_Ver_WhatsApp_Exec = True
    Reg_Abrir_WhatsApp = True
    Reg_Reabrir_WhatsApp = True
    Reg_Id_Maior_Monitor = True
    Reg_Posicionar_WhatsApp = True
    Reg_Maximizar_WhatsApp = True
    Reg_Pesquisa_Global_WhatsApp = True
    Reg_Abrir_Img_Ilustrativa = True
    Reg_Mesa_Virtual_Frente = True
    Reg_Capturar_Painel_Pesquisa = True
    Reg_Pesquisa_Termo_Inicial = True
    Reg_Calibrar_Rolagem = True
    Reg_Capturas_OCR_Tratamento = True
    Reg_Conclusao_Programa = True          

log_file = "debug_log.json"

def salvar_log(ativado, **kwargs):
    """
    Salva um log no arquivo JSON. 
    Registra logs no JSON apenas se 'ativado' for True.
    Aceita múltiplas variáveis como argumentos nomeados.
    """

    if not ativado:
        return  # Se o log estiver desativado, sai sem fazer nada

    log_entry = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        **kwargs  # Adiciona todas as variáveis passadas à função
    }
    
    # Lê o arquivo JSON atual para não sobrescrever os logs antigos
    if os.path.exists(log_file):
        with open(log_file, "r") as file:
            try:
                logs = json.load(file)
                if not isinstance(logs, list):
                    logs = []
            except json.JSONDecodeError:
                logs = []
    else:
        logs = []

    # Adiciona o novo log e salva o arquivo atualizado
    logs.append(log_entry)
    with open(log_file, "w") as file:
        json.dump(logs, file, indent=4)
        file.write(",\n")  # Mantém formato JSON de lista

    print("🔹 Log salvo:", log_entry)  # Apenas para depuração

# ------------------------------------------------
# Gera uma caixa de informações
# ------------------------------------------------

import sys
import platform
import subprocess
import tkinter as tk
from tkinter import messagebox

def show_message(title, message):
    # macOS: Diálogo nativo com título via osascript
    if platform.system() == 'Darwin':
        script = f'display dialog "{message}" with title "{title}" buttons {{"OK"}}'
        subprocess.run(['osascript', '-e', script])
    # Windows/Linux: Tkinter padrão
    else:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(title, message)
        root.destroy()

def show_3option_message(title, message):
    # macOS: Diálogo nativo com três botões via osascript
    if platform.system() == 'Darwin':
        script = f'''
        set response to button returned of (display dialog "{message}" with title "{title}" buttons {{"Prosseguir", "Retornar", "Sair"}} default button "Prosseguir")
        return response
        '''
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        return result.stdout.strip()  # Retorna a opção escolhida

    # Windows/Linux: Tkinter padrão
    else:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()
        response = messagebox.askquestion(title, message)  # Retorna 'yes' ou 'no'
        root.destroy()
        return response

# Encerra o script se o sistema for Windows
if platform.system() == "Windows":
    show_message("Busca por Listas no Whatsapp", "\U0001F6A7 Este sistema ainda está em desenvolvimento para o Windows!")
    sys.exit()  # Encerra o script imediatamente

# Exemplo de uso:
show_message("Busca por Listas no Whatsapp", "\u26A0\uFE0F Para sua maior segurança salve seus arquivos abertos!")

# ---------------------------------------------------------------
# Verifica se o WhatsApp Desktop está em execução (MacOs apenas)
# ---------------------------------------------------------------

def is_whatsapp_running():
    """
    Verifica se o WhatsApp Desktop está em execução no macOS.
    Retorna True se estiver aberto, False caso contrário.
    """
    try:
        # Comando para listar aplicativos em execução via AppleScript
        result = subprocess.check_output([
            'osascript',
            '-e', 'tell application "System Events"',
            '-e', 'get name of every application process',
            '-e', 'end tell'
        ]).decode('utf-8').strip()

        # Verifica se "WhatsApp" está na lista de processos
        return "WhatsApp" in result.split(', ')
    
    except subprocess.CalledProcessError:
        return False
        
# -----------------------------------------------------------------
# Chama a função para abrir o WhatsApp (MacOs apenas)
# -----------------------------------------------------------------

def open_whatsapp_if_needed():
    if not is_whatsapp_running():
        salvar_log(Reg_Abrir_WhatsApp, evento="WhatsApp não estava em execução. Tentando abrir...")
        print("❌ WhatsApp Desktop não está em execução. Abrindo agora...")
        subprocess.run(["open", "-a", "WhatsApp"])
        salvar_log(Reg_Abrir_WhatsApp, evento="Comando para abrir o WhatsApp foi executado.")

# -----------------------------------------------------------------
# Código para fechar e reabrir o WhatsApp se necessário (MacOs apenas)
# -----------------------------------------------------------------

import time

Delay = 2  # 2 segundos de atraso

def close_whatsapp():
    """Fecha o WhatsApp caso esteja aberto."""
    if is_whatsapp_running():
        salvar_log(Reg_Reabrir_WhatsApp, evento="WhatsApp já estava aberto.")
        print("✅ WhatsApp Desktop está aberto!")
        show_message("Busca por Listas no Whatsapp", f"\u26A0\uFE0F Aguarde!\n\nO WhatsApp será reiniciado automaticamente após {Delay}s!")
        os.system('osascript -e \'quit app "WhatsApp"\'')
        time.sleep(Delay)  # Aguarda um tempo para garantir que fechou

def restart_whatsapp():
    """Reinicia o WhatsApp e aguarda ele abrir completamente."""
    close_whatsapp()
    time.sleep(Delay)  # Aguarda um tempo para garantir que fechou  
    open_whatsapp_if_needed()
        
    # Aguarda o WhatsApp abrir completamente antes de tentar mover a janela
    for _ in range(5*Delay):  # Tentativas (10s no total)
        if is_whatsapp_running():
            print("✅ WhatsApp detectado! Aguarde estabilização...")
            time.sleep(Delay)  # Pequena espera adicional
            break
        time.sleep(Delay)
    else:
        print("❌ WhatsApp não abriu corretamente.")
        return  # Sai da função se o WhatsApp não abriu

    time.sleep(Delay+1)  # Aguarde mais um pouco para garantir que a janela seja detectável

# Agora, sempre que quiser garantir que o WhatsApp está no estado inicial:
restart_whatsapp()

time.sleep(Delay)  # Aguarda um tempo para garantir que abriu

# -----------------------------------------------------------------------------------------------------
# Identifica o maior monitor disponível (MacOs apenas)
# -----------------------------------------------------------------------------------------------------

import AppKit
from AppKit import NSScreen

screens = NSScreen.screens() # Lista de monitores disponíveis

print(screens)

def get_largest_screen():
    salvar_log(Reg_Id_Maior_Monitor, evento="Iniciando busca pela maior tela", funcao_executada="get_largest_screen", qtd_telas=len(screens))
    if not screens:
        salvar_log(Reg_Id_Maior_Monitor, evento="Nenhuma tela detectada", resultado=None)
        print("Nenhuma tela detectada.")
        return None

    largest_screen = max(screens, key=lambda s: s.frame().size.width * s.frame().size.height)

    salvar_log(Reg_Id_Maior_Monitor, evento="Tela detectada", width=largest_screen.frame().size.width, height=largest_screen.frame().size.height, screen_id=id(largest_screen))

    return largest_screen

# Teste
largest_screen = get_largest_screen()
if largest_screen:
    print(f"Maior monitor: {largest_screen.frame().size.width}x{largest_screen.frame().size.height}")
else:
    print("Erro ao detectar os monitores.")

# -----------------------------------------------------------------------------------------------------
# Posiciona a janela do Whatsappp para Desktop no maior monitor disponível (MacOs apenas)
# -----------------------------------------------------------------------------------------------------

import Quartz
from AppKit import NSWorkspace
from Quartz import CGWindowListCopyWindowInfo, kCGWindowListOptionOnScreenOnly, kCGNullWindowID

def get_whatsapp_window():
    """ Encontra a janela do WhatsApp e retorna suas informações """
    app_name = "WhatsApp"
    window_list = CGWindowListCopyWindowInfo(kCGWindowListOptionOnScreenOnly, kCGNullWindowID)

    for window in window_list:
        if app_name.lower() in window.get("kCGWindowOwnerName", "").lower():
            salvar_log(Reg_Posicionar_WhatsApp, evento="WhatsApp encontrado", window_id=window.get("kCGWindowNumber"))
            return window

    salvar_log(Reg_Posicionar_WhatsApp, evento="WhatsApp não encontrado")
    return None

    # Teste (Debugger)
    show_message("Debug", f"WhatsApp encontrado? {get_whatsapp_window() is not None}")

def move_whatsapp_to_largest_screen():
    """ Move a janela do WhatsApp para o maior monitor """
    salvar_log(Reg_Posicionar_WhatsApp, evento="Iniciando reposicionamento do WhatsApp")

    whatsapp_window = get_whatsapp_window()
    if not whatsapp_window:
        salvar_log(Reg_Posicionar_WhatsApp, evento="Erro: WhatsApp não está aberto")
        print("❌ WhatsApp não está aberto.")
        return
    
    # Obtém a tela maior
    largest_screen = get_largest_screen()
    if not largest_screen:
        salvar_log(Reg_Posicionar_WhatsApp, evento="Erro ao detectar o maior monitor")
        print("❌ Erro ao detectar o maior monitor.")
        return
    
    screen_frame = largest_screen.frame()

    # Pegamos a posição atual e tamanho da janela
    current_x = whatsapp_window.get("kCGWindowBounds", {}).get("X", 0)
    current_y = whatsapp_window.get("kCGWindowBounds", {}).get("Y", 0)

    salvar_log(Reg_Posicionar_WhatsApp, evento="Posição atual do WhatsApp", x=current_x, y=current_y)

    # Se a janela já estiver no maior monitor, não movemos
    if screen_frame.origin.x <= current_x <= screen_frame.origin.x + screen_frame.size.width:
        salvar_log(Reg_Posicionar_WhatsApp, evento="WhatsApp já está no maior monitor")
        print("✅ WhatsApp já está no maior monitor.")
        return
    
    # Nova posição no maior monitor
    new_x = int(screen_frame.origin.x + 50)  # Pequeno deslocamento para evitar bugs
    new_y = int(screen_frame.origin.y + 50)

    # Comando para mover a janela
    script = f'''
    tell application "System Events"
        tell process "WhatsApp"
            set position of window 1 to {{{new_x}, {new_y}}}
        end tell
    end tell
    '''
    os.system(f"osascript -e '{script}'")

    salvar_log(Reg_Posicionar_WhatsApp, evento="WhatsApp movido", new_x=new_x, new_y=new_y, monitor_width=screen_frame.size.width, monitor_height=screen_frame.size.height)
    print(f"✅ WhatsApp movido para o monitor {screen_frame.size.width}x{screen_frame.size.height}.")

# Execução
move_whatsapp_to_largest_screen()  # Move o WhatsApp para o maior monitor

time.sleep(1)  # Pequeno delay para garantir que o foco foi aplicado

# Debuger:
show_message("Busca por Listas no Whatsapp", "✅ WhatsApp movido para o maior monitor disponível!")

# -----------------------------------------------------------------------------------------------------
# Maximiza a janela do Whatsappp para Desktop no maior monitor disponível (MacOs apenas)
# -----------------------------------------------------------------------------------------------------    

import pyautogui

def maximize_whatsapp():
    try:
        """ Maximiza o WhatsApp via AppleScript """

        salvar_log(Reg_Maximizar_WhatsApp, evento="Iniciando maximização do WhatsApp")

        script = '''
        tell application "WhatsApp" to activate
        delay 1
        tell application "System Events"
            tell process "WhatsApp"
                set frontmost to true
                set value of attribute "AXFullScreen" of window 1 to true
            end tell
        end tell
        '''
       
        result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)

        # Verifica se houve mensagem de erro no output
        if result.stderr:
            salvar_log(Reg_Maximizar_WhatsApp, evento="Erro no AppleScript", detalhe=result.stderr)
            print(f"❌ Erro do AppleScript: {result.stderr}")
        else:
            salvar_log(Reg_Maximizar_WhatsApp, evento="WhatsApp maximizado com sucesso")
            print("✅ WhatsApp maximizado com sucesso via AppleScript!")

    except subprocess.CalledProcessError as e:
        salvar_log(Reg_Maximizar_WhatsApp, evento="Falha técnica no subprocess", codigo_erro=e.returncode, mensagem=e.output, detalhes=e.stderr)
        print(f"❌ Falha técnica:")
        print(f"Código de erro: {e.returncode}")
        print(f"Mensagem: {e.output}")
        print(f"Detalhes: {e.stderr}")

    except Exception as e:
        salvar_log(Reg_Maximizar_WhatsApp, evento="Erro inesperado", detalhe=str(e))
        print(f"❌ Erro inesperado: {str(e)}")

maximize_whatsapp()  # Maximiza a janela

time.sleep(1)  # Pequeno delay para garantir que a maximização foi aplicada

# Debuger:
show_message("Busca por Listas no Whatsapp", "✅ WhatsApp maximizado no maior monitor disponível!")

# -----------------------------------------------------------------------------------------------------
# Seleciona, limpa e insere termos na Barra de Pesquisa Global do Whatsapp (MacOs apenas)
# -----------------------------------------------------------------------------------------------------  

import pyperclip  # Biblioteca para manipular a área de transferência

def limpar_e_escrever_na_barra_pesquisa(termo):

    salvar_log(Reg_Pesquisa_Global_WhatsApp, evento="Iniciando interação com a barra de pesquisa")

    try:
        # Aguarda alguns segundos para garantir que o WhatsApp está em foco
        time.sleep(2)

        # 1. Foca na barra de pesquisa usando o atalho Cmd + F (macOS)
        pyautogui.hotkey('command', 'f')  # No Windows, use: pyautogui.hotkey('ctrl', 'f')
        
        # 2. Seleciona todo o texto na barra de pesquisa
        pyautogui.hotkey('command', 'a')  # No macOS: Command + A (selecionar tudo)
        # No Windows, use: pyautogui.hotkey('ctrl', 'a')

        # 3. Apaga o texto selecionado
        pyautogui.press('backspace')  # Apaga tudo o que estiver selecionado

        # 4. Prepara o texto para colar
        pyperclip.copy(termo)  # Copia o termo para a área de transferência

        # 4.1. Aguarda alguns segundos para garantir que a o WhatsApp está em foco
        time.sleep(0.5)

        # 5. Foca novamente na barra de pesquisa (garantia extra)
        pyautogui.hotkey('command', 'f')
   
        # 6. Cola o texto da área de transferência
        pyautogui.hotkey('command', 'v')  # No macOS: Command + V (colar)
        # No Windows, use: pyautogui.hotkey('ctrl', 'v')
        
        time.sleep(0.5)  # Aguarda a digitação ser concluída

        salvar_log(Reg_Pesquisa_Global_WhatsApp, evento="Termo inserido com sucesso", detalhe=termo)
        print(f"Termo '{termo}' escrito na barra de pesquisa com sucesso!")

    except Exception as e:
        salvar_log(Reg_Pesquisa_Global_WhatsApp, evento="Erro ao interagir com a barra de pesquisa", detalhe=str(e))
        print(f"Erro ao interagir com a barra de pesquisa: {e}")

# Teste da função
limpar_e_escrever_na_barra_pesquisa("Teste: termo de busca")
time.sleep(2)  # Aguarda 2 segundos antes da próxima ação
limpar_e_escrever_na_barra_pesquisa("teste concluído")

# -----------------------------------------------------------------------------------------------------
#  Abre janela com a imagem ilustrativa (MacOs apenas)
# -----------------------------------------------------------------------------------------------------

# Captura o caminho do diretório onde o script está instalado
caminho_script = os.path.dirname(os.path.abspath(__file__))

# Remove a pasta 'src' e insere 'data' no lugar
diretorio_final = caminho_script.replace("/src", "/data")

# Altera o cursor do mouse para facilitar a visualização
def abrir_janela_de_imagem(nome_imagem = ""):
    """ Abre uma janela com a imagem ilustrativa """
    try:
        # Adiciona o nome da imagem ao diretório final
        caminho_completoA = os.path.join(diretorio_final, nome_imagem)

        if not os.path.exists(caminho_completoA):
            salvar_log(Reg_Abrir_Img_Ilustrativa, evento="Imagem não encontrada", detalhe=caminho_completoA)
            print(f"❌ Imagem '{nome_imagem}' não encontrada!")
            return

        subprocess.run(["open", caminho_completoA])

        salvar_log(Reg_Abrir_Img_Ilustrativa, evento="Imagem aberta com sucesso", detalhe=caminho_completoA)
        print(f"✅ Imagem '{nome_imagem}' aberta com sucesso!")

    except Exception as e:
        salvar_log(Reg_Abrir_Img_Ilustrativa, evento="Erro ao abrir imagem", detalhe=f"Arquivo: {nome_imagem} | Caminho: {caminho_completoA} | Erro: {e}")
        print(f"❌ Erro ao abrir a imagem '{nome_imagem}': {e}")

# Fecha apenas a imagem específica
def fechar_janela_de_imagem(nome_imagem = ""):
    """ Fecha a janela com a imagem ilustrativa """
    try:
        # Adiciona o nome da imagem ao diretório final
        caminho_completoF = os.path.join(diretorio_final, nome_imagem)
        
        if not os.path.exists(caminho_completoF):       
            salvar_log(Reg_Abrir_Img_Ilustrativa, evento="Imagem não encontrada", detalhe=caminho_completoF)
            print(f"❌ Imagem '{nome_imagem}' não encontrada!")
            return 
        
        script_fechar = f'''
        tell application "Preview"
            close (every window whose name contains "{os.path.basename(caminho_completoF)}")
        end tell
        '''

        subprocess.run(["osascript", "-e", script_fechar])

        salvar_log(Reg_Abrir_Img_Ilustrativa, evento="Imagem fechada com sucesso", detalhe=f"Arquivo: {nome_imagem} | Caminho: {caminho_completoF}")
        print(f"✅ Imagem '{nome_imagem}' fechada com sucesso!")

    except Exception as e:
        salvar_log(Reg_Abrir_Img_Ilustrativa, evento="Erro ao fechar imagem", detalhe=f"Arquivo: {nome_imagem} | Caminho: {caminho_completoF} | Erro: {e}")
        print(f"❌ Erro ao fechar a imagem '{nome_imagem}': {e}")       

# -----------------------------------------------------------------------------------------------------
#  Traz o Mesa Virtual do WhatsApp para frente (MacOs apenas)
# -----------------------------------------------------------------------------------------------------

def trazer_mesa_do_whatsapp_para_frente_um_monitor():
    """Traz o WhatsApp para frente em sistemas com apenas um monitor."""
    try:
        salvar_log(Reg_Trazer_WhatsApp_Frente, evento="Iniciando", detalhe="Tentando trazer WhatsApp para frente (1 monitor).")
        print("🔄 Trazendo o WhatsApp para frente (sistema com um monitor)...")
        
        script = '''
        tell application "System Events"
            tell application "WhatsApp" to activate
            delay 0.5
            key code 0 using {command down} -- Simula Cmd + Tab
        end tell
        '''

        #subprocess.run(["osascript", "-e", script])

        resultado = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)

        if resultado.stderr:
            salvar_log(Reg_Mesa_Virtual_Frente, evento="Erro", detalhe=f"Falha ao trazer WhatsApp (1 monitor): {resultado.stderr}")
            print(f"❌ Erro ao trazer WhatsApp para frente: {resultado.stderr}")
        else:
            salvar_log(Reg_Mesa_Virtual_Frente, evento="Sucesso", detalhe="WhatsApp trazido para frente (1 monitor).")
            print("✅ WhatsApp trazido para frente com sucesso!")


    except Exception as e:
        salvar_log(Reg_Mesa_Virtual_Frente, evento="Erro crítico", detalhe=f"Exceção ao trazer WhatsApp (1 monitor): {e}")
        print(f"❌ Erro inesperado: {e}")

def trazer_mesa_do_whatsapp_para_frente_dois_monitores():
    """ Traz o WhatsApp para frente """
    try:
        
        salvar_log(Reg_Mesa_Virtual_Frente, evento="Iniciando", detalhe="Tentando trazer WhatsApp para frente (2 monitores).")
        print("🔄 Trazendo o WhatsApp para frente (sistema com mais de um monitor)...")
        
        script = '''
        
        tell application "WhatsApp" to activate
        delay 1
        tell application "System Events"
            tell process "WhatsApp"
                set frontmost to true
            end tell
        end tell
        '''
        
        resultado = subprocess.run(["osascript", "-e", script])
    
        if resultado.stderr:
            salvar_log(Reg_Mesa_Virtual_Frente, evento="Erro", detalhe=f"Falha ao trazer WhatsApp (2 monitores): {resultado.stderr}")
            print(f"❌ Erro ao trazer WhatsApp para frente: {resultado.stderr}")
        else:
            salvar_log(Reg_Mesa_Virtual_Frente, evento="Sucesso", detalhe="WhatsApp trazido para frente (2 monitores).")
            print("✅ WhatsApp trazido para frente com sucesso!")

    except Exception as e:
        salvar_log(Reg_Mesa_Virtual_Frente, evento="Erro crítico", detalhe=f"Exceção ao trazer WhatsApp (2 monitores): {e}")
        print(f"❌ Erro inesperado: {e}")


# Traz o WhatsApp para frente com base no número de monitores
def trazer_mesa_do_whatsapp_para_frente():

    try:
        salvar_log(Reg_Mesa_Virtual_Frente, evento="Iniciando", detalhe="Detectando número de monitores...")

        # Detecta quantos monitores estão conectados
        num_monitores = len(screens)  # # Detecta quantos monitores estão conectados. Certifique-se de que a variável `screens` esteja definida corretamente

        salvar_log(Reg_Mesa_Virtual_Frente, evento="Monitores detectados", detalhe=f"Número de monitores: {num_monitores}")

        # Escolhe a versão correta da função com base no número de monitores
        if num_monitores == 1:
            trazer_mesa_do_whatsapp_para_frente_um_monitor() # Se for um monitor
        else:
            trazer_mesa_do_whatsapp_para_frente_dois_monitores() # Se forem dois monitores

    except Exception as e:
        salvar_log(Reg_Mesa_Virtual_Frente, evento="Erro crítico", detalhe=f"Falha ao detectar monitores: {e}")
        print(f"❌ Erro inesperado: {e}")

# -----------------------------------------------------------------------------------------------------
#  Identifica e Captura o Painel de Pesquisa do WhatsApp (MacOs apenas)
# -----------------------------------------------------------------------------------------------------

from pynput import mouse

# Captura de tela
def tirar_screenshot(x, y, largura, altura):
    salvar_log(Reg_Capturar_Painel_Pesquisa, evento = "Capturando tela.", Coord_x = x, Coord_y = y, var_largura = largura, var_altura = altura)
    return pyautogui.screenshot(region=(x, y, largura, altura))

# Altera o cursor do mouse para facilitar a visualização
def abrir_ferramenta_de_captura():
    """Simula a ação de pressionar Cmd+Shift+4 no macOS para abrir a ferramenta de captura de tela."""
    # Pressiona "Esc" para sair do campo de pesquisa
    pyautogui.press('esc')
    # Pequeno delay para garantir que o foco foi removido
    time.sleep(0.5)
    # Agora ativa a captura de tela
    pyautogui.hotkey('command', 'shift', '4')
    salvar_log(Reg_Capturar_Painel_Pesquisa, evento = "Ferramenta de captura de tela ativada.")

# Captura as coordenadas do clique do mouse
def capturar_coordenadas(NomeImg="", TextoMsg="Selecione a área para capturar as coordenadas iniciais..."):
    coordenadas = []  # Lista para armazenar as coordenadas

    def on_click(x, y, button, pressed):
        if pressed:  # Só captura no clique, não no soltar do botão
            resultado = subprocess.run(["cliclick", "p"], capture_output=True, text=True)
            coords = resultado.stdout.strip().split(",")  # Divide em X e Y
            coordenadas.extend([int(coords[0]), int(coords[1])])  # Converte para inteiros e salva
            print(f"Coordenadas capturadas: {coordenadas[0]}, {coordenadas[1]}")
            salvar_log(Reg_Capturar_Painel_Pesquisa, evento = f"Coordenadas capturadas: {coordenadas[0]}, {coordenadas[1]}")
            listener.stop()  # Para a escuta após capturar um clique
            
    print(TextoMsg)
    abrir_janela_de_imagem(NomeImg)
    time.sleep(0.5) # Aguarda 0,5 segundo para garantir a conclusão de atividades na janela em foco
    show_message("Busca por Listas no Whatsapp", TextoMsg)  
    time.sleep(1.5) # Aguarda 1,5 segundos para garantir a conclusão de atividades na janela em foco
    trazer_mesa_do_whatsapp_para_frente()  # Traz a janela do WhatsApp para frente
    time.sleep(1) # Aguarda 1 segundo para garantir que a janela está em foco

    # Executando a captura
    abrir_ferramenta_de_captura()
    time.sleep(2) # Aguarda 1 segundo para garantir que a janela está em foco

    with mouse.Listener(on_click=on_click) as listener:
        listener.join()  # Mantém o script aguardando o clique

    return tuple(coordenadas) if coordenadas else None  # Retorna as coordenadas capturadas

def conf_captura_coord():
    
    nome_imagem = "temp/teste_screenshot.png"
    caminho_completoT = os.path.join(diretorio_final, nome_imagem)
    
    while True:  # Loop infinito até o usuário decidir sair ou prosseguir

        # Simulação de captura de coordenadas
        x_min, y_min, largura, altura = None, None, None, None  # Valores vazios

        # Abre a janela de imagem para o usuário visualizar
        abrir_janela_de_imagem("referencia_cursor_captura.jpg")

        # Mensagem de instrução
        mensagem = ("\u26A0\uFE0F Para capturar as coordenadas da região a ser pesquisada você terá que selecioná-la manualmente em dois passos."
                    "\n\n✅ Este procedimento só necessita ser realizado uma vez!"   
                    "\n✅ Aguarde até o cursor tomar a forma de mira, conforme mostrado na janela da pré-visualização ao lado."
                    "\n✅ Aguarde até sua janela do WhatsApp ser maximizada no maior monitor."
                    "\n✅ Capture as coordenadas da região exemplificada no SEU WhatsaApp e ❌ não na janela de exemplo da pré-visualização!"
                )
        show_message("Busca por Listas no Whatsapp", mensagem)
        time.sleep(1) # Aguarda 1 segundo para garantir que a janela está em foco

        # Captura as coordenadas da região de pesquisa
        x, y = capturar_coordenadas("referencia_painel_CSE.bmp","\u26A0\uFE0F Clique no\nCANTO SUPERIOR ESQUERDO\ne arraste até o canto inferior direito para capturar as coordenadas...")    # Primeiro clique
        time.sleep(0.5) # Aguarda 0.5 segundo para garantir que a janela está em foco
        x1, y1 = capturar_coordenadas("referencia_painel_CID.bmp","\u26A0\uFE0F Clique no\nCANTO INFERIOR DIREITO\ne arraste até o canto superior esquerdo para capturar as coordenadas...")  # Segundo clique
        time.sleep(0.5) # Aguarda 0,5 segundo para garantir que a janela está em foco
        
        fechar_janela_de_imagem("referencia_painel_CID.bmp")  # Fecha a janela de imagem
        fechar_janela_de_imagem("referencia_painel_CSE.bmp")  # Fecha a janela de imagem
        fechar_janela_de_imagem("referencia_cursor_captura.jpg")  # Fecha a janela de imagem

        # Definindo os cantos com base na sua regra:
        x_min = min(x, x1)  # Menor X (esquerda)
        y_max = max(y, y1)  # Maior Y (superior)
        x_max = max(x, x1)  # Maior X (direita)
        y_min = min(y, y1)  # Menor Y (inferior)

        # Calculando largura e altura
        largura = x_max - x_min
        altura = y_max - y_min

        salvar_log(Reg_Capturar_Painel_Pesquisa, evento=f"Região capturada: ({x_min}, {y_min}, {largura}, {altura})")
        print(f"Região selecionada: ({x_min}, {y_min}, {largura}, {altura})")

        # Capturando a região slecionada da tela
        screenshot = tirar_screenshot(x, y, largura, altura)
        screenshot.save(caminho_completoT)
        salvar_log(Reg_Capturar_Painel_Pesquisa, evento="Captura salva.", arquivo=caminho_completoT)
        print("Captura concluída! Verifique o arquivo 'teste_screenshot.png'")

        # Abre a imagem capturada
        abrir_janela_de_imagem("temp/teste_screenshot.png")

        # Verifica se a imagem está correta
        salvar_log(Reg_Capturar_Painel_Pesquisa, evento="Validando coordenadas com o usuário.")
        resposta = show_3option_message("Busca por Listas no Whatsapp", "Captura concluída! Verifique se a imagem está correta.\n\nSe sim ✅ aperte em 'Prosseguir'\nSe ❌ não, aperte em retornar para tentar novamente\nOu, se preferir, aperte em 'Sair' para desistir ❌ da autobusca.")

        # Lógica condicional baseada na resposta do usuário

        if resposta == "Prosseguir":
            if None in (x_min, y_min, largura, altura):  # Verifica se algum valor é None
                salvar_log(Reg_Capturar_Painel_Pesquisa, evento="Usuário escolheu: Prosseguir", detalhe="Coordenadas inválidas.")
                print("⚠️ Erro: Coordenadas inválidas. Tente novamente.")
                print("🔄 Repetindo captura de coordenadas...\n")
                fechar_janela_de_imagem("temp/teste_screenshot.png")  # Fecha a janela de imagem
                continue  # Volta ao início do loop
            print("➡️ Continuando com o processo...")
            fechar_janela_de_imagem("temp/teste_screenshot.png")  # Fecha a janela de imagem
            salvar_log(Reg_Capturar_Painel_Pesquisa, evento="Usuário escolheu: Prosseguir", detalhe=f"Coordenadas validadas: ({x_min}, {y_min}, {largura}, {altura})")
            return x_min, y_min, largura, altura  # Retorna os valores apenas se forem válidos

        elif resposta == "Retornar":
            salvar_log(Reg_Capturar_Painel_Pesquisa, evento="Usuário escolheu: Retornar.")
            print("⬅️ Voltando para a etapa anterior...") 
            fechar_janela_de_imagem("temp/teste_screenshot.png")  # Fecha a janela de imagem 
            continue  # Reinicia o loop, pedindo a escolha novamente

        elif resposta == "Sair":
            salvar_log(Reg_Capturar_Painel_Pesquisa, evento="Usuário escolheu: Sair do programa.")
            print("❌ Encerrando o programa...")
            fechar_janela_de_imagem("temp/teste_screenshot.png")  # Fecha a janela de imagem
            exit()  # Sai completamente do programa

        else:
            salvar_log(Reg_Capturar_Painel_Pesquisa, evento="Usuário escolheu uma opção inválida.")
            print("⚠️ Opção inesperada! Por favor, tente novamente.")  
            fechar_janela_de_imagem("temp/teste_screenshot.png")  # Fecha a janela de imagem
            continue  # Se houver erro, repete a caixa de mensagem

# Chamando a função
x_painel_pesq_WA, y_painel_pesq_WA, larg_painel_pesq_WA, alt_painel_pesq_WA  = conf_captura_coord()

# -----------------------------------------------------------------------------------------------------------------------------------------------
# Pesquisar por um termo básico antes de rolar a tela (1. Caso contrário a tela poderá estar sem conteúdo suficiente para rolar; 2. MacOs apenas)
# -----------------------------------------------------------------------------------------------------------------------------------------------

trazer_mesa_do_whatsapp_para_frente() # Traz a janela do WhatsApp para frente
time.sleep(1)  # Pequeno delay para garantir que a janela
limpar_e_escrever_na_barra_pesquisa("oi")  # Escreve "oi" na barra de pesquisa
time.sleep(2)  # Aguarda 2 segundos antes da próxima ação

# -----------------------------------------------------------------------------------------------------
#  Calibração da rolagem automática (MacOs apenas)
# -----------------------------------------------------------------------------------------------------

# ============================
# 1️⃣ O usuário faz uma rolagem manual e o script captura o número de "ticks" (eventos de rolagem).
# ============================
# - O script exibe uma mensagem explicando o critério de rolagem.
# - O usuário faz a rolagem até que o conteúdo inicial desapareça completamente.
# - O script monitora os eventos de rolagem do mouse.
# - Armazena o número total de "ticks" até o usuário parar.
# - Exibe o número total de "ticks" capturados.

from Quartz import CGEventCreateMouseEvent

scroll_count = 0  # Variável para contar os eventos de rolagem

def on_scroll(x, y, dx, dy):
    """Incrementa a contagem de rolagem."""
    global scroll_count
    salvar_log(Reg_Calibrar_Rolagem, evento="Rolagem inicial ({scroll_count})")
    scroll_count += dy  # dy captura rolagens verticais
    salvar_log(Reg_Calibrar_Rolagem, evento="Incrementa a contagem de rolagem.", coordendas=f"Posição atual do mouse ({x}, {y}), deslocamento horizontal ({dx}), deslocamento vertical ({dy}) e rolagem final ({scroll_count})")

def monitor_scroll():
    """Monitora a rolagem do mouse enquanto o usuário interage."""
    with mouse.Listener(on_scroll=on_scroll) as listener:
        input("Role manualmente até que o conteúdo inicial desapareça COMPLETAMENTE e depois pressione ENTER...")
        salvar_log(Reg_Calibrar_Rolagem, evento="Início do monitoramento do mouse para calibração da rolagem.")
        listener.stop()  # Para a captura ao pressionar ENTER
        salvar_log(Reg_Calibrar_Rolagem, evento="Pressionado ENTER para a captura do valor de rolagem.")

# Inicia a captura de rolagem
monitor_scroll()

# Exibe a quantidade de eventos de rolagem capturados
print(f"Número de eventos de rolagem registrados: {scroll_count}")

# ============================
# 2️⃣ Armazenar essa unidade de rolagem
# ============================
# - Esse valor será usado para automatizar a rolagem futura na captura de telas.
# - O número de ticks é salvo para referência.

calibrated_scroll = abs(scroll_count)  # Usa o valor absoluto para evitar negativos
salvar_log(Reg_Calibrar_Rolagem, evento=f"Unidade de rolagem calibrada: {calibrated_scroll} ticks")
print(f"Unidade de rolagem calibrada: {calibrated_scroll} ticks")

# ============================
# 3️⃣ Ajuste de margens de segurança
# ============================
# - Se o último item não tiver espaço abaixo, reduzimos um pouco a unidade de rolagem.
# - Isso evita cortes na captura de tela.

final_scroll = max(1, calibrated_scroll)
salvar_log(Reg_Calibrar_Rolagem, evento=f"Unidade de rolagem mínima: {final_scroll} ticks")

# -----------------------------------------------------------------------------------------------------
#  4️⃣ Rolar a tela automaticamente (MacOs apenas)
# -----------------------------------------------------------------------------------------------------

import math  # Importa o módulo math
import threading
from pynput import keyboard

#import keyboard  # Instale com: pip install keyboard (Não utilizada no macOS em razão de limitações de permissão…)

# Calcula a posição de destino do mouse antes de iniciar a rolagem automática (Centro do painel de pesquisa do WhatsApp)
mouse_x, mouse_y = x_painel_pesq_WA + larg_painel_pesq_WA/2, y_painel_pesq_WA + alt_painel_pesq_WA/2
salvar_log(Reg_Calibrar_Rolagem, evento=f"Posição de destino do mouse (Centro do painel de pesquisa do WhatsApp): {mouse_x} x {mouse_y}.")

# Usamos um Event para controlar o travamento do mouse
mouse_locked = threading.Event()

# Função para detectar a tecla 'Esc'
def on_press(key):
    global mouse_locked  # Declara que estamos modificando a variável global
    if key == keyboard.Key.esc:
        salvar_log(Reg_Calibrar_Rolagem, evento="Mouse desbloqueado pelo usuário via tecla ESC.")
        print("Mouse desbloqueado pelo usuário.")
        mouse_locked[0] = False  # Modifica o valor da referência
        return False  # Interrompe o listener

# Função para manter o mouse fixo
def lock_mouse(Input_mouse_x, Input_mouse_y): 
    
    # Mensagem de bloqueio do mouse
    salvar_log(Reg_Calibrar_Rolagem, evento="Mouse bloqueado bloqueado automaticamente na posição {Input_mouse_x} x {Input_mouse_y}.")
    print("\u26A0\uFE0F Mouse bloqueado!\n\nPressione 'Esc' para desbloquear (não recomendado).")
   
    # Inicia o listener do teclado em uma thread separada
    listener = keyboard.Listener(on_press=on_press)
    listener.start()

    # Loop para manter o mouse fixo
    while mouse_locked.is_set():  # Verifica se o mouse está travado
        pyautogui.moveTo(Input_mouse_x, Input_mouse_y)
        time.sleep(0.05)  # Atualiza a posição a cada 50ms

    # Encerra o listener quando o loop termina
    listener.stop()
    salvar_log(Reg_Calibrar_Rolagem, evento="Mouse desbloqueado automaticamente.")

def travar_mouse(x, y):
    """Trava o mouse na posição (x, y) e retorna a thread de travamento."""
    global mouse_locked  # Declara que estamos modificando a variável global
    mouse_locked.set()  # Trava o mouse
    mouse_thread = threading.Thread(target=lock_mouse, args=(x, y), daemon=True)
    mouse_thread.start()
    salvar_log(Reg_Calibrar_Rolagem, evento=f"Mouse travado na posição ({x}, {y}).")
    return mouse_thread

def destravar_mouse(mouse_thread=None):
    """Destrava o mouse e aguarda a thread de travamento terminar."""
    global mouse_locked  # Declara que estamos modificando a variável global
    mouse_locked.clear()  # Desbloqueia o mouse
    salvar_log(Reg_Calibrar_Rolagem, evento="Mouse desbloqueado via sistema.")
    print("Mouse desbloqueado.")
    
    if mouse_thread and mouse_thread.is_alive():
        mouse_thread.join()  # Aguarda a thread terminar
        salvar_log(Reg_Calibrar_Rolagem, evento="Thread de travamento encerrada.")
        print("Thread de travamento encerrada.")

def rolar_pagina(tics, TicDelay):  
    pyautogui.scroll(tics)
    salvar_log(Reg_Calibrar_Rolagem, evento=f"Rolagem realizada: {tics} tics.")
    time.sleep(TicDelay)  # Pequeno delay para suavizar a rolagem

# Inicia a thread para travar o mouse
mouse_thread = travar_mouse(mouse_x, mouse_y)

# Retornar a pesquisa para o estado inicial
salvar_log(Reg_Calibrar_Rolagem, evento="Iniciando retorno da pesquisa para o estado inicial.")
print("Retornando a pesquisa para o estado inicial. Aguarde...")
for _ in range(math.floor(calibrated_scroll*1.5)):
    rolar_pagina(1, 0) # Rola para cima um tic

time.sleep(1)  # Pequeno delay para garantir a rolagem

# Iniciando a rolagem automática de confirmação
scroll_tics = math.floor(0.6111*final_scroll)  # Supondo que `scroll_tics` armazene a quantidade de tics capturados durante a calibração com a curva y = 0,6111x
salvar_log(Reg_Calibrar_Rolagem, evento=f"Iniciando rolagem automática com {scroll_tics} tics.")
print(f"Iniciando rolagem de {scroll_tics} tics...")
for _ in range(scroll_tics):
    rolar_pagina(-1, 0.05)  # Rola para baixo um tic com pequeno delay para suavizar a rolagem

# Desbloqueia o mouse
destravar_mouse(mouse_thread)

# Pede confirmação para o usuário de que a rolagem foi bem-sucedida
salvar_log(Reg_Calibrar_Rolagem, evento="Rolagem concluída com sucesso.")
show_message("Busca por Listas no Whatsapp", "✅ Rolagem concluída com sucesso!\nVerifique se a rolagem foi realizada corretamente.")

MARGIN_ADJUSTMENT = 30  # Ajuste de segurança
salvar_log(Reg_Calibrar_Rolagem, evento=f"Ajuste de segurança: {MARGIN_ADJUSTMENT} ticks.")
print(f"Margem de segurança: {MARGIN_ADJUSTMENT} ticks")
safe_scroll_tics = math.floor((scroll_tics-MARGIN_ADJUSTMENT)*0.75)  # Multiplica por 3/4 para garantir que não ultrapasse
salvar_log(Reg_Calibrar_Rolagem, evento=f"Unidade de rolagem segura: {safe_scroll_tics} ticks.")
print(f"Unidade de rolagem segura: {safe_scroll_tics} ticks")

# -----------------------------------------------------------------------------------------------------
#  Bloco de Capturas, OCR e Tratamento de Dados (MacOs apenas)
# -----------------------------------------------------------------------------------------------------

import pytesseract
import pandas as pd
from pynput.mouse import Controller, Button
import cv2
import numpy as np

caffeinate_process = None  # Variável global para armazenar o processo

def ativar_caffeinate():
    global caffeinate_process
    caffeinate_process = subprocess.Popen(["caffeinate", "-d"])
    salvar_log(Reg_Capturas_OCR_Tratamento, evento="Iniciado o caffeinate.")

def desativar_caffeinate():
    global caffeinate_process
    if caffeinate_process:
        caffeinate_process.terminate()  # Encerra o processo
        caffeinate_process = None
        salvar_log(Reg_Capturas_OCR_Tratamento, evento="Caffeinate desativado.")

# def posicionar_mouse(x, y): # Já usada antes, compatibilizar!
#     mouse = Controller()
#     mouse.position = (x, y)

def processar_ocr(imagem):
    """
    Processa a imagem com OCR, separando corretamente os itens e identificando linhas em branco.
    """
    salvar_log(Reg_Capturas_OCR_Tratamento, evento="Iniciando processamento OCR.")
    
    texto_bruto = pytesseract.image_to_string(imagem) # Extrai texto da imagem
    salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Texto bruto extraído pelo OCR:\n{texto_bruto}")

    linhas = texto_bruto.split('\n') # Separa o texto por quebras de linha
    salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Linhas extraídas do texto OCR: {linhas}")
    
    itens = []
    item_atual = []
    ultima_linha_em_branco = False  # Bandeira para indicar se a última linha foi em branco
    
    for linha in linhas:

        linha_limpa = linha.strip() # remove qualquer espaço, tabulação ou quebra de linha antes e depois do texto da variável linha
        salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Linha original: '{linha}' -> Linha limpa: '{linha_limpa}'")

        if not linha_limpa:  # Se a linha está em branco (Em Python, strings não vazias são avaliadas como True, enquanto strings vazias - "" - são avaliadas como False.)
            if item_atual:  # Se havia um item em construção, ele é concluído (Em Python, listas vazias - [] - são avaliadas como False)
                itens.append("\n".join(item_atual)) # Junta os elementos da lista item_atual em uma única string, separando-os por um espaço (" ") e adiciona a string resultante ao final da lista itens.
                salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Item capturado e adicionado à lista: {itens[-1]}")
                item_atual = [] # item_atual esvaziado para iniciar um novo item
            ultima_linha_em_branco = True  # Linha em branco detectada
        else: # Se a linha não está em branco...
            item_atual.append(linha_limpa)
            salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Item em construção: {item_atual[-1]}")
            ultima_linha_em_branco = False  # Linha não está vazia
    
    if item_atual:  # Adiciona o último item caso o loop termine sem linha em branco
        itens.append("\n".join(item_atual)) # Garante que um último bloco de texto não seja perdido, caso o texto_bruto não termine com uma linha vazia.
        salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Último item capturado e adicionado à lista: {itens[-1]}")
    
    # Criar matriz N x 2 com a classificação de "completo" ou "incompleto"
    matriz = [[item, "completo"] for item in itens]  # Assume "completo" para todos
    if matriz and not ultima_linha_em_branco:
        matriz[-1][1] = "incompleto"  # O último item será incompleto se a última linha não foi em branco

    return matriz

def classificar_itens(matriz, novos_itens, num_interacao):
    """
    Classifica os novos itens identificando se são 'completos' ou 'incompletos',
    e se são 'novos' ou 'repetidos' com base nas interações anteriores.

    Parâmetros:
    - matriz: Lista de listas contendo as interações anteriores. Formato: [interacao, nome, completo, repetido]
    - novos_itens: Lista de listas contendo os itens da interação atual. Formato: [nome, completo]
    - num_interacao: Número da interação atual.

    Retorna:
    - Lista de listas representando os novos itens classificados. Formato: [interacao, nome, completo, repetido]
    """

    salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Iniciando classificação - Iteração: {num_interacao}")

    # Criar um conjunto com os nomes dos itens completos de interações anteriores
    itens_completos_anteriores = {linha[1] for linha in matriz if linha[2] == "completo"}

    salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Itens anteriores completos: {itens_completos_anteriores}")

    novos_resultados = []

    for item in novos_itens:
        nome_item, status_completo = item[0], item[1]  # Nome do item e se é completo/incompleto

        # Verificar se o item é repetido (só se já foi completo em uma interação anterior)
        item_repetido = "repetido" if status_completo == "completo" and nome_item in itens_completos_anteriores else "novo"

        # Adicionar à lista de resultados
        novos_resultados.append([num_interacao, nome_item, status_completo, item_repetido])

        salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Interação: {num_interacao} | Item: {nome_item} | Completo: {status_completo} | Repetido: {item_repetido}")

    salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Classificação finalizada - Iteração {num_interacao}")

    return novos_resultados

def verificar_potencial_ultimo_print(historico, num_interacao, num_int_check):

    salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Verificando potencial último print - num_int_check: {num_int_check}")

    # Se ainda não houve interações suficientes, não pode encerrar
    if num_interacao < num_int_check:
        salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Interpretação de interações insuficientes para verificação de último print.")
        return False  
    else:
        salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Interpretação de interações suficientes para verificação de último print.")
        if not historico: # Se o histórico estiver vazio, não há como verificar nada
            salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Todas as {num_int_check} capturas retornaram vazias. Encerrando pesquisa do termo.")
            return True

    # Pegamos as últimas `num_int_check` interações inteiras
    ultimos = []
    interacoes_vistas = set()

    for linha in reversed(historico):  # Percorre o histórico de trás para frente
        num_interacao_linha = linha[0]  # Número da interação salvo na coluna 0
        
        if num_interacao_linha not in interacoes_vistas:
            interacoes_vistas.add(num_interacao_linha)

        if len(interacoes_vistas) > num_int_check:
            break  # Já coletamos as últimas `num_int_check` interações, podemos parar

        ultimos.append(linha)
    
    # Precisamos inverter a ordem porque coletamos de trás para frente
    ultimos.reverse()

    # Filtra apenas os itens completos
    itens_completos = [linha for linha in ultimos if linha[2] == "completo"]

    # Verifica se todos os itens completos são repetidos
    resultado_final = all(linha[3] == "repetido" for linha in itens_completos)

    return resultado_final  

def executar_ocr_loop(safe_scroll_tics, over_scroll_check_ticks, num_int_check, x_painel_pesq_WA, y_painel_pesq_WA, larg_painel_pesq_WA, alt_painel_pesq_WA):
    
    global diretorio_final
    
    trazer_mesa_do_whatsapp_para_frente()
    time.sleep(1)  # Pequeno delay para garantir o foco
    ativar_caffeinate()
    historico = []
    num_interacao = 0 # Inicializa o contador de interações
    
    while True:

        salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Iniciando interação {num_interacao}.")

        nome_imagem = f"temp/screenshot_OCR_{num_interacao:04}.png"
        caminho_completoOCR = os.path.join(diretorio_final, nome_imagem)

        salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Imagens serão salvas em {caminho_completoOCR}.")

        num_interacao += 1 # Incrementa o contador de interações
        print(f"Interacao {num_interacao} iniciada...")
        
        # 🔹 Grupo de ações 1: Captura de Tela. Mouse travado no canto superior direito de painel de pesquisa...
        mouse_x, mouse_y = x_painel_pesq_WA + larg_painel_pesq_WA, y_painel_pesq_WA + alt_painel_pesq_WA/2
        # Inicia a thread para travar o mouse
        mouse_thread = travar_mouse(mouse_x, mouse_y)

        print("Executando ações do Grupo 1: Captura de Tela...")
        
        screenshot = tirar_screenshot(x_painel_pesq_WA, y_painel_pesq_WA, larg_painel_pesq_WA, alt_painel_pesq_WA)
        screenshot.save(caminho_completoOCR)
        lista_itens = processar_ocr(screenshot)

        salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"Classificando {len(lista_itens)} itens na interação {num_interacao}.")

        matriz = classificar_itens(historico, lista_itens, num_interacao)
        historico.extend(matriz)

        # 🔹 Desbloqueia o mouse para mudar de posição
        destravar_mouse(mouse_thread)  # Desbloqueia o mouse e aguarda a thread terminar
        
        # 🔹 Grupo de ações 2: Verificar se o último print é repetido. Mouse travado no centro de painel de pesquisa...
        mouse_x, mouse_y = x_painel_pesq_WA + larg_painel_pesq_WA/2, y_painel_pesq_WA + alt_painel_pesq_WA/2
        # Inicia a thread para travar o mouse
        print("Mudando a posição do mouse...")
        mouse_thread = travar_mouse(mouse_x, mouse_y)

        print("Executando ações do Grupo 2: Verificação de repetição de print...")

        if verificar_potencial_ultimo_print(historico, num_interacao, num_int_check):
            salvar_log(Reg_Capturas_OCR_Tratamento, evento="Print repetido detectado. Verificando novamente após rolagem.")
            rolar_pagina(-over_scroll_check_ticks, 0)
            screenshot = tirar_screenshot(x_painel_pesq_WA, y_painel_pesq_WA, larg_painel_pesq_WA, alt_painel_pesq_WA)
            lista_itens = processar_ocr(screenshot)
            matriz = classificar_itens(historico, lista_itens, num_interacao)
            if verificar_potencial_ultimo_print(historico, num_interacao, num_int_check):
                salvar_log(Reg_Capturas_OCR_Tratamento, evento="Prints repetidos confirmados. Encerrando OCR loop.")
                # Sobrescreve a última linha do histórico como "completo"
                if historico:
                    historico[-1][2] = "completo"
                break
            rolar_pagina(over_scroll_check_ticks, 0)  # Volta à posição anterior

        # 🔹 Grupo de ações 3: Rolar a pesquisa para o próximo print. Mouse travado no centro de painel de pesquisa...
        rolar_pagina(-safe_scroll_tics, 0.05)  # Rola para baixo um pouco
        print("Executando ações do Grupo 3: Rolar a pesquisa...")
        
        # 🔹 Desbloqueia o mouse novamente antes de repetir o loop
        destravar_mouse(mouse_thread)  # Desbloqueia o mouse e aguarda a thread terminar
        print("Loop de captura e rolagem reiniciando...\n")

        salvar_log(Reg_Capturas_OCR_Tratamento, evento="Rolagem realizada para próximo print.")

    df = pd.DataFrame(historico, columns=["Interacao", "Item", "Completo", "Repetido"])
    df_filtrado = df[(df["Completo"] == "completo") & (df["Repetido"] == "novo")]  # Limpa os dados
    df.to_csv(caminho_completoOCR, index=False, quoting=csv.QUOTE_NONNUMERIC)  # Salva o arquivo CSV

    salvar_log(Reg_Capturas_OCR_Tratamento, evento=f"OCR loop concluído. Dados salvos em {caminho_completoOCR}.")
    print(f"OCR loop concluído. Dados salvos em {caminho_completoOCR}.")
    desativar_caffeinate()  # ✅ Desativa fora do loop, apenas uma vez, antes do return
    salvar_log(Reg_Capturas_OCR_Tratamento, evento="Caffeinate desativado.")

    return df

# Execução do loop de captura, OCR e rolagem.
num_int_check = math.ceil((scroll_tics+3*MARGIN_ADJUSTMENT)/safe_scroll_tics) # Número de interações para verificar se o último print é repetido (3 vezes a margem de segurança)
salvar_log(Reg_Capturas_OCR_Tratamento,  evento=f"Número de interações para verificar se o último print é repetido {num_int_check}.")
over_scroll_check_ticks = safe_scroll_tics*num_int_check*2  # Número de tics para verificar se o último print é repetido 
salvar_log(Reg_Capturas_OCR_Tratamento,  evento=f"Número de tics para verificar se o último print é repetido {over_scroll_check_ticks}.")

trazer_mesa_do_whatsapp_para_frente() # Traz a janela do WhatsApp para frente
time.sleep(1)  # Pequeno delay para garantir que a janela
limpar_e_escrever_na_barra_pesquisa("contrato")  # Escreve o termo pesquisado na barra de pesquisa
time.sleep(2)  # Aguarda 2 segundos antes da próxima ação

df = executar_ocr_loop(safe_scroll_tics, over_scroll_check_ticks, num_int_check, x_painel_pesq_WA, y_painel_pesq_WA, larg_painel_pesq_WA, alt_painel_pesq_WA)

# -----------------------------------------------------------------------------------------------------
# Aviso de conclusão do programa
# -----------------------------------------------------------------------------------------------------  

# Exemplo de uso:
show_message("Busca por Listas no Whatsapp", "✅ Programa concluído com sucesso!\nVocê já pode voltar a operar o sistema normalmente!")