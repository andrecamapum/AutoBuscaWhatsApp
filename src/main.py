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
        print("❌ WhatsApp Desktop não está em execução. Abrindo agora...")
        subprocess.run(["open", "-a", "WhatsApp"])

Delay = 2  # 2 segundos

if __name__ == "__main__":
    if is_whatsapp_running():
        print("✅ WhatsApp Desktop está aberto!")
        show_message("Busca por Listas no Whatsapp", f"\u26A0\uFE0F Aguarde!\n\nO WhatsApp será reiniciado automaticamente após {Delay}s!")

# -----------------------------------------------------------------
# Código para fechar e reabrir o WhatsApp se necessário (MacOs apenas)
# -----------------------------------------------------------------

import time
import os

def close_whatsapp():
    """Fecha o WhatsApp caso esteja aberto."""
    if is_whatsapp_running():
        os.system('osascript -e \'quit app "WhatsApp"\'')
        time.sleep(Delay)  # Aguarda um tempo para garantir que fechou

def restart_whatsapp():
    """Reinicia o WhatsApp."""
    close_whatsapp()
    open_whatsapp_if_needed()  # Supondo que esta função já esteja implementada

# Agora, sempre que quiser garantir que o WhatsApp está no estado inicial:
restart_whatsapp()

# -----------------------------------------------------------------------------------------------------
# Identifica o maior monitor disponível (MacOs apenas)
# -----------------------------------------------------------------------------------------------------

import AppKit
from AppKit import NSScreen

def get_largest_screen():
    screens = NSScreen.screens()
    if not screens:
        print("Nenhuma tela detectada.")
        return None

    largest_screen = max(screens, key=lambda s: s.frame().size.width * s.frame().size.height)
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
            return window

    return None

def move_whatsapp_to_largest_screen():
    """ Move a janela do WhatsApp para o maior monitor """
    whatsapp_window = get_whatsapp_window()
    if not whatsapp_window:
        print("❌ A janela do WhatsApp não está localizada no monitor maior.")
        return

    largest_screen = max(NSScreen.screens(), key=lambda s: s.frame().size.width * s.frame().size.height)
    screen_frame = largest_screen.frame()
    
    # Pegamos a posição atual e tamanho da janela
    current_x = whatsapp_window.get("kCGWindowBounds", {}).get("X", 0)
    current_y = whatsapp_window.get("kCGWindowBounds", {}).get("Y", 0)

    # Se a janela já estiver no maior monitor, não movemos
    if screen_frame.origin.x <= current_x <= screen_frame.origin.x + screen_frame.size.width:
        print("✅ WhatsApp já está no maior monitor.")
        return
    
    # Calculamos a nova posição da janela
    new_x = screen_frame.origin.x + 50  # Pequeno deslocamento
    new_y = screen_frame.origin.y + 50

    # Comando para mover a janela
    script = f'''
    tell application "System Events"
        tell process "WhatsApp"
            set position of window 1 to {{{new_x}, {new_y}}}
        end tell
    end tell
    '''
    os.system(f"osascript -e '{script}'")

    print(f"✅ WhatsApp movido para o monitor {screen_frame.size.width}x{screen_frame.size.height}.")

# -----------------------------------------------------------------------------------------------------
# Maximiza a janela do Whatsappp para Desktop no maior monitor disponível (MacOs apenas)
# -----------------------------------------------------------------------------------------------------    

def maximize_whatsapp():
    script = """
    tell application "System Events"
        tell process "WhatsApp"
            if exists window 1 then
                set frontmost to true
                set value of attribute "AXFullScreen" of window 1 to true
            end if
        end tell
    end tell
    """
    os.system(f"osascript -e '{script}'")
    print("✅ WhatsApp maximizado.")

# Executa a movimentação e maximização
move_whatsapp_to_largest_screen()
time.sleep(1)  # Pequena pausa para evitar falhas
maximize_whatsapp()

# -----------------------------------------------------------------------------------------------------
# Seleciona e Limpa e Insere termos na Barra de Pesquisa Global do Whatsapp (MacOs apenas)
# -----------------------------------------------------------------------------------------------------  

import pyautogui
import pyperclip  # Biblioteca para manipular a área de transferência

def limpar_e_escrever_na_barra_pesquisa(termo):
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

        print(f"Termo '{termo}' escrito na barra de pesquisa com sucesso!")
    except Exception as e:
        print(f"Erro ao interagir com a barra de pesquisa: {e}")

# Teste da função
limpar_e_escrever_na_barra_pesquisa("Teste: termo de busca")
time.sleep(2)  # Aguarda 2 segundos antes da próxima ação
limpar_e_escrever_na_barra_pesquisa("teste concluído")

# -----------------------------------------------------------------------------------------------------
# Aviso de conclusão do programa
# -----------------------------------------------------------------------------------------------------  

# Exemplo de uso:
show_message("Busca por Listas no Whatsapp", "✅ Programa concluído com sucesso!\nVocê já pode voltar a operar o sistema normalmente!")