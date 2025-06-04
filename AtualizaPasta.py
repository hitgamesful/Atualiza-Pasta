import tkinter as tk
from tkinter import messagebox
import pyautogui
import pytesseract
from PIL import Image
import subprocess
import time
import os
import re
import shutil
import json

# -------- CONFIGURAÇÕES INICIAIS --------

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ['TESSDATA_PREFIX'] = r"C:\Program Files\Tesseract-OCR\tessdata"
ARQUIVO_COORDENADAS = "coordenadas.json"

# -------- LER COORDENADAS DO JSON --------

def ler_coordenadas():
    with open(ARQUIVO_COORDENADAS, "r") as f:
        return json.load(f)

def clicar_elementos(coords):
    time.sleep(1)
    pyautogui.click(coords["baixar"][0], coords["baixar"][1])
    time.sleep(1)
    pyautogui.click(coords["apilocal"][0], coords["apilocal"][1])
    time.sleep(0.5)
    # pyautogui.click(coords["apiweb"][0], coords["apiweb"][1])  # Linha removida!
    pyautogui.click(coords["sisplanweb"][0], coords["sisplanweb"][1])
    time.sleep(0.5)
    pyautogui.click(coords["download"][0], coords["download"][1])
    time.sleep(0.5)

def versao_existe_na_pasta(versao, pasta):
    try:
        for nome in os.listdir(pasta):
            if versao in nome:
                return True
        return False
    except Exception as e:
        messagebox.showerror("Erro ao acessar pasta", str(e))
        return False

def executar_rotina_quando_nao_acha(versao):
    try:
        shutil.rmtree(r"C:\Apache24\htdocs\sisplan_web_old", ignore_errors=True)
        shutil.rmtree(r"C:\Sisplan\estrutura_api", ignore_errors=True)
        time.sleep(3)

        nomePasta = versao
        caminhoPasta = os.path.join(r"C:\VersoesWEB", nomePasta)
        os.makedirs(caminhoPasta, exist_ok=True)

        time.sleep(15)
        origem = r"C:\AtualizarPasta\AtualizaPasta"
        for arquivo in ["ApiLocal.zip", "apiWeb.zip", "sisplan_web.zip"]:
            origem_arquivo = os.path.join(origem, arquivo)
            if os.path.exists(origem_arquivo):
                shutil.copy(origem_arquivo, caminhoPasta)

        subprocess.call('taskkill /f /im atualizador_loja_web.exe', shell=True)

        root.lift()
        root.attributes('-topmost', True)
        root.after_idle(root.attributes, '-topmost', False)
        messagebox.showinfo("Cópia finalizada", "Arquivos copiados com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro na rotina", str(e))

def buscar_versao():
    try:
        exe_path = r'C:\Apache24\htdocs\atualizador_sisplan_web\atualizador_loja_web.exe'
        subprocess.Popen(exe_path)
        time.sleep(3)

        x1 = 1000
        y1 = 550
        width = 700
        height = 55

        screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
        screenshot.save("recorte_versao.png")

        texto = pytesseract.image_to_string(Image.open("recorte_versao.png"), lang='por')
        print(f"TEXTO CAPTURADO: {texto}")

        match = re.search(r'[\d]+\.[\d]+\.[\d]+\.[\d]+', texto)
        if match:
            versao = match.group(0)
            pasta_versoes = r'C:\VersoesWEB'
            if versao_existe_na_pasta(versao, pasta_versoes):
                resultado = f"Última versão encontrada: {versao}\nStatus: VERSÃO ENCONTRADA NA PASTA! Nenhuma ação será tomada."
                lbl_resultado.config(text=resultado)
                # Fecha o atualizador e o app após 1,5 segundos
                subprocess.call('taskkill /f /im atualizador_loja_web.exe', shell=True)
                root.after(1500, root.destroy)
            else:
                resultado = f"Última versão encontrada: {versao}\nStatus: NÃO encontrada em C:\\VersoesWEB\nIniciando download e criação da nova pasta."
                lbl_resultado.config(text=resultado)
                clicar_elementos(coords)
                executar_rotina_quando_nao_acha(versao)
        else:
            lbl_resultado.config(text="Não foi possível encontrar a versão na imagem.\nVeja o print no arquivo 'recorte_versao.png'.")
    except Exception as e:
        messagebox.showerror("Erro", str(e))

# -------- INICIALIZAÇÃO DO SCRIPT --------

coords = ler_coordenadas()

root = tk.Tk()
root.title("Buscar Última Versão")

lbl_resultado = tk.Label(root, text="Buscando a versão, aguarde...")
lbl_resultado.pack(padx=10, pady=10)

# --- INÍCIO AUTOMÁTICO ---
root.after(100, buscar_versao)  # Chama a função buscar_versao() após 100ms da janela abrir

root.mainloop()
