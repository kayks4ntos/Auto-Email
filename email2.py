import pandas as pd
import pyautogui
import time
import webbrowser
import pyperclip
import ssl
import os

# Ignorar certificados SSL se necessário
ssl._create_default_https_context = ssl._create_unverified_context

# Link da planilha em formato CSV
url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQLNV87-CA9wsG9KFXEN5v1DbEmnPy0AP4DNLITjOYKR0NNgdeteWVtchsj8M_ldsmS0uxjY0M9QE8x/pub?output=csv"

# Lê os dados da planilha
df = pd.read_csv(url)
df.columns = df.columns.str.strip()  # Remove espaços nos nomes das colunas

# Caminho para salvar os e-mails já enviados
log_path = "emails_enviados.txt"
if os.path.exists(log_path):
    with open(log_path, "r") as f:
        enviados = set(linha.strip() for linha in f.readlines())
else:
    enviados = set()

# Filtra apenas os que ainda não receberam
novos_emails = df[~df["email"].isin(enviados)]

# Verifica se há algo para enviar
if novos_emails.empty:
    print("✅ Todos os e-mails já foram enviados.")
else:
    print(f"📧 Existem {len(novos_emails)} novos e-mails para enviar.")
    
    # Abre o Gmail
    webbrowser.open("https://mail.google.com")
    time.sleep(9)  # tempo para carregar

    for index, row in novos_emails.iterrows():
        nome = row["nome"]
        email = row["email"]

        print(f"➡️ Enviando email para: {email}")

        assunto = "Obrigado pela resposta"
        mensagem = f"Olá {nome},\n\nObrigado por responder! Em breve entraremos em contato.\n\nAtenciosamente, Alunos do senac."

        # Abrir novo e-mail
        pyautogui.click(x=140,y=178)
        time.sleep(3)

        # Preencher campos do e-mail
        pyperclip.copy(email)
        pyautogui.hotkey("ctrl", "v")
        pyautogui.press("tab")

        pyperclip.copy(assunto)
        pyautogui.hotkey("ctrl", "v")
        pyautogui.press("tab")

        pyperclip.copy(mensagem)
        pyautogui.hotkey("ctrl", "v")

        # Enviar
        pyautogui.hotkey("ctrl", "enter")
        time.sleep(3)

        # Registrar no log
        with open(log_path, "a") as f:
            f.write(email + "\n")

    print("✅ Todos os novos e-mails foram enviados com sucesso.")
