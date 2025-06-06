import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import threading
import time
import json
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# === Aparência do app ===
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# === Arquivo de logins ===
LOGIN_JSON_PATH = Path("logins.json")

def carregar_logins():
    if LOGIN_JSON_PATH.exists():
        with open(LOGIN_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def salvar_logins():
    with open(LOGIN_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(logins_disponiveis, f, indent=2, ensure_ascii=False)

logins_disponiveis = carregar_logins()
login_selecionado = {"nome": None, "login": None, "senha": None}
pausa_event = threading.Event()
pausa_event.set()
processo_em_execucao = {"ativo": False}

def iniciar_driver():
    options = Options()
    options.add_argument("--start-maximized")
    return webdriver.Chrome(options=options)

def login(driver):
    driver.get("https://clarobrasil.etadirect.com/")
    time.sleep(2)
    driver.find_element(By.ID, "username").send_keys(login_selecionado["login"])
    driver.find_element(By.ID, "password").send_keys(login_selecionado["senha"])
    driver.find_element(By.ID, "sign-in").click()
    time.sleep(5)
    driver.find_element(By.XPATH, "/html/body/div/div[3]/form/div[2]/input").send_keys(login_selecionado["login"])
    driver.find_element(By.XPATH, "/html/body/div/div[3]/form/div[3]/input").send_keys(login_selecionado["senha"])
    driver.find_element(By.XPATH, "/html/body/div/div[3]/form/div[4]/button").click()
    time.sleep(8)

def elemento_existe(driver, by, value):
    try:
        driver.find_element(by, value)
        return True
    except:
        return False

def mover_tecnico(driver, nome_tecnico):
    try:
        driver.find_element(By.CLASS_NAME, "hint_links_mobility").find_element(By.XPATH, ".//span[text()='Mover']").click()
        time.sleep(1)
        driver.find_element(By.CLASS_NAME, "oj-switch-track").click()
        campo_pesquisa = driver.find_element(By.CLASS_NAME, "oj-inputsearch-filter")
        campo_pesquisa.click()
        campo_pesquisa.send_keys(nome_tecnico)
        time.sleep(1)
        tecnicos = driver.find_elements(By.XPATH, "//div[contains(@class, 'resource-name-wrapper')]")
        if len(tecnicos) > 1:
            tecnicos[1].click()
            time.sleep(1)
            driver.find_element(By.XPATH, "//button[text()='Mover']").click()
        else:
            raise Exception("Técnico não encontrado")
    except Exception as e:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)
        raise Exception(f"Erro ao mover técnico: {str(e)}")

def processar_planilha(filepath, aba_escolhida):
    try:
        df = pd.read_excel(filepath, sheet_name=aba_escolhida, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Erro ao abrir arquivo", f"Não foi possível abrir o arquivo:\n{str(e)}")
        processo_em_execucao["ativo"] = False
        return

    colunas_necessarias = {
    "contrato": "Contrato",
    "novo tec":"Novo Tec",
    "novo login": "Novo Login",
    "se": "Se"
}


    colunas_planilha = {col.lower().strip(): col for col in df.columns}

    colunas_mapeadas = {}
    colunas_faltando = []

    for nome_normalizado, nome_padrao in colunas_necessarias.items():
        if nome_normalizado in colunas_planilha:
            colunas_mapeadas[nome_padrao] = colunas_planilha[nome_normalizado]
        else:
            colunas_faltando.append(nome_padrao)

    if colunas_faltando:
        messagebox.showerror("Erro na Planilha", f"Faltando colunas obrigatórias:\n\n{', '.join(colunas_faltando)}")
        processo_em_execucao["ativo"] = False
        return

    df.rename(columns={v: k for k, v in colunas_mapeadas.items()}, inplace=True)

    colunas_esperadas = [
        "Contrato", "Novo Téc", "Novo Login", "Se", "Nome", "Técnico atual", "Resultado"
    ]
    for coluna in colunas_esperadas:
        if coluna not in df.columns:
            df[coluna] = ""

    driver = iniciar_driver()

    progresso_win = ctk.CTkToplevel()
    progresso_win.attributes("-topmost", True)
    progresso_win.title("Progresso")
    progresso_win.geometry("300x150")
    progresso_win.resizable(False, False)

    progresso_label = ctk.CTkLabel(progresso_win, text="Iniciando...", font=ctk.CTkFont(size=14))
    progresso_label.pack(pady=(20, 5))

    progresso_bar = ctk.CTkProgressBar(progresso_win, width=200, progress_color="#d9534f", fg_color="#2a2a2a")
    progresso_bar.pack(pady=10)
    progresso_bar.set(0)

    btns_frame = ctk.CTkFrame(progresso_win)
    btns_frame.pack(pady=10)

    ctk.CTkButton(btns_frame, text="Pausar", command=lambda: pausa_event.clear(),
                  fg_color="#d9534f", hover_color="#c9302c").pack(side="left", padx=5)
    ctk.CTkButton(btns_frame, text="Retomar", command=lambda: pausa_event.set(),
                  fg_color="#d9534f", hover_color="#c9302c").pack(side="left", padx=5)

    try:
        login(driver)
        total = len(df)
        for index, row in df.iterrows():
            pausa_event.wait()
            progresso_label.configure(text=f"Processando: {row['Contrato']}")

            if row["Se"] != "Trocar":
                df.at[index, "Resultado"] = "Ignorado"
                continue

            try:
                contrato = ''.join(filter(str.isdigit, str(row["Contrato"])))
                tecnico_destino = str(row["Novo Login"]).strip()

                campo_busca = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='search-bar-container']//input[contains(@class, 'search-bar-input')]"))
                )
                driver.execute_script("arguments[0].value = arguments[1];", campo_busca, contrato)
                driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", campo_busca)
                driver.execute_script("arguments[0].dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));", campo_busca)
                time.sleep(1)

                try:
                    driver.find_element(By.CLASS_NAME, "buttons-panel").click()
                except:
                    pass

                time.sleep(2)

                if not elemento_existe(driver, By.CLASS_NAME, "oj-collapsible-wrapper"):
                    df.at[index, "Resultado"] = "Contrato não localizado"
                    continue

                atividades = driver.find_elements(By.CLASS_NAME, "found-item-activity")
                if not atividades:
                    df.at[index, "Resultado"] = "Contrato não localizado"
                    continue

                atividades[0].find_element(By.TAG_NAME, "div").click()
                time.sleep(2)

                if elemento_existe(driver, By.XPATH, "//span[text()='iniciado']"):
                    df.at[index, "Resultado"] = "Contrato já iniciado"
                    continue

                if elemento_existe(driver, By.XPATH, "//*[contains(text(),'cancelado')]"):
                    df.at[index, "Resultado"] = "Contrato cancelado"
                    continue

                mover_tecnico(driver, tecnico_destino)
                df.at[index, "Resultado"] = "Troca realizada com sucesso"

            except Exception as e:
                df.at[index, "Resultado"] = f"Erro: {str(e)}"

            progresso_bar.set((index + 1) / total)

    except Exception as erro_geral:
        messagebox.showerror("Erro geral", f"Ocorreu um erro:\n{str(erro_geral)}")
    finally:
        driver.quit()
        df.to_excel("resultado_troca.xlsx", index=False)
        messagebox.showinfo("Fim", "Processo concluído! Resultado salvo como 'resultado_troca.xlsx'")
        try:
            progresso_win.destroy()
        except:
            pass
        processo_em_execucao["ativo"] = False

def escolher_aba_e_processar(filepath):
    try:
        xls = pd.ExcelFile(filepath, engine="openpyxl")
        abas = xls.sheet_names
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler abas da planilha:\n{str(e)}")
        return

    janela_abas = ctk.CTkToplevel()
    janela_abas.attributes("-topmost", True)
    janela_abas.title("Escolher Aba")
    janela_abas.geometry("300x200")
    janela_abas.configure(fg_color="#1a1a1a")

    ctk.CTkLabel(janela_abas, text="Selecione a aba:", font=ctk.CTkFont(size=14)).pack(pady=10)

    aba_var = ctk.StringVar(value=abas[0])
    aba_menu = ctk.CTkOptionMenu(janela_abas, values=abas, variable=aba_var,
                                 fg_color="#d9534f", button_color="#d9534f", dropdown_fg_color="#2a2a2a")
    aba_menu.pack(pady=10)

    def confirmar_aba():
        janela_abas.destroy()
        processo_em_execucao["ativo"] = True
        threading.Thread(target=processar_planilha, args=(filepath, aba_var.get()), daemon=True).start()

    ctk.CTkButton(janela_abas, text="Confirmar", command=confirmar_aba,
                  fg_color="#d9534f", hover_color="#c9302c").pack(pady=10)


def criar_interface():
    def definir_login(nome):
        if nome == "Escolha um login":
            login_selecionado["nome"] = None
            login_selecionado["login"] = None
            login_selecionado["senha"] = None
            login_label.configure(text="")
        else:
            login_selecionado["nome"] = nome
            login_selecionado["login"], login_selecionado["senha"] = logins_disponiveis[nome]
            login_label.configure(text=f"Login: {login_selecionado['login']}")

    def escolher_arquivo_thread():
        if processo_em_execucao["ativo"]:
            messagebox.showwarning("Aviso", "Um processo já está em execução.")
            return
        filepath = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xls *.xlsx *.xlsm")])
        if not filepath:
            return
        if not login_selecionado["login"]:
            messagebox.showerror("Erro", "Você precisa selecionar um login.")
            return
        processo_em_execucao["ativo"] = True
        escolher_aba_e_processar(filepath)

    def abrir_gerenciador_logins():
        janela = ctk.CTkToplevel()
        janela.attributes("-topmost", True)
        janela.title("Gerenciar Logins")
        janela.geometry("400x420")
        janela.configure(fg_color="#1a1a1a")
        lista_usuarios = ctk.CTkComboBox(janela, values=list(logins_disponiveis.keys()),
                                         fg_color="#ff2c25", button_color="#fa221b", dropdown_fg_color="#2a2a2a")
        lista_usuarios.pack(pady=10)

        entry_nome = ctk.CTkEntry(janela, placeholder_text="Nome completo", fg_color="#2a2a2a")
        entry_nome.pack(pady=8)

        entry_login = ctk.CTkEntry(janela, placeholder_text="Login", fg_color="#2a2a2a")
        entry_login.pack(pady=8)

        entry_senha = ctk.CTkEntry(janela, placeholder_text="Senha", show="*", fg_color="#2a2a2a")
        entry_senha.pack(pady=8)

        def carregar_dados_usuario():
            nome = lista_usuarios.get()
            if nome in logins_disponiveis:
                login, senha = logins_disponiveis[nome]
                entry_nome.delete(0, "end")
                entry_login.delete(0, "end")
                entry_senha.delete(0, "end")
                entry_nome.insert(0, nome)
                entry_login.insert(0, login)
                entry_senha.insert(0, senha)

        def salvar_usuario():
            nome = entry_nome.get().strip()
            login = entry_login.get().strip()
            senha = entry_senha.get().strip()
            if nome and login and senha:
                logins_disponiveis[nome] = [login, senha]
                salvar_logins()
                atualizar_option_menu()
                messagebox.showinfo("Sucesso", "Login salvo com sucesso.")
            else:
                messagebox.showerror("Erro", "Preencha todos os campos.")

        def excluir_usuario():
            nome = lista_usuarios.get()
            if nome in logins_disponiveis:
                confirm = messagebox.askyesno("Confirmar", f"Remover {nome}?")
                if confirm:
                    del logins_disponiveis[nome]
                    salvar_logins()
                    atualizar_option_menu()
                    entry_nome.delete(0, "end")
                    entry_login.delete(0, "end")
                    entry_senha.delete(0, "end")

        ctk.CTkButton(janela, text="Carregar", command=carregar_dados_usuario,
                      fg_color="#bd2e29", hover_color="#c9302c").pack(pady=(10, 6))
        ctk.CTkButton(janela, text="Salvar", command=salvar_usuario,
                      fg_color="#4B4B4B", hover_color="#c9302c").pack(pady=6)
        ctk.CTkButton(janela, text="Excluir", command=excluir_usuario,
                      fg_color="#ff0800", hover_color="#c9302c").pack(pady=6)
        ctk.CTkLabel(janela, text="CTB© 2025", font=ctk.CTkFont(size=12), text_color="gray").pack(side="bottom", pady=10)
    def atualizar_option_menu():
        values = ["Escolha um login"] + list(logins_disponiveis.keys())
        option_menu.configure(values=values)
        atual = login_selecionado["nome"]
        if atual and atual in logins_disponiveis:
            option_menu.set(atual)
        else:
            option_menu.set("Escolha um login")

    app = ctk.CTk()
    app.title("Troca de Técnicos - Claro")
    app.geometry("560x400")
    app.resizable(False, False)
    app.configure(fg_color="#1a1a1a")

    ctk.CTkLabel(app, text="Troca de Técnicos", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(20, 10))

    ctk.CTkButton(app, text="Gerenciar Logins", command=abrir_gerenciador_logins,
                  fg_color="#fa221b", hover_color="#c9302c").pack(pady=(0, 10))

    option_menu = ctk.CTkOptionMenu(app, values=["Escolha um login"] + list(logins_disponiveis.keys()),
                                    command=definir_login,
                                    fg_color="#797979", button_color="#797979", button_hover_color="#c9302c",
                                    dropdown_fg_color="#2a2a2a", dropdown_hover_color="#404040")
    option_menu.set("Escolha um login")
    option_menu.pack(pady=10)

    login_label = ctk.CTkLabel(app, text="", font=ctk.CTkFont(size=14))
    login_label.pack(pady=(0, 10))

    ctk.CTkButton(app, text="Selecionar Planilha", command=escolher_arquivo_thread,
                  font=ctk.CTkFont(size=24), fg_color="#fa221b", hover_color="#c9302c").pack(pady=20)

    ctk.CTkLabel(app, text="CTB© 2025", font=ctk.CTkFont(size=12), text_color="gray").pack(side="bottom", pady=10)

    app.mainloop()

if __name__ == "__main__":
    criar_interface()
