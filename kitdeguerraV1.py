import customtkinter as ctk
from tkinter import END
import subprocess
import threading
import os
import platform
import shutil
from PIL import Image

# --- CONFIGURAÇÕES ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

class TechToolApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Kit de Guerra v2.0 - N1")
        self.geometry("1000x700")
        
        # Configuração do Grid Principal
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- 1. BANNER (Topo) ---
        # Aumentei height para 150 e removi o padx/pady para ele encostar nas bordas
        self.banner_frame = ctk.CTkFrame(self, height=150, fg_color="transparent")
        self.banner_frame.grid(row=0, column=0, columnspan=2, sticky="new", padx=0, pady=0)
        
        try:
            # Tenta carregar banner.png
            img_source = Image.open("banner.png")
            
            # --- CONFIGURAÇÃO DE TAMANHO ---
            # Define a altura desejada do banner (Ex: 150 pixels)
            target_height = 150 
            
            # Calcula a nova largura mantendo a proporção (Aspect Ratio)
            # Se a imagem for muito larga, ela vai preencher bem a tela
            aspect_ratio = target_height / float(img_source.size[1])
            target_width = int(float(img_source.size[0]) * float(aspect_ratio))
            
            # Se a imagem for muito estreita, forçamos ela a ter pelo menos a largura da janela (1000)
            # Nota: Isso pode cortar um pouco a imagem se ela não for panorâmica
            if target_width < 1000:
                target_width = 1000

            self.banner_img = ctk.CTkImage(light_image=img_source, dark_image=img_source, size=(target_width, target_height))
            
            self.lbl_banner = ctk.CTkLabel(self.banner_frame, text="", image=self.banner_img)
            self.lbl_banner.pack(fill="both", expand=True) # Manda a imagem ocupar todo o espaço do frame
            
        except Exception as e:
            # Fallback caso não tenha imagem
            print(f"Erro no banner: {e}")
            self.banner_frame.configure(fg_color="#1f538d") # Cor de fundo azul se falhar imagem
            self.lbl_banner = ctk.CTkLabel(self.banner_frame, text="VeloTrack IT - Dashboard", font=("Impact", 36), text_color="white")
            self.lbl_banner.pack(pady=40)

        # --- 2. MENU LATERAL (Esquerda) ---
        self.sidebar = ctk.CTkScrollableFrame(self, width=220, label_text="Comandos")
        self.sidebar.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        # --- 3. ÁREA DE LOG E STATUS (Direita) ---
        self.right_frame = ctk.CTkFrame(self)
        self.right_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=5)
        
        # Dashboard Header
        self.lbl_log = ctk.CTkLabel(self.right_frame, text="Monitoramento em Tempo Real", font=("Arial", 16, "bold"))
        self.lbl_log.pack(pady=5, anchor="w", padx=10)

        # Caixa de Log (Com tags de cor)
        self.log_box = ctk.CTkTextbox(self.right_frame, font=("Consolas", 12))
        self.log_box.pack(expand=True, fill="both", padx=10, pady=5)
        
        # Configurando Cores (Hack para acessar o Widget Tkinter por baixo do CustomTkinter)
        self.log_box._textbox.tag_config("SUCCESS", foreground="#00E676") # Verde Matrix
        self.log_box._textbox.tag_config("ERROR", foreground="#FF5252")   # Vermelho Alerta
        self.log_box._textbox.tag_config("INFO", foreground="#40C4FF")    # Azul Info
        self.log_box._textbox.tag_config("WARNING", foreground="#FFD740") # Amarelo Aviso

        # --- 4. BARRA DE DESCRIÇÃO (Rodapé) ---
        self.desc_frame = ctk.CTkFrame(self, height=40, fg_color="#2b2b2b")
        self.desc_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        
        self.lbl_desc = ctk.CTkLabel(self.desc_frame, text="Passe o mouse sobre um botão para ver a ajuda...", text_color="#aaaaaa", font=("Arial", 12, "italic"))
        self.lbl_desc.pack(pady=5)

        # --- CRIAÇÃO DOS BOTÕES ---
        self.create_buttons()

    def create_buttons(self):
        # Grupo: Diagnóstico
        self.add_section("🔍 DIAGNÓSTICO")
        self.add_btn("Ping Teste (Google)", self.cmd_ping, 
                     "Testa conectividade com a internet (Google DNS). Verifica perda de pacotes e latência.")
        self.add_btn("Verificar IP/Rede", self.cmd_ipconfig, 
                     "Exibe IP, Gateway e DNS atuais. Útil para ver se o PC pegou IP correto.")
        self.add_btn("Info do Sistema (Serial)", self.cmd_sysinfo, 
                     "Busca Hostname e Serial Number (Service Tag) para abertura de chamados na Dell/HP/Lenovo.")

        # Grupo: Reparo Rápido
        self.add_section("🛠️ REPARO RÁPIDO")
        self.add_btn("Limpar Cache DNS", lambda: self.run_thread("ipconfig /flushdns", "DNS Limpo com sucesso", check_error=False), 
                     "Resolve problemas de sites que não carregam ou erro de 'Página não encontrada'.")
        self.add_btn("Limpar Temp (%TEMP%)", self.cmd_clean_temp, 
                     "Remove arquivos temporários do usuário. Resolve lentidão e erros de instalação.")
        self.add_btn("Atualizar Políticas (GPO)", lambda: self.run_thread("gpupdate /force", "Políticas atualizadas"), 
                     "Força o Windows a baixar as regras mais recentes da empresa (mapeamentos, permissões).")

        # Grupo: Aplicativos
        self.add_section("💻 APLICATIVOS")
        self.add_btn("Reparar/Abrir Outlook", self.cmd_open_outlook, 
                     "Tenta forçar a abertura do Outlook. Use se o ícone não responder.")
        self.add_btn("Reiniciar OneDrive", self.cmd_restart_onedrive, 
                     "Fecha forçado e reabre o OneDrive. Resolve erros de sincronização (X vermelho).")
        self.add_btn("Gerenciar Credenciais", lambda: self.run_thread("control /name Microsoft.CredentialManager", "Gerenciador Aberto"), 
                     "Abre o cofre de senhas. Use para apagar senhas antigas do Outlook/Teams.")

        # Grupo: Impressão
        self.add_section("🖨️ IMPRESSÃO")
        self.add_btn("Fila de Impressão", lambda: self.run_thread("rundll32 printui.dll,PrintUIEntry /v", "Fila aberta"), 
                     "Abre a fila de impressão para cancelar documentos travados.")
        self.add_btn("Instalar Impressora", lambda: self.run_thread("rundll32 printui.dll,PrintUIEntry /il", "Wizard Aberto"), 
                     "Abre o assistente nativo do Windows para adicionar impressora de rede.")

        # Grupo: Info
        self.add_section("ℹ️ SOBRE")
        self.add_btn("Créditos & Versão", self.cmd_about, 
                     "Exibe informações do desenvolvedor e dedicatória à equipe.")

    def add_section(self, text):
        lbl = ctk.CTkLabel(self.sidebar, text=text, font=("Arial", 12, "bold"), anchor="w", text_color="#888")
        lbl.pack(pady=(15, 2), padx=5, fill="x")

    def add_btn(self, text, command, description):
        btn = ctk.CTkButton(self.sidebar, text=text, command=command, height=32, anchor="w", fg_color="#1f538d", hover_color="#14375e")
        btn.pack(pady=2, padx=5, fill="x")
        
        # Eventos de Hover (Mouse em cima/fora)
        btn.bind("<Enter>", lambda e: self.lbl_desc.configure(text=f"ℹ️ {description}", text_color="white"))
        btn.bind("<Leave>", lambda e: self.lbl_desc.configure(text="...", text_color="#aaaaaa"))

    # --- LÓGICA DE LOG E ANÁLISE ---

    def write_log(self, text, tag="INFO"):
        self.log_box.configure(state="normal")
        self.log_box.insert(END, f"\n[{tag}] ", tag) # Insere a tag colorida
        self.log_box.insert(END, f"{text}")          # Insere o texto normal
        self.log_box.see(END)
        self.log_box.configure(state="disabled")

    def run_thread(self, command, success_msg="", check_error=True):
        def task():
            self.write_log(f"Executando: {command}...", "INFO")
            try:
                result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='cp850', errors='ignore')
                output = result.stdout + result.stderr
                
                # Análise simples de erro no texto de saída
                if check_error and ("erro" in output.lower() or "falha" in output.lower() or "access denied" in output.lower()):
                     self.write_log(f"Comando finalizado com alertas:\n{output}", "WARNING")
                else:
                    if success_msg: self.write_log(success_msg, "SUCCESS")
                    else: self.write_log(output.strip(), "INFO")
            except Exception as e:
                self.write_log(f"Erro Crítico: {str(e)}", "ERROR")
        threading.Thread(target=task).start()

    # --- FUNÇÕES ESPECÍFICAS (INTELIGENTES) ---

    def cmd_ping(self):
        def task():
            self.write_log("Testando conectividade (Google IPv4)...", "INFO")
            try:
                if platform.system() == "Windows":
                    # Adicionei o '-4' para forçar IPv4. Evita endereços longos que quebram linha.
                    # Adicionei criação de startupinfo para garantir que não pisque janelas
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    
                    cmd = "ping -n 2 -4 google.com"
                    
                    result = subprocess.run(cmd, startupinfo=startupinfo, capture_output=True, text=True, encoding='cp850', errors='ignore')
                    output = result.stdout
                    
                    # --- CORREÇÃO AQUI ---
                    # Agora procuramos por "Perdidos = 0" (PT-BR) ou "Lost = 0" (EN-US)
                    # Isso evita o erro de quebra de linha
                    if "Perdidos = 0" in output or "Lost = 0" in output:
                        # Tenta extrair o tempo médio usando Regex simples para ficar bonito
                        import re
                        match = re.search(r"(Média|Average) = (\d+ms)", output)
                        time_ms = match.group(2) if match else "<10ms"
                        
                        self.write_log(f"✅ CONEXÃO ESTÁVEL. Latência: {time_ms}", "SUCCESS")
                    else:
                        self.write_log("❌ FALHA DE CONEXÃO. Houve perda de pacotes.", "ERROR")
                        self.write_log(output, "WARNING") # Mostra o log completo para análise
                else:
                    self.write_log("OS não compatível.", "ERROR")
            except Exception as e:
                self.write_log(f"Erro Crítico no Ping: {str(e)}", "ERROR")
        threading.Thread(target=task).start()

    def cmd_clean_temp(self):
        def task():
            self.write_log("Iniciando limpeza segura do %TEMP%...", "INFO")
            temp_path = os.getenv('TEMP')
            deleted = 0
            errors = 0
            
            for root, dirs, files in os.walk(temp_path):
                for name in files:
                    try:
                        file_path = os.path.join(root, name)
                        os.remove(file_path)
                        deleted += 1
                    except:
                        errors += 1
            
            self.write_log(f"✅ LIMPEZA CONCLUÍDA.", "SUCCESS")
            self.write_log(f"Arquivos removidos: {deleted}", "INFO")
            self.write_log(f"Arquivos em uso (ignorados): {errors}", "WARNING")
            
        threading.Thread(target=task).start()

    def cmd_ipconfig(self):
        def task():
            self.write_log("Verificando Rede...", "INFO")
            output = subprocess.check_output("ipconfig", shell=True, encoding='cp850')
            # Filtra apenas linhas importantes para não poluir
            relevant = [line for line in output.split('\n') if "IPv4" in line or "Gateway" in line or "Adaptador" in line]
            formatted = "\n".join(relevant)
            self.write_log(formatted, "INFO")
            
            if "169.254" in output:
                 self.write_log("⚠️ ALERTA: IP APIPA detectado (169.254.x.x). O PC não está pegando IP do DHCP.", "ERROR")
        threading.Thread(target=task).start()

    def cmd_sysinfo(self):
        def task():
            self.write_log("Buscando Serial Number...", "INFO")
            try:
                # Comando PowerShell mais robusto que WMIC antigo
                cmd = "powershell \"Get-CimInstance -ClassName Win32_BIOS | Select-Object -ExpandProperty SerialNumber\""
                serial = subprocess.check_output(cmd, shell=True, text=True).strip()
                hostname = platform.node()
                
                self.write_log(f"Hostname: {hostname}", "SUCCESS")
                self.write_log(f"Serial Number (Service Tag): {serial}", "SUCCESS")
            except Exception as e:
                self.write_log(f"Não foi possível ler o serial: {e}", "ERROR")
        threading.Thread(target=task).start()

    def cmd_open_outlook(self):
        def task():
            self.write_log("Procurando Outlook no sistema...", "INFO")
            # Lista de caminhos comuns do Office
            paths = [
                r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
                r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
                r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
                r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
            ]
            found = False
            for p in paths:
                if os.path.exists(p):
                    self.write_log(f"Outlook encontrado em: {p}", "SUCCESS")
                    subprocess.Popen([p]) # Abre sem travar
                    found = True
                    break
            
            if not found:
                # Tenta chamar pelo comando executar padrão
                try:
                    subprocess.Popen(["start", "outlook"], shell=True)
                    self.write_log("Comando 'start outlook' enviado.", "INFO")
                except:
                    self.write_log("❌ Outlook não encontrado nos caminhos padrão.", "ERROR")
        threading.Thread(target=task).start()

    def cmd_restart_onedrive(self):
        def task():
            self.write_log("Encerrando OneDrive...", "WARNING")
            subprocess.run("taskkill /f /im OneDrive.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            import time
            time.sleep(2)
            
            # Caminho padrão do OneDrive do usuário
            user_onedrive = os.path.join(os.getenv('LOCALAPPDATA'), r"Microsoft\OneDrive\OneDrive.exe")
            
            if os.path.exists(user_onedrive):
                self.write_log("Iniciando OneDrive...", "INFO")
                subprocess.Popen(user_onedrive)
                self.write_log("✅ OneDrive reiniciado com sucesso.", "SUCCESS")
            else:
                self.write_log(f"❌ Executável não encontrado em: {user_onedrive}", "ERROR")
        threading.Thread(target=task).start()

    def cmd_about(self):
        # Cria uma janela flutuante (Toplevel)
        about_window = ctk.CTkToplevel(self)
        about_window.title("Sobre")
        about_window.geometry("400x350")
        about_window.resizable(False, False)
        
        # Garante que a janela fique no topo
        about_window.attributes("-topmost", True)

        # Título do App
        lbl_title = ctk.CTkLabel(about_window, text="Kit de Guerra", font=("Impact", 24), text_color="#1f538d")
        lbl_title.pack(pady=(20, 5))
        
        lbl_ver = ctk.CTkLabel(about_window, text="Versão 2.0.1 - Build 2025", font=("Arial", 12))
        lbl_ver.pack(pady=0)

        # Divisória
        line = ctk.CTkFrame(about_window, height=2, width=300, fg_color="#555")
        line.pack(pady=15)

        # Créditos do Desenvolvedor
        lbl_dev_title = ctk.CTkLabel(about_window, text="Desenvolvido por:", font=("Arial", 12, "bold"))
        lbl_dev_title.pack()
        
        lbl_dev_name = ctk.CTkLabel(about_window, text="Gleison Andrade dos Santos", font=("Arial", 14, "bold"), text_color="#40C4FF")
        lbl_dev_name.pack(pady=2)

        # Dedicatória (Caixa de Texto para caber bastante coisa se quiser)
        dedication_text = (
            "Esta ferramenta foi desenvolvida pensando na agilidade \n"
            "e eficiência do nosso time de T.I.\n\n"
            "Dedicado a todos os meus colegas de trabalho que \n"
            "batalham diariamente para manter tudo funcionando.\n"
            "Juntos somos mais fortes! 🚀"
        )
        
        lbl_dedication = ctk.CTkLabel(about_window, text=dedication_text, font=("Arial", 11, "italic"), text_color="#aaaaaa", justify="center")
        lbl_dedication.pack(pady=20)

        # Botão Fechar
        btn_close = ctk.CTkButton(about_window, text="Fechar", command=about_window.destroy, width=100, fg_color="#333", hover_color="#444")
        btn_close.pack(pady=10)

if __name__ == "__main__":
    app = TechToolApp()
    app.mainloop()