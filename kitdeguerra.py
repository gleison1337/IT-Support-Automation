import customtkinter as ctk
from tkinter import END
import subprocess
import threading
import os
import platform
import time
from PIL import Image

# --- CONFIGURAÇÕES GERAIS ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

class TechToolApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Kit de Guerra v3.0 - Enterprise Design")
        self.geometry("1100x750")
        
        # Configuração do Grid Principal (1x1 para sobreposição de telas)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- DADOS E ESTRUTURA DOS MENUS ---
        # Aqui organizamos o que cada categoria tem. 
        # Fica fácil adicionar coisas novas sem mexer no layout.
        self.menus = {
            "Diagnóstico": [
                {"name": "Ping Teste (Google)", "cmd": self.cmd_ping, "desc": "Testa conectividade e latência IPv4."},
                {"name": "Verificar IP/Rede", "cmd": self.cmd_ipconfig, "desc": "Exibe IP, Gateway e alerta de APIPA (169.254)."},
                {"name": "Info do Sistema", "cmd": self.cmd_sysinfo, "desc": "Captura Hostname e Serial Number (Dell/HP/Lenovo)."}
            ],
            "Reparo Rápido": [
                {"name": "Limpar Cache DNS", "cmd": lambda: self.run_thread("ipconfig /flushdns", "DNS Limpo", check_error=False), "desc": "Resolve falhas de resolução de nomes."},
                {"name": "Limpar Temp (%TEMP%)", "cmd": self.cmd_clean_temp, "desc": "Limpa arquivos temporários do usuário atual."},
                {"name": "Atualizar Políticas", "cmd": lambda: self.run_thread("gpupdate /force", "Políticas atualizadas"), "desc": "Força atualização de GPO."}
            ],
            "Aplicativos": [
                {"name": "Reparar Outlook", "cmd": self.cmd_open_outlook, "desc": "Tenta abrir Outlook ou forçar modo de segurança."},
                {"name": "Reiniciar OneDrive", "cmd": self.cmd_restart_onedrive, "desc": "Mata o processo e reinicia o executável local."},
                {"name": "Credenciais", "cmd": lambda: self.run_thread("control /name Microsoft.CredentialManager", "Cofre Aberto"), "desc": "Abre gerenciador de senhas do Windows."}
            ],
            "Impressão": [
                {"name": "Fila de Impressão", "cmd": lambda: self.run_thread("rundll32 printui.dll,PrintUIEntry /v", "Fila aberta"), "desc": "Abre fila para cancelar jobs travados."},
                {"name": "Wizard Instalação", "cmd": lambda: self.run_thread("rundll32 printui.dll,PrintUIEntry /il", "Wizard Aberto"), "desc": "Assistente nativo para adicionar impressora IP."}
            ]
        }

        # Variável para controlar qual categoria está ativa
        self.current_category_items = []

        # --- INICIALIZAÇÃO DAS TELAS ---
        self.home_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.dashboard_frame = ctk.CTkFrame(self, fg_color="transparent")

        # Constrói as telas (mas só mostra a Home)
        self.build_home_screen()
        self.build_dashboard_screen()
        
        # Mostra a Home inicialmente
        self.show_home()

    # =========================================================================
    #                           CONSTRUÇÃO DA HOME
    # =========================================================================
    def build_home_screen(self):
        # 1. Banner Grande
        self.banner_container = ctk.CTkFrame(self.home_frame, height=200, fg_color="transparent")
        self.banner_container.pack(fill="x", pady=(20, 10))
        
        try:
            img_source = Image.open("banner.png")
            # Lógica de redimensionamento para banner grande
            target_height = 200
            aspect_ratio = target_height / float(img_source.size[1])
            target_width = int(float(img_source.size[0]) * float(aspect_ratio))
            if target_width < 800: target_width = 800
            
            self.banner_img = ctk.CTkImage(light_image=img_source, dark_image=img_source, size=(target_width, target_height))
            lbl_banner = ctk.CTkLabel(self.banner_container, text="", image=self.banner_img)
            lbl_banner.pack()
        except:
            lbl_banner = ctk.CTkLabel(self.banner_container, text="KIT DE GUERRA", font=("Impact", 60), text_color="white")
            lbl_banner.pack(pady=40)

        # 2. Informações / Apresentação
        info_frame = ctk.CTkFrame(self.home_frame, fg_color="transparent")
        info_frame.pack(fill="x", padx=50, pady=10)
        
        lbl_welcome = ctk.CTkLabel(info_frame, text="Bem-vindo ao Painel de Controle N1", font=("Arial", 22, "bold"))
        lbl_welcome.pack()
        lbl_sub = ctk.CTkLabel(info_frame, text="Selecione uma categoria abaixo para iniciar o atendimento", font=("Arial", 14), text_color="#aaaaaa")
        lbl_sub.pack()

        # 3. Grade de Botões Gigantes (Categorias)
        menu_grid = ctk.CTkFrame(self.home_frame, fg_color="transparent")
        menu_grid.pack(expand=True, fill="both", padx=50, pady=30)
        
        # Define grid 2x2
        menu_grid.grid_columnconfigure(0, weight=1)
        menu_grid.grid_columnconfigure(1, weight=1)

        # Botão Diagnóstico
        self.create_big_btn(menu_grid, "🔍 DIAGNÓSTICO", "Ping, IP, Sistema", 0, 0, lambda: self.go_to_category("Diagnóstico"))
        # Botão Reparo
        self.create_big_btn(menu_grid, "🛠️ REPARO RÁPIDO", "Cache, Temp, GPO", 0, 1, lambda: self.go_to_category("Reparo Rápido"))
        # Botão Apps
        self.create_big_btn(menu_grid, "💻 APLICATIVOS", "Outlook, OneDrive", 1, 0, lambda: self.go_to_category("Aplicativos"))
        # Botão Impressão
        self.create_big_btn(menu_grid, "🖨️ IMPRESSÃO", "Fila, Spooler", 1, 1, lambda: self.go_to_category("Impressão"))
        
        # Botão Sobre (Pequeno embaixo)
        btn_about = ctk.CTkButton(self.home_frame, text="ℹ️ Sobre & Créditos", command=self.cmd_about, fg_color="transparent", border_width=1, text_color="#aaa")
        btn_about.pack(side="bottom", pady=20)

    def create_big_btn(self, parent, title, subtitle, row, col, command):
        # Frame clicável (gambiarra visual para botão grande com subtítulo)
        btn = ctk.CTkButton(parent, text=f"{title}\n\n{subtitle}", command=command, 
                            font=("Arial", 16, "bold"), height=100, corner_radius=15,
                            fg_color="#1f538d", hover_color="#14375e")
        btn.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")

    # =========================================================================
    #                           CONSTRUÇÃO DO DASHBOARD
    # =========================================================================
    def build_dashboard_screen(self):
        self.dashboard_frame.grid_columnconfigure(1, weight=1)
        self.dashboard_frame.grid_rowconfigure(0, weight=1)

        # --- SIDEBAR (Esquerda) ---
        self.sidebar = ctk.CTkFrame(self.dashboard_frame, width=250, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        # Botão VOLTAR (Topo da Sidebar)
        self.btn_back = ctk.CTkButton(self.sidebar, text="⬅ VOLTAR AO MENU", command=self.show_home, 
                                      fg_color="#333", hover_color="#444", height=40)
        self.btn_back.pack(pady=(20, 10), padx=10, fill="x")
        
        # Título da Categoria Atual
        self.lbl_cat_title = ctk.CTkLabel(self.sidebar, text="CATEGORIA", font=("Impact", 20), text_color="#40C4FF")
        self.lbl_cat_title.pack(pady=10)

        # Botão Mágico: EXECUTE ALL
        self.btn_exec_all = ctk.CTkButton(self.sidebar, text="⚡ EXECUTAR TUDO", command=self.run_all_in_category,
                                          fg_color="#00C853", hover_color="#009624", height=40, font=("Arial", 12, "bold"))
        self.btn_exec_all.pack(pady=(0, 20), padx=10, fill="x")

        # Container para os botões dinâmicos
        self.sidebar_buttons_frame = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent")
        self.sidebar_buttons_frame.pack(expand=True, fill="both")

        # --- ÁREA DE LOG (Direita) ---
        right_panel = ctk.CTkFrame(self.dashboard_frame, fg_color="transparent")
        right_panel.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        
        ctk.CTkLabel(right_panel, text="TERMINAL DE MONITORAMENTO", font=("Arial", 14, "bold"), anchor="w").pack(fill="x")
        
        self.log_box = ctk.CTkTextbox(right_panel, font=("Consolas", 12), state="disabled")
        self.log_box.pack(expand=True, fill="both", pady=5)
        
        # Cores do Log
        self.log_box._textbox.tag_config("SUCCESS", foreground="#00E676")
        self.log_box._textbox.tag_config("ERROR", foreground="#FF5252")
        self.log_box._textbox.tag_config("INFO", foreground="#40C4FF")
        self.log_box._textbox.tag_config("WARNING", foreground="#FFD740")
        
        # Rodapé de Ajuda no Dashboard
        self.dash_desc = ctk.CTkLabel(right_panel, text="Selecione uma função à esquerda...", text_color="#888", font=("Arial", 12, "italic"))
        self.dash_desc.pack(fill="x")

    # =========================================================================
    #                           NAVEGAÇÃO
    # =========================================================================
    def show_home(self):
        self.dashboard_frame.grid_forget()
        self.home_frame.grid(row=0, column=0, sticky="nsew")

    def go_to_category(self, category_name):
        self.home_frame.grid_forget()
        self.dashboard_frame.grid(row=0, column=0, sticky="nsew")
        
        # Atualiza Título
        self.lbl_cat_title.configure(text=category_name.upper())
        
        # Limpa botões antigos
        for widget in self.sidebar_buttons_frame.winfo_children():
            widget.destroy()
            
        # Gera novos botões baseado no dicionário self.menus
        items = self.menus.get(category_name, [])
        self.current_category_items = items # Salva para o "Executar Tudo" usar depois
        
        for item in items:
            btn = ctk.CTkButton(self.sidebar_buttons_frame, text=item['name'], command=item['cmd'], 
                                height=35, anchor="w", fg_color="#1f538d", hover_color="#14375e")
            btn.pack(pady=2, padx=5, fill="x")
            
            # Hover Effect para Descrição
            desc = item['desc']
            btn.bind("<Enter>", lambda e, d=desc: self.dash_desc.configure(text=f"ℹ️ {d}", text_color="white"))
            btn.bind("<Leave>", lambda e: self.dash_desc.configure(text="...", text_color="#888"))

        # Log inicial
        self.write_log(f"--- Entrando em {category_name} ---", "INFO")

    # =========================================================================
    #                           LÓGICA "EXECUTAR TUDO"
    # =========================================================================
    def run_all_in_category(self):
        def task_sequence():
            self.write_log("⚡ INICIANDO EXECUÇÃO EM LOTE...", "WARNING")
            self.write_log("Por favor, aguarde o fim de cada processo.", "WARNING")
            
            total = len(self.current_category_items)
            for i, item in enumerate(self.current_category_items):
                # Aviso visual
                self.write_log(f"[{i+1}/{total}] Executando: {item['name']}...", "INFO")
                
                # Executa a função associada ao botão
                # Nota: Como as funções originais já rodam em thread, aqui chamamos direto.
                # Se elas não bloqueiam, o delay abaixo ajuda a organizar visualmente.
                item['cmd']() 
                
                # Delay estético para não encavalar logs
                time.sleep(2.5) 
            
            self.write_log("🏁 EXECUÇÃO EM LOTE FINALIZADA.", "SUCCESS")

        threading.Thread(target=task_sequence).start()

    # =========================================================================
    #                           LÓGICA DE SISTEMA (MANTIDA)
    # =========================================================================
    def write_log(self, text, tag="INFO"):
        self.log_box.configure(state="normal")
        self.log_box.insert(END, f"\n[{tag}] ", tag)
        self.log_box.insert(END, f"{text}")
        self.log_box.see(END)
        self.log_box.configure(state="disabled")

    def run_thread(self, command, success_msg="", check_error=True):
        def task():
            self.write_log(f"Executando: {command}...", "INFO")
            try:
                result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='cp850', errors='ignore')
                output = result.stdout + result.stderr
                if check_error and ("erro" in output.lower() or "falha" in output.lower() or "access denied" in output.lower()):
                     self.write_log(f"Alerta:\n{output}", "WARNING")
                else:
                    if success_msg: self.write_log(success_msg, "SUCCESS")
                    else: self.write_log(output.strip(), "INFO")
            except Exception as e:
                self.write_log(f"Erro: {str(e)}", "ERROR")
        threading.Thread(target=task).start()

    # --- FUNÇÕES NATIVAS (Mantidas iguais à V2) ---
    def cmd_ping(self):
        def task():
            self.write_log("Testando conectividade (Google IPv4)...", "INFO")
            try:
                if platform.system() == "Windows":
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    cmd = "ping -n 2 -4 google.com"
                    result = subprocess.run(cmd, startupinfo=startupinfo, capture_output=True, text=True, encoding='cp850', errors='ignore')
                    output = result.stdout
                    
                    if "Perdidos = 0" in output or "Lost = 0" in output:
                        import re
                        match = re.search(r"(Média|Average) = (\d+ms)", output)
                        time_ms = match.group(2) if match else "<10ms"
                        self.write_log(f"✅ CONEXÃO ESTÁVEL. Latência: {time_ms}", "SUCCESS")
                    else:
                        self.write_log("❌ FALHA DE CONEXÃO.", "ERROR")
            except Exception as e:
                self.write_log(f"Erro: {str(e)}", "ERROR")
        threading.Thread(target=task).start()

    def cmd_clean_temp(self):
        def task():
            self.write_log("Limpando %TEMP%...", "INFO")
            temp_path = os.getenv('TEMP')
            deleted = 0
            for root, dirs, files in os.walk(temp_path):
                for name in files:
                    try:
                        os.remove(os.path.join(root, name))
                        deleted += 1
                    except: pass
            self.write_log(f"✅ LIMPEZA CONCLUÍDA. Arquivos removidos: {deleted}", "SUCCESS")
        threading.Thread(target=task).start()

    def cmd_ipconfig(self):
        def task():
            self.write_log("Verificando Rede...", "INFO")
            output = subprocess.check_output("ipconfig", shell=True, encoding='cp850')
            relevant = [line for line in output.split('\n') if "IPv4" in line or "Gateway" in line or "Adaptador" in line]
            self.write_log("\n".join(relevant), "INFO")
            if "169.254" in output: self.write_log("⚠️ ALERTA: IP APIPA (169.254) DETECTADO.", "ERROR")
        threading.Thread(target=task).start()

    def cmd_sysinfo(self):
        def task():
            self.write_log("Buscando Serial...", "INFO")
            try:
                cmd = "powershell \"Get-CimInstance -ClassName Win32_BIOS | Select-Object -ExpandProperty SerialNumber\""
                serial = subprocess.check_output(cmd, shell=True, text=True).strip()
                self.write_log(f"Hostname: {platform.node()}", "SUCCESS")
                self.write_log(f"Serial: {serial}", "SUCCESS")
            except: self.write_log("Falha ao ler serial.", "ERROR")
        threading.Thread(target=task).start()

    def cmd_open_outlook(self):
        def task():
            self.write_log("Buscando Outlook...", "INFO")
            subprocess.Popen(["start", "outlook"], shell=True) # Simplificado para brevidade
            self.write_log("Comando enviado.", "INFO")
        threading.Thread(target=task).start()

    def cmd_restart_onedrive(self):
        def task():
            self.write_log("Reiniciando OneDrive...", "WARNING")
            subprocess.run("taskkill /f /im OneDrive.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(1)
            user_onedrive = os.path.join(os.getenv('LOCALAPPDATA'), r"Microsoft\OneDrive\OneDrive.exe")
            if os.path.exists(user_onedrive):
                subprocess.Popen(user_onedrive)
                self.write_log("✅ Reiniciado.", "SUCCESS")
            else: self.write_log("Executável não achado.", "ERROR")
        threading.Thread(target=task).start()

    def cmd_about(self):
        about = ctk.CTkToplevel(self)
        about.title("Sobre")
        about.geometry("400x300")
        about.attributes("-topmost", True)
        ctk.CTkLabel(about, text="Kit de Guerra v3.0", font=("Impact", 20)).pack(pady=20)
        ctk.CTkLabel(about, text="Dev: Gleison Andrade dos Santos", text_color="#40C4FF").pack()
        ctk.CTkLabel(about, text="Dedicado ao time de T.I.", font=("Arial", 10, "italic")).pack(pady=20)

if __name__ == "__main__":
    app = TechToolApp()
    app.mainloop()