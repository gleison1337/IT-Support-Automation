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

        # --- MAPA DE SERVIDORES DE IMPRESSÃO ---
        self.print_servers = {
            "Selecione uma localidade...": "",
            "Araraquara": r"\\ARapp01",
            "Barra Bonita": r"\\BAAPP01",
            "Benalcool": r"\\BEapp01",
            "Bloco 5": r"\\B5ads01",
            "Bom Retiro": r"\\BRapp01",
            "Bonfim": r"\\BOapp01",
            "Caraapó": r"\\CAapp01",
            "Comgas JK": r"\\FGPRT02V",
            "Comgas Santos": r"\\FGPRP01V",
            "Continental": r"\\coapp01",
            "Costa Pinto (CP03)": r"\\CP03",
            "Costa Pinto (CP05)": r"\\CP05",
            "Costa Pinto (RZPRT)": r"\\RZPRT01V",
            "CSC (CAR) - 01": r"\\CSads01",
            "CSC (CAR) - 02": r"\\CSads02",
            "Destivale": r"\\DVapp01",
            "Diamante": r"\\Diapp01",
            "Dois Córregos": r"\\DCapp01",
            "Fábrica de Lubrificantes": r"\\slads02",
            "Faria Lima": r"\\spads01",
            "Figueira": r"\\FGPRT01V",
            "Gasa": r"\\GAapp01",
            "Ilha do Governador": r"\\SLAds02",
            "Ipaussu": r"\\IPapp01",
            "Jatai": r"\\JAapp01",
            "Junqueira": r"\\JUapp01",
            "Lagoa da Prata": r"\\lpapp01",
            "Leme": r"\\leapp01",
            "Maracai": r"\\MAapp01",
            "Moove": r"\\SLAPP01",
            "Mundial": r"\\MUapp01",
            "OFFICE": r"\\ANapp01",
            "Paraguaçu": r"\\PRapp01",
            "Passa Tempo": r"\\ptapp01",
            "Paulinia": r"\\PLapp01",
            "Portuária Santos": r"\\POapp01",
            "Rafard": r"\\RAapp01",
            "Rio Brilhante": r"\\rbapp01",
            "Santa Elisa": r"\\stapp01",
            "Santa Helena": r"\\SHapp01",
            "São Francisco": r"\\SFapp01",
            "São Jose dos Campos COMGAS": r"\\sj_base_r01",
            "São Paulo - JK (01)": r"\\SPads01",
            "São Paulo - JK (02)": r"\\SPads02",
            "Serra": r"\\SEapp01",
            "Solutec": r"\\slads02",
            "Tamoio": r"\\TAapp01",
            "Taruma": r"\\TMapp01",
            "UMB/MB": r"\\lpapp01",
            "Univalem": r"\\UNapp01",
            "Vale do Rosário": r"\\vrapp01",
            "Victor Civita": r"\\VCapp01"
        }

        self.title("Kit de Guerra v3.0 - Enterprise Design")
        self.geometry("1100x750")
        
        # Configuração do Grid Principal (1x1 para sobreposição de telas)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- DADOS E ESTRUTURA DOS MENUS ---
        
        self.menus = {
            "Diagnóstico": [
                {"name": "🌐 Ping Teste (Google)", "cmd": self.cmd_ping, "desc": "Testa conectividade e latência IPv4."},
                {"name": "🔌 Verificar IP/Rede", "cmd": self.cmd_ipconfig, "desc": "Exibe IP, Gateway e alerta de APIPA (169.254)."},
                {"name": "💻 Info do Sistema", "cmd": self.cmd_sysinfo, "desc": "Captura Hostname e Serial Number (Dell/HP/Lenovo)."}
            ],
            "Reparo Rápido": [
                {"name": "📂 Mapear Rede (Manual)", "cmd": self.open_network_mapper, "desc": "Escolha a letra e conecte uma pasta de rede manualmente."},
                {"name": "🧹 Limpar Cache DNS", "cmd": lambda: self.run_thread("ipconfig /flushdns", "DNS Limpo", check_error=False), "desc": "Resolve falhas de resolução de nomes."},
                {"name": "🗑️ Limpar Temp (%TEMP%)", "cmd": self.cmd_clean_temp, "desc": "Limpa arquivos temporários do usuário atual."},
                {"name": "🔄 Atualizar Políticas", "cmd": lambda: self.run_thread("gpupdate /force", "Políticas atualizadas"), "desc": "Força atualização de GPO."}
            ],
            "Aplicativos": [
                {"name": "📧 Reparar Perfil Outlook", "cmd": self.cmd_open_outlook, "desc": "Tenta abrir Outlook ou forçar modo de segurança."},
                {"name": "☁️ Resetar OneDrive", "cmd": self.cmd_restart_onedrive, "desc": "Mata o processo e reinicia o executável local."},
                {"name": "🔑 Cofre de Credenciais", "cmd": lambda: self.run_thread("control /name Microsoft.CredentialManager", "Cofre Aberto"), "desc": "Abre gerenciador de senhas do Windows."},
                {"name": "🛡️ Modo Seguro Office", "cmd": self.cmd_office_safe_menu, "desc": "Menu para abrir Excel/Word/PPT em modo seguro."}
            ],
            "Impressão": [
                {"name": "✨ Instalador Inteligente", "cmd": self.open_smart_printer_installer, "desc": "Mapeia servidores, lista impressoras e instala automaticamente."},
                {"name": "📄 Fila de Impressão", "cmd": lambda: self.run_thread("rundll32 printui.dll,PrintUIEntry /v", "Fila aberta"), "desc": "Abre fila para cancelar jobs travados."},
                {"name": "⚙️ Wizard Nativo", "cmd": lambda: self.run_thread("rundll32 printui.dll,PrintUIEntry /il", "Wizard Aberto"), "desc": "Assistente padrão do Windows (Backup)."}
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
            lbl_banner = ctk.CTkLabel(self.banner_container, text="N1 Tools", font=("Impact", 60), text_color="white")
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
        self.log_box._textbox.tag_config("SUCESSO", foreground="#00E676")
        self.log_box._textbox.tag_config("ERRO", foreground="#FF5252")
        self.log_box._textbox.tag_config("INFO", foreground="#40C4FF")
        self.log_box._textbox.tag_config("AVISO", foreground="#FFD740")
        
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
            self.write_log("⚡ INICIANDO EXECUÇÃO EM LOTE...", "AVISO")
            self.write_log("Por favor, aguarde o fim de cada processo.", "AVISO")
            
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
                     self.write_log(f"Alerta:\n{output}", "AVISO")
                else:
                    if success_msg: self.write_log(success_msg, "SUCESSO")
                    else: self.write_log(output.strip(), "INFO")
            except Exception as e:
                self.write_log(f"Erro: {str(e)}", "ERRO")
        threading.Thread(target=task).start()

    # --- FUNÇÕES NATIVAS ---
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
                        self.write_log(f"✅ CONEXÃO ESTÁVEL. Latência: {time_ms}", "SUCESSO")
                    else:
                        self.write_log("❌ FALHA DE CONEXÃO.", "ERRO")
            except Exception as e:
                self.write_log(f"Erro: {str(e)}", "ERRO")
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
            self.write_log(f"✅ LIMPEZA CONCLUÍDA. Arquivos removidos: {deleted}", "SUCESSO")
        threading.Thread(target=task).start()

    def cmd_ipconfig(self):
        def task():
            self.write_log("Verificando Rede...", "INFO")
            output = subprocess.check_output("ipconfig", shell=True, encoding='cp850')
            relevant = [line for line in output.split('\n') if "IPv4" in line or "Gateway" in line or "Adaptador" in line]
            self.write_log("\n".join(relevant), "INFO")
            if "169.254" in output: self.write_log("⚠️ ALERTA: IP APIPA (169.254) DETECTADO.", "ERRO")
        threading.Thread(target=task).start()

    def cmd_sysinfo(self):
        def task():
            self.write_log("Buscando Serial...", "INFO")
            try:
                cmd = "powershell \"Get-CimInstance -ClassName Win32_BIOS | Select-Object -ExpandProperty SerialNumber\""
                serial = subprocess.check_output(cmd, shell=True, text=True).strip()
                self.write_log(f"Hostname: {platform.node()}", "SUCESSO")
                self.write_log(f"Serial: {serial}", "SUCESSO")
            except: self.write_log("Falha ao ler serial.", "ERRO")
        threading.Thread(target=task).start()

    def cmd_open_outlook(self):
        def task():
            self.write_log("🔪 Iniciando Reset de Perfil do Outlook...", "AVISO")
            
            # 1. Matar o Outlook se estiver aberto (para não travar o arquivo)
            self.write_log("Fechando Outlook...", "INFO")
            subprocess.run("taskkill /f /im outlook.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(2)
            
            # 2. Deletar a chave de registro onde ficam os perfis
            # Isso equivale a ir no Painel de Controle e excluir todos os perfis manualmente.
            # Caminho padrão para Office 2016, 2019 e 365
            reg_path = r"HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles"
            
            # Comando para deletar a chave silenciosamente (/f)
            cmd_delete = f'reg delete "{reg_path}" /f'
            
            self.write_log("⚙️ Excluindo perfis antigos no Registro...", "AVISO")
            result = subprocess.run(cmd_delete, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            if result.returncode == 0:
                self.write_log("✅ Perfil removido com sucesso.", "SUCESSO")
            else:
                # Se der erro, pode ser que a chave não exista (já estava limpo) ou seja uma versão antiga (15.0)
                self.write_log("Aviso: Perfil padrão não encontrado ou já removido.", "INFO")
            
            time.sleep(1)

            # 3. Abrir o Outlook
            # Como não tem perfil, ele vai abrir AUTOMATICAMENTE a tela de "Adicionar Conta"
            self.write_log("🚀 Abrindo Outlook para nova configuração...", "INFO")
            subprocess.Popen("start outlook", shell=True)
            
            self.write_log("🏁 Pronto. A tela de configuração do novo e-mail foi aberta.", "SUCESSO")

        threading.Thread(target=task).start()

    def cmd_restart_onedrive(self):
        def task():
            self.write_log("🔍 Rastreando executável do OneDrive...", "INFO")
            
            # Lista de caça: Vamos procurar em todos os buracos onde ele costuma se esconder
            locais_possiveis = [
                # 1. O local EXATO do seu print (Program Files)
                r"C:\Program Files\Microsoft OneDrive\OneDrive.exe",
                # 2. O local para instalações 32-bits
                r"C:\Program Files (x86)\Microsoft OneDrive\OneDrive.exe",
                # 3. O local padrão de instalação por usuário (AppData)
                os.path.join(os.getenv('LOCALAPPDATA'), r"Microsoft\OneDrive\OneDrive.exe")
            ]
            
            onedrive_path = None
            
            # Varredura
            for caminho in locais_possiveis:
                if os.path.exists(caminho):
                    onedrive_path = caminho
                    break # Achou, para de procurar
            
            if onedrive_path:
                try:
                    self.write_log(f"✅ Localizado em: {onedrive_path}", "SUCESSO")
                    self.write_log("⚠️ Executando Reset (redefinir)...", "AVISO")
                    
                    # Executa o reset.
                    subprocess.run(f'"{onedrive_path}" /reset', shell=True, check=False)
                    
                    self.write_log("⏳ Aguardando 5s para reiniciar...", "INFO")
                    time.sleep(5)
                    
                    # Ressuscita o OneDrive
                    self.write_log("🚀 Reiniciando OneDrive...", "INFO")
                    subprocess.Popen(f'"{onedrive_path}"', shell=True)
                    
                    self.write_log("🏁 Processo de Reset concluído. Aguarde algumas horas para sincronizar completamente.", "SUCCESS")
                    
                except Exception as e:
                    self.write_log(f"Erro na execução: {str(e)}", "ERRO")
            else:
                self.write_log("❌ OneDrive.exe não encontrado nem no Program Files nem no AppData.", "ERRO")

        threading.Thread(target=task).start()

    def cmd_office_safe_menu(self):
        import tkinter as tk
        
        # --- FUNÇÃO CAÇADORA DE OFFICE ---
        def find_office_path(exe_name):
            # Lista de esconderijos comuns do Office (do mais novo para o mais antigo)
            possible_paths = [
                r"C:\Program Files\Microsoft Office\root\Office16",       # Office 365 / 2019 / 2021 (64-bit)
                r"C:\Program Files (x86)\Microsoft Office\root\Office16", # Office 365 / 2019 / 2021 (32-bit)
                r"C:\Program Files\Microsoft Office\Office16",            # Office 2016 MSI
                r"C:\Program Files (x86)\Microsoft Office\Office16",      # Office 2016 MSI (32-bit)
                r"C:\Program Files\Microsoft Office\Office15",            # Office 2013
                r"C:\Program Files (x86)\Microsoft Office\Office15"       # Office 2013 (32-bit)
            ]
            
            for path in possible_paths:
                full_path = os.path.join(path, exe_name)
                if os.path.exists(full_path):
                    return full_path
            return None

        # Função que executa o comando
        def run_safe(exe_name, nome_bonito):
            self.write_log(f"🔍 Procurando {nome_bonito} no sistema...", "INFO")
            
            # 1. Tenta achar o caminho completo
            caminho_real = find_office_path(exe_name)
            
            if caminho_real:
                try:
                    self.write_log(f"✅ Encontrado em: {caminho_real}", "SUCESSO")
                    # Aspas são importantes caso tenha espaço no caminho
                    subprocess.Popen(f'"{caminho_real}" /safe', shell=True)
                    self.write_log(f"🚀 {nome_bonito} iniciando em Modo Seguro...", "SUCESSO")
                    janela_safe.destroy()
                except Exception as e:
                    self.write_log(f"❌ Erro ao executar: {str(e)}", "ERRO")
            else:
                # 2. Tentativa de desespero: Tentar rodar direto (vai que está no PATH)
                self.write_log(f"⚠️ Caminho padrão não encontrado. Tentando execução direta...", "AVISO")
                try:
                    subprocess.Popen(f"{exe_name} /safe", shell=True)
                    self.write_log("Comando genérico enviado.", "INFO")
                    janela_safe.destroy()
                except:
                    self.write_log(f"❌ {nome_bonito} não encontrado. Verifique se o Office está instalado.", "ERRO")

        # --- UI DA JANELA ---
        janela_safe = tk.Toplevel()
        janela_safe.title("Modo de Segurança Office")
        janela_safe.geometry("300x280")
        janela_safe.configure(bg="#2b2b2b")
        janela_safe.attributes("-topmost", True) # Mantém no topo
        
        tk.Label(janela_safe, text="Selecione para abrir sem plugins:", fg="#ccc", bg="#2b2b2b", font=("Arial", 10)).pack(pady=(15, 10))

        # Botões (Configurados com cores oficiais dos apps)
        tk.Button(janela_safe, text="Excel (/safe)", bg="#1D6F42", fg="white", font=("Arial", 10, "bold"), width=25,
                  command=lambda: run_safe("EXCEL.EXE", "Excel")).pack(pady=5)

        tk.Button(janela_safe, text="Word (/safe)", bg="#2B579A", fg="white", font=("Arial", 10, "bold"), width=25,
                  command=lambda: run_safe("WINWORD.EXE", "Word")).pack(pady=5)

        tk.Button(janela_safe, text="PowerPoint (/safe)", bg="#D24726", fg="white", font=("Arial", 10, "bold"), width=25,
                  command=lambda: run_safe("POWERPNT.EXE", "PowerPoint")).pack(pady=5)

        tk.Button(janela_safe, text="Outlook (/safe)", bg="#0078D4", fg="white", font=("Arial", 10, "bold"), width=25,
                  command=lambda: run_safe("OUTLOOK.EXE", "Outlook")).pack(pady=5)
        
        tk.Button(janela_safe, text="Cancelar", bg="#444", fg="white", width=15, command=janela_safe.destroy).pack(pady=15)

    def cmd_about(self):
        about = ctk.CTkToplevel(self)
        about.title("Sobre")
        about.geometry("400x300")
        about.attributes("-topmost", True)
        ctk.CTkLabel(about, text="Kit de Guerra v3.0", font=("Impact", 20)).pack(pady=20)
        ctk.CTkLabel(about, text="Dev: Gleison Andrade dos Santos", text_color="#40C4FF").pack()
        ctk.CTkLabel(about, text="Dedicado ao time de T.I.", font=("Arial", 10, "italic")).pack(pady=20)


    def open_smart_printer_installer(self):
        # Janela Flutuante
        win = ctk.CTkToplevel(self)
        win.title("Instalador Inteligente de Impressoras")
        win.geometry("500x500")
        win.attributes("-topmost", True)

        # 1. Seleção do Servidor
        ctk.CTkLabel(win, text="1. Selecione o Servidor de Impressão:", font=("Arial", 12, "bold")).pack(pady=(20, 5), anchor="w", padx=20)
        
        # Variável para guardar o IP/Nome escolhido
        self.selected_server = ctk.StringVar(value="")

        def on_server_change(choice):
            # Pega o caminho real do dicionário (Ex: \\anapp01)
            path = self.print_servers.get(choice, "")
            self.entry_manual_server.delete(0, END)
            self.entry_manual_server.insert(0, path)

        combo_servers = ctk.CTkComboBox(win, values=list(self.print_servers.keys()), command=on_server_change, width=400)
        combo_servers.pack(padx=20)

        ctk.CTkLabel(win, text="Ou digite manualmente (Ex: \\\\192.168.0.10):", font=("Arial", 10)).pack(pady=(5, 2), anchor="w", padx=20)
        self.entry_manual_server = ctk.CTkEntry(win, placeholder_text="\\\\servidor", width=400)
        self.entry_manual_server.pack(padx=20)

        # 2. Listagem (Scan)
        ctk.CTkLabel(win, text="2. Buscar Impressoras no Servidor:", font=("Arial", 12, "bold")).pack(pady=(20, 5), anchor="w", padx=20)
        
        self.combo_printers = ctk.CTkComboBox(win, values=["Aguardando busca..."], width=400)
        self.combo_printers.pack(padx=20)

        def cmd_scan_printers():
            server = self.entry_manual_server.get().strip()
            if not server or server == "\\":
                self.write_log("Erro: Nenhum servidor especificado para busca.", "ERRO")
                return
            
            # Botão de feedback visual
            btn_scan.configure(text="Buscando...", state="disabled")
            self.combo_printers.configure(values=["Buscando..."])
            win.update()

            def task_scan():
                try:
                    # Usa o comando 'net view' para listar compartilhamentos
                    # Isso é nativo e rápido, não precisa de PowerShell complexo
                    cmd = f"net view {server}"
                    self.write_log(f"Varrendo servidor: {server}...", "INFO")
                    
                    # Roda o comando escondido
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    result = subprocess.run(cmd, startupinfo=startupinfo, capture_output=True, text=True, encoding='cp850', errors='ignore')
                    
                    lines = result.stdout.split('\n')
                    found_printers = []
                    
                    # Filtra o que é impressora na saída do comando
                    for line in lines:
                        # O output do net view geralmente é: "Nome  Tipo  Comentário"
                        # Procuramos linhas que tenham "Impress" ou "Print"
                        if "Impress" in line or "Print" in line:
                            # Pega a primeira palavra da linha (o nome do compartilhamento)
                            parts = line.split()
                            if parts:
                                found_printers.append(parts[0])

                    if found_printers:
                        self.combo_printers.configure(values=found_printers)
                        self.combo_printers.set(found_printers[0])
                        self.write_log(f"Encontradas {len(found_printers)} impressoras em {server}.", "SUCESSO")
                    else:
                        self.combo_printers.configure(values=["Nenhuma encontrada"])
                        self.write_log(f"Conectou em {server}, mas não listou impressoras. Verifique permissão.", "AVISO")

                except Exception as e:
                    self.write_log(f"Erro ao buscar: {str(e)}", "ERRO")
                finally:
                    btn_scan.configure(text="🔍 Listar Impressoras Disponíveis", state="normal")

            threading.Thread(target=task_scan).start()

        btn_scan = ctk.CTkButton(win, text="🔍 Listar Impressoras Disponíveis", command=cmd_scan_printers, fg_color="#E65100", hover_color="#BF360C")
        btn_scan.pack(pady=10)

        # 3. Instalação
        ctk.CTkLabel(win, text="3. Finalizar:", font=("Arial", 12, "bold")).pack(pady=(20, 5), anchor="w", padx=20)

        def cmd_install_printer():
            server = self.entry_manual_server.get().strip()
            printer_name = self.combo_printers.get()
            
            if not server or not printer_name or "..." in printer_name:
                self.write_log("Selecione servidor e impressora válidos.", "ERRO")
                return

            # Monta o caminho completo: \\servidor\impressora
            # Removemos barras extras caso o usuário tenha digitado errado
            clean_server = server.rstrip("\\")
            full_path = f"{clean_server}\\{printer_name}"
            
            def task_install():
                self.write_log(f"Instalando: {full_path}...", "INFO")
                try:
                    # Comando MÁGICO do Windows para mapear impressora silenciosamente
                    # /in = install network printer
                    # /n = name
                    cmd = f'rundll32 printui.dll,PrintUIEntry /in /n "{full_path}"'
                    
                    subprocess.run(cmd, shell=True)
                    
                    # Infelizmente o rundll32 não retorna erro fácil, mas se rodou, geralmente funciona.
                    # Podemos tentar checar se ela apareceu
                    time.sleep(3)
                    self.write_log(f"Comando de instalação enviado para {printer_name}.", "SUCESSO")
                    self.write_log("Verifique se a impressora apareceu no painel de controle.", "INFO")
                    win.destroy() # Fecha janela ao terminar
                except Exception as e:
                    self.write_log(f"Erro na instalação: {e}", "ERRO")

            threading.Thread(target=task_install).start()

        btn_install = ctk.CTkButton(win, text="✅ INSTALAR IMPRESSORA", command=cmd_install_printer, 
                                    height=50, font=("Arial", 14, "bold"), fg_color="#00C853", hover_color="#009624")
        btn_install.pack(pady=10, padx=20, fill="x")
        

    def open_network_mapper(self):
        win = ctk.CTkToplevel(self)
        win.title("Mapear Unidade de Rede")
        win.geometry("450x350")
        win.attributes("-topmost", True)

        # Título
        ctk.CTkLabel(win, text="Conectar Nova Unidade de Rede", font=("Arial", 14, "bold"), text_color="#40C4FF").pack(pady=(20, 10))

        # 1. Campo de Caminho
        ctk.CTkLabel(win, text="Caminho da Pasta (Ex: \\\\servidor\\pasta):", font=("Arial", 12)).pack(anchor="w", padx=30)
        
        entry_path = ctk.CTkEntry(win, placeholder_text="\\\\servidor\\compartilhamento", width=390)
        entry_path.pack(pady=5, padx=30)
        
        # Botão de colar
        def paste_clipboard():
            try:
                entry_path.delete(0, END)
                entry_path.insert(0, win.clipboard_get())
            except: pass
            
        btn_paste = ctk.CTkButton(win, text="📋 Colar", command=paste_clipboard, width=60, height=20, fg_color="#444", hover_color="#555")
        btn_paste.pack(anchor="e", padx=30)

        # 2. Escolha da Letra (DROPDOWN)
        ctk.CTkLabel(win, text="Escolha a Letra da Unidade (Será substituída):", font=("Arial", 12, "bold")).pack(pady=(10, 0), anchor="w", padx=30)
        
        # Gera lista de Z: até E:
        letras = [f"{chr(i)}:" for i in range(90, 68, -1)] 
        combo_drive = ctk.CTkComboBox(win, values=letras, width=120)
        combo_drive.set("Z:") # Padrão
        combo_drive.pack(anchor="w", padx=30, pady=5)

        # 3. Executar
        def cmd_map_drive():
            caminho = entry_path.get().strip()
            letra = combo_drive.get() # PEGA A LETRA QUE O USUÁRIO ESCOLHEU

            if not caminho or len(caminho) < 3 or "\\" not in caminho:
                self.write_log("Erro: Caminho inválido. Use o formato \\\\servidor\\pasta", "ERRO")
                return

            def task_map():
                btn_connect.configure(state="disabled", text="Conectando...")
                self.write_log(f"Limpando e mapeando {letra} para {caminho}...", "INFO")
                
                try:
                    # 1. REMOVE A LETRA ESCOLHIDA (Limpeza)
                    subprocess.run(f"net use {letra} /delete /y", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    
                    # 2. MAPEA A LETRA ESCOLHIDA
                    cmd = f'net use {letra} "{caminho}" /PERSISTENT:YES'
                    
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    
                    result = subprocess.run(cmd, startupinfo=startupinfo, capture_output=True, text=True, encoding='cp850', errors='ignore')
                    
                    if result.returncode == 0:
                        self.write_log(f"✅ Sucesso! Unidade {letra} conectada.", "SUCESSO")
                        subprocess.Popen(f"explorer {letra}", shell=True) # Abre a pasta
                        win.destroy()
                    else:
                        self.write_log("❌ Falha ao mapear.", "ERRO")
                        self.write_log(f"Erro: {result.stderr.strip()}", "AVISO")
                
                except Exception as e:
                    self.write_log(f"Erro crítico: {e}", "ERRO")
                finally:
                    if win.winfo_exists():
                        btn_connect.configure(state="normal", text="🔗 CONECTAR AGORA")

            threading.Thread(target=task_map).start()

        btn_connect = ctk.CTkButton(win, text="🔗 CONECTAR AGORA", command=cmd_map_drive, 
                                    height=45, font=("Arial", 13, "bold"), fg_color="#00C853", hover_color="#009624")
        btn_connect.pack(pady=25, padx=30, fill="x")

if __name__ == "__main__":
    app = TechToolApp()
    app.mainloop()