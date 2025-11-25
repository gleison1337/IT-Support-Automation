import customtkinter as ctk
import subprocess
import os
import threading
import platform

# Configuração da Aparência
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class TechToolApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("TechTool Kit - N1 Automation")
        self.geometry("900x600")

        # Layout: 2 Colunas (Botões à esquerda, Log à direita)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- FRAME DE BOTÕES (Esquerda) ---
        self.left_frame = ctk.CTkScrollableFrame(self, width=250, label_text="Ferramentas")
        self.left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Seção: Diagnóstico
        self.add_section_label("Diagnóstico & Info")
        self.add_button("Verificar IP / Rede", self.check_network_info)
        self.add_button("Ping Google (Teste Net)", lambda: self.run_command_thread("ping google.com"))
        self.add_button("Obter Serial/Hostname", self.get_system_info)

        # Seção: Correções Rápidas
        self.add_section_label("Correções Comuns")
        self.add_button("Reiniciar OneDrive", self.restart_onedrive)
        self.add_button("Limpar Temp (%TEMP%)", self.clear_temp_files)
        self.add_button("Flush DNS (Limpar Cache)", lambda: self.run_command_thread("ipconfig /flushdns"))
        self.add_button("Reiniciar Spooler Impressão", self.restart_spooler)
        self.add_button("Forçar GPUPDATE", lambda: self.run_command_thread("gpupdate /force"))

        # Seção: Outlook/Office
        self.add_section_label("Office & Outlook")
        self.add_button("Abrir Gerenciador Credenciais", lambda: self.run_command_thread("control /name Microsoft.CredentialManager"))
        self.add_button("Reparar Outlook (Safe Mode)", lambda: self.run_command_thread("outlook.exe /safe"))

        # Seção: Ferramentas
        self.add_section_label("Utilitários")
        self.add_button("Mapear Unidade de Rede", self.map_network_drive)
        self.add_button("Instalar Impressora (Wizard)", lambda: self.run_command_thread("rundll32 printui.dll,PrintUIEntry /il"))

        # --- FRAME DE LOG (Direita) ---
        self.right_frame = ctk.CTkFrame(self)
        self.right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        
        self.log_label = ctk.CTkLabel(self.right_frame, text="Log de Saída", font=ctk.CTkFont(size=16, weight="bold"))
        self.log_label.pack(pady=5)

        self.log_textbox = ctk.CTkTextbox(self.right_frame, width=500)
        self.log_textbox.pack(expand=True, fill="both", padx=5, pady=5)

    def add_section_label(self, text):
        label = ctk.CTkLabel(self.left_frame, text=text, font=ctk.CTkFont(size=14, weight="bold"), anchor="w")
        label.pack(pady=(15, 5), padx=5, fill="x")

    def add_button(self, text, command):
        btn = ctk.CTkButton(self.left_frame, text=text, command=command, height=35, anchor="w")
        btn.pack(pady=2, padx=5, fill="x")

    def log(self, message):
        self.log_textbox.insert("end", f"\n> {message}\n")
        self.log_textbox.see("end")

    # --- FUNÇÕES LÓGICAS ---

    def run_command_thread(self, command):
        """Roda comandos sem travar a interface"""
        def task():
            self.log(f"Executando: {command}...")
            try:
                # shell=True permite rodar comandos como se fosse no CMD
                result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='cp850') 
                self.log(result.stdout)
                if result.stderr:
                    self.log(f"ERRO/INFO: {result.stderr}")
            except Exception as e:
                self.log(f"Erro ao executar: {str(e)}")
        
        threading.Thread(target=task).start()

    def check_network_info(self):
        self.run_command_thread("ipconfig /all")

    def get_system_info(self):
        self.log("Coletando dados do sistema...")
        def task():
            try:
                serial = subprocess.check_output("wmic bios get serialnumber", shell=True).decode().split("\n")[1].strip()
                hostname = platform.node()
                self.log(f"--------------------------")
                self.log(f"HOSTNAME: {hostname}")
                self.log(f"SERIAL NUMBER: {serial}")
                self.log(f"--------------------------")
            except Exception as e:
                self.log(f"Erro: {e}")
        threading.Thread(target=task).start()

    def restart_onedrive(self):
        self.log("Reiniciando OneDrive...")
        def task():
            subprocess.run("taskkill /f /im OneDrive.exe", shell=True)
            # Tenta encontrar o executável do OneDrive no local padrão
            user_path = os.path.expanduser("~")
            onedrive_path = os.path.join(user_path, r"AppData\Local\Microsoft\OneDrive\OneDrive.exe")
            
            if os.path.exists(onedrive_path):
                subprocess.Popen(onedrive_path)
                self.log("OneDrive iniciado com sucesso.")
            else:
                self.log(f"Executável não encontrado em: {onedrive_path}. Tente abrir manualmente.")
        threading.Thread(target=task).start()

    def clear_temp_files(self):
        self.log("Limpando arquivos temporários...")
        # Nota: Alguns arquivos não serão apagados se estiverem em uso. Isso é normal.
        self.run_command_thread('del /q/f/s %TEMP%\\*')

    def restart_spooler(self):
        self.log("Tentando reiniciar Spooler (Pode pedir Admin)...")
        # O Spooler geralmente precisa de admin. Se o usuário não for admin local, isso falhará.
        # Mas muitos devs e analistas têm admin local em suas máquinas.
        self.run_command_thread("net stop spooler && net start spooler")

    def map_network_drive(self):
        # Abre um popup para pedir o caminho
        dialog = ctk.CTkInputDialog(text="Digite o caminho da rede (Ex: \\\\servidor\\pasta):", title="Mapear Rede")
        path = dialog.get_input()
        if path:
            drive_letter = "Z:" # Simplificação. Ideal seria verificar letra livre.
            self.run_command_thread(f"net use {drive_letter} {path} /PERSISTENT:YES")

if __name__ == "__main__":
    app = TechToolApp()
    app.mainloop()