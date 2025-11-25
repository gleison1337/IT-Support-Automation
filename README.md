# 🛠️ TechTool Kit - Automação para Service Desk

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Status](https://img.shields.io/badge/Status-Em%20Desenvolvimento-green)

Ferramenta desenvolvida para agilizar o atendimento de **Service Desk N1/N2**, automatizando diagnósticos e correções comuns no ambiente Windows Corporativo. Focado em reduzir o TMA (Tempo Médio de Atendimento) com soluções de um clique.

## 🚀 Funcionalidades Principais

### 🔧 Reparos Rápidos
- **Reset do OneDrive:** Localiza a instalação (User ou Machine wide), força o reset e reinicia o processo automaticamente.
- **Reparo do Outlook:** Tenta abrir em modo de segurança ou resetar o perfil de navegação.
- **Gerenciador de Credenciais:** Atalho rápido para limpeza de senhas do Windows.

### 🛡️ Modo de Segurança (Office)
Menu exclusivo para iniciar aplicativos do pacote Office sem suplementos (plugins) para diagnóstico de travamentos:
- Excel, Word, PowerPoint e Outlook (`/safe`).

### ⚡ Utilitários de Rede & Sistema
- **Limpeza de Cache DNS:** `ipconfig /flushdns`.
- **Atualização de Políticas (GPO):** `gpupdate /force`.
- **Limpeza de Temporários:** Esvazia a pasta `%TEMP%`.

## 💻 Tecnologias Utilizadas
- **Python 3.x**
- **CustomTkinter:** Para uma interface gráfica moderna e escura (Dark Mode nativo).
- **Subprocess & Threading:** Para execução de comandos do sistema sem travar a interface.

## 📦 Como Rodar Localmente

1. Clone o repositório:
   ```bash
   git clone [https://github.com/gleison1337/IT-Support-Automation.git](https://github.com/gleison1337/IT-Support-Automation.git)
