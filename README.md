# 🛠️ TechTool Kit v3.1 - Enterprise Edition

![Status](https://img.shields.io/badge/Status-Stable-green) ![Python](https://img.shields.io/badge/Python-3.x-blue) ![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

Uma aplicação desktop robusta desenvolvida para agilizar o atendimento de **Service Desk Nível 1**, automatizando diagnósticos, configurações de rede e instalações complexas sem a necessidade de credenciais administrativas globais.

## 🚀 Funcionalidades Principais

### 🔍 Diagnóstico & Monitoramento
- **Ping Inteligente:** Validação de conectividade (IPv4) com análise de latência.
- **Scanner de Rede:** Exibe IP/Gateway e alerta automaticamente sobre APIPA (169.254.x.x).
- **Asset Info:** Captura automática de Serial Number (Dell/Lenovo/HP) e Hostname.

### 🖨️ Gestão de Impressão (Novo!)
- **Instalador Inteligente:** Mapeamento automático de servidores de impressão por localidade (SP, RJ, MG, BA, etc.).
- **Scanner de Drivers:** Lista impressoras disponíveis no servidor remoto via `net view`.
- **Instalação Silenciosa:** Adiciona a impressora ao Windows sem wizards demorados.

### 📂 Rede & Arquivos (Novo!)
- **Mapeador Persistente:** Conecta unidades de rede (Z:, Y:) com limpeza automática de conexões antigas.
- **Limpeza de Cache:** Flush DNS e remoção segura de arquivos `%TEMP%`.
- **Políticas:** Atualização forçada de GPO (`gpupdate`).

### 💻 Aplicativos & Office
- **Office Safe Mode Hunter:** Detecta automaticamente a versão instalada do Office (365/2016/2019) e abre apps em Modo Seguro.
- **OneDrive Reset:** Mata processos travados e força redefinição do executável local.
- **Outlook Fix:** Recriação de perfil e abertura em modo de segurança.

## 💻 Tecnologias
- **Python 3**
- **CustomTkinter** (Interface Gráfica Moderna Dark Mode)
- **Win32 API & Subprocess** (Automação nativa do Windows)
- **Threading** (Execução assíncrona para fluidez da UI)

## 📦 Instalação e Uso

1. **Clone o repositório:**
   ```bash
   git clone [https://github.com/gleison1337/IT-Support-Automation.git](https://github.com/gleison1337/IT-Support-Automation.git)
