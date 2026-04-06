<div align="center">

<img src="assets/icon.png" width="120" alt="Network Configurator Icon"/>

# Network Configurator

**Gerenciador de rede para Windows** — aplica IPv4 estático ou DHCP sem precisar de prompt de administrador na interface.  
Roda como serviço Windows, inicia com o sistema e fica acessível pelo ícone na bandeja.

![Platform](https://img.shields.io/badge/platform-Windows-0078D4?logo=windows)
![.NET](https://img.shields.io/badge/.NET-8.0-512BD4?logo=dotnet)
![License](https://img.shields.io/github/license/shanksyko/network-configurator)
![Release](https://img.shields.io/github/v/release/shanksyko/network-configurator)

</div>

---

## Funcionalidades

| Área | Recursos |
|------|----------|
| **Configuração** | IPv4 estático ou DHCP, gateway, prefixo, DNS |
| **Diagnósticos** | Ping, Traceroute, DNS Lookup, Port Scan, SSL/TLS, WHOIS, ARP, Geo IP e muito mais  |
| **Ferramentas** | Calculadora de Sub-rede, MAC Vendor, Wake-on-LAN, CIDR Range, IP Converter |
| **Perfis** | Salve e carregue configurações de rede rapidamente |
| **Serviço** | Aplica configurações pelo serviço local sem exigir UAC na interface |
| **Bandeja** | Inicia oculto com o Windows, acessível pela system tray |

---

## Instalação (recomendado)

Baixe o instalador `.msi` na [página de releases](https://github.com/shanksyko/network-configurator/releases) e execute-o.

O instalador:

- Copia o `network-configurator.exe` para `Program Files\Network Configurator`
- Registra o serviço Windows **NetworkConfigurator** (inicia automaticamente)
- Cria atalhos no Menu Iniciar e na Área de Trabalho
- Registra inicialização automática no logon do Windows
- Suporta **Reparar** e **Desinstalar** via Painel de Controle → Programas

Para gerar um MSI localmente:

```powershell
.\build-msi.ps1
# Saída: artifacts\installer\msi\win-x64\network-configurator-installer-v1.0.x-<data>.msi
```

---

## Compilar e executar

**Pré-requisitos:** .NET SDK 8.0+

```powershell
# Executar em modo debug
dotnet run

# Compilar Release
dotnet build -c Release

# Publicar executável único
dotnet publish -c Release -r win-x64 --self-contained false
```

---

## Arquitetura

```
┌─────────────────────────────────┐        named pipe
│  Interface (WinForms)            │ ──────────────────▶ │  Serviço Windows  │
│  network-configurator.exe        │                      │  NetworkConfigurator│
└─────────────────────────────────┘                      └───────────────────┘
```

- Se o usuário já for **administrador**, a configuração é aplicada diretamente.  
- Caso contrário, o app envia a solicitação ao serviço local pelo named pipe `network-configurator-service`.  
- Logs em `%ProgramData%\network-configurator\secure-logs\network-configurator-YYYYMMDD.txt`

---

## Estrutura do projeto

```
network-configurator.csproj   ← projeto principal (WinForms .NET 8)
MainForm.cs                   ← toda a UI + lógica de diagnósticos
NetworkConfigurationService.cs← serviço Windows
ServiceHost.cs / ServiceContracts.cs
LocalServiceManager.cs
assets/
  icon.ico / icon.png         ← ícone do aplicativo
installer/
  Package.wxs                 ← definição WiX do MSI
  ipchange.installer.wixproj
build-msi.ps1                 ← script de build do MSI
```

---

## Requisitos

- Windows 10/11 (x64)
- .NET 8 Runtime (ou self-contained)
- Privilégio administrativo **apenas na instalação do serviço** (via MSI ou primeira execução)

---

<div align="center">
Feito com ♥ por <a href="https://github.com/shanksyko">shanksyko</a>
</div>