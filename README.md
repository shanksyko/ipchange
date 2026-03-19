# ipchange

Aplicativo WinForms em C# para trocar o IPv4 de um adaptador de rede por meio do serviço Windows local instalado pelo MSI, sem PowerShell e sem fluxo de diagnóstico.

## Requisitos

- .NET SDK 8.0+ para compilar ou executar o aplicativo
- Windows com `netsh`
- para uso sem prompt de elevação, instale o próprio `ipchange.exe` como serviço Windows local uma única vez com privilégios administrativos

## O que a interface faz

- carrega os adaptadores de rede visíveis em uma lista
- permite alternar entre IPv4 estático e DHCP
- no modo estático, permite informar IPv4, prefixo, gateway e DNS
- usa o serviço local do próprio `ipchange.exe` quando o usuário atual não está em modo administrador
- inicia em segundo plano com ícone na bandeja quando executado com `--background`
- quando instalado via MSI, registra inicialização automática no logon do Windows
- executa a alteração real via named pipe para o serviço Windows local
- permite redimensionar a janela e adapta os campos ao tamanho disponível
- mostra o andamento básico da operação na própria janela
- grava logs em arquivo `.txt`

## Como executar

```powershell
dotnet run
```

Ou pelo executável compilado:

```powershell
.\bin\Debug\net8.0-windows\ipchange.exe
```

## Instalador

O fluxo de instalação suportado pelo repositório é o MSI via WiX em [installer/ipchange.installer.wixproj](c:/Repositorio/ipchange/installer/ipchange.installer.wixproj).

O fluxo recomendado de distribuição agora é o MSI:

1. publicar o `ipchange.exe` em single-file
2. gerar o MSI com [build-msi.ps1](c:/Repositorio/ipchange/build-msi.ps1)
3. distribuir o MSI versionado gerado em `artifacts/installer/msi/win-x64`

Para gerar o MSI:

```powershell
.\build-msi.ps1
```

Cada novo build do MSI recebe:

- um nome de arquivo único com versão + timestamp
- uma versão interna crescente do Windows Installer

Isso faz com que um MSI novo atualize ou substitua a instalação anterior do `ipchange` automaticamente.

O MSI:

- copia o `ipchange.exe` para `Program Files\ipchange`
- registra o serviço Windows `Ipchange`
- registra a inicialização automática em segundo plano no logon do Windows
- inicia o serviço automaticamente durante a instalação
- remove o serviço automaticamente na desinstalação
- remove a entrada de inicialização automática na desinstalação
- remove os logs protegidos em `%ProgramData%\ipchange\secure-logs` na desinstalação
- remove a pasta de instalação e as pastas de dados criadas pelo app quando estiverem vazias
- deixa a interface disponível em `Program Files\\ipchange\\ipchange.exe`
- cria atalho no menu Iniciar
- cria atalho na área de trabalho

A pasta final de distribuição passa a incluir o MSI versionado a cada build, por exemplo:

- `ipchange-installer-v1.0.1-20260318-224500.msi`

## Modo de execução

- se o usuário atual já estiver em modo administrador, o app aplica a configuração diretamente
- se o usuário atual não estiver em modo administrador, o app envia a solicitação ao serviço Windows local `Ipchange` pelo named pipe `ipchange-service`
- quando iniciado com o Windows, o app sobe em segundo plano e permanece acessível pela bandeja do sistema

## Observações

- o aplicativo é Windows-only
- a aplicação real do IP ainda exige privilégio administrativo, mas isso fica concentrado no serviço Windows local
- a inicialização automática do app instalado passa a ser controlada pelo MSI para permitir atualização e remoção completas
- os logs passam a ficar em `%ProgramData%\ipchange\secure-logs` com extensão `.txt`
- quando o serviço cria a pasta de logs, ele restringe o acesso para Administrators e LocalSystem
