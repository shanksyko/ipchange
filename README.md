# ipchange

Aplicativo em C# que chama um script PowerShell para listar os adaptadores de rede e trocar o IPv4 de forma autônoma usando um usuário com permissão administrativa.

## Linguagem

Este projeto foi feito em **C#**, chamando o script **PowerShell** `ipchange.ps1`.

## Requisitos

- .NET SDK 8.0+ para compilar ou executar o aplicativo C#
- Windows com `Get-NetAdapter`, `Get-NetIPAddress` e `netsh`
- PowerShell 5.1+ ou PowerShell 7+
- Usuário e senha com privilégio administrativo no computador

## Segurança das credenciais

- A forma mais segura disponível neste projeto continua sendo passar `-Password` como `SecureString`
- Para evitar prompts, você pode usar `-PlainTextPassword` ou as variáveis `IPCHANGE_ADMIN_USERNAME` e `IPCHANGE_ADMIN_PASSWORD`
- Ambos os caminhos usam senha em texto puro em algum momento, então trate esse uso com cuidado
- Argumentos de linha de comando podem ficar visíveis em listagens de processo
- Variáveis de ambiente também ficam em texto puro no processo atual até serem limpas e são herdadas automaticamente pelo processo PowerShell iniciado pelo C#
- Prefira definir `IPCHANGE_ADMIN_USERNAME` e `IPCHANGE_ADMIN_PASSWORD` apenas no processo/sessão atual, não como variáveis persistentes de usuário ou sistema
- Mesmo após converter a senha para `SecureString`, o texto puro pode permanecer na memória por algum tempo
- Se você montar um `SecureString` a partir de uma senha em texto puro por conta própria, essa mesma limitação existe antes da conversão

## O que o script faz

- lista todos os adaptadores de rede visíveis
- permite escolher qual adaptador será configurado
- solicita usuário e senha quando necessário
- aplica IPv4 estático, gateway e DNS
- confirma se o IP foi realmente alterado

## Como listar os adaptadores

### Via C#

```powershell
dotnet run -- -ListAdapters
```

### Via PowerShell

```powershell
.\ipchange.ps1 -ListAdapters
```

## Como executar de forma interativa

### Via C#

```powershell
dotnet run --
```

### Via PowerShell

```powershell
.\ipchange.ps1
```

O script vai:

1. mostrar os adaptadores disponíveis
2. pedir o `InterfaceIndex` do adaptador
3. pedir o novo IPv4, prefixo, gateway e DNS
4. pedir usuário e senha administrativa
5. reaplicar a execução com a credencial informada e validar a alteração

## Como executar informando os dados

### Via C#

```powershell
dotnet run -- `
  -Username "DOMINIO\usuario" `
  -PlainTextPassword "SenhaAdminAqui" `
  -AdapterName "Ethernet" `
  -IPAddress "192.168.0.50" `
  -PrefixLength 24 `
  -DefaultGateway "192.168.0.1" `
  -DnsServers "1.1.1.1","8.8.8.8"
```

> O parâmetro `-PlainTextPassword` evita o prompt, mas expõe a senha na linha de comando.

### Via C# com variáveis de ambiente

```powershell
$env:IPCHANGE_ADMIN_USERNAME = "DOMINIO\usuario"
$env:IPCHANGE_ADMIN_PASSWORD = "SenhaAdminAqui"

dotnet run -- `
  -AdapterName "Ethernet" `
  -IPAddress "192.168.0.50" `
  -PrefixLength 24 `
  -DefaultGateway "192.168.0.1" `
  -DnsServers "1.1.1.1","8.8.8.8"

Remove-Item Env:IPCHANGE_ADMIN_USERNAME
Remove-Item Env:IPCHANGE_ADMIN_PASSWORD
```

> Quando as variáveis de ambiente estiverem definidas, o processo iniciado pelo C# herda esses valores e o script não pede usuário nem senha.
>
> Depois do uso, limpe essas variáveis para reduzir a exposição da credencial no terminal atual.

### Via PowerShell

```powershell
$password = Read-Host "Senha" -AsSecureString

.\ipchange.ps1 `
  -Username "DOMINIO\usuario" `
  -Password $password `
  -AdapterName "Ethernet" `
  -IPAddress "192.168.0.50" `
  -PrefixLength 24 `
  -DefaultGateway "192.168.0.1" `
  -DnsServers "1.1.1.1","8.8.8.8"
```

### Via PowerShell sem prompt de senha

```powershell
.\ipchange.ps1 `
  -Username "DOMINIO\usuario" `
  -PlainTextPassword "SenhaAdminAqui" `
  -AdapterName "Ethernet" `
  -IPAddress "192.168.0.50" `
  -PrefixLength 24 `
  -DefaultGateway "192.168.0.1" `
  -DnsServers "1.1.1.1","8.8.8.8"
```

## Observações

- o executável C# apenas orquestra a chamada do `ipchange.ps1`
- execute o script em um host Windows
- o usuário informado precisa ter permissão administrativa
- você pode evitar o prompt de credencial usando `-Username` com `-PlainTextPassword` ou com `IPCHANGE_ADMIN_USERNAME` + `IPCHANGE_ADMIN_PASSWORD`
- para simular a alteração sem aplicar, use `-WhatIf`
