# ipchange

Aplicativo em C# que chama um script PowerShell para listar os adaptadores de rede e trocar o IPv4 de forma autônoma usando um usuário com permissão administrativa.

## Linguagem

Este projeto foi feito em **C#**, chamando o script **PowerShell** `ipchange.ps1`.

## Requisitos

- .NET SDK 8.0+ para compilar ou executar o aplicativo C#
- Windows com `Get-NetAdapter`, `Get-NetIPAddress` e `netsh`
- PowerShell 5.1+ ou PowerShell 7+
- Usuário e senha com privilégio administrativo no computador

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
  -AdapterName "Ethernet" `
  -IPAddress "192.168.0.50" `
  -PrefixLength 24 `
  -DefaultGateway "192.168.0.1" `
  -DnsServers "1.1.1.1","8.8.8.8"
```

> Quando você precisar informar uma `SecureString` para `-Password`, execute o `ipchange.ps1` diretamente no PowerShell para montar a senha com `Read-Host -AsSecureString`.

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

## Observações

- o executável C# apenas orquestra a chamada do `ipchange.ps1`
- execute o script em um host Windows
- o usuário informado precisa ter permissão administrativa
- para simular a alteração sem aplicar, use `-WhatIf`
