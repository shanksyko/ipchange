# ipchange

Script PowerShell para listar os adaptadores de rede e trocar o IPv4 de forma autônoma usando um usuário com permissão administrativa.

## Requisitos

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

```powershell
.\ipchange.ps1 -ListAdapters
```

## Como executar de forma interativa

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

- execute o script em um host Windows
- o usuário informado precisa ter permissão administrativa
- para simular a alteração sem aplicar, use `-WhatIf`
