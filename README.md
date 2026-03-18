# ipchange

Aplicativo em C# que chama um script PowerShell para listar os adaptadores de rede e trocar o IPv4 de forma autônoma usando um usuário com permissão administrativa.

## Linguagem

Este projeto foi feito em **C#**, chamando o script **PowerShell** `ipchange.ps1`.

## Requisitos

- .NET SDK 8.0+ para compilar ou executar o aplicativo C#
- Windows com `Get-NetAdapter`, `Get-NetIPAddress` e `netsh`
- PowerShell 5.1+ ou PowerShell 7+
- A conta local `.\support` deve existir com privilégio administrativo no computador; se ela não existir, sobrescreva o usuário padrão com `-Username` ou `IPCHANGE_ADMIN_USERNAME`

## Segurança das credenciais

- A forma mais segura disponível neste projeto continua sendo passar `-Password` como `SecureString`
- O usuário padrão é o local `.\support`; para evitar prompt de senha, você pode usar `-PlainTextPassword` ou as variáveis `IPCHANGE_ADMIN_USERNAME` e `IPCHANGE_ADMIN_PASSWORD`
- Se preferir manter um valor local fixo na máquina, copie `ipchange.local.example.psd1` para `ipchange.local.psd1`; esse arquivo é ignorado pelo Git e pode definir `Username` e `PlainTextPassword`
- Ambos os caminhos usam senha em texto puro em algum momento, então trate esse uso com cuidado
- Argumentos de linha de comando podem ficar visíveis em listagens de processo
- Variáveis de ambiente também ficam em texto puro no processo atual até serem limpas e são herdadas automaticamente pelo processo PowerShell iniciado pelo C#
- Prefira definir `IPCHANGE_ADMIN_USERNAME` e `IPCHANGE_ADMIN_PASSWORD` apenas no processo/sessão atual, não como variáveis persistentes de usuário ou sistema
- Mesmo após converter a senha para `SecureString`, o texto puro pode permanecer na memória por algum tempo
- Se você montar um `SecureString` a partir de uma senha em texto puro por conta própria, essa mesma limitação existe antes da conversão

## O que o script faz

- lista todos os adaptadores de rede visíveis
- permite escolher qual adaptador será configurado
- usa por padrão o usuário local `.\support` e solicita apenas a senha quando necessário
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
4. pedir apenas a senha administrativa do usuário local `.\support`
5. reaplicar a execução com a credencial informada e validar a alteração

## Como executar informando os dados

### Via C#

```powershell
dotnet run -- `
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
$env:IPCHANGE_ADMIN_PASSWORD = "SenhaAdminAqui"

dotnet run -- `
  -AdapterName "Ethernet" `
  -IPAddress "192.168.0.50" `
  -PrefixLength 24 `
  -DefaultGateway "192.168.0.1" `
  -DnsServers "1.1.1.1","8.8.8.8"

Remove-Item Env:IPCHANGE_ADMIN_PASSWORD
```

> Quando a variável de ambiente de senha estiver definida, o processo iniciado pelo C# herda esse valor e o script não pede a senha. O usuário padrão continua sendo `.\support`.
>
> Se você também definir `IPCHANGE_ADMIN_USERNAME`, ela sobrescreve o usuário padrão. Depois do uso, limpe essas variáveis para reduzir a exposição da credencial no terminal atual.

### Via PowerShell

```powershell
$password = Read-Host "Senha" -AsSecureString

.\ipchange.ps1 `
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
  -PlainTextPassword "SenhaAdminAqui" `
  -AdapterName "Ethernet" `
  -IPAddress "192.168.0.50" `
  -PrefixLength 24 `
  -DefaultGateway "192.168.0.1" `
  -DnsServers "1.1.1.1","8.8.8.8"
```

### Via arquivo local ignorado pelo Git

Copie o exemplo para um arquivo local não versionado:

```powershell
Copy-Item .\ipchange.local.example.psd1 .\ipchange.local.psd1
```

Depois edite `ipchange.local.psd1` e preencha sua senha:

```powershell
@{
  Username = '.\support'
  PlainTextPassword = 'SuaSenhaAqui'
}
```

Com esse arquivo presente, o script passa a usar esses valores como padrão e deixa de pedir a senha interativamente.

## Observações

- o executável C# apenas orquestra a chamada do `ipchange.ps1`
- execute o script em um host Windows
- o usuário local padrão `.\support` precisa ter permissão administrativa, a menos que você sobrescreva `-Username` ou `IPCHANGE_ADMIN_USERNAME`
- você pode evitar o prompt de credencial usando `-PlainTextPassword`, `IPCHANGE_ADMIN_PASSWORD` ou um arquivo local `ipchange.local.psd1`
- para simular a alteração sem aplicar, use `-WhatIf`
