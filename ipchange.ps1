[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [string]$Username,

    [Parameter()]
    [SecureString]$Password,

    [Parameter()]
    [string]$PlainTextPassword,

    [Parameter()]
    [string]$AdapterName,

    [Parameter()]
    [string]$IPAddress,

    [Parameter()]
    [ValidateRange(0, 32)]
    [int]$PrefixLength,

    [Parameter()]
    [string]$DefaultGateway,

    [Parameter()]
    [string[]]$DnsServers,

    [Parameter()]
    [switch]$ListAdapters,

    [Parameter()]
    [switch]$UseCurrentCredential
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-Windows {
    if (-not ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT)) {
        throw 'Este script só pode ser executado no Windows, pois depende de Get-NetAdapter/Get-NetIPAddress e netsh.'
    }
}

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'O usuário informado precisa executar o script com privilégios administrativos para alterar o IP.'
    }
}

function Test-IPv4Address {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Address
    )

    $parsed = $null
    if (-not [System.Net.IPAddress]::TryParse($Address, [ref]$parsed)) {
        return $false
    }

    return $parsed.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork
}

function ConvertTo-SubnetMask {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, 32)]
        [int]$Prefix
    )

    $mask = [uint32]0
    for ($bit = 0; $bit -lt $Prefix; $bit++) {
        $mask = $mask -bor ([uint32]1 -shl (31 - $bit))
    }

    return [System.Net.IPAddress]::new([BitConverter]::GetBytes([System.Net.IPAddress]::HostToNetworkOrder([int]$mask))).ToString()
}

function Get-VisibleAdapters {
    Assert-Windows

    return Get-NetAdapter -ErrorAction Stop |
        Sort-Object -Property InterfaceIndex |
        Select-Object InterfaceIndex, Name, InterfaceDescription, Status, MacAddress, LinkSpeed
}

function Show-AdapterTable {
    $adapters = Get-VisibleAdapters

    if (-not $adapters) {
        throw 'Nenhum adaptador de rede foi encontrado.'
    }

    $adapters | Format-Table -AutoSize | Out-Host
    return $adapters
}

function Read-RequiredValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter()]
        [scriptblock]$Validator,

        [Parameter()]
        [string]$ValidationMessage = 'Valor inválido.'
    )

    while ($true) {
        $value = Read-Host -Prompt $Prompt
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        if ($null -eq $Validator -or (& $Validator $value)) {
            return $value
        }

        Write-Warning $ValidationMessage
    }
}

function Resolve-AdapterName {
    param(
        [Parameter()]
        [string]$SelectedAdapterName
    )

    $adapters = Show-AdapterTable

    if ($SelectedAdapterName) {
        $match = $adapters | Where-Object { $_.Name -eq $SelectedAdapterName }
        if (-not $match) {
            throw "Adaptador '$SelectedAdapterName' não foi encontrado."
        }

        return $match.Name
    }

    while ($true) {
        $selection = Read-Host -Prompt 'Digite o InterfaceIndex do adaptador que deseja configurar'
        $parsedSelection = 0
        if (-not [int]::TryParse($selection, [ref]$parsedSelection)) {
            Write-Warning 'Informe um número válido.'
            continue
        }

        $selected = $adapters | Where-Object { $_.InterfaceIndex -eq $parsedSelection }
        if ($selected) {
            return $selected.Name
        }

        Write-Warning 'Adaptador não encontrado para o índice informado.'
    }
}

function Get-ProcessExecutable {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        return 'pwsh.exe'
    }

    return 'powershell.exe'
}

function Invoke-WithCredential {
    param(
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential,

        [Parameter(Mandatory = $true)]
        [string]$ChosenAdapterName,

        [Parameter(Mandatory = $true)]
        [string]$ChosenIPAddress,

        [Parameter(Mandatory = $true)]
        [int]$ChosenPrefixLength,

        [Parameter()]
        [string]$ChosenDefaultGateway,

        [Parameter()]
        [string[]]$ChosenDnsServers
    )

    $argumentList = @(
        '-NoLogo',
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $PSCommandPath,
        '-UseCurrentCredential',
        '-AdapterName', $ChosenAdapterName,
        '-IPAddress', $ChosenIPAddress,
        '-PrefixLength', $ChosenPrefixLength.ToString()
    )

    if ($ChosenDefaultGateway) {
        $argumentList += @('-DefaultGateway', $ChosenDefaultGateway)
    }

    $normalizedDnsServers = $ChosenDnsServers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if ($normalizedDnsServers) {
        $argumentList += '-DnsServers'
        $argumentList += $normalizedDnsServers
    }

    if ($WhatIfPreference) {
        $argumentList += '-WhatIf'
    }

    $process = Start-Process -FilePath (Get-ProcessExecutable) -Credential $Credential -ArgumentList $argumentList -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "A execução com a credencial informada falhou com código de saída $($process.ExitCode)."
    }
}

function Invoke-Netsh {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $output = & netsh @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw ($output | Out-String).Trim()
    }

    return $output
}

function Set-AdapterIPv4Configuration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ChosenAdapterName,

        [Parameter(Mandatory = $true)]
        [string]$ChosenIPAddress,

        [Parameter(Mandatory = $true)]
        [int]$ChosenPrefixLength,

        [Parameter()]
        [string]$ChosenDefaultGateway,

        [Parameter()]
        [string[]]$ChosenDnsServers
    )

    Assert-Windows
    Assert-Administrator

    $adapter = Get-NetAdapter -Name $ChosenAdapterName -ErrorAction Stop
    $subnetMask = ConvertTo-SubnetMask -Prefix $ChosenPrefixLength
    $gatewayValue = if ($ChosenDefaultGateway) { $ChosenDefaultGateway } else { 'none' }

    if ($PSCmdlet.ShouldProcess($ChosenAdapterName, "Configurar IPv4 estático $ChosenIPAddress/$ChosenPrefixLength")) {
        Invoke-Netsh -Arguments @(
            'interface', 'ipv4', 'set', 'address',
            "name=$ChosenAdapterName",
            'source=static',
            "address=$ChosenIPAddress",
            "mask=$subnetMask",
            "gateway=$gatewayValue"
        ) | Out-Null

        if ($ChosenDnsServers.Count -gt 0) {
            Invoke-Netsh -Arguments @(
                'interface', 'ipv4', 'set', 'dnsservers',
                "name=$ChosenAdapterName",
                'source=static',
                "address=$($ChosenDnsServers[0])",
                'register=primary',
                'validate=no'
            ) | Out-Null

            for ($index = 1; $index -lt $ChosenDnsServers.Count; $index++) {
                Invoke-Netsh -Arguments @(
                    'interface', 'ipv4', 'add', 'dnsservers',
                    "name=$ChosenAdapterName",
                    "address=$($ChosenDnsServers[$index])",
                    "index=$($index + 1)",
                    'validate=no'
                ) | Out-Null
            }
        }
    }
    else {
        return [pscustomobject]@{
            AdapterName    = $ChosenAdapterName
            IPAddress      = $ChosenIPAddress
            PrefixLength   = $ChosenPrefixLength
            DefaultGateway = $ChosenDefaultGateway
            DnsServers     = $ChosenDnsServers
        }
    }

    $maxVerificationAttempts = 10
    $verificationRetryDelaySeconds = 2

    for ($attempt = 0; $attempt -lt $maxVerificationAttempts; $attempt++) {
        $currentIp = Get-NetIPAddress -InterfaceIndex $adapter.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -eq $ChosenIPAddress -and $_.PrefixLength -eq $ChosenPrefixLength }

        if ($currentIp) {
            if ($ChosenDefaultGateway) {
                $defaultRoute = Get-NetRoute -InterfaceIndex $adapter.InterfaceIndex -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
                    Where-Object { $_.NextHop -eq $ChosenDefaultGateway }

                if (-not $defaultRoute) {
                    Start-Sleep -Seconds $verificationRetryDelaySeconds
                    continue
                }
            }

            return [pscustomobject]@{
                AdapterName    = $ChosenAdapterName
                IPAddress      = $ChosenIPAddress
                PrefixLength   = $ChosenPrefixLength
                DefaultGateway = $ChosenDefaultGateway
                DnsServers     = $ChosenDnsServers
            }
        }

        Start-Sleep -Seconds $verificationRetryDelaySeconds
    }

    throw "Não foi possível confirmar a alteração do IP para '$ChosenIPAddress/$ChosenPrefixLength' no adaptador '$ChosenAdapterName'."
}

Assert-Windows

if ($ListAdapters) {
    Show-AdapterTable | Out-Null
    return
}

if (-not $AdapterName) {
    $AdapterName = Resolve-AdapterName -SelectedAdapterName $null
}
else {
    $null = Resolve-AdapterName -SelectedAdapterName $AdapterName
}

if (-not $IPAddress) {
    $IPAddress = Read-RequiredValue -Prompt 'Digite o novo IPv4' -Validator { param($value) Test-IPv4Address -Address $value } -ValidationMessage 'Informe um IPv4 válido.'
}
elseif (-not (Test-IPv4Address -Address $IPAddress)) {
    throw 'O endereço IPv4 informado é inválido.'
}

if (-not $PSBoundParameters.ContainsKey('PrefixLength')) {
    $PrefixLength = [int](Read-RequiredValue -Prompt 'Digite o prefixo de rede (ex.: 24)' -Validator {
        param($value)
        $parsed = 0
        [int]::TryParse($value, [ref]$parsed) -and $parsed -ge 0 -and $parsed -le 32
    } -ValidationMessage 'Informe um prefixo entre 0 e 32.')
}

if ([string]::IsNullOrWhiteSpace($DefaultGateway)) {
    $DefaultGateway = Read-Host -Prompt 'Digite o gateway padrão (ou deixe em branco para não configurar)'
}
elseif (-not (Test-IPv4Address -Address $DefaultGateway)) {
    throw 'O gateway informado é inválido.'
}

if ([string]::IsNullOrWhiteSpace($DefaultGateway)) {
    $DefaultGateway = $null
}

if (-not $DnsServers) {
    $dnsInput = Read-Host -Prompt 'Digite os DNS IPv4 separados por vírgula (ou deixe em branco para não configurar)'
    if (-not [string]::IsNullOrWhiteSpace($dnsInput)) {
        $DnsServers = $dnsInput.Split(',').ForEach({ $_.Trim() }) | Where-Object { $_ }
    }
}

if ($DnsServers) {
    foreach ($dnsServer in $DnsServers) {
        if (-not (Test-IPv4Address -Address $dnsServer)) {
            throw "DNS inválido informado: '$dnsServer'."
        }
    }
}

if ($UseCurrentCredential) {
    $result = Set-AdapterIPv4Configuration `
        -ChosenAdapterName $AdapterName `
        -ChosenIPAddress $IPAddress `
        -ChosenPrefixLength $PrefixLength `
        -ChosenDefaultGateway $DefaultGateway `
        -ChosenDnsServers $DnsServers

    $result | Format-List | Out-Host
    return
}

if (-not $Username) {
    $Username = $env:IPCHANGE_ADMIN_USERNAME
}

if (-not $Username) {
    $Username = Read-RequiredValue -Prompt 'Digite o usuário com permissão administrativa'
}

if (-not $Password) {
    if ([string]::IsNullOrWhiteSpace($PlainTextPassword)) {
        $PlainTextPassword = $env:IPCHANGE_ADMIN_PASSWORD
    }

    if (-not [string]::IsNullOrWhiteSpace($PlainTextPassword)) {
        $Password = ConvertTo-SecureString -String $PlainTextPassword -AsPlainText -Force
        $PlainTextPassword = $null
    }
    else {
        $Password = Read-Host -Prompt 'Digite a senha do usuário informado' -AsSecureString
    }
}

$credential = [pscredential]::new($Username, $Password)

Invoke-WithCredential `
    -Credential $credential `
    -ChosenAdapterName $AdapterName `
    -ChosenIPAddress $IPAddress `
    -ChosenPrefixLength $PrefixLength `
    -ChosenDefaultGateway $DefaultGateway `
    -ChosenDnsServers $DnsServers

Write-Host "IP do adaptador '$AdapterName' alterado com sucesso para $IPAddress/$PrefixLength."
