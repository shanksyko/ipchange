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
    [switch]$UseCurrentCredential,

    [Parameter()]
    [switch]$ListAdaptersJson,

    [Parameter()]
    [switch]$DiagnosticsJson,

    [Parameter()]
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$DefaultUsername = '.\support'
$LocalCredentialDefaultsPath = Join-Path -Path $PSScriptRoot -ChildPath 'ipchange.local.psd1'

function Get-DefaultLogPath {
    $baseDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    if ([string]::IsNullOrWhiteSpace($baseDirectory)) {
        $baseDirectory = [System.IO.Path]::GetTempPath()
    }

    return Join-Path -Path $baseDirectory -ChildPath ('ipchange\logs\ipchange-{0:yyyyMMdd}.log' -f (Get-Date))
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Get-DefaultLogPath
}

$script:CurrentLogPath = $LogPath

function Initialize-LogFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $directory = Split-Path -Path $Path -Parent
        if (-not [string]::IsNullOrWhiteSpace($directory)) {
            $null = New-Item -ItemType Directory -Path $directory -Force
        }

        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            $null = New-Item -ItemType File -Path $Path -Force
        }
    }
    catch {
        # Não interrompe a execução por falha de log.
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    if ([string]::IsNullOrWhiteSpace($script:CurrentLogPath)) {
        return
    }

    try {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        Add-Content -LiteralPath $script:CurrentLogPath -Encoding UTF8 -Value "[$timestamp] [$Level] $Message"
    }
    catch {
        # Não interrompe a execução por falha de log.
    }
}

function Get-IsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [Security.Principal.WindowsPrincipal]::new($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Log -Level 'WARN' -Message "Não foi possível verificar privilégios administrativos: $($_.Exception.Message)"
        return $false
    }
}

function Get-CurrentIdentityName {
    try {
        return [Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    catch {
        return [Environment]::UserName
    }
}

function Assert-Windows {
    if (-not ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT)) {
        Write-Log -Level 'ERROR' -Message 'Execução bloqueada fora do Windows.'
        throw 'Este script só pode ser executado no Windows, pois depende de Get-NetAdapter/Get-NetIPAddress e netsh.'
    }
}

function Assert-Administrator {
    if (-not (Get-IsAdministrator)) {
        Write-Log -Level 'ERROR' -Message "Permissão insuficiente para o usuário atual '$([Environment]::UserName)'."
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

    Write-Log -Message 'Consultando adaptadores visíveis.'

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

function Get-LocalCredentialDefaults {
    if (-not (Test-Path -LiteralPath $LocalCredentialDefaultsPath -PathType Leaf)) {
        Write-Log -Message 'Arquivo local de credenciais padrão não encontrado.'
        return @{}
    }

    Write-Log -Message "Carregando credenciais padrão locais de '$LocalCredentialDefaultsPath'."
    $settings = Import-PowerShellDataFile -Path $LocalCredentialDefaultsPath
    if ($null -eq $settings) {
        return @{}
    }

    return $settings
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

    if ($script:CurrentLogPath) {
        $argumentList += @('-LogPath', $script:CurrentLogPath)
    }

    Write-Log -Message "Reiniciando o script com a credencial '$($Credential.UserName)'."

    try {
        $process = Start-Process -FilePath (Get-ProcessExecutable) -Credential $Credential -ArgumentList $argumentList -Wait -PassThru
    }
    catch {
        Write-Log -Level 'ERROR' -Message "Falha ao iniciar o processo com credenciais alternativas: $($_.Exception.Message)"
        throw "Falha ao iniciar o PowerShell com a credencial '$($Credential.UserName)'. Verifique se a conta tem permissão administrativa e se a senha está correta. Detalhes: $($_.Exception.Message)"
    }

    if ($process.ExitCode -ne 0) {
        Write-Log -Level 'ERROR' -Message "A execução com a credencial '$($Credential.UserName)' retornou código $($process.ExitCode)."
        throw "A execução com a credencial informada falhou com código de saída $($process.ExitCode)."
    }
}

function Invoke-Netsh {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    Write-Log -Message ("Executando netsh: {0}" -f ($Arguments -join ' '))
    $output = & netsh @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Log -Level 'ERROR' -Message ("Falha no netsh: {0}" -f (($output | Out-String).Trim()))
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

    Write-Log -Message "Solicitada configuração IPv4 para o adaptador '$ChosenAdapterName' com IP '$ChosenIPAddress/$ChosenPrefixLength'."

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
        Write-Log -Level 'DEBUG' -Message "Validando configuração aplicada, tentativa $($attempt + 1) de $maxVerificationAttempts."
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

            Write-Log -Message "Configuração confirmada para '$ChosenAdapterName'."
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

function Get-PermissionDiagnostics {
    $adapterEnumerationError = $null
    $canEnumerateAdapters = $false

    if ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT) {
        try {
            $null = Get-VisibleAdapters
            $canEnumerateAdapters = $true
        }
        catch {
            $adapterEnumerationError = $_.Exception.Message
        }
    }

    return [pscustomobject]@{
        OSSupported                 = ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT)
        CurrentUser                 = Get-CurrentIdentityName
        IsAdministrator             = Get-IsAdministrator
        PowerShellEdition           = $PSVersionTable.PSEdition
        PowerShellVersion           = $PSVersionTable.PSVersion.ToString()
        DefaultUsername             = $DefaultUsername
        LocalCredentialDefaultsPath = $LocalCredentialDefaultsPath
        LocalCredentialDefaultsUsed = (Test-Path -LiteralPath $LocalCredentialDefaultsPath -PathType Leaf)
        CanEnumerateAdapters        = $canEnumerateAdapters
        AdapterEnumerationError     = $adapterEnumerationError
        LogPath                     = $script:CurrentLogPath
    }
}

Initialize-LogFile -Path $script:CurrentLogPath
Write-Log -Message "Iniciando execução do ipchange.ps1. Usuário atual: '$(Get-CurrentIdentityName)'."

try {
    if ($DiagnosticsJson) {
        Write-Log -Message 'Gerando diagnóstico de permissões em JSON.'
        Get-PermissionDiagnostics | ConvertTo-Json -Depth 4
        return
    }

    Assert-Windows

    if ($ListAdaptersJson) {
        Write-Log -Message 'Listando adaptadores em JSON.'
        Get-VisibleAdapters | ConvertTo-Json -Depth 4
        return
    }

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
        Write-Log -Message "Aplicando configuração diretamente com a credencial atual ao adaptador '$AdapterName'."

        $result = Set-AdapterIPv4Configuration `
            -ChosenAdapterName $AdapterName `
            -ChosenIPAddress $IPAddress `
            -ChosenPrefixLength $PrefixLength `
            -ChosenDefaultGateway $DefaultGateway `
            -ChosenDnsServers $DnsServers

        $result | Format-List | Out-Host
        return
    }

    # Use IPCHANGE_ADMIN_USERNAME/IPCHANGE_ADMIN_PASSWORD apenas na sessão/processo atual e limpe essas variáveis após o uso.
    $localCredentialDefaults = Get-LocalCredentialDefaults

    if (-not $Username) {
        $Username = $env:IPCHANGE_ADMIN_USERNAME
    }

    if (-not $Username) {
        $Username = $localCredentialDefaults.Username
    }

    if (-not $Username) {
        $Username = $DefaultUsername
    }

    if (-not $Password) {
        if ([string]::IsNullOrWhiteSpace($PlainTextPassword)) {
            $PlainTextPassword = $env:IPCHANGE_ADMIN_PASSWORD
        }

        if ([string]::IsNullOrWhiteSpace($PlainTextPassword)) {
            $PlainTextPassword = $localCredentialDefaults.PlainTextPassword
        }

        if (-not [string]::IsNullOrWhiteSpace($PlainTextPassword)) {
            Write-Warning 'A senha administrativa está sendo usada em texto puro. Prefira -Password com SecureString quando isso for possível.'
            Write-Log -Level 'WARN' -Message "Senha administrativa recebida em texto puro para o usuário '$Username'."
            $Password = ConvertTo-SecureString -String $PlainTextPassword -AsPlainText -Force
        }
        else {
            Write-Log -Message "Solicitando a senha do usuário '$Username'."
            $Password = Read-Host -Prompt "Digite a senha do usuário $Username" -AsSecureString
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

    Write-Log -Message "Fluxo concluído com sucesso para o adaptador '$AdapterName'."
    Write-Host "IP do adaptador '$AdapterName' alterado com sucesso para $IPAddress/$PrefixLength."
}
catch {
    Write-Log -Level 'ERROR' -Message $_.Exception.Message
    throw
}
