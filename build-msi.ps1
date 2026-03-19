[CmdletBinding()]
param(
    [string]$RuntimeIdentifier = 'win-x64',
    [string]$Configuration = 'Release',
    [switch]$SkipPublish,
    [switch]$AllowExistingMsiFallback = $true
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectFile = Join-Path $root 'ipchange.csproj'
$publishDir = Join-Path $root "artifacts\release\$RuntimeIdentifier"
$publishedExe = Join-Path $publishDir 'ipchange.exe'
$installerProject = Join-Path $root 'installer\ipchange.installer.wixproj'
$installerOutputDir = Join-Path $root "artifacts\installer\msi\$RuntimeIdentifier"
$installerBuildOutputDir = Join-Path $root 'installer\bin\Release'
$builtInstallerMsi = Join-Path $installerBuildOutputDir 'ipchange-installer.msi'

[xml]$projectXml = Get-Content -LiteralPath $projectFile
$projectVersion = [string]($projectXml.Project.PropertyGroup.Version | Select-Object -First 1)
if ([string]::IsNullOrWhiteSpace($projectVersion)) {
    $projectVersion = '0.0.0'
}

function Get-MsiProductVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseVersion,

        [datetime]$Now = (Get-Date)
    )

    $versionMatch = [regex]::Match($BaseVersion, '^(?<major>\d+)\.(?<minor>\d+)')
    if (-not $versionMatch.Success) {
        return '1.0.0'
    }

    $major = [int]$versionMatch.Groups['major'].Value
    $minor = [int]$versionMatch.Groups['minor'].Value
    $minor = [Math]::Min($minor, 255)

    $quarterHourBuild = (($Now.DayOfYear - 1) * 96) + ($Now.Hour * 4) + [int][Math]::Floor($Now.Minute / 15)
    return "$major.$minor.$quarterHourBuild"
}

$msiProductVersion = Get-MsiProductVersion -BaseVersion $projectVersion

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$artifactLabel = "v$projectVersion-$timestamp"
$installerMsi = Join-Path $installerOutputDir "ipchange-installer-$artifactLabel.msi"

function Get-LatestArtifact {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    return Get-ChildItem -Path $installerOutputDir -Filter $Pattern -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTimeUtc -Descending |
        Select-Object -First 1
}

if (-not $SkipPublish) {
    dotnet publish (Join-Path $root 'ipchange.csproj') `
        -c $Configuration `
        -r $RuntimeIdentifier `
        --self-contained true `
        -p:PublishSingleFile=true `
        -p:EnableCompressionInSingleFile=true `
        -p:DebugType=None `
        -p:DebugSymbols=false `
        -o $publishDir
}

if (-not (Test-Path -LiteralPath $publishedExe -PathType Leaf)) {
    throw "Executavel publicado nao encontrado em: $publishedExe"
}

New-Item -ItemType Directory -Path $installerOutputDir -Force | Out-Null

& dotnet build $installerProject `
    -c $Configuration `
    -p:AppVersion=$msiProductVersion `
    -p:PublishedExePath=$publishedExe

$msiBuildSucceeded = $LASTEXITCODE -eq 0

if (-not $msiBuildSucceeded) {
    $latestInstaller = Get-LatestArtifact -Pattern 'ipchange-installer-*.msi'
    if ($AllowExistingMsiFallback -and $null -ne $latestInstaller) {
        Write-Warning "Falha ao gerar um novo MSI nesta maquina. O ultimo MSI valido sera mantido em: $($latestInstaller.FullName)"
    }
    else {
        throw "Falha ao gerar o MSI e nenhum arquivo versionado existente foi encontrado em: $installerOutputDir"
    }
}
elseif (-not (Test-Path -LiteralPath $builtInstallerMsi -PathType Leaf)) {
    throw "O build do WiX terminou sem gerar o arquivo esperado em: $builtInstallerMsi"
}

if ($msiBuildSucceeded) {
    Copy-Item -LiteralPath $builtInstallerMsi -Destination $installerMsi -Force
}

if ($msiBuildSucceeded) {
    Write-Host "MSI versionado gerado em: $installerMsi"
    Write-Host "Versao interna do MSI: $msiProductVersion"
}
