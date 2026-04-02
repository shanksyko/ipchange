using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Json;

namespace NetworkConfigurator;

internal static class AppJson
{
    public static readonly JsonSerializerOptions Options = new() { WriteIndented = true };
    public static readonly JsonSerializerOptions CompactOptions = new() { WriteIndented = false };
}

internal static class AppLogger
{
    private static readonly object Sync = new();
    private static string _path = BuildDefaultLogPath();
    private static string? _securedDirectory;

    public static string CurrentPath => _path;

    public static void SetPath(string path)
    {
        if (!string.IsNullOrWhiteSpace(path))
        {
            _path = path;
        }
    }

    public static void Write(string level, string message)
    {
        try
        {
            lock (Sync)
            {
                EnsureDirectoryReady();
                File.AppendAllText(
                    _path,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}",
                    Encoding.UTF8);
            }
        }
        catch
        {
        }
    }

    private static string BuildDefaultLogPath()
    {
        var baseDirectory = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        if (string.IsNullOrWhiteSpace(baseDirectory))
        {
            baseDirectory = Path.GetTempPath();
        }

        return Path.Combine(baseDirectory, "network-configurator", "secure-logs", $"network-configurator-{DateTime.Now:yyyyMMdd}.txt");
    }

    private static void EnsureDirectoryReady()
    {
        var directory = Path.GetDirectoryName(_path)
            ?? throw new InvalidOperationException("Diretório de log inválido.");

        Directory.CreateDirectory(directory);

        if (!OperatingSystem.IsWindows() ||
            !NetworkConfigurationService.GetIsAdministrator() ||
            string.Equals(_securedDirectory, directory, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var security = new DirectorySecurity();
        security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);

        var inheritanceFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;
        security.AddAccessRule(new FileSystemAccessRule(
            new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null),
            FileSystemRights.FullControl,
            inheritanceFlags,
            PropagationFlags.None,
            AccessControlType.Allow));

        security.AddAccessRule(new FileSystemAccessRule(
            new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null),
            FileSystemRights.FullControl,
            inheritanceFlags,
            PropagationFlags.None,
            AccessControlType.Allow));

        var directoryInfo = new DirectoryInfo(directory);
        directoryInfo.SetAccessControl(security);
        _securedDirectory = directory;
    }
}

internal static class NetworkConfigurationService
{
    public static IReadOnlyList<AdapterInfo> GetVisibleAdapters()
    {
        return GetVisibleNetworkInterfaces()
            .Select(CreateAdapterInfo)
            .OrderBy(adapter => adapter.InterfaceIndex < 0 ? int.MaxValue : adapter.InterfaceIndex)
            .ThenBy(adapter => adapter.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static async Task<AdapterConfigurationResult> ApplyConfigurationAsync(ApplyRequest request)
    {
        ValidateRequest(request);

        var networkInterface = FindVisibleNetworkInterface(request.AdapterName!)
            ?? throw new InvalidOperationException($"Adaptador '{request.AdapterName}' não foi encontrado.");
        var adapter = CreateAdapterInfo(networkInterface);

        if (adapter.InterfaceIndex < 0)
        {
            throw new InvalidOperationException($"O adaptador '{adapter.Name}' não possui InterfaceIndex IPv4 utilizável.");
        }

        if (request.UseDhcp)
        {
            await InvokeNetshAsync(new[]
            {
                "interface", "ipv4", "set", "address",
                $"name={adapter.Name}",
                "source=dhcp"
            });

            await InvokeNetshAsync(new[]
            {
                "interface", "ipv4", "set", "dnsservers",
                $"name={adapter.Name}",
                "source=dhcp"
            });

            return await VerifyDhcpConfigurationAsync(adapter);
        }

        var ipAddress = request.IPAddress!;
        var prefixLength = request.PrefixLength ?? throw new InvalidOperationException("Informe um prefixo entre 0 e 32.");
        var subnetMask = ConvertToSubnetMask(prefixLength);
        var gatewayValue = string.IsNullOrWhiteSpace(request.DefaultGateway) ? "none" : request.DefaultGateway;

        await InvokeNetshAsync(new[]
        {
            "interface", "ipv4", "set", "address",
            $"name={adapter.Name}",
            "source=static",
            $"address={ipAddress}",
            $"mask={subnetMask}",
            $"gateway={gatewayValue}"
        });

        if (request.DnsServers.Count > 0)
        {
            await InvokeNetshAsync(new[]
            {
                "interface", "ipv4", "set", "dnsservers",
                $"name={adapter.Name}",
                "source=static",
                $"address={request.DnsServers[0]}",
                "register=primary",
                "validate=no"
            });

            for (var index = 1; index < request.DnsServers.Count; index++)
            {
                await InvokeNetshAsync(new[]
                {
                    "interface", "ipv4", "add", "dnsservers",
                    $"name={adapter.Name}",
                    $"address={request.DnsServers[index]}",
                    $"index={index + 1}",
                    "validate=no"
                });
            }
        }

        return await VerifyStaticConfigurationAsync(adapter, ipAddress, prefixLength, request.DefaultGateway);
    }

    public static bool GetIsAdministrator()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    public static void ValidateRequest(ApplyRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.AdapterName))
        {
            throw new InvalidOperationException("Selecione um adaptador.");
        }

        if (request.UseDhcp)
        {
            return;
        }

        if (!IsValidIPv4(request.IPAddress))
        {
            throw new InvalidOperationException("Informe um IPv4 válido.");
        }

        if (!request.PrefixLength.HasValue || request.PrefixLength.Value is < 0 or > 32)
        {
            throw new InvalidOperationException("Informe um prefixo entre 0 e 32.");
        }

        if (!string.IsNullOrWhiteSpace(request.DefaultGateway) && !IsValidIPv4(request.DefaultGateway))
        {
            throw new InvalidOperationException("O gateway informado é inválido.");
        }

        foreach (var dnsServer in request.DnsServers)
        {
            if (!IsValidIPv4(dnsServer))
            {
                throw new InvalidOperationException($"DNS inválido informado: '{dnsServer}'.");
            }
        }
    }

    private static bool IsValidIPv4(string? value) => IPAddress.TryParse(value, out var address) && address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork;

    private static async Task InvokeNetshAsync(IEnumerable<string> arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "netsh",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        AppLogger.Write("INFO", $"Executando netsh: {string.Join(' ', arguments)}");

        using var process = Process.Start(startInfo);
        if (process is null)
        {
            throw new InvalidOperationException("Não foi possível iniciar o netsh.");
        }

        var standardOutputTask = process.StandardOutput.ReadToEndAsync();
        var standardErrorTask = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();

        var standardOutput = await standardOutputTask;
        var standardError = await standardErrorTask;
        if (process.ExitCode != 0)
        {
            var output = string.IsNullOrWhiteSpace(standardError) ? standardOutput : standardError;
            AppLogger.Write("ERROR", $"Falha no netsh: {output.Trim()}");

            if (IsPermissionRelatedFailure(output))
            {
                throw new InvalidOperationException("Não foi possível alterar o IP com a credencial configurada.");
            }

            throw new InvalidOperationException(output.Trim());
        }
    }

    private static bool IsPermissionRelatedFailure(string? output)
    {
        if (string.IsNullOrWhiteSpace(output))
        {
            return false;
        }

        return output.Contains("requires elevation", StringComparison.OrdinalIgnoreCase)
            || output.Contains("run as administrator", StringComparison.OrdinalIgnoreCase)
            || output.Contains("access is denied", StringComparison.OrdinalIgnoreCase)
            || output.Contains("acesso negado", StringComparison.OrdinalIgnoreCase)
            || output.Contains("privilégios administrativos", StringComparison.OrdinalIgnoreCase)
            || output.Contains("privilegios administrativos", StringComparison.OrdinalIgnoreCase);
    }

    private static async Task<AdapterConfigurationResult> VerifyStaticConfigurationAsync(AdapterInfo adapter, string ipAddress, int prefixLength, string? defaultGateway)
    {
        for (var attempt = 0; attempt < 12; attempt++)
        {
            await Task.Delay(TimeSpan.FromMilliseconds(500));

            var networkInterface = NetworkInterface.GetAllNetworkInterfaces()
                .FirstOrDefault(item => string.Equals(item.Name, adapter.Name, StringComparison.OrdinalIgnoreCase));
            if (networkInterface is null)
            {
                continue;
            }

            var properties = networkInterface.GetIPProperties();
            var ipMatch = properties.UnicastAddresses.Any(address =>
                address.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork &&
                address.Address.ToString() == ipAddress &&
                address.PrefixLength == prefixLength);

            if (!ipMatch)
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(defaultGateway))
            {
                var gatewayMatch = properties.GatewayAddresses.Any(address => address.Address.ToString() == defaultGateway);
                if (!gatewayMatch)
                {
                    continue;
                }
            }

            return BuildConfigurationResult(networkInterface, useDhcp: false);
        }

        throw new InvalidOperationException($"Não foi possível confirmar a alteração do IP para '{ipAddress}/{prefixLength}' no adaptador '{adapter.Name}'.");
    }

    private static async Task<AdapterConfigurationResult> VerifyDhcpConfigurationAsync(AdapterInfo adapter)
    {
        for (var attempt = 0; attempt < 12; attempt++)
        {
            await Task.Delay(TimeSpan.FromMilliseconds(500));

            var networkInterface = NetworkInterface.GetAllNetworkInterfaces()
                .FirstOrDefault(item => string.Equals(item.Name, adapter.Name, StringComparison.OrdinalIgnoreCase));
            if (networkInterface is null)
            {
                continue;
            }

            var ipv4Properties = networkInterface.GetIPProperties().GetIPv4Properties();
            if (ipv4Properties?.IsDhcpEnabled == true)
            {
                return BuildConfigurationResult(networkInterface, useDhcp: true);
            }
        }

        throw new InvalidOperationException($"Não foi possível confirmar a ativação do DHCP no adaptador '{adapter.Name}'.");
    }

    private static AdapterConfigurationResult BuildConfigurationResult(NetworkInterface networkInterface, bool useDhcp)
    {
        var properties = networkInterface.GetIPProperties();
        var ipv4Address = properties.UnicastAddresses
            .FirstOrDefault(address => address.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);

        var gateway = properties.GatewayAddresses
            .FirstOrDefault(address => address.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
            ?.Address
            .ToString();

        var dnsServers = properties.DnsAddresses
            .Where(address => address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
            .Select(address => address.ToString())
            .ToArray();

        return new AdapterConfigurationResult(
            networkInterface.Name,
            useDhcp,
            ipv4Address?.Address.ToString(),
            ipv4Address?.PrefixLength,
            gateway,
            dnsServers);
    }

    private static IEnumerable<NetworkInterface> GetVisibleNetworkInterfaces()
    {
        return NetworkInterface.GetAllNetworkInterfaces()
            .Where(networkInterface => networkInterface.NetworkInterfaceType != NetworkInterfaceType.Loopback);
    }

    private static NetworkInterface? FindVisibleNetworkInterface(string adapterName)
    {
        return GetVisibleNetworkInterfaces()
            .FirstOrDefault(item => string.Equals(item.Name, adapterName, StringComparison.OrdinalIgnoreCase));
    }

    private static AdapterInfo CreateAdapterInfo(NetworkInterface networkInterface)
    {
        int? interfaceIndex = null;
        string? currentIPAddress = null;
        int? currentPrefixLength = null;
        string? currentDefaultGateway = null;
        string[] currentDnsServers = Array.Empty<string>();
        var isDhcpEnabled = false;

        try
        {
            var properties = networkInterface.GetIPProperties();
            var ipv4Props = properties.GetIPv4Properties();
            interfaceIndex = ipv4Props?.Index;
            isDhcpEnabled = ipv4Props?.IsDhcpEnabled ?? false;

            var unicastAddr = properties.UnicastAddresses
                .FirstOrDefault(a => a.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
            currentIPAddress = unicastAddr?.Address.ToString();
            currentPrefixLength = unicastAddr?.PrefixLength;

            currentDefaultGateway = properties.GatewayAddresses
                .FirstOrDefault(g => g.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                ?.Address.ToString();

            currentDnsServers = properties.DnsAddresses
                .Where(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                .Select(a => a.ToString())
                .ToArray();
        }
        catch
        {
        }

        return new AdapterInfo(
            interfaceIndex ?? -1,
            networkInterface.Name,
            networkInterface.Description,
            networkInterface.OperationalStatus.ToString(),
            FormatMacAddress(networkInterface.GetPhysicalAddress()),
            FormatLinkSpeed(networkInterface.Speed),
            currentIPAddress,
            currentPrefixLength,
            currentDefaultGateway,
            currentDnsServers,
            isDhcpEnabled);
    }

    private static string ConvertToSubnetMask(int prefixLength)
    {
        var mask = prefixLength == 0 ? 0u : uint.MaxValue << (32 - prefixLength);
        return new IPAddress(BitConverter.GetBytes((int)IPAddress.HostToNetworkOrder((int)mask))).ToString();
    }

    private static string FormatMacAddress(PhysicalAddress address)
    {
        var bytes = address.GetAddressBytes();
        return bytes.Length == 0 ? string.Empty : string.Join('-', bytes.Select(item => item.ToString("X2")));
    }

    private static string FormatLinkSpeed(long speedBitsPerSecond)
    {
        if (speedBitsPerSecond <= 0)
        {
            return string.Empty;
        }

        if (speedBitsPerSecond >= 1_000_000_000)
        {
            return $"{speedBitsPerSecond / 1_000_000_000d:0.##} Gbps";
        }

        if (speedBitsPerSecond >= 1_000_000)
        {
            return $"{speedBitsPerSecond / 1_000_000d:0.##} Mbps";
        }

        return $"{speedBitsPerSecond} bps";
    }
}

internal sealed class ApplyRequest
{
    public string? AdapterName { get; set; }
    public bool UseDhcp { get; set; }
    public string? IPAddress { get; set; }
    public int? PrefixLength { get; set; }
    public string? DefaultGateway { get; set; }
    public List<string> DnsServers { get; } = new();
    public string? LogPath { get; set; }
    public string? ResultPath { get; set; }
    public string? ErrorPath { get; set; }
}

internal sealed record AdapterInfo(
    int InterfaceIndex,
    string Name,
    string InterfaceDescription,
    string Status,
    string MacAddress,
    string LinkSpeed,
    string? CurrentIPAddress,
    int? CurrentPrefixLength,
    string? CurrentDefaultGateway,
    string[] CurrentDnsServers,
    bool IsDhcpEnabled)
{
    public string DisplayName => string.IsNullOrWhiteSpace(CurrentIPAddress)
        ? Name
        : IsDhcpEnabled
            ? $"{Name}  [{CurrentIPAddress}/{CurrentPrefixLength} – DHCP]"
            : $"{Name}  [{CurrentIPAddress}/{CurrentPrefixLength}]";
}
internal sealed record AdapterConfigurationResult(string AdapterName, bool UseDhcp, string? IPAddress, int? PrefixLength, string? DefaultGateway, string[] DnsServers);
internal sealed record RouteEntry(string Destination, string Mask, string Gateway, string Interface, int Metric);
internal sealed record ArpEntry(string Interface, string IPAddress, string MacAddress, string Type);
