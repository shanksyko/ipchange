namespace NetworkConfigurator;

internal static class NetworkConfiguratorServiceProtocol
{
  public const string PipeName = "network-configurator-service";
  public const string ServiceName = "NetworkConfigurator";
}

internal sealed class ServiceApplyRequest
{
  public ApplyRequest Request { get; set; } = new();
}

internal sealed class ServiceApplyResponse
{
  public bool Success { get; set; }
  public AdapterConfigurationResult? Result { get; set; }
  public string? ErrorMessage { get; set; }
}

internal sealed class NetworkProfile
{
  public string Name { get; set; } = string.Empty;
  public string AdapterName { get; set; } = string.Empty;
  public bool UseDhcp { get; set; }
  public string IPAddress { get; set; } = string.Empty;
  public int PrefixLength { get; set; } = 24;
  public string DefaultGateway { get; set; } = string.Empty;
  public string DnsServers { get; set; } = string.Empty;
}