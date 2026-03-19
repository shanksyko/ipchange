namespace ipchange;

internal static class IpChangeServiceProtocol
{
  public const string PipeName = "ipchange-service";
  public const string ServiceName = "Ipchange";
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