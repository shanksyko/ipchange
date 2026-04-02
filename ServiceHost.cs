using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace NetworkConfigurator;

internal static class NetworkConfiguratorServiceRuntime
{
  public static async Task<int> RunAsync(string[] args)
  {
    var filteredArgs = args.Where(arg => !string.Equals(arg, "--service", StringComparison.OrdinalIgnoreCase)).ToArray();
    var builder = Host.CreateApplicationBuilder(filteredArgs);
    builder.Services.AddWindowsService(options => options.ServiceName = NetworkConfiguratorServiceProtocol.ServiceName);
    builder.Services.AddHostedService<NetworkConfiguratorWorkerService>();

    using var host = builder.Build();
    await host.RunAsync();
    return 0;
  }
}

internal sealed class NetworkConfiguratorWorkerService(ILogger<NetworkConfiguratorWorkerService> logger) : BackgroundService
{
  private readonly ILogger<NetworkConfiguratorWorkerService> _logger = logger;

protected override async Task ExecuteAsync(CancellationToken stoppingToken)
{
  AppLogger.Write("INFO", "Serviço Network Configurator iniciado.");

  while (!stoppingToken.IsCancellationRequested)
  {
    using var pipe = CreateServerPipe();

    try
    {
      await pipe.WaitForConnectionAsync(stoppingToken);
      await HandleClientAsync(pipe, stoppingToken);
    }
    catch (OperationCanceledException)
    {
      break;
    }
    catch (Exception ex)
    {
      _logger.LogError(ex, "Falha no serviço Network Configurator.");
      AppLogger.Write("ERROR", $"Falha no serviço Network Configurator: {ex.Message}");
      await Task.Delay(1000, stoppingToken);
    }
  }
}

private static NamedPipeServerStream CreateServerPipe()
{
  var security = new PipeSecurity();
  security.AddAccessRule(new PipeAccessRule(
      new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null),
      PipeAccessRights.ReadWrite | PipeAccessRights.CreateNewInstance,
      AccessControlType.Allow));

  security.AddAccessRule(new PipeAccessRule(
      new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null),
      PipeAccessRights.FullControl,
      AccessControlType.Allow));

  return NamedPipeServerStreamAcl.Create(
      NetworkConfiguratorServiceProtocol.PipeName,
      PipeDirection.InOut,
      1,
      PipeTransmissionMode.Byte,
      PipeOptions.Asynchronous,
      4096,
      4096,
      security);
}

private static async Task HandleClientAsync(NamedPipeServerStream pipe, CancellationToken cancellationToken)
{
  using var reader = new StreamReader(pipe, Encoding.UTF8, false, 1024, leaveOpen: true);
  using var writer = new StreamWriter(pipe, new UTF8Encoding(false), 1024, leaveOpen: true) { AutoFlush = true };

  var line = await reader.ReadLineAsync(cancellationToken);
  if (string.IsNullOrWhiteSpace(line))
  {
    return;
  }

  ServiceApplyResponse response;

  try
  {
    var request = JsonSerializer.Deserialize<ServiceApplyRequest>(line, AppJson.CompactOptions)
      ?? throw new InvalidOperationException("A requisição enviada ao serviço é inválida.");

    AppLogger.Write(
        "INFO",
        request.Request.UseDhcp
            ? $"Solicitação recebida do cliente. Adaptador='{request.Request.AdapterName}', Modo='dhcp'."
            : $"Solicitação recebida do cliente. Adaptador='{request.Request.AdapterName}', IP='{request.Request.IPAddress}', Prefixo='{request.Request.PrefixLength}', Gateway='{request.Request.DefaultGateway ?? "none"}', DNS='{string.Join(", ", request.Request.DnsServers)}'.");

    var result = await NetworkConfigurationService.ApplyConfigurationAsync(request.Request);
    AppLogger.Write(
        "INFO",
        result.UseDhcp
            ? $"DHCP aplicado com sucesso. Adaptador='{result.AdapterName}', IP='{result.IPAddress ?? "pending"}/{result.PrefixLength?.ToString() ?? "-"}', Gateway='{result.DefaultGateway ?? "none"}', DNS='{string.Join(", ", result.DnsServers)}'."
            : $"Configuração aplicada com sucesso. Adaptador='{result.AdapterName}', IP='{result.IPAddress}/{result.PrefixLength}', Gateway='{result.DefaultGateway ?? "none"}', DNS='{string.Join(", ", result.DnsServers)}'.");
    response = new ServiceApplyResponse
    {
      Success = true,
      Result = result
    };
  }
  catch (Exception ex)
  {
    AppLogger.Write("ERROR", ex.Message);
    response = new ServiceApplyResponse
    {
      Success = false,
      ErrorMessage = ex.Message
    };
  }

  var payload = JsonSerializer.Serialize(response, AppJson.CompactOptions);
  await writer.WriteLineAsync(payload);
}
}