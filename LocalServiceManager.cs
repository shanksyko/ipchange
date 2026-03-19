using System.Diagnostics;

namespace ipchange;

internal static class LocalServiceManager
{
  public static async Task<LocalServiceState> GetStatusAsync()
  {
    var result = await RunScAsync($"query {IpChangeServiceProtocol.ServiceName}", allowFailure: true);
    if (result.ExitCode == 1060 || result.Output.Contains("does not exist", StringComparison.OrdinalIgnoreCase))
    {
      return LocalServiceState.NotInstalled;
    }

    if (result.ExitCode != 0)
    {
      throw new InvalidOperationException(string.IsNullOrWhiteSpace(result.Output)
          ? "Não foi possível consultar o serviço local do ipchange."
          : result.Output);
    }

    if (result.Output.Contains("RUNNING", StringComparison.OrdinalIgnoreCase))
    {
      return LocalServiceState.Running;
    }

    return LocalServiceState.Installed;
  }

  private static async Task<ScCommandResult> RunScAsync(string arguments, bool allowFailure)
  {
    var startInfo = new ProcessStartInfo
    {
      FileName = "sc.exe",
      Arguments = arguments,
      UseShellExecute = false,
      RedirectStandardOutput = true,
      RedirectStandardError = true,
      CreateNoWindow = true
    };

    using var process = Process.Start(startInfo);
    if (process is null)
    {
      throw new InvalidOperationException("Não foi possível iniciar o sc.exe.");
    }

    var standardOutputTask = process.StandardOutput.ReadToEndAsync();
    var standardErrorTask = process.StandardError.ReadToEndAsync();
    await process.WaitForExitAsync();

    var standardOutput = (await standardOutputTask).Trim();
    var standardError = (await standardErrorTask).Trim();
    var output = string.IsNullOrWhiteSpace(standardError) ? standardOutput : standardError;
    AppLogger.Write("INFO", $"sc.exe {arguments}");

    if (process.ExitCode != 0 && !allowFailure)
    {
      throw new InvalidOperationException(string.IsNullOrWhiteSpace(output)
          ? "Falha ao executar o sc.exe."
          : output);
    }

    return new ScCommandResult(process.ExitCode, output);
  }

  private readonly record struct ScCommandResult(int ExitCode, string Output);
}

internal enum LocalServiceState
{
  NotInstalled,
  Installed,
  Running
}