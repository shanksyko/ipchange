using System.Diagnostics;

namespace ipchange;

internal static class Program
{
    private static int Main(string[] args)
    {
        if (!OperatingSystem.IsWindows())
        {
            Console.Error.WriteLine("Este aplicativo C# precisa ser executado no Windows para chamar o script PowerShell de configuração de IP.");
            return 1;
        }

        var scriptPath = Path.Combine(AppContext.BaseDirectory, "ipchange.ps1");
        if (!File.Exists(scriptPath))
        {
            Console.Error.WriteLine($"Script PowerShell não encontrado em '{scriptPath}'.");
            return 1;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = false
        };

        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-ExecutionPolicy");
        startInfo.ArgumentList.Add("Bypass");
        startInfo.ArgumentList.Add("-File");
        startInfo.ArgumentList.Add(scriptPath);

        foreach (var arg in args)
        {
            startInfo.ArgumentList.Add(arg);
        }

        using var process = Process.Start(startInfo);
        if (process is null)
        {
            Console.Error.WriteLine("Não foi possível iniciar o PowerShell.");
            return 1;
        }

        process.WaitForExit();
        return process.ExitCode;
    }
}
