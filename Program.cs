using System.Diagnostics;
using System.Text;

namespace ipchange;

internal static class Program
{
    private static readonly object LogSync = new();
    private static readonly string LogPath = BuildLogPath();

    private static async Task<int> Main(string[] args)
    {
        var forceUi = HasArgument(args, "--ui");
        var filteredArgs = args.Where(arg => !string.Equals(arg, "--cli", StringComparison.OrdinalIgnoreCase) &&
                                             !string.Equals(arg, "--ui", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        if (forceUi)
        {
            Console.WriteLine("A interface HTTP/web foi removida. Iniciando o aplicativo .exe normal no console.");
        }

        return await RunCliAsync(filteredArgs);
    }

    private static async Task<int> RunCliAsync(string[] args)
    {
        if (!OperatingSystem.IsWindows())
        {
            Console.Error.WriteLine("Este aplicativo em C# precisa ser executado no Windows para alterar o IP.");
            return 1;
        }

        var runtime = ResolveRuntime();
        if (!runtime.IsValid)
        {
            Console.Error.WriteLine(runtime.ErrorMessage);
            return 1;
        }

        var startInfo = CreateBaseStartInfo(runtime);
        foreach (var arg in args)
        {
            startInfo.ArgumentList.Add(arg);
        }

        if (!HasArgument(args, "-LogPath"))
        {
            startInfo.ArgumentList.Add("-LogPath");
            startInfo.ArgumentList.Add(LogPath);
        }

        Log("INFO", $"CLI iniciada com argumentos: {SanitizeArguments(startInfo.ArgumentList)}");

        try
        {
            using var process = Process.Start(startInfo);
            if (process is null)
            {
                Console.Error.WriteLine("Não foi possível iniciar o PowerShell.");
                Log("ERROR", "O PowerShell não pôde ser iniciado no modo CLI.");
                return 1;
            }

            await process.WaitForExitAsync();
            Log("INFO", $"CLI finalizada com código {process.ExitCode}.");

            if (process.ExitCode != 0)
            {
                Console.Error.WriteLine($"Falha ao executar o script. Consulte o log em: {LogPath}");
            }

            return process.ExitCode;
        }
        catch (Exception ex)
        {
            Log("ERROR", $"Falha ao executar o modo CLI: {ex.Message}");
            Console.Error.WriteLine($"Falha ao executar o PowerShell: {ex.Message}");
            Console.Error.WriteLine($"Log disponível em: {LogPath}");
            return 1;
        }
    }

    private static ProcessStartInfo CreateBaseStartInfo(RuntimeResolution runtime)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = runtime.PowerShellExecutable!,
            UseShellExecute = false
        };

        startInfo.ArgumentList.Add("-NoLogo");
        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-ExecutionPolicy");
        startInfo.ArgumentList.Add("Bypass");
        startInfo.ArgumentList.Add("-File");
        startInfo.ArgumentList.Add(runtime.ScriptPath!);

        return startInfo;
    }

    private static RuntimeResolution ResolveRuntime()
    {
        var scriptPath = Path.Combine(AppContext.BaseDirectory, "ipchange.ps1");
        if (!File.Exists(scriptPath))
        {
            return new RuntimeResolution(false, scriptPath, null, $"Script PowerShell não encontrado em '{scriptPath}'.");
        }

        var powerShellExecutable = FindPowerShellExecutable();
        if (powerShellExecutable is null)
        {
            return new RuntimeResolution(false, scriptPath, null, "Nenhum executável do PowerShell foi encontrado. Instale o PowerShell 5.1+ ou o PowerShell 7+.");
        }

        return new RuntimeResolution(true, scriptPath, powerShellExecutable, null);
    }

    private static string? FindPowerShellExecutable()
    {
        foreach (var executable in new[] { "pwsh.exe", "powershell.exe" })
        {
            if (IsExecutableOnPath(executable))
            {
                return executable;
            }
        }

        return null;
    }

    private static bool IsExecutableOnPath(string executable)
    {
        var path = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        foreach (var directory in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            if (File.Exists(Path.Combine(directory, executable)))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasArgument(IEnumerable<string> arguments, string value) =>
        arguments.Any(argument => string.Equals(argument, value, StringComparison.OrdinalIgnoreCase));

    private static string BuildLogPath()
    {
        var baseDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(baseDirectory))
        {
            baseDirectory = Path.GetTempPath();
        }

        return Path.Combine(baseDirectory, "ipchange", "logs", $"ipchange-{DateTime.Now:yyyyMMdd}.log");
    }

    private static void Log(string level, string message)
    {
        try
        {
            lock (LogSync)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
                File.AppendAllText(
                    LogPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}",
                    Encoding.UTF8);
            }
        }
        catch
        {
        }
    }

    private static string SanitizeArguments(IEnumerable<string> arguments)
    {
        var sanitized = new List<string>();
        var redactNext = false;

        foreach (var argument in arguments)
        {
            if (redactNext)
            {
                sanitized.Add("<redacted>");
                redactNext = false;
                continue;
            }

            sanitized.Add(argument);
            if (string.Equals(argument, "-PlainTextPassword", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(argument, "-Password", StringComparison.OrdinalIgnoreCase))
            {
                redactNext = true;
            }
        }

        return string.Join(' ', sanitized);
    }

    private sealed record RuntimeResolution(bool IsValid, string? ScriptPath, string? PowerShellExecutable, string? ErrorMessage);
}
