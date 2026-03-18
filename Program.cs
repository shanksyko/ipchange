using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Hosting;

namespace ipchange;

internal static class Program
{
    private static readonly object LogSync = new();
    private static readonly string LogPath = BuildLogPath();
    private const string DefaultLocalUsername = @".\support";

    private static async Task<int> Main(string[] args)
    {
        var forceCli = HasArgument(args, "--cli");
        var forceGui = HasArgument(args, "--ui");
        var filteredArgs = args.Where(arg => !string.Equals(arg, "--cli", StringComparison.OrdinalIgnoreCase) &&
                                             !string.Equals(arg, "--ui", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        if (forceGui || (!forceCli && filteredArgs.Length == 0))
        {
            return await RunGuiAsync();
        }

        return await RunCliAsync(filteredArgs);
    }

    private static async Task<int> RunCliAsync(string[] args)
    {
        if (!OperatingSystem.IsWindows())
        {
            Console.Error.WriteLine("Este aplicativo precisa ser executado no Windows para alterar o IP. Use --ui para abrir apenas a interface local de diagnóstico.");
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

    private static async Task<int> RunGuiAsync()
    {
        var port = GetAvailableLocalPort();
        var url = $"http://127.0.0.1:{port}";
        var builder = WebApplication.CreateBuilder(new WebApplicationOptions
        {
            Args = Array.Empty<string>(),
            ContentRootPath = AppContext.BaseDirectory
        });

        var app = builder.Build();
        app.Urls.Add(url);

        app.MapGet("/", async () =>
        {
            var model = await BuildDashboardPageModelAsync();
            return Results.Content(RenderDashboardHtml(model), "text/html; charset=utf-8");
        });

        app.MapPost("/actions/adapters", async (HttpRequest request) =>
        {
            var form = await ReadApplyRequestAsync(request);
            var model = await BuildDashboardPageModelAsync(
                form,
                resultMessage: "Adaptadores atualizados.",
                resultOutput: "A lista de adaptadores foi recarregada pelo servidor C#.");
            return Results.Content(RenderDashboardHtml(model), "text/html; charset=utf-8");
        });

        app.MapPost("/actions/diagnostics", async (HttpRequest request) =>
        {
            var form = await ReadApplyRequestAsync(request);
            var result = await InvokePowerShellAsync(new[] { "-DiagnosticsJson" });
            var output = FormatJsonOrText(result.StandardOutput, result.StandardError, "Clique em “Diagnosticar permissões” para carregar os detalhes.");
            var model = await BuildDashboardPageModelAsync(
                form,
                diagnosticsOutput: output,
                resultMessage: result.ExitCode == 0 ? "Diagnóstico executado com sucesso." : "Falha ao executar o diagnóstico.",
                resultIsError: result.ExitCode != 0,
                resultOutput: output);
            return Results.Content(RenderDashboardHtml(model), "text/html; charset=utf-8");
        });

        app.MapPost("/actions/logs", async (HttpRequest request) =>
        {
            var form = await ReadApplyRequestAsync(request);
            var model = await BuildDashboardPageModelAsync(
                form,
                resultMessage: "Logs atualizados.",
                resultOutput: "O conteúdo do log foi recarregado pelo servidor C#.");
            return Results.Content(RenderDashboardHtml(model), "text/html; charset=utf-8");
        });

        app.MapPost("/actions/apply", async (HttpRequest request) =>
        {
            var form = await ReadApplyRequestAsync(request);
            var validationError = ValidateApplyRequest(form);
            if (validationError is not null)
            {
                var invalidModel = await BuildDashboardPageModelAsync(
                    form,
                    resultMessage: validationError,
                    resultIsError: true,
                    resultOutput: "Corrija os campos destacados e tente novamente.");
                return Results.Content(RenderDashboardHtml(invalidModel), "text/html; charset=utf-8");
            }

            var arguments = BuildApplyArguments(form);
            var environment = BuildEnvironmentVariables(form);
            var result = await InvokePowerShellAsync(arguments, environment);
            var output = FormatJsonOrText(
                result.StandardOutput,
                result.StandardError,
                "Execução concluída sem saída textual.");
            var model = await BuildDashboardPageModelAsync(
                form with { Password = string.Empty },
                resultMessage: result.ExitCode == 0 ? $"Configuração aplicada com sucesso. Log: {result.LogPath}" : "Falha ao aplicar a configuração.",
                resultIsError: result.ExitCode != 0,
                resultOutput: output);
            return Results.Content(RenderDashboardHtml(model), "text/html; charset=utf-8");
        });

        app.MapGet("/api/status", () =>
        {
            var runtime = ResolveRuntime();
            return Results.Json(new
            {
                isWindows = OperatingSystem.IsWindows(),
                currentUser = $"{Environment.UserDomainName}\\{Environment.UserName}",
                powerShellExecutable = runtime.PowerShellExecutable,
                scriptPath = runtime.ScriptPath,
                isReady = runtime.IsValid,
                errorMessage = runtime.ErrorMessage,
                logPath = LogPath
            });
        });

        app.MapGet("/api/adapters", async () =>
        {
            var result = await InvokePowerShellAsync(new[] { "-ListAdaptersJson" });
            return BuildApiResult(result);
        });

        app.MapGet("/api/diagnostics", async () =>
        {
            var result = await InvokePowerShellAsync(new[] { "-DiagnosticsJson" });
            return BuildApiResult(result);
        });

        app.MapGet("/api/logs", () =>
        {
            return Results.Json(new
            {
                logPath = LogPath,
                content = ReadLogTail(LogPath, 250)
            });
        });

        app.MapPost("/api/apply", async (ApplyRequest request) =>
        {
            var validationError = ValidateApplyRequest(request);
            if (validationError is not null)
            {
                return Results.BadRequest(new { message = validationError, logPath = LogPath });
            }

            var arguments = BuildApplyArguments(request);
            var environment = BuildEnvironmentVariables(request);
            var result = await InvokePowerShellAsync(arguments, environment);

            return Results.Json(new
            {
                success = result.ExitCode == 0,
                exitCode = result.ExitCode,
                output = result.StandardOutput,
                error = result.StandardError,
                logPath = result.LogPath
            }, statusCode: result.ExitCode == 0 ? StatusCodes.Status200OK : StatusCodes.Status500InternalServerError);
        });

        Log("INFO", $"Interface gráfica iniciada em {url}.");

        if (OperatingSystem.IsWindows())
        {
            _ = Task.Run(async () =>
            {
                await Task.Delay(800);
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = url,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    Log("WARN", $"Não foi possível abrir o navegador automaticamente: {ex.Message}");
                }
            });
        }

        Console.WriteLine($"Interface gráfica disponível em {url}");
        Console.WriteLine($"Logs em: {LogPath}");
        await app.RunAsync(url);
        return 0;
    }

    private static IResult BuildApiResult(PowerShellRunResult result)
    {
        if (result.ExitCode != 0)
        {
            return Results.Json(new
            {
                message = "A execução falhou. Verifique a saída e o log para detalhes.",
                output = result.StandardOutput,
                error = result.StandardError,
                logPath = result.LogPath
            }, statusCode: StatusCodes.Status500InternalServerError);
        }

        if (string.IsNullOrWhiteSpace(result.StandardOutput))
        {
            return Results.Json(Array.Empty<object>());
        }

        return Results.Content(result.StandardOutput, "application/json; charset=utf-8");
    }

    private static string? ValidateApplyRequest(ApplyRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.AdapterName))
        {
            return "Selecione um adaptador antes de aplicar a configuração.";
        }

        if (string.IsNullOrWhiteSpace(request.IPAddress))
        {
            return "Informe um endereço IPv4.";
        }

        if (request.PrefixLength is < 0 or > 32)
        {
            return "Informe um prefixo entre 0 e 32.";
        }

        if (!request.UseCurrentCredential && string.IsNullOrWhiteSpace(request.Password))
        {
            return "Informe a senha administrativa ou marque a opção para usar a credencial atual.";
        }

        return null;
    }

    private static IReadOnlyList<string> BuildApplyArguments(ApplyRequest request)
    {
        var arguments = new List<string>
        {
            "-AdapterName", request.AdapterName!.Trim(),
            "-IPAddress", request.IPAddress!.Trim(),
            "-PrefixLength", request.PrefixLength.ToString()
        };

        if (!string.IsNullOrWhiteSpace(request.DefaultGateway))
        {
            arguments.Add("-DefaultGateway");
            arguments.Add(request.DefaultGateway.Trim());
        }

        var dnsServers = SplitCsv(request.DnsServers);
        if (dnsServers.Count > 0)
        {
            arguments.Add("-DnsServers");
            arguments.AddRange(dnsServers);
        }

        if (request.UseCurrentCredential)
        {
            arguments.Add("-UseCurrentCredential");
        }

        return arguments;
    }

    private static Dictionary<string, string?> BuildEnvironmentVariables(ApplyRequest request)
    {
        var environment = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

        if (!request.UseCurrentCredential)
        {
            environment["IPCHANGE_ADMIN_USERNAME"] = string.IsNullOrWhiteSpace(request.Username)
                ? DefaultLocalUsername
                : request.Username.Trim();

            if (!string.IsNullOrWhiteSpace(request.Password))
            {
                environment["IPCHANGE_ADMIN_PASSWORD"] = request.Password;
            }
        }

        return environment;
    }

    private static async Task<PowerShellRunResult> InvokePowerShellAsync(
        IReadOnlyList<string> scriptArguments,
        IReadOnlyDictionary<string, string?>? environmentVariables = null)
    {
        if (!OperatingSystem.IsWindows())
        {
            return new PowerShellRunResult(
                1,
                string.Empty,
                "A alteração de IP e os diagnósticos reais de adaptador só funcionam no Windows.",
                LogPath);
        }

        var runtime = ResolveRuntime();
        if (!runtime.IsValid)
        {
            return new PowerShellRunResult(1, string.Empty, runtime.ErrorMessage ?? "Runtime inválido.", LogPath);
        }

        var startInfo = CreateBaseStartInfo(runtime, redirectOutput: true);
        foreach (var argument in scriptArguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        if (!HasArgument(scriptArguments, "-LogPath"))
        {
            startInfo.ArgumentList.Add("-LogPath");
            startInfo.ArgumentList.Add(LogPath);
        }

        if (environmentVariables is not null)
        {
            foreach (var pair in environmentVariables)
            {
                startInfo.Environment[pair.Key] = pair.Value ?? string.Empty;
            }
        }

        Log("INFO", $"Iniciando PowerShell: {SanitizeArguments(startInfo.ArgumentList)}");

        try
        {
            using var process = Process.Start(startInfo);
            if (process is null)
            {
                return new PowerShellRunResult(1, string.Empty, "Não foi possível iniciar o PowerShell.", LogPath);
            }

            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();

            await process.WaitForExitAsync();

            var standardOutput = await outputTask;
            var standardError = await errorTask;

            if (!string.IsNullOrWhiteSpace(standardOutput))
            {
                Log("INFO", $"Saída do PowerShell:{Environment.NewLine}{standardOutput.Trim()}");
            }

            if (!string.IsNullOrWhiteSpace(standardError))
            {
                Log("WARN", $"Erros do PowerShell:{Environment.NewLine}{standardError.Trim()}");
            }

            Log("INFO", $"PowerShell finalizado com código {process.ExitCode}.");
            return new PowerShellRunResult(process.ExitCode, standardOutput, standardError, LogPath);
        }
        catch (Exception ex)
        {
            Log("ERROR", $"Falha ao executar o PowerShell: {ex.Message}");
            return new PowerShellRunResult(1, string.Empty, ex.Message, LogPath);
        }
    }

    private static ProcessStartInfo CreateBaseStartInfo(RuntimeResolution runtime, bool redirectOutput = false)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = runtime.PowerShellExecutable!,
            UseShellExecute = false,
            RedirectStandardOutput = redirectOutput,
            RedirectStandardError = redirectOutput
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

    private static int GetAvailableLocalPort()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();
        return port;
    }

    private static IReadOnlyList<string> SplitCsv(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return Array.Empty<string>();
        }

        return input.Split(new[] { ',', ';', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
    }

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

    private static string ReadLogTail(string path, int maxLines)
    {
        if (!File.Exists(path))
        {
            return "Nenhum log foi gravado ainda.";
        }

        var lines = File.ReadAllLines(path);
        return string.Join(Environment.NewLine, lines.Skip(Math.Max(0, lines.Length - maxLines)));
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

    private static async Task<DashboardPageModel> BuildDashboardPageModelAsync(
        ApplyRequest? request = null,
        string? diagnosticsOutput = null,
        string? resultMessage = null,
        bool resultIsError = false,
        string? resultOutput = null)
    {
        var runtime = ResolveRuntime();
        var adapterResult = await InvokePowerShellAsync(new[] { "-ListAdaptersJson" });
        var adapters = adapterResult.ExitCode == 0
            ? ParseAdapters(adapterResult.StandardOutput)
            : Array.Empty<AdapterOption>();

        return new DashboardPageModel(
            BuildStatusSnapshot(runtime),
            adapters,
            request ?? new ApplyRequest { PrefixLength = 24, Username = DefaultLocalUsername },
            diagnosticsOutput ?? "Clique em “Diagnosticar permissões” para carregar os detalhes.",
            resultMessage ?? "Aguardando ação.",
            resultIsError,
            resultOutput ?? "Toda a interface abaixo é renderizada diretamente em C#, sem depender de TypeScript no navegador.",
            ReadLogTail(LogPath, 250),
            adapterResult.ExitCode == 0 ? null : FormatJsonOrText(adapterResult.StandardOutput, adapterResult.StandardError, "Não foi possível carregar os adaptadores."));
    }

    private static StatusSnapshot BuildStatusSnapshot(RuntimeResolution runtime)
    {
        return new StatusSnapshot(
            OperatingSystem.IsWindows(),
            $"{Environment.UserDomainName}\\{Environment.UserName}",
            runtime.PowerShellExecutable,
            runtime.ScriptPath,
            runtime.IsValid,
            runtime.ErrorMessage,
            LogPath);
    }

    private static IReadOnlyList<AdapterOption> ParseAdapters(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return Array.Empty<AdapterOption>();
        }

        try
        {
            using var document = JsonDocument.Parse(json);
            if (document.RootElement.ValueKind == JsonValueKind.Array)
            {
                return document.RootElement
                    .EnumerateArray()
                    .Select(BuildAdapterOption)
                    .Where(option => !string.IsNullOrWhiteSpace(option.Name))
                    .ToArray();
            }

            if (document.RootElement.ValueKind == JsonValueKind.Object)
            {
                var single = BuildAdapterOption(document.RootElement);
                return string.IsNullOrWhiteSpace(single.Name) ? Array.Empty<AdapterOption>() : new[] { single };
            }
        }
        catch (JsonException ex)
        {
            Log("WARN", $"Não foi possível interpretar a lista de adaptadores: {ex.Message}");
        }

        return Array.Empty<AdapterOption>();
    }

    private static AdapterOption BuildAdapterOption(JsonElement element)
    {
        var name = element.TryGetProperty("Name", out var nameElement) ? nameElement.GetString() : null;
        var status = element.TryGetProperty("Status", out var statusElement) ? statusElement.GetString() : null;

        int? interfaceIndex = null;
        if (element.TryGetProperty("InterfaceIndex", out var indexElement))
        {
            if (indexElement.ValueKind == JsonValueKind.Number && indexElement.TryGetInt32(out var numericIndex))
            {
                interfaceIndex = numericIndex;
            }
            else if (indexElement.ValueKind == JsonValueKind.String &&
                     int.TryParse(indexElement.GetString(), out var stringIndex))
            {
                interfaceIndex = stringIndex;
            }
        }

        return new AdapterOption(name, interfaceIndex, status);
    }

    private static async Task<ApplyRequest> ReadApplyRequestAsync(HttpRequest request)
    {
        var form = await request.ReadFormAsync();
        return new ApplyRequest
        {
            Username = ReadFormValue(form, "Username"),
            Password = ReadFormValue(form, "Password"),
            AdapterName = ReadFormValue(form, "AdapterName"),
            IPAddress = ReadFormValue(form, "IPAddress"),
            PrefixLength = int.TryParse(ReadFormValue(form, "PrefixLength"), out var prefixLength) ? prefixLength : -1,
            DefaultGateway = ReadFormValue(form, "DefaultGateway"),
            DnsServers = ReadFormValue(form, "DnsServers"),
            UseCurrentCredential = form.ContainsKey("UseCurrentCredential")
        };
    }

    private static string? ReadFormValue(IFormCollection form, string key)
    {
        if (!form.TryGetValue(key, out var value))
        {
            return null;
        }

        var text = value.ToString();
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    private static string FormatJsonOrText(string? output, string? error, string emptyFallback)
    {
        var combined = string.Join(
            Environment.NewLine + Environment.NewLine,
            new[] { output?.Trim(), error?.Trim() }.Where(value => !string.IsNullOrWhiteSpace(value)));

        if (string.IsNullOrWhiteSpace(combined))
        {
            return emptyFallback;
        }

        try
        {
            using var document = JsonDocument.Parse(combined);
            return JsonSerializer.Serialize(document.RootElement, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (JsonException)
        {
            return combined;
        }
    }

    private static string RenderDashboardHtml(DashboardPageModel model)
    {
        var builder = new StringBuilder();
        builder.AppendLine("""
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>ipchange</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #0f172a;
      --panel: rgba(15, 23, 42, 0.82);
      --panel-soft: rgba(30, 41, 59, 0.82);
      --line: rgba(148, 163, 184, 0.2);
      --text: #e2e8f0;
      --muted: #94a3b8;
      --accent: #38bdf8;
      --danger: #fb7185;
      --shadow: 0 24px 60px rgba(15, 23, 42, 0.45);
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background:
        radial-gradient(circle at top left, rgba(56, 189, 248, 0.3), transparent 32%),
        radial-gradient(circle at top right, rgba(34, 197, 94, 0.18), transparent 24%),
        linear-gradient(135deg, #020617 0%, #0f172a 45%, #111827 100%);
      color: var(--text);
    }

    .page {
      width: min(1180px, calc(100% - 32px));
      margin: 32px auto;
      display: grid;
      gap: 20px;
    }

    .hero, .panel {
      background: var(--panel);
      backdrop-filter: blur(18px);
      border: 1px solid var(--line);
      border-radius: 24px;
      box-shadow: var(--shadow);
    }

    .hero {
      padding: 28px;
      display: flex;
      flex-wrap: wrap;
      gap: 20px;
      justify-content: space-between;
      align-items: center;
    }

    .hero h1 {
      margin: 0 0 8px;
      font-size: clamp(2rem, 4vw, 3rem);
    }

    .hero p, .muted {
      margin: 0;
      color: var(--muted);
      line-height: 1.6;
    }

    .badge {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 10px 16px;
      border-radius: 999px;
      background: rgba(56, 189, 248, 0.12);
      border: 1px solid rgba(56, 189, 248, 0.25);
      color: #bae6fd;
      font-weight: 600;
    }

    .panel { padding: 24px; }
    .panel h2 { margin: 0 0 18px; font-size: 1.1rem; }

    .cards, .grid {
      display: grid;
      gap: 14px;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    }

    .layout {
      display: grid;
      gap: 20px;
      grid-template-columns: 1.2fr 0.9fr;
    }

    .card, fieldset {
      padding: 16px;
      border-radius: 18px;
      background: var(--panel-soft);
      border: 1px solid var(--line);
    }

    fieldset {
      margin: 0;
      display: grid;
      gap: 14px;
    }

    .label {
      color: var(--muted);
      font-size: 0.85rem;
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }

    .card strong {
      display: block;
      margin-top: 6px;
      font-size: 1rem;
      word-break: break-word;
    }

    form { display: grid; gap: 14px; }

    .field {
      display: grid;
      gap: 8px;
    }

    .field label {
      font-weight: 600;
      color: #cbd5e1;
    }

    .field input, .field select, .field textarea {
      width: 100%;
      border: 1px solid rgba(148, 163, 184, 0.18);
      background: rgba(15, 23, 42, 0.7);
      color: var(--text);
      border-radius: 14px;
      padding: 13px 14px;
      outline: none;
    }

    .toolbar, .actions {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
    }

    button {
      border: 0;
      border-radius: 14px;
      padding: 13px 16px;
      font-weight: 700;
      cursor: pointer;
    }

    .primary {
      background: linear-gradient(135deg, #38bdf8, #2563eb);
      color: white;
      box-shadow: 0 14px 30px rgba(37, 99, 235, 0.32);
    }

    .secondary {
      background: rgba(148, 163, 184, 0.14);
      color: var(--text);
      border: 1px solid rgba(148, 163, 184, 0.18);
    }

    .status {
      border-left: 4px solid var(--accent);
      padding: 14px 16px;
      border-radius: 16px;
      background: rgba(56, 189, 248, 0.08);
      color: #e0f2fe;
    }

    .status.error {
      border-left-color: var(--danger);
      background: rgba(251, 113, 133, 0.08);
      color: #ffe4e6;
    }

    .mono {
      margin: 0;
      padding: 16px;
      border-radius: 16px;
      background: rgba(2, 6, 23, 0.65);
      border: 1px solid rgba(148, 163, 184, 0.16);
      font-family: "Cascadia Code", Consolas, monospace;
      font-size: 0.88rem;
      white-space: pre-wrap;
      word-break: break-word;
      min-height: 160px;
      max-height: 420px;
      overflow: auto;
    }

    .hint {
      color: var(--muted);
      line-height: 1.55;
    }

    @media (max-width: 980px) {
      .layout { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <main class="page">
""");

        builder.AppendLine($"""
    <section class="hero">
      <div>
        <div class="badge">✨ ipchange em C# direto</div>
        <h1>Configuração de IP em app local do Windows</h1>
        <p>Esta interface é gerada diretamente em C#, sem TypeScript no navegador. O formulário abaixo reaproveita a lógica existente do PowerShell para listar adaptadores, diagnosticar permissões e aplicar IPv4.</p>
      </div>
      <div class="hint">Log atual: {HtmlEncode(model.Status.LogPath)}</div>
    </section>
""");

        builder.AppendLine("""
    <section class="panel">
      <h2>Resumo do ambiente</h2>
      <div class="cards">
""");

        AppendCard(builder, "Sistema operacional", model.Status.IsWindows ? "Windows" : "Não Windows");
        AppendCard(builder, "Usuário atual", model.Status.CurrentUser);
        AppendCard(builder, "PowerShell", model.Status.PowerShellExecutable ?? "Não encontrado");
        AppendCard(builder, "Script", model.Status.ScriptPath ?? "Não encontrado");

        builder.AppendLine("""
      </div>
""");

        builder.AppendLine($"""
      <div class="status{(model.Status.IsReady ? string.Empty : " error")}" style="margin-top: 16px;">
        {HtmlEncode(model.Status.IsReady
            ? "Ambiente carregado. A interface agora depende apenas de renderização em C# no servidor local."
            : model.Status.ErrorMessage ?? "O aplicativo ainda não está pronto para executar o script.")}
      </div>
""");

        builder.AppendLine("""
    </section>
    <section class="layout">
      <section class="panel">
        <h2>Alterar configuração IPv4</h2>
""");

        if (!string.IsNullOrWhiteSpace(model.AdapterMessage))
        {
            builder.AppendLine($"""
        <div class="status error" style="margin-bottom: 16px;">{HtmlEncode(model.AdapterMessage)}</div>
""");
        }

        builder.AppendLine($"""
        <form method="post" action="/actions/apply">
          <div class="toolbar">
            <button class="secondary" type="submit" formaction="/actions/adapters">Atualizar adaptadores</button>
            <button class="secondary" type="submit" formaction="/actions/diagnostics">Diagnosticar permissões</button>
            <button class="secondary" type="submit" formaction="/actions/logs">Atualizar logs</button>
          </div>

          <div class="field">
            <label for="AdapterName">Adaptador</label>
            <select id="AdapterName" name="AdapterName" required>
              <option value="">Selecione um adaptador</option>
              {RenderAdapterOptions(model.Adapters, model.Request.AdapterName)}
            </select>
          </div>

          <div class="grid">
            <div class="field">
              <label for="IPAddress">IPv4</label>
              <input id="IPAddress" name="IPAddress" value="{HtmlAttributeEncode(model.Request.IPAddress)}" placeholder="192.168.0.50" required />
            </div>
            <div class="field">
              <label for="PrefixLength">Prefixo</label>
              <input id="PrefixLength" name="PrefixLength" type="number" min="0" max="32" value="{HtmlAttributeEncode(model.Request.PrefixLength <= 0 ? "24" : model.Request.PrefixLength.ToString())}" required />
            </div>
            <div class="field">
              <label for="DefaultGateway">Gateway padrão</label>
              <input id="DefaultGateway" name="DefaultGateway" value="{HtmlAttributeEncode(model.Request.DefaultGateway)}" placeholder="192.168.0.1" />
            </div>
            <div class="field">
              <label for="DnsServers">DNS</label>
              <input id="DnsServers" name="DnsServers" value="{HtmlAttributeEncode(model.Request.DnsServers)}" placeholder="1.1.1.1, 8.8.8.8" />
            </div>
          </div>

          <label class="status" style="display: flex; gap: 10px; align-items: center;">
            <input id="UseCurrentCredential" name="UseCurrentCredential" type="checkbox" {(model.Request.UseCurrentCredential ? "checked" : string.Empty)} />
            Usar a credencial atual em vez de fornecer usuário e senha administrativa
          </label>

          <fieldset {(model.Request.UseCurrentCredential ? "disabled" : string.Empty)}>
            <legend class="label">Credenciais administrativas</legend>
            <div class="grid">
              <div class="field">
                <label for="Username">Usuário administrativo</label>
                <input id="Username" name="Username" value="{HtmlAttributeEncode(model.Request.Username ?? DefaultLocalUsername)}" placeholder=".\support" />
              </div>
              <div class="field">
                <label for="Password">Senha administrativa</label>
                <input id="Password" name="Password" type="password" placeholder="Digite a senha" />
              </div>
            </div>
            <p class="muted">Por segurança, a senha nunca é preenchida novamente após um POST.</p>
          </fieldset>

          <div class="actions">
            <button class="primary" type="submit">Aplicar configuração</button>
            <button class="secondary" type="reset">Limpar campos</button>
          </div>
        </form>
      </section>

      <section class="panel">
        <h2>Diagnóstico e retorno técnico</h2>
        <p class="muted">Use os diagnósticos para verificar Windows, PowerShell, conta atual, privilégios administrativos e possíveis pistas de permissão.</p>
        <pre class="mono">{HtmlEncode(model.DiagnosticsOutput)}</pre>
        <div class="status{(model.ResultIsError ? " error" : string.Empty)}" style="margin-top: 16px;">{HtmlEncode(model.ResultMessage)}</div>
        <pre class="mono" style="margin-top: 16px;">{HtmlEncode(model.ResultOutput)}</pre>
      </section>
    </section>

    <section class="panel">
      <h2>Log de verificação</h2>
      <p class="muted">O log abaixo reúne o fluxo do wrapper C# e do script PowerShell para facilitar a análise de falhas, especialmente de permissão.</p>
      <pre class="mono">{HtmlEncode(model.LogContent)}</pre>
    </section>
  </main>
</body>
</html>
""");

        return builder.ToString();
    }

    private static void AppendCard(StringBuilder builder, string label, string value)
    {
        builder.AppendLine($"""
        <article class="card">
          <span class="label">{HtmlEncode(label)}</span>
          <strong>{HtmlEncode(value)}</strong>
        </article>
""");
    }

    private static string RenderAdapterOptions(IReadOnlyList<AdapterOption> adapters, string? selectedAdapter)
    {
        var builder = new StringBuilder();
        var selectedExists = false;

        foreach (var adapter in adapters)
        {
            var isSelected = string.Equals(adapter.Name, selectedAdapter, StringComparison.OrdinalIgnoreCase);
            selectedExists |= isSelected;
            var indexText = adapter.InterfaceIndex?.ToString() ?? "—";
            builder.AppendLine($"""<option value="{HtmlAttributeEncode(adapter.Name)}" {(isSelected ? "selected" : string.Empty)}>{HtmlEncode($"{indexText} · {adapter.Name} ({adapter.Status ?? "sem status"})")}</option>""");
        }

        if (!selectedExists && !string.IsNullOrWhiteSpace(selectedAdapter))
        {
            builder.AppendLine($"""<option value="{HtmlAttributeEncode(selectedAdapter)}" selected>{HtmlEncode(selectedAdapter)}</option>""");
        }

        return builder.ToString();
    }

    private static string HtmlEncode(string? value) => WebUtility.HtmlEncode(value ?? string.Empty);

    private static string HtmlAttributeEncode(string? value) => WebUtility.HtmlEncode(value ?? string.Empty);

    private sealed record RuntimeResolution(bool IsValid, string? ScriptPath, string? PowerShellExecutable, string? ErrorMessage);

    private sealed record StatusSnapshot(
        bool IsWindows,
        string CurrentUser,
        string? PowerShellExecutable,
        string? ScriptPath,
        bool IsReady,
        string? ErrorMessage,
        string LogPath);

    private sealed record AdapterOption(string? Name, int? InterfaceIndex, string? Status);

    private sealed record DashboardPageModel(
        StatusSnapshot Status,
        IReadOnlyList<AdapterOption> Adapters,
        ApplyRequest Request,
        string DiagnosticsOutput,
        string ResultMessage,
        bool ResultIsError,
        string ResultOutput,
        string LogContent,
        string? AdapterMessage);

    private sealed record PowerShellRunResult(int ExitCode, string StandardOutput, string StandardError, string LogPath);

    private sealed record ApplyRequest
    {
        public string? Username { get; init; }
        public string? Password { get; init; }
        public string? AdapterName { get; init; }
        public string? IPAddress { get; init; }
        public int PrefixLength { get; init; }
        public string? DefaultGateway { get; init; }
        public string? DnsServers { get; init; }
        public bool UseCurrentCredential { get; init; }
    }
}
