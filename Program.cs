using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Hosting;

namespace ipchange;

internal static class Program
{
    private static readonly object LogSync = new();
    private static readonly string LogPath = BuildLogPath();

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

        app.MapGet("/", () => Results.Content(UserInterfaceHtml, "text/html; charset=utf-8"));

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
            if (!string.IsNullOrWhiteSpace(request.Username))
            {
                environment["IPCHANGE_ADMIN_USERNAME"] = request.Username.Trim();
            }

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

    private static readonly string UserInterfaceHtml = """
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
      --accent-2: #22c55e;
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

    .hero p {
      margin: 0;
      color: var(--muted);
      max-width: 740px;
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

    .layout {
      display: grid;
      gap: 20px;
      grid-template-columns: 1.25fr 0.95fr;
    }

    .panel {
      padding: 24px;
    }

    .panel h2 {
      margin: 0 0 18px;
      font-size: 1.1rem;
    }

    .cards {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 14px;
    }

    .card {
      padding: 16px;
      border-radius: 18px;
      background: var(--panel-soft);
      border: 1px solid var(--line);
    }

    .card strong {
      display: block;
      margin-top: 6px;
      font-size: 1rem;
      word-break: break-word;
    }

    .label {
      color: var(--muted);
      font-size: 0.85rem;
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }

    form {
      display: grid;
      gap: 14px;
    }

    .grid {
      display: grid;
      gap: 14px;
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }

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
      transition: border-color 160ms ease, transform 160ms ease, box-shadow 160ms ease;
    }

    .field input:focus, .field select:focus, .field textarea:focus {
      border-color: rgba(56, 189, 248, 0.9);
      box-shadow: 0 0 0 4px rgba(56, 189, 248, 0.15);
    }

    .field textarea {
      min-height: 92px;
      resize: vertical;
    }

    .actions, .toolbar {
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
      transition: transform 160ms ease, opacity 160ms ease, box-shadow 160ms ease;
    }

    button:hover { transform: translateY(-1px); }
    button:disabled { opacity: 0.65; cursor: wait; transform: none; }

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

    .success {
      background: linear-gradient(135deg, #22c55e, #16a34a);
      color: white;
      box-shadow: 0 14px 30px rgba(22, 163, 74, 0.28);
    }

    .hint, .muted {
      color: var(--muted);
      line-height: 1.55;
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

    .toggle {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 12px 14px;
      border-radius: 14px;
      background: rgba(15, 23, 42, 0.58);
      border: 1px solid var(--line);
      font-weight: 600;
    }

    .toggle input {
      width: 18px;
      height: 18px;
      accent-color: var(--accent);
    }

    @media (max-width: 980px) {
      .layout { grid-template-columns: 1fr; }
      .grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <main class="page">
    <section class="hero">
      <div>
        <div class="badge">✨ Interface gráfica local do ipchange</div>
        <h1>Configuração de IP com visual mais amigável</h1>
        <p>Selecione o adaptador, informe IPv4, prefixo, gateway e DNS, rode um diagnóstico de permissão e acompanhe o log técnico da execução no mesmo lugar.</p>
      </div>
      <div class="hint" id="logPathHero">Preparando interface...</div>
    </section>

    <section class="panel">
      <h2>Resumo do ambiente</h2>
      <div class="cards" id="statusCards"></div>
      <div id="statusMessage" class="status" style="margin-top: 16px;">Carregando status do aplicativo...</div>
    </section>

    <section class="layout">
      <section class="panel">
        <h2>Alterar configuração IPv4</h2>
        <form id="configForm">
          <div class="toolbar">
            <button class="secondary" type="button" id="loadAdaptersButton">Atualizar adaptadores</button>
            <button class="secondary" type="button" id="diagnosticsButton">Diagnosticar permissões</button>
            <button class="secondary" type="button" id="refreshLogsButton">Atualizar logs</button>
          </div>

          <div class="field">
            <label for="adapterName">Adaptador</label>
            <select id="adapterName" required>
              <option value="">Selecione um adaptador</option>
            </select>
          </div>

          <div class="grid">
            <div class="field">
              <label for="ipAddress">IPv4</label>
              <input id="ipAddress" placeholder="192.168.0.50" required />
            </div>
            <div class="field">
              <label for="prefixLength">Prefixo</label>
              <input id="prefixLength" type="number" min="0" max="32" value="24" required />
            </div>
            <div class="field">
              <label for="defaultGateway">Gateway padrão</label>
              <input id="defaultGateway" placeholder="192.168.0.1" />
            </div>
            <div class="field">
              <label for="dnsServers">DNS</label>
              <input id="dnsServers" placeholder="1.1.1.1, 8.8.8.8" />
            </div>
          </div>

          <label class="toggle">
            <input id="useCurrentCredential" type="checkbox" />
            Usar a credencial atual em vez de fornecer usuário e senha administrativa
          </label>

          <div class="grid" id="credentialFields">
            <div class="field">
              <label for="username">Usuário administrativo</label>
              <input id="username" placeholder=".\support" />
            </div>
            <div class="field">
              <label for="password">Senha administrativa</label>
              <input id="password" type="password" placeholder="Digite a senha" />
            </div>
          </div>

          <div class="actions">
            <button class="primary" type="submit" id="applyButton">Aplicar configuração</button>
            <button class="secondary" type="reset" id="clearButton">Limpar campos</button>
          </div>
        </form>
      </section>

      <section class="panel">
        <h2>Diagnóstico e retorno técnico</h2>
        <p class="muted">Use os diagnósticos para verificar Windows, PowerShell, conta atual, privilégios administrativos e possíveis pistas de permissão.</p>
        <pre class="mono" id="diagnosticsOutput">Clique em “Diagnosticar permissões” para carregar os detalhes.</pre>
        <div id="resultMessage" class="status" style="margin-top: 16px;">Aguardando ação.</div>
        <pre class="mono" id="resultOutput" style="margin-top: 16px;">Saída da execução aparecerá aqui.</pre>
      </section>
    </section>

    <section class="panel">
      <h2>Log de verificação</h2>
      <p class="muted">O log abaixo reúne o fluxo do wrapper C# e do script PowerShell para facilitar a análise de falhas, especialmente de permissão.</p>
      <pre class="mono" id="logOutput">Nenhum log carregado ainda.</pre>
    </section>
  </main>

  <script>
    const statusCards = document.getElementById('statusCards');
    const statusMessage = document.getElementById('statusMessage');
    const diagnosticsOutput = document.getElementById('diagnosticsOutput');
    const resultMessage = document.getElementById('resultMessage');
    const resultOutput = document.getElementById('resultOutput');
    const logOutput = document.getElementById('logOutput');
    const logPathHero = document.getElementById('logPathHero');
    const credentialFields = document.getElementById('credentialFields');
    const useCurrentCredential = document.getElementById('useCurrentCredential');
    const form = document.getElementById('configForm');
    const adapterName = document.getElementById('adapterName');
    const applyButton = document.getElementById('applyButton');
    const loadAdaptersButton = document.getElementById('loadAdaptersButton');
    const diagnosticsButton = document.getElementById('diagnosticsButton');
    const refreshLogsButton = document.getElementById('refreshLogsButton');

    function setMessage(target, text, isError = false) {
      target.textContent = text;
      target.classList.toggle('error', isError);
    }

    function setBusy(button, busy) {
      if (!button) return;
      button.disabled = busy;
    }

    function renderCards(items) {
      statusCards.innerHTML = '';
      for (const [label, value] of items) {
        const card = document.createElement('article');
        card.className = 'card';
        card.innerHTML = `<span class="label">${label}</span><strong>${value ?? '—'}</strong>`;
        statusCards.appendChild(card);
      }
    }

    function toggleCredentialFields() {
      credentialFields.style.display = useCurrentCredential.checked ? 'none' : 'grid';
    }

    async function readJson(response) {
      const text = await response.text();
      if (!text) {
        return null;
      }
      try {
        return JSON.parse(text);
      } catch {
        return text;
      }
    }

    async function loadStatus() {
      const response = await fetch('/api/status');
      const data = await response.json();
      renderCards([
        ['Sistema operacional', data.isWindows ? 'Windows' : 'Não Windows'],
        ['Usuário atual', data.currentUser],
        ['PowerShell', data.powerShellExecutable || 'Não encontrado'],
        ['Script', data.scriptPath || 'Não encontrado']
      ]);
      logPathHero.textContent = `Log atual: ${data.logPath}`;
      setMessage(
        statusMessage,
        data.isReady
          ? 'Ambiente carregado. Você já pode verificar adaptadores, diagnósticos e logs.'
          : (data.errorMessage || 'O aplicativo ainda não está pronto para executar o script.'),
        !data.isReady
      );
    }

    async function loadAdapters() {
      setBusy(loadAdaptersButton, true);
      try {
        const response = await fetch('/api/adapters');
        const data = await readJson(response);
        if (!response.ok) {
          throw new Error(data?.error || data?.message || 'Não foi possível listar os adaptadores.');
        }

        const items = Array.isArray(data) ? data : [data];
        adapterName.innerHTML = '<option value="">Selecione um adaptador</option>';
        for (const item of items) {
          if (!item?.Name) continue;
          const option = document.createElement('option');
          option.value = item.Name;
          option.textContent = `${item.InterfaceIndex} · ${item.Name} (${item.Status || 'sem status'})`;
          adapterName.appendChild(option);
        }

        setMessage(resultMessage, `${items.length} adaptador(es) carregado(s) para seleção.`);
      } catch (error) {
        setMessage(resultMessage, error.message, true);
      } finally {
        setBusy(loadAdaptersButton, false);
      }
    }

    async function runDiagnostics() {
      setBusy(diagnosticsButton, true);
      try {
        const response = await fetch('/api/diagnostics');
        const data = await readJson(response);
        if (!response.ok) {
          throw new Error(data?.error || data?.message || 'Falha ao executar o diagnóstico.');
        }

        diagnosticsOutput.textContent = JSON.stringify(data, null, 2);
        setMessage(resultMessage, 'Diagnóstico executado com sucesso.');
      } catch (error) {
        diagnosticsOutput.textContent = String(error.message || error);
        setMessage(resultMessage, diagnosticsOutput.textContent, true);
      } finally {
        setBusy(diagnosticsButton, false);
      }
    }

    async function refreshLogs() {
      setBusy(refreshLogsButton, true);
      try {
        const response = await fetch('/api/logs');
        const data = await response.json();
        logOutput.textContent = data.content || 'Nenhum log disponível.';
        logPathHero.textContent = `Log atual: ${data.logPath}`;
      } finally {
        setBusy(refreshLogsButton, false);
      }
    }

    async function applyConfiguration(event) {
      event.preventDefault();
      setBusy(applyButton, true);
      const payload = {
        adapterName: adapterName.value,
        ipAddress: document.getElementById('ipAddress').value,
        prefixLength: Number(document.getElementById('prefixLength').value),
        defaultGateway: document.getElementById('defaultGateway').value,
        dnsServers: document.getElementById('dnsServers').value,
        useCurrentCredential: useCurrentCredential.checked,
        username: document.getElementById('username').value,
        password: document.getElementById('password').value
      };

      try {
        const response = await fetch('/api/apply', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        const data = await response.json();
        const combinedOutput = [data.output, data.error].filter(Boolean).join('\n\n');
        resultOutput.textContent = combinedOutput || 'Execução concluída sem saída textual.';
        setMessage(
          resultMessage,
          response.ok
            ? `Configuração aplicada com sucesso. Log: ${data.logPath}`
            : (data.error || data.message || 'Falha ao aplicar a configuração.'),
          !response.ok
        );
        await refreshLogs();
      } catch (error) {
        resultOutput.textContent = String(error.message || error);
        setMessage(resultMessage, resultOutput.textContent, true);
      } finally {
        setBusy(applyButton, false);
      }
    }

    useCurrentCredential.addEventListener('change', toggleCredentialFields);
    loadAdaptersButton.addEventListener('click', loadAdapters);
    diagnosticsButton.addEventListener('click', runDiagnostics);
    refreshLogsButton.addEventListener('click', refreshLogs);
    form.addEventListener('submit', applyConfiguration);
    form.addEventListener('reset', () => setTimeout(toggleCredentialFields, 0));

    toggleCredentialFields();
    loadStatus();
    refreshLogs();
  </script>
</body>
</html>
""";

    private sealed record RuntimeResolution(bool IsValid, string? ScriptPath, string? PowerShellExecutable, string? ErrorMessage);

    private sealed record PowerShellRunResult(int ExitCode, string StandardOutput, string StandardError, string LogPath);

    private sealed class ApplyRequest
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
