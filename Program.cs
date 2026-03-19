using System.Text.Json;
using System.Windows.Forms;

namespace ipchange;

internal static class Program
{
    [STAThread]
    private static async Task<int> Main(string[] args)
    {
        if (!OperatingSystem.IsWindows())
        {
            MessageBox.Show(
                "Este aplicativo precisa ser executado no Windows para alterar o IP.",
                "ipchange",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        try
        {
            if (args.Any(arg => string.Equals(arg, "--service", StringComparison.OrdinalIgnoreCase)))
            {
                return await IpChangeServiceRuntime.RunAsync(args);
            }

            if (InternalCommandParser.TryParse(args, out var command))
            {
                return await InternalCommandRunner.RunAsync(command!);
            }

            ApplicationConfiguration.Initialize();
            var startHidden = args.Any(arg => string.Equals(arg, "--background", StringComparison.OrdinalIgnoreCase));
            Application.Run(new MainForm(startHidden));
            return 0;
        }
        catch (Exception ex)
        {
            AppLogger.Write("ERROR", ex.Message);
            MessageBox.Show(
                ex.Message,
                "ipchange",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }
    }
}

internal static class InternalCommandParser
{
    public static bool TryParse(string[] args, out InternalCommand? command)
    {
        command = null;
        if (args.Length == 0)
        {
            return false;
        }

        if (!string.Equals(args[0], "--apply", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var request = new ApplyRequest();
        for (var index = 1; index < args.Length; index++)
        {
            var argument = args[index];
            switch (argument.ToLowerInvariant())
            {
                case "--dhcp":
                    request.UseDhcp = true;
                    break;
                case "--adapter-name":
                    request.AdapterName = ReadValue(args, ref index, argument);
                    break;
                case "--ip-address":
                    request.IPAddress = ReadValue(args, ref index, argument);
                    break;
                case "--prefix-length":
                    if (!int.TryParse(ReadValue(args, ref index, argument), out var prefixLength) || prefixLength is < 0 or > 32)
                    {
                        throw new ArgumentException("O valor de --prefix-length deve estar entre 0 e 32.");
                    }

                    request.PrefixLength = prefixLength;
                    break;
                case "--default-gateway":
                    request.DefaultGateway = ReadValue(args, ref index, argument);
                    break;
                case "--dns-server":
                    request.DnsServers.Add(ReadValue(args, ref index, argument));
                    break;
                case "--log-path":
                    request.LogPath = ReadValue(args, ref index, argument);
                    break;
                case "--result-path":
                    request.ResultPath = ReadValue(args, ref index, argument);
                    break;
                case "--error-path":
                    request.ErrorPath = ReadValue(args, ref index, argument);
                    break;
                default:
                    throw new ArgumentException($"Argumento interno não reconhecido: {argument}");
            }
        }

        if (string.IsNullOrWhiteSpace(request.AdapterName) ||
            (!request.UseDhcp &&
             (string.IsNullOrWhiteSpace(request.IPAddress) || !request.PrefixLength.HasValue)))
        {
            throw new ArgumentException(request.UseDhcp
                ? "O modo interno --apply com DHCP exige ao menos o adaptador."
                : "O modo interno --apply exige adapter, ip e prefixo.");
        }

        command = new ApplyCommand(request);
        return true;
    }

    private static string ReadValue(string[] args, ref int index, string argumentName)
    {
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
        {
            throw new ArgumentException($"O argumento {argumentName} exige um valor.");
        }

        index++;
        return args[index];
    }
}

internal static class InternalCommandRunner
{
    public static async Task<int> RunAsync(InternalCommand command)
    {
        switch (command)
        {
            case ApplyCommand applyCommand:
                if (!string.IsNullOrWhiteSpace(applyCommand.Request.LogPath))
                {
                    AppLogger.SetPath(applyCommand.Request.LogPath!);
                }

                try
                {
                    var result = await NetworkConfigurationService.ApplyConfigurationAsync(applyCommand.Request);
                    var payload = JsonSerializer.Serialize(result, AppJson.Options);
                    if (!string.IsNullOrWhiteSpace(applyCommand.Request.ResultPath))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(applyCommand.Request.ResultPath!)!);
                        await File.WriteAllTextAsync(applyCommand.Request.ResultPath!, payload);
                    }
                    else
                    {
                        Console.WriteLine(payload);
                    }

                    return 0;
                }
                catch (Exception ex)
                {
                    AppLogger.Write("ERROR", ex.Message);
                    if (!string.IsNullOrWhiteSpace(applyCommand.Request.ErrorPath))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(applyCommand.Request.ErrorPath!)!);
                        await File.WriteAllTextAsync(applyCommand.Request.ErrorPath!, ex.Message);
                    }
                    else
                    {
                        Console.Error.WriteLine(ex.Message);
                    }

                    return 1;
                }
            default:
                throw new InvalidOperationException("Comando interno inválido.");
        }
    }
}

internal abstract record InternalCommand;
internal sealed record ApplyCommand(ApplyRequest Request) : InternalCommand;
