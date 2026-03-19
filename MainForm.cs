using Microsoft.Win32;
using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;

namespace ipchange;

internal sealed class MainForm : Form
{
    private readonly ComboBox _adapterComboBox = new();
    private readonly RadioButton _staticModeRadioButton = new();
    private readonly RadioButton _dhcpModeRadioButton = new();
    private readonly TextBox _ipAddressTextBox = new();
    private readonly NumericUpDown _prefixLengthUpDown = new();
    private readonly TextBox _defaultGatewayTextBox = new();
    private readonly TextBox _dnsServersTextBox = new();
    private readonly Button _refreshButton = new();
    private readonly Button _applyButton = new();
    private readonly Label _statusLabel = new();
    private readonly Label _modeHintLabel = new();
    private readonly NotifyIcon _notifyIcon = new();
    private readonly ContextMenuStrip _trayMenu = new();
    private readonly bool _startHidden;
    private bool _allowExit;

    public MainForm(bool startHidden)
    {
        _startHidden = startHidden;

        Text = "ipchange";
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.Dpi;
        MinimumSize = new Size(760, 460);
        Width = 940;
        Height = 560;
        BackColor = Color.FromArgb(242, 245, 247);

        InitializeComponent();
        Load += async (_, _) => await InitializeRuntimeAsync();
        Shown += (_, _) =>
        {
            if (_startHidden)
            {
                HideToTray();
            }
        };
        Resize += (_, _) =>
        {
            if (WindowState == FormWindowState.Minimized)
            {
                HideToTray();
            }
        };
        FormClosing += MainForm_FormClosing;
    }

    private void InitializeComponent()
    {
        var mainLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(18)
        };
        mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 120,
            BackColor = Color.FromArgb(22, 74, 92),
            Padding = new Padding(24, 18, 24, 18),
            Margin = new Padding(0, 0, 0, 18)
        };

        var titleLabel = new Label
        {
            AutoSize = true,
            Text = "Troca de IPv4",
            ForeColor = Color.White,
            Font = new Font("Segoe UI Semibold", 18, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 10)
        };

        var subtitleLabel = new Label
        {
            AutoSize = true,
            Text = "Roda em segundo plano, inicia com o Windows e aplica IPv4 estatico ou DHCP pelo servico local.",
            ForeColor = Color.FromArgb(220, 233, 239),
            MaximumSize = new Size(820, 0),
            Margin = new Padding(0)
        };

        headerPanel.Controls.Add(titleLabel);
        headerPanel.Controls.Add(subtitleLabel);
        subtitleLabel.Top = titleLabel.Bottom + 6;

        var configCard = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(20),
            Margin = new Padding(0),
            BorderStyle = BorderStyle.FixedSingle
        };

        var sectionTitleLabel = new Label
        {
            AutoSize = true,
            Text = "Configuracao do adaptador",
            Font = new Font("Segoe UI Semibold", 12.5F, FontStyle.Bold),
            ForeColor = Color.FromArgb(24, 39, 52),
            Margin = new Padding(0, 0, 0, 6)
        };

        var sectionSubtitleLabel = new Label
        {
            AutoSize = true,
            Text = "Escolha o adaptador, defina o modo e aplique a configuracao desejada.",
            ForeColor = Color.FromArgb(92, 104, 115),
            Margin = new Padding(0, 0, 0, 18)
        };

        var formLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            AutoSize = true,
            AutoScroll = true,
            Padding = new Padding(0)
        };
        formLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
        formLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        formLayout.RowStyles.Clear();
        for (var index = 0; index < 6; index++)
        {
            formLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        ConfigureButtons();
        ConfigureInputs();

        AddLabeledControl(formLayout, 0, "Modo", BuildModeRow());
        AddLabeledControl(formLayout, 1, "Adaptador", BuildAdapterRow());
        AddLabeledControl(formLayout, 2, "IPv4", _ipAddressTextBox);
        AddLabeledControl(formLayout, 3, "Prefixo", BuildPrefixRow());
        AddLabeledControl(formLayout, 4, "Gateway", _defaultGatewayTextBox);
        AddLabeledControl(formLayout, 5, "DNS", _dnsServersTextBox);

        configCard.Controls.Add(formLayout);
        configCard.Controls.Add(sectionSubtitleLabel);
        configCard.Controls.Add(sectionTitleLabel);

        sectionTitleLabel.Location = new Point(20, 20);
        sectionSubtitleLabel.Location = new Point(20, sectionTitleLabel.Bottom + 4);
        formLayout.Location = new Point(20, sectionSubtitleLabel.Bottom + 18);
        formLayout.Width = Math.Max(0, configCard.ClientSize.Width - 40);
        configCard.Resize += (_, _) => formLayout.Width = Math.Max(0, configCard.ClientSize.Width - 40);

        var buttonRow = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Margin = new Padding(0, 16, 0, 0)
        };

        buttonRow.Controls.Add(_applyButton);
        buttonRow.Controls.Add(_refreshButton);

        _statusLabel.AutoSize = true;
        _statusLabel.Text = "Pronto.";
        _statusLabel.ForeColor = Color.FromArgb(83, 96, 107);
        _statusLabel.Margin = new Padding(0, 8, 0, 0);

        mainLayout.Controls.Add(headerPanel, 0, 0);
        mainLayout.Controls.Add(configCard, 0, 1);
        mainLayout.Controls.Add(buttonRow, 0, 2);
        mainLayout.Controls.Add(_statusLabel, 0, 3);
        Controls.Add(mainLayout);

        _trayMenu.Items.Add("Abrir", null, (_, _) => ShowMainWindow());
        _trayMenu.Items.Add(new ToolStripSeparator());
        _trayMenu.Items.Add("Sair", null, (_, _) => ExitApplication());

        _notifyIcon.Text = "ipchange";
        _notifyIcon.Icon = SystemIcons.Application;
        _notifyIcon.ContextMenuStrip = _trayMenu;
        _notifyIcon.Visible = true;
        _notifyIcon.DoubleClick += (_, _) => ShowMainWindow();

        AcceptButton = _applyButton;
        UpdateModeState();
    }

    private void ConfigureButtons()
    {
        _refreshButton.Text = "Atualizar adaptadores";
        _refreshButton.AutoSize = true;
        _refreshButton.Padding = new Padding(12, 7, 12, 7);
        _refreshButton.FlatStyle = FlatStyle.Flat;
        _refreshButton.FlatAppearance.BorderColor = Color.FromArgb(180, 190, 197);
        _refreshButton.BackColor = Color.White;
        _refreshButton.Click += async (_, _) => await LoadAdaptersAsync();

        _applyButton.Text = "Aplicar";
        _applyButton.AutoSize = true;
        _applyButton.Padding = new Padding(18, 7, 18, 7);
        _applyButton.FlatStyle = FlatStyle.Flat;
        _applyButton.FlatAppearance.BorderSize = 0;
        _applyButton.BackColor = Color.FromArgb(15, 114, 121);
        _applyButton.ForeColor = Color.White;
        _applyButton.Click += async (_, _) => await ApplyAsync();
    }

    private void ConfigureInputs()
    {
        _ipAddressTextBox.PlaceholderText = "192.168.0.50";
        _defaultGatewayTextBox.PlaceholderText = "192.168.0.1";
        _dnsServersTextBox.PlaceholderText = "1.1.1.1, 8.8.8.8";
        _prefixLengthUpDown.Minimum = 0;
        _prefixLengthUpDown.Maximum = 32;
        _prefixLengthUpDown.Value = 24;

        _staticModeRadioButton.Text = "IPv4 estatico";
        _staticModeRadioButton.AutoSize = true;
        _staticModeRadioButton.Checked = true;
        _staticModeRadioButton.CheckedChanged += (_, _) => UpdateModeState();

        _dhcpModeRadioButton.Text = "DHCP automatico";
        _dhcpModeRadioButton.AutoSize = true;
        _dhcpModeRadioButton.CheckedChanged += (_, _) => UpdateModeState();

        _modeHintLabel.AutoSize = true;
        _modeHintLabel.ForeColor = Color.FromArgb(96, 108, 118);
        _modeHintLabel.Margin = new Padding(16, 6, 0, 0);
    }

    private Control BuildModeRow()
    {
        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Margin = new Padding(0)
        };

        panel.Controls.Add(_staticModeRadioButton);
        panel.Controls.Add(_dhcpModeRadioButton);
        panel.Controls.Add(_modeHintLabel);
        return panel;
    }

    private Control BuildAdapterRow()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            Margin = new Padding(0)
        };

        _adapterComboBox.Dock = DockStyle.Fill;
        _adapterComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
        panel.Controls.Add(_adapterComboBox);
        return panel;
    }

    private Control BuildPrefixRow()
    {
        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false
        };
        _prefixLengthUpDown.Width = 100;
        panel.Controls.Add(_prefixLengthUpDown);
        panel.Controls.Add(new Label { AutoSize = true, Text = "bits", Margin = new Padding(8, 6, 0, 0) });
        return panel;
    }

    private static void AddLabeledControl(TableLayoutPanel layout, int rowIndex, string labelText, Control control)
    {
        var label = new Label
        {
            Text = labelText,
            Anchor = AnchorStyles.Left,
            AutoSize = true,
            Margin = new Padding(0, 8, 12, 8)
        };

        control.Margin = new Padding(0, 4, 0, 4);
        control.Dock = DockStyle.Fill;

        layout.Controls.Add(label, 0, rowIndex);
        layout.Controls.Add(control, 1, rowIndex);
    }

    private async Task InitializeRuntimeAsync()
    {
        SetStatus(NetworkConfigurationService.GetIsAdministrator()
            ? "Executando com privilégio administrativo local."
            : "Pronto. As alterações serão enviadas ao serviço local.");

        CleanupLegacyStartupRegistration();
        await RefreshServiceStatusAsync();
        await LoadAdaptersAsync();
    }

    private void CleanupLegacyStartupRegistration()
    {
        try
        {
            using var startupKey = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", writable: true);
            if (startupKey?.GetValue("ipchange") is null)
            {
                return;
            }

            startupKey.DeleteValue("ipchange", throwOnMissingValue: false);
            AppLogger.Write("INFO", "Entrada legada de inicialização automática do usuário atual foi removida.");
        }
        catch (Exception ex)
        {
            AppLogger.Write("ERROR", $"Falha ao limpar inicialização automática legada: {ex.Message}");
        }
    }

    private async Task RefreshServiceStatusAsync()
    {
        try
        {
            var status = await LocalServiceManager.GetStatusAsync();
            AppLogger.Write("INFO", $"Status do serviço local: {status}.");

            if (NetworkConfigurationService.GetIsAdministrator())
            {
                return;
            }

            if (status == LocalServiceState.Running)
            {
                SetStatus("Serviço local pronto.");
            }
            else if (status == LocalServiceState.NotInstalled)
            {
                SetStatus("Serviço local não instalado. Use o instalador para preparar este PC.");
            }
            else
            {
                SetStatus("Serviço local instalado, mas não iniciado.");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Write("ERROR", $"Falha ao consultar serviço local: {ex.Message}");

            if (!NetworkConfigurationService.GetIsAdministrator())
            {
                SetStatus("Falha ao consultar o serviço local.");
            }
        }
    }

    private async Task LoadAdaptersAsync()
    {
        try
        {
            SetBusy(true);
            var adapters = await Task.Run(NetworkConfigurationService.GetVisibleAdapters);
            _adapterComboBox.BeginUpdate();
            _adapterComboBox.DataSource = adapters.ToList();
            _adapterComboBox.DisplayMember = nameof(AdapterInfo.Name);
            _adapterComboBox.ValueMember = nameof(AdapterInfo.Name);
            _adapterComboBox.EndUpdate();
            SetStatus($"Adaptadores carregados: {adapters.Count}");
        }
        catch (Exception ex)
        {
            SetStatus("Falha ao carregar adaptadores.");
            MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task ApplyAsync()
    {
        try
        {
            SetBusy(true);
            var request = BuildRequest();
            var result = NetworkConfigurationService.GetIsAdministrator()
                ? await NetworkConfigurationService.ApplyConfigurationAsync(request)
                : await RunWithServiceAsync(request);

            var successMessage = FormatResultMessage(result);
            SetStatus(successMessage);
            AppLogger.Write("INFO", successMessage);

            MessageBox.Show(
                result.UseDhcp
                    ? "DHCP aplicado com sucesso."
                    : "Configuração aplicada com sucesso.",
                Text,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            SetStatus("Falha ao aplicar a configuração.");
            AppLogger.Write("ERROR", ex.Message);
            MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private ApplyRequest BuildRequest()
    {
        var adapter = _adapterComboBox.SelectedItem as AdapterInfo;
        var request = new ApplyRequest
        {
            AdapterName = adapter?.Name,
            UseDhcp = _dhcpModeRadioButton.Checked,
            IPAddress = _ipAddressTextBox.Text.Trim(),
            PrefixLength = (int)_prefixLengthUpDown.Value,
            DefaultGateway = string.IsNullOrWhiteSpace(_defaultGatewayTextBox.Text) ? null : _defaultGatewayTextBox.Text.Trim()
        };

        foreach (var dnsValue in _dnsServersTextBox.Text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            request.DnsServers.Add(dnsValue);
        }

        NetworkConfigurationService.ValidateRequest(request);
        return request;
    }

    private void UpdateModeState()
    {
        var useDhcp = _dhcpModeRadioButton.Checked;
        _ipAddressTextBox.Enabled = !useDhcp;
        _prefixLengthUpDown.Enabled = !useDhcp;
        _defaultGatewayTextBox.Enabled = !useDhcp;
        _dnsServersTextBox.Enabled = !useDhcp;

        _modeHintLabel.Text = useDhcp
            ? "IP e DNS serao obtidos automaticamente do DHCP."
            : "Informe IP, prefixo, gateway e DNS manualmente.";
    }

    private static string FormatResultMessage(AdapterConfigurationResult result)
    {
        if (result.UseDhcp)
        {
            return string.IsNullOrWhiteSpace(result.IPAddress)
                ? $"DHCP habilitado no adaptador {result.AdapterName}."
                : $"DHCP habilitado: {result.AdapterName} agora usa {result.IPAddress}/{result.PrefixLength}.";
        }

        return $"IP aplicado: {result.IPAddress}/{result.PrefixLength}";
    }

    private async Task<AdapterConfigurationResult> RunWithServiceAsync(ApplyRequest request)
    {
        SetStatus("Conectando ao serviço local...");

        using var client = new NamedPipeClientStream(".", IpChangeServiceProtocol.PipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
        try
        {
            await client.ConnectAsync(5000);
        }
        catch (TimeoutException)
        {
            throw new InvalidOperationException("O serviço do ipchange não está em execução. Instale e inicie o serviço usando o mesmo ipchange.exe com o parâmetro --service.");
        }

        using var reader = new StreamReader(client, Encoding.UTF8, false, 1024, leaveOpen: true);
        using var writer = new StreamWriter(client, new UTF8Encoding(false), 1024, leaveOpen: true) { AutoFlush = true };

        var payload = JsonSerializer.Serialize(new ServiceApplyRequest { Request = request }, AppJson.CompactOptions);
        await writer.WriteLineAsync(payload);

        SetStatus("Aguardando resposta do serviço...");
        var responseLine = await reader.ReadLineAsync();
        if (string.IsNullOrWhiteSpace(responseLine))
        {
            throw new InvalidOperationException("O serviço do ipchange encerrou a conexão sem retornar resposta.");
        }

        var response = JsonSerializer.Deserialize<ServiceApplyResponse>(responseLine, AppJson.CompactOptions)
            ?? throw new InvalidOperationException("A resposta do serviço não pôde ser interpretada.");

        if (!response.Success || response.Result is null)
        {
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(response.ErrorMessage)
                ? "O serviço do ipchange não conseguiu aplicar a configuração."
                : response.ErrorMessage);
        }

        return response.Result;
    }

    private void SetBusy(bool busy)
    {
        UseWaitCursor = busy;
        _refreshButton.Enabled = !busy;
        _applyButton.Enabled = !busy;
    }

    private void SetStatus(string message)
    {
        _statusLabel.Text = message;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _notifyIcon.Visible = false;
            _notifyIcon.Dispose();
            _trayMenu.Dispose();
        }

        base.Dispose(disposing);
    }

    private void ShowMainWindow()
    {
        Show();
        ShowInTaskbar = true;
        WindowState = FormWindowState.Normal;
        Activate();
    }

    private void HideToTray()
    {
        Hide();
        ShowInTaskbar = false;
        WindowState = FormWindowState.Minimized;
    }

    private void ExitApplication()
    {
        _allowExit = true;
        _notifyIcon.Visible = false;
        Application.Exit();
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_allowExit ||
            e.CloseReason == CloseReason.ApplicationExitCall ||
            e.CloseReason == CloseReason.WindowsShutDown)
        {
            return;
        }

        e.Cancel = true;
        HideToTray();
    }

}