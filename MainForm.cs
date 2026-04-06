using Microsoft.Win32;
using System.IO.Pipes;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Authentication;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;

namespace NetworkConfigurator;

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
    private readonly Label _statusDot = new();
    private readonly Label _modeHintLabel = new();
    private readonly NotifyIcon _notifyIcon = new();
    private readonly ContextMenuStrip _trayMenu = new();
    private readonly bool _startHidden;
    private bool _allowExit;

    // Diagnósticos
    private readonly TextBox _diagHostTextBox = new();
    private readonly Button _pingButton = new();
    private readonly Button _tracerouteButton = new();
    private readonly Button _dnsLookupButton = new();
    private readonly Button _fqdnButton = new();
    private readonly Button _routeButton = new();
    private readonly Button _dot1xButton = new();
    private readonly Button _portScanButton = new();
    private readonly Button _arpButton = new();

    // Subnet Calculator
    private readonly TextBox _subnetIpTextBox = new();
    private readonly NumericUpDown _subnetPrefixUpDown = new();
    private readonly RichTextBox _subnetOutputBox = new();
    private readonly TextBox _portRangeTextBox = new();
    private readonly Button _cancelDiagButton = new();
    private readonly Button _clearDiagButton = new();
    private readonly RichTextBox _diagOutputBox = new();
    private CancellationTokenSource? _diagCts;

    // Diagnósticos: novos botões
    private readonly Button _httpButton = new();
    private readonly Button _sslButton = new();
    private readonly Button _whoisButton = new();
    private readonly Button _bandwidthButton = new();
    private readonly Button _viewLogButton = new();
    private readonly Button _netstatButton = new();
    private readonly Button _wifiScanButton = new();
    private readonly Button _geoIpButton = new();
    private readonly Button _ntpButton = new();
    private readonly Button _dnsFlushButton = new();
    private readonly Button _firewallButton = new();
    private readonly Button _exportDiagButton = new();

    // Novos diagnósticos: Proxy, Firewall Profile, TCP Test, SMB/Shares, Certificados
    private readonly Button _proxyButton = new();
    private readonly Button _fwProfileButton = new();
    private readonly Button _tcpTestButton = new();
    private readonly Button _smbButton = new();
    private readonly Button _certButton = new();

    // Diagnósticos: MTU, IPCONFIG, VPN, Subnet Scan, SMTP Test, Interface Stats
    private readonly Button _mtuButton = new();
    private readonly Button _ipconfigButton = new();
    private readonly Button _vpnButton = new();
    private readonly Button _subnetScanButton = new();
    private readonly Button _smtpTestButton = new();
    private readonly Button _ifStatsButton = new();
    private System.Threading.Timer? _connectivityTimer;
    private bool _lastConnectivityState = true;

    // Dark mode
    private bool _darkMode;
    private readonly Button _darkModeButton = new();

    // Histórico de comandos
    private readonly List<string> _hostHistory = new();
    private int _hostHistoryCursor = -1;
    private readonly Button _copyOutputButton = new();

    // Persistência de janela
    private string WindowSettingsPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "NetworkConfigurator", "window.json");

    // Ferramentas: MAC Vendor
    private readonly TextBox _macInputTextBox = new();
    private readonly RichTextBox _macOutputBox = new();

    // Ferramentas: Wake-on-LAN
    private readonly TextBox _wolIpTextBox = new();
    private readonly TextBox _wolMacTextBox = new();
    private readonly RichTextBox _wolOutputBox = new();

    // Ferramentas: CIDR Range
    private readonly TextBox _cidrInputTextBox = new();
    private readonly RichTextBox _cidrOutputBox = new();

    // Ferramentas: IP Converter
    private readonly TextBox _ipConvTextBox = new();
    private readonly RichTextBox _ipConvOutputBox = new();

    // Perfis de rede
    private readonly ComboBox _profileComboBox = new();
    private readonly TextBox _profileNameTextBox = new();
    private List<NetworkProfile> _profiles = new();

    public MainForm(bool startHidden)
    {
        _startHidden = startHidden;

        Text = "Network Configurator";
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Segoe UI", 9.5F);
        MinimumSize = new Size(960, 580);
        Width = 1300;
        Height = 800;
        BackColor = Color.FromArgb(235, 238, 243);

        // Load app icon from assets folder next to the executable
        var iconPath = Path.Combine(AppContext.BaseDirectory, "assets", "icon.ico");
        if (File.Exists(iconPath))
            Icon = new Icon(iconPath);

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
            RowCount = 3,
            Padding = new Padding(18)
        };
        mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
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
            Text = "Network Configurator",
            ForeColor = Color.White,
            Font = new Font("Segoe UI Semibold", 18, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 10)
        };

        var subtitleLabel = new Label
        {
            AutoSize = true,
            Text = "Roda em segundo plano, inicia com o Windows e aplica IPv4 estático ou DHCP pelo serviço local.",
            ForeColor = Color.FromArgb(220, 233, 239),
            MaximumSize = new Size(820, 0),
            Margin = new Padding(0)
        };

        headerPanel.Controls.Add(titleLabel);
        headerPanel.Controls.Add(subtitleLabel);
        subtitleLabel.Top = titleLabel.Bottom + 6;

        // Dark mode toggle no canto superior direito do header
        _darkModeButton.Text = "🌙 Dark";
        _darkModeButton.AutoSize = true;
        _darkModeButton.FlatStyle = FlatStyle.Flat;
        _darkModeButton.FlatAppearance.BorderColor = Color.FromArgb(100, 160, 180);
        _darkModeButton.BackColor = Color.FromArgb(22, 74, 92);
        _darkModeButton.ForeColor = Color.White;
        _darkModeButton.Padding = new Padding(8, 4, 8, 4);
        _darkModeButton.Click += (_, _) => ToggleDarkMode();
        headerPanel.Controls.Add(_darkModeButton);
        headerPanel.Resize += (_, _) =>
            _darkModeButton.Location = new Point(headerPanel.ClientSize.Width - _darkModeButton.Width - 16, 16);
        headerPanel.Layout += (_, _) =>
            _darkModeButton.Location = new Point(headerPanel.ClientSize.Width - _darkModeButton.Width - 16, 16);

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
            Text = "Configuração do adaptador",
            Font = new Font("Segoe UI Semibold", 12.5F, FontStyle.Bold),
            ForeColor = Color.FromArgb(24, 39, 52),
            Margin = new Padding(0, 0, 0, 6)
        };

        var sectionSubtitleLabel = new Label
        {
            AutoSize = true,
            Text = "Escolha o adaptador, defina o modo e aplique a configuração desejada.",
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
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 0)
        };

        buttonRow.Controls.Add(_applyButton);
        buttonRow.Controls.Add(_refreshButton);

        var configTabLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(0)
        };
        configTabLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        configTabLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        configTabLayout.Controls.Add(configCard, 0, 0);
        configTabLayout.Controls.Add(buttonRow, 0, 1);

        var configTab = new TabPage { Text = "Configuração", Padding = new Padding(6) };
        configTab.Controls.Add(configTabLayout);

        var diagTab = new TabPage { Text = "Diagnósticos", Padding = new Padding(6) };
        diagTab.Controls.Add(BuildDiagnosticsPanel());

        var toolsTab = new TabPage { Text = "Ferramentas", Padding = new Padding(6) };
        toolsTab.Controls.Add(BuildToolsPanel());

        var profilesTab = new TabPage { Text = "Perfis", Padding = new Padding(6) };
        profilesTab.Controls.Add(BuildProfilesPanel());

        var tabControl = new TabControl { Dock = DockStyle.Fill, Margin = new Padding(0) };
        tabControl.TabPages.Add(configTab);
        tabControl.TabPages.Add(diagTab);
        tabControl.TabPages.Add(toolsTab);
        tabControl.TabPages.Add(profilesTab);

        _statusLabel.AutoSize = true;
        _statusLabel.Text = "Pronto.";
        _statusLabel.ForeColor = Color.FromArgb(83, 96, 107);
        _statusLabel.Margin = new Padding(0, 8, 0, 0);

        mainLayout.Controls.Add(headerPanel, 0, 0);
        mainLayout.Controls.Add(tabControl, 0, 1);
        mainLayout.Controls.Add(_statusLabel, 0, 2);
        Controls.Add(mainLayout);

        _trayMenu.Items.Add("Abrir", null, (_, _) => ShowMainWindow());
        _trayMenu.Items.Add(new ToolStripSeparator());
        _trayMenu.Items.Add("Sair", null, (_, _) => ExitApplication());

        _notifyIcon.Text = "Network Configurator";
        var trayIconPath = Path.Combine(AppContext.BaseDirectory, "assets", "icon.ico");
        _notifyIcon.Icon = File.Exists(trayIconPath) ? new Icon(trayIconPath) : SystemIcons.Application;
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

        _staticModeRadioButton.Text = "IPv4 estático";
        _staticModeRadioButton.AutoSize = true;
        _staticModeRadioButton.Checked = true;
        _staticModeRadioButton.CheckedChanged += (_, _) => UpdateModeState();

        _dhcpModeRadioButton.Text = "DHCP automático";
        _dhcpModeRadioButton.AutoSize = true;
        _dhcpModeRadioButton.CheckedChanged += (_, _) => UpdateModeState();

        _modeHintLabel.AutoSize = true;
        _modeHintLabel.ForeColor = Color.FromArgb(96, 108, 118);
        _modeHintLabel.Margin = new Padding(16, 6, 0, 0);
    }

    private Control BuildDiagnosticsPanel()
    {
        // 4-row layout: input bar | category tabs | action bar | output
        var panel = new TableLayoutPanel
        {
            Dock        = DockStyle.Fill,
            ColumnCount = 1,
            RowCount    = 4,
            Padding     = new Padding(0)
        };
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // input bar
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));  // category tabs
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // action bar
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  // output box

        // --- Style all buttons ---
        StyleDiagButton(_pingButton,       "Ping");
        StyleDiagButton(_tracerouteButton, "Traceroute");
        StyleDiagButton(_dnsLookupButton,  "DNS Lookup");
        StyleDiagButton(_fqdnButton,       "FQDN");
        StyleDiagButton(_routeButton,      "Rotas");
        StyleDiagButton(_dot1xButton,      "802.1X");
        StyleDiagButton(_ntpButton,        "NTP");
        StyleDiagButton(_dnsFlushButton,   "Flush DNS");
        StyleDiagButton(_portScanButton,   "Port Scan");
        StyleDiagButton(_tcpTestButton,    "TCP Teste");
        StyleDiagButton(_arpButton,        "ARP");
        StyleDiagButton(_subnetScanButton, "Subnet Scan");
        StyleDiagButton(_httpButton,       "HTTP Headers");
        StyleDiagButton(_sslButton,        "SSL/TLS");
        StyleDiagButton(_whoisButton,      "WHOIS");
        StyleDiagButton(_bandwidthButton,  "Bandwidth");
        StyleDiagButton(_mtuButton,        "MTU");
        StyleDiagButton(_netstatButton,    "NetStat");
        StyleDiagButton(_ipconfigButton,   "IPConfig");
        StyleDiagButton(_ifStatsButton,    "IF Stats");
        StyleDiagButton(_vpnButton,        "VPN");
        StyleDiagButton(_wifiScanButton,   "Wi-Fi");
        StyleDiagButton(_geoIpButton,      "Geo IP");
        StyleDiagButton(_firewallButton,   "Firewall");
        StyleDiagButton(_fwProfileButton,  "FW Perfil");
        StyleDiagButton(_proxyButton,      "Proxy");
        StyleDiagButton(_certButton,       "Certificados");
        StyleDiagButton(_smbButton,        "SMB/Shares");
        StyleDiagButton(_smtpTestButton,   "SMTP Test");
        StyleDiagButton(_viewLogButton,    "Ver Log");
        StyleDiagButton(_exportDiagButton, "Exportar");
        StyleDiagButton(_copyOutputButton, "Copiar");
        StyleDiagButton(_cancelDiagButton, "Cancelar");
        StyleDiagButton(_clearDiagButton,  "Limpar");

        _cancelDiagButton.Enabled = false;
        _cancelDiagButton.BackColor = Color.FromArgb(180, 40, 40);
        _cancelDiagButton.ForeColor = Color.White;
        _cancelDiagButton.FlatAppearance.BorderSize = 0;

        // --- Wire click events ---
        _pingButton.Click       += async (_, _) => await RunPingAsync();
        _tracerouteButton.Click += async (_, _) => await RunTracerouteAsync();
        _dnsLookupButton.Click  += async (_, _) => await RunDnsLookupAsync();
        _fqdnButton.Click       += async (_, _) => await RunFqdnAsync();
        _routeButton.Click      += async (_, _) => await RunRouteTableAsync();
        _dot1xButton.Click      += async (_, _) => await RunDot1xAsync();
        _ntpButton.Click        += async (_, _) => await RunNtpCheckAsync();
        _dnsFlushButton.Click   += async (_, _) => await RunDnsFlushAsync();
        _portScanButton.Click   += async (_, _) => await RunPortScanAsync();
        _tcpTestButton.Click    += async (_, _) => await RunTcpTestAsync();
        _arpButton.Click        += async (_, _) => await RunArpTableAsync();
        _subnetScanButton.Click += async (_, _) => await RunSubnetScanAsync();
        _httpButton.Click       += async (_, _) => await RunHttpHeadersAsync();
        _sslButton.Click        += async (_, _) => await RunSslInspectorAsync();
        _whoisButton.Click      += async (_, _) => await RunWhoisAsync();
        _bandwidthButton.Click  += async (_, _) => await RunBandwidthAsync();
        _mtuButton.Click        += async (_, _) => await RunMtuDiscoveryAsync();
        _netstatButton.Click    += async (_, _) => await RunNetStatAsync();
        _ipconfigButton.Click   += async (_, _) => await RunIpConfigAsync();
        _ifStatsButton.Click    += async (_, _) => await RunIfStatsAsync();
        _vpnButton.Click        += async (_, _) => await RunVpnStatusAsync();
        _wifiScanButton.Click   += async (_, _) => await RunWifiScanAsync();
        _geoIpButton.Click      += async (_, _) => await RunGeoIpAsync();
        _firewallButton.Click   += async (_, _) => await RunFirewallRulesAsync();
        _fwProfileButton.Click  += async (_, _) => await RunFirewallProfileAsync();
        _proxyButton.Click      += async (_, _) => await RunProxyDiagAsync();
        _certButton.Click       += async (_, _) => await RunCertStoreDiagAsync();
        _smbButton.Click        += async (_, _) => await RunSmbDiagAsync();
        _smtpTestButton.Click   += async (_, _) => await RunSmtpTestAsync();
        _viewLogButton.Click    += (_, _) => ShowLog();
        _exportDiagButton.Click += (_, _) => ExportDiagnosticOutput();
        _copyOutputButton.Click += (_, _) => CopyDiagnosticOutput();
        _cancelDiagButton.Click += (_, _) => _diagCts?.Cancel();
        _clearDiagButton.Click  += (_, _) => _diagOutputBox.Clear();

        // --- Input controls (shared across all tabs) ---
        _diagHostTextBox.Width = 200;
        _diagHostTextBox.PlaceholderText = "google.com ou 8.8.8.8";
        _diagHostTextBox.Margin = new Padding(0, 3, 8, 0);
        _diagHostTextBox.KeyDown  += DiagHostTextBox_KeyDown;
        _diagHostTextBox.KeyPress += (_, e) =>
        {
            if (e.KeyChar == (char)13) { e.Handled = true; _pingButton.PerformClick(); }
        };

        _portRangeTextBox.Width = 155;
        _portRangeTextBox.PlaceholderText = "80,443 ou 1-1024";
        _portRangeTextBox.Margin = new Padding(0, 3, 6, 0);

        // --- Row 1: input bar always visible above tabs ---
        var inputBar = new FlowLayoutPanel
        {
            Dock          = DockStyle.Fill,
            AutoSize      = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents  = false,
            BackColor     = Color.FromArgb(240, 243, 248),
            Padding       = new Padding(8, 6, 8, 6)
        };
        inputBar.Controls.Add(new Label { Text = "Host / IP:", AutoSize = true, Margin = new Padding(0, 5, 6, 0), ForeColor = Color.FromArgb(55, 65, 80) });
        inputBar.Controls.Add(_diagHostTextBox);
        inputBar.Controls.Add(new Label { Text = "Portas:", AutoSize = true, Margin = new Padding(18, 5, 6, 0), ForeColor = Color.FromArgb(55, 65, 80) });
        inputBar.Controls.Add(_portRangeTextBox);

        // --- Row 2: category tabs (each with 6-9 buttons, no wrapping) ---
        var catTabs = new TabControl
        {
            Dock    = DockStyle.Fill,
            Padding = new Point(14, 4),
        };

        static TabPage MakeCatTab(string title) => new()
        {
            Text                  = title,
            Padding               = new Padding(6, 6, 6, 4),
            UseVisualStyleBackColor = true
        };

        var tabRede = MakeCatTab("Rede");
        var rowRede = MakeDiagRow();
        rowRede.Controls.AddRange(new Control[]
        {
            _pingButton, _tracerouteButton, _dnsLookupButton, _fqdnButton,
            _routeButton, _dot1xButton, _ntpButton, _dnsFlushButton
        });
        tabRede.Controls.Add(rowRede);

        var tabVarredura = MakeCatTab("Varredura");
        var rowVarredura = MakeDiagRow();
        rowVarredura.Controls.AddRange(new Control[]
        {
            _portScanButton, _tcpTestButton, _arpButton, _subnetScanButton,
            _httpButton, _sslButton, _whoisButton, _bandwidthButton, _mtuButton
        });
        tabVarredura.Controls.Add(rowVarredura);

        var tabSistema = MakeCatTab("Sistema");
        var rowSistema = MakeDiagRow();
        rowSistema.Controls.AddRange(new Control[]
        {
            _netstatButton, _ipconfigButton, _ifStatsButton,
            _vpnButton, _wifiScanButton, _geoIpButton
        });
        tabSistema.Controls.Add(rowSistema);

        var tabSeg = MakeCatTab("Segurança");
        var rowSeg = MakeDiagRow();
        rowSeg.Controls.AddRange(new Control[]
        {
            _firewallButton, _fwProfileButton, _proxyButton,
            _certButton, _smbButton, _smtpTestButton
        });
        tabSeg.Controls.Add(rowSeg);

        catTabs.TabPages.AddRange(new TabPage[] { tabRede, tabVarredura, tabSistema, tabSeg });

        // --- Row 3: action bar ---
        var actionBar = new TableLayoutPanel
        {
            Dock        = DockStyle.Fill,
            ColumnCount = 3,
            AutoSize    = true,
            BackColor   = Color.FromArgb(240, 243, 248),
            Padding     = new Padding(6, 4, 6, 4)
        };
        actionBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        actionBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        actionBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var outputTools = new FlowLayoutPanel { AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
        outputTools.Controls.AddRange(new Control[] { _viewLogButton, _exportDiagButton, _copyOutputButton });

        var cancelTools = new FlowLayoutPanel { AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
        cancelTools.Controls.AddRange(new Control[] { _cancelDiagButton, _clearDiagButton });

        actionBar.Controls.Add(outputTools, 0, 0);
        actionBar.Controls.Add(cancelTools, 2, 0);

        // --- Row 4: output box ---
        _diagOutputBox.Dock        = DockStyle.Fill;
        _diagOutputBox.ReadOnly    = true;
        _diagOutputBox.BackColor   = Color.FromArgb(15, 17, 23);
        _diagOutputBox.ForeColor   = Color.FromArgb(180, 230, 180);
        _diagOutputBox.Font        = new Font("Consolas", 9.5F);
        _diagOutputBox.BorderStyle = BorderStyle.None;
        _diagOutputBox.ScrollBars  = RichTextBoxScrollBars.Vertical;

        panel.Controls.Add(inputBar,       0, 0);
        panel.Controls.Add(catTabs,        0, 1);
        panel.Controls.Add(actionBar,      0, 2);
        panel.Controls.Add(_diagOutputBox, 0, 3);
        return panel;
    }

    private static FlowLayoutPanel MakeDiagRow() => new()
    {
        Dock          = DockStyle.Fill,
        AutoSize      = true,
        FlowDirection = FlowDirection.LeftToRight,
        WrapContents  = false,
        Padding       = new Padding(0, 0, 0, 2),
        Margin        = new Padding(0)
    };

    private static void StyleDiagButton(Button button, string text)
    {
        button.Text = text;
        button.AutoSize = true;
        button.Padding = new Padding(10, 5, 10, 5);
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = Color.FromArgb(180, 190, 197);
        button.BackColor = Color.White;
        button.Margin = new Padding(0, 2, 6, 0);
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
        _adapterComboBox.SelectedIndexChanged += OnAdapterSelectionChanged;
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

    private Control BuildToolsPanel()
    {
        var tc = new TabControl { Dock = DockStyle.Fill };

        var subnetTab = new TabPage { Text = "Sub-rede", Padding = new Padding(6) };
        subnetTab.Controls.Add(BuildSubnetPanel());

        var macTab = new TabPage { Text = "MAC Vendor", Padding = new Padding(6) };
        macTab.Controls.Add(BuildMacVendorPanel());

        var wolTab = new TabPage { Text = "Wake-on-LAN", Padding = new Padding(6) };
        wolTab.Controls.Add(BuildWakeOnLanPanel());

        var cidrTab = new TabPage { Text = "CIDR Range", Padding = new Padding(6) };
        cidrTab.Controls.Add(BuildCidrRangePanel());

        var ipConvTab = new TabPage { Text = "IP Converter", Padding = new Padding(6) };
        ipConvTab.Controls.Add(BuildIpConverterPanel());

        tc.TabPages.Add(subnetTab);
        tc.TabPages.Add(macTab);
        tc.TabPages.Add(wolTab);
        tc.TabPages.Add(cidrTab);
        tc.TabPages.Add(ipConvTab);
        return tc;
    }

    private Control BuildSubnetPanel()
    {
        var outer = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(0)
        };
        outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var titleLabel = new Label
        {
            Text = "Calculadora de Sub-rede",
            Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
            ForeColor = Color.FromArgb(24, 39, 52),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 4)
        };
        var subtitleLabel = new Label
        {
            Text = "Informe um IPv4 e o comprimento do prefixo para calcular os parâmetros da rede.",
            ForeColor = Color.FromArgb(92, 104, 115),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 10)
        };

        // ── Linha de entrada ──────────────────────────────────────────────
        var inputRow = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Margin = new Padding(0, 0, 0, 10)
        };

        var ipLabel = new Label { Text = "IPv4:", AutoSize = true, Margin = new Padding(0, 7, 6, 0) };
        _subnetIpTextBox.Width = 160;
        _subnetIpTextBox.PlaceholderText = "192.168.1.50";
        _subnetIpTextBox.Margin = new Padding(0, 4, 8, 0);

        var prefixLabel = new Label { Text = "/", AutoSize = true, Margin = new Padding(0, 7, 4, 0), Font = new Font("Segoe UI Semibold", 11F) };
        _subnetPrefixUpDown.Minimum = 0;
        _subnetPrefixUpDown.Maximum = 32;
        _subnetPrefixUpDown.Value = 24;
        _subnetPrefixUpDown.Width = 60;
        _subnetPrefixUpDown.Margin = new Padding(0, 4, 10, 0);

        var calcButton = new Button
        {
            Text = "Calcular",
            AutoSize = true,
            Padding = new Padding(14, 6, 14, 6),
            FlatStyle = FlatStyle.Flat,
            FlatAppearance = { BorderSize = 0 },
            BackColor = Color.FromArgb(15, 114, 121),
            ForeColor = Color.White,
            Margin = new Padding(0, 2, 0, 0)
        };
        calcButton.Click += (_, _) => RunSubnetCalc();

        inputRow.Controls.Add(ipLabel);
        inputRow.Controls.Add(_subnetIpTextBox);
        inputRow.Controls.Add(prefixLabel);
        inputRow.Controls.Add(_subnetPrefixUpDown);
        inputRow.Controls.Add(calcButton);

        var headerPanel = new Panel { AutoSize = true, Dock = DockStyle.Top };
        headerPanel.Controls.Add(inputRow);
        headerPanel.Controls.Add(subtitleLabel);
        headerPanel.Controls.Add(titleLabel);
        titleLabel.Location = new Point(0, 0);
        subtitleLabel.Location = new Point(0, titleLabel.Bottom + 2);
        inputRow.Location = new Point(0, subtitleLabel.Bottom + 4);

        // ── Saída ─────────────────────────────────────────────────────────
        _subnetOutputBox.Dock = DockStyle.Fill;
        _subnetOutputBox.ReadOnly = true;
        _subnetOutputBox.BackColor = Color.FromArgb(15, 17, 23);
        _subnetOutputBox.ForeColor = Color.FromArgb(180, 230, 180);
        _subnetOutputBox.Font = new Font("Consolas", 10F);
        _subnetOutputBox.BorderStyle = BorderStyle.None;
        _subnetOutputBox.ScrollBars = RichTextBoxScrollBars.Vertical;

        outer.Controls.Add(headerPanel, 0, 0);
        outer.Controls.Add(_subnetOutputBox, 0, 1);
        return outer;
    }

    private void RunSubnetCalc()
    {
        _subnetOutputBox.Clear();

        var ipText = _subnetIpTextBox.Text.Trim();
        if (!IPAddress.TryParse(ipText, out var ipAddr) ||
            ipAddr.AddressFamily != System.Net.Sockets.AddressFamily.InterNetwork)
        {
            AppendSubnetLine("IPv4 inválido.", Color.OrangeRed);
            return;
        }

        var prefix = (int)_subnetPrefixUpDown.Value;

        // Máscara e wildcard
        var maskUint = prefix == 0 ? 0u : uint.MaxValue << (32 - prefix);
        var wildcardUint = ~maskUint;

        var maskBytes = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((int)maskUint));
        var mask = new IPAddress(maskBytes);

        var wildcardBytes = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((int)wildcardUint));
        var wildcard = new IPAddress(wildcardBytes);

        // Endereço de rede e broadcast
        var addrBytes = ipAddr.GetAddressBytes();
        var maskBytesArr = mask.GetAddressBytes();

        var networkBytes = new byte[4];
        var broadcastBytes = new byte[4];
        for (var i = 0; i < 4; i++)
        {
            networkBytes[i] = (byte)(addrBytes[i] & maskBytesArr[i]);
            broadcastBytes[i] = (byte)(addrBytes[i] | ~maskBytesArr[i]);
        }

        var network = new IPAddress(networkBytes);
        var broadcast = new IPAddress(broadcastBytes);

        // Primeiro e último host
        var firstHostBytes = (byte[])networkBytes.Clone();
        firstHostBytes[3] = (byte)(networkBytes[3] + (prefix < 32 ? 1 : 0));

        var lastHostBytes = (byte[])broadcastBytes.Clone();
        lastHostBytes[3] = (byte)(broadcastBytes[3] - (prefix < 32 ? 1 : 0));

        var firstHost = new IPAddress(firstHostBytes);
        var lastHost = new IPAddress(lastHostBytes);

        // Número de hosts
        var totalHosts = prefix >= 32 ? 1L : (long)Math.Pow(2, 32 - prefix);
        var usableHosts = prefix >= 31 ? totalHosts : Math.Max(0, totalHosts - 2);

        // Classe da rede
        var firstOctet = networkBytes[0];
        var networkClass = firstOctet switch
        {
            < 128 => "A",
            < 192 => "B",
            < 224 => "C",
            < 240 => "D (Multicast)",
            _ => "E (Reservado)"
        };

        // RFC 1918 (privado)?
        var isPrivate =
            (networkBytes[0] == 10) ||
            (networkBytes[0] == 172 && networkBytes[1] >= 16 && networkBytes[1] <= 31) ||
            (networkBytes[0] == 192 && networkBytes[1] == 168);

        AppendSubnetLine($"=== Subnet Calculator: {ipText}/{prefix} ===", Color.Cyan);
        AppendSubnetLine(string.Empty, Color.White);
        AppendSubnetLine(string.Format("  {0,-26} {1}", "Endereço IP:", ipAddr), Color.White);
        AppendSubnetLine(string.Format("  {0,-26} {1}", "CIDR:", $"{ipAddr}/{prefix}"), Color.White);
        AppendSubnetLine(string.Format("  {0,-26} {1}", "Máscara de Sub-rede:", mask), Color.FromArgb(180, 230, 180));
        AppendSubnetLine(string.Format("  {0,-26} {1}", "Wildcard (inversa):", wildcard), Color.FromArgb(180, 230, 180));
        AppendSubnetLine(string.Empty, Color.White);
        AppendSubnetLine(string.Format("  {0,-26} {1}", "Rede:", network), Color.FromArgb(255, 220, 140));
        AppendSubnetLine(string.Format("  {0,-26} {1}", "Broadcast:", broadcast), Color.FromArgb(255, 180, 100));
        AppendSubnetLine(string.Format("  {0,-26} {1}", "Primeiro host:", firstHost), Color.LightGreen);
        AppendSubnetLine(string.Format("  {0,-26} {1}", "Último host:", lastHost), Color.LightGreen);
        AppendSubnetLine(string.Format("  {0,-26} {1}", "Range utilizável:", $"{firstHost} – {lastHost}"), Color.LightGreen);
        AppendSubnetLine(string.Empty, Color.White);
        AppendSubnetLine(string.Format("  {0,-26} {1:N0}", "Total de endereços:", totalHosts), Color.White);
        AppendSubnetLine(string.Format("  {0,-26} {1:N0}", "Hosts utilizáveis:", usableHosts), Color.White);
        AppendSubnetLine(string.Format("  {0,-26} Classe {1}", "Classe:", networkClass), Color.White);
        AppendSubnetLine(string.Format("  {0,-26} {1}", "Espaço:", isPrivate ? "Privado (RFC 1918)" : "Público"), isPrivate ? Color.FromArgb(140, 200, 255) : Color.FromArgb(200, 200, 200));

        // Subnets disponíveis para split visual se prefix < 30
        if (prefix <= 28)
        {
            AppendSubnetLine(string.Empty, Color.White);
            AppendSubnetLine("  Subnets /28 mais próximas dentro desta rede:", Color.FromArgb(140, 180, 255));
            var subPrefix = Math.Min(prefix + 4, 28);
            var subCount = (int)Math.Pow(2, subPrefix - prefix);
            var subSize = (int)Math.Pow(2, 32 - subPrefix);
            var baseNet = BitConverter.ToUInt32(networkBytes.Reverse().ToArray(), 0);
            for (var s = 0; s < Math.Min(subCount, 16); s++)
            {
                var subNetUint = baseNet + (uint)(s * subSize);
                var subNetBytes = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((int)subNetUint));
                AppendSubnetLine($"    {new IPAddress(subNetBytes)}/{subPrefix}", Color.FromArgb(160, 210, 160));
            }
            if (subCount > 16)
                AppendSubnetLine($"    ... (+{subCount - 16} mais)", Color.Gray);
        }
    }

    private void AppendSubnetLine(string text, Color color)
    {
        _subnetOutputBox.SelectionStart = _subnetOutputBox.TextLength;
        _subnetOutputBox.SelectionLength = 0;
        _subnetOutputBox.SelectionColor = color;
        _subnetOutputBox.AppendText(text + Environment.NewLine);
        _subnetOutputBox.ScrollToCaret();
    }

    private async Task InitializeRuntimeAsync()
    {
        SetStatus(NetworkConfigurationService.GetIsAdministrator()
            ? "Executando com privilégio administrativo local."
            : "Pronto. As alterações serão enviadas ao serviço local.");

        CleanupLegacyStartupRegistration();
        await RefreshServiceStatusAsync();
        await LoadAdaptersAsync();
        await LoadProfilesAsync();
        StartConnectivityMonitor();
        LoadWindowSettings();
    }

    private void CleanupLegacyStartupRegistration()
    {
        try
        {
            using var startupKey = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", writable: true);
            if (startupKey?.GetValue("ipchange") is null && startupKey?.GetValue("Network-Configurator") is null)
            {
                return;
            }

            startupKey?.DeleteValue("ipchange", throwOnMissingValue: false);
            startupKey?.DeleteValue("Network-Configurator", throwOnMissingValue: false);
            AppLogger.Write("INFO", "Entradas legadas de inicialização automática do usuário atual foram removidas.");
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
            _adapterComboBox.DisplayMember = nameof(AdapterInfo.DisplayName);
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

    private void OnAdapterSelectionChanged(object? sender, EventArgs e)
    {
        if (_adapterComboBox.SelectedItem is not AdapterInfo adapter)
        {
            return;
        }

        if (adapter.IsDhcpEnabled)
        {
            _dhcpModeRadioButton.Checked = true;
            return;
        }

        if (!string.IsNullOrWhiteSpace(adapter.CurrentIPAddress))
        {
            _staticModeRadioButton.Checked = true;
            _ipAddressTextBox.Text = adapter.CurrentIPAddress;
            _prefixLengthUpDown.Value = adapter.CurrentPrefixLength ?? 24;
            _defaultGatewayTextBox.Text = adapter.CurrentDefaultGateway ?? string.Empty;
            _dnsServersTextBox.Text = adapter.CurrentDnsServers.Length > 0
                ? string.Join(", ", adapter.CurrentDnsServers)
                : string.Empty;
        }
    }

    private void UpdateModeState()
    {
        var useDhcp = _dhcpModeRadioButton.Checked;
        _ipAddressTextBox.Enabled = !useDhcp;
        _prefixLengthUpDown.Enabled = !useDhcp;
        _defaultGatewayTextBox.Enabled = !useDhcp;
        _dnsServersTextBox.Enabled = !useDhcp;

        _modeHintLabel.Text = useDhcp
            ? "IP e DNS serão obtidos automaticamente do DHCP."
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

        using var client = new NamedPipeClientStream(".", NetworkConfiguratorServiceProtocol.PipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
        try
        {
            await client.ConnectAsync(5000);
        }
        catch (TimeoutException)
        {
            throw new InvalidOperationException("O serviço do Network Configurator não está em execução. Instale e inicie o serviço usando o mesmo network-configurator.exe com o parâmetro --service.");
        }

        using var reader = new StreamReader(client, Encoding.UTF8, false, 1024, leaveOpen: true);
        using var writer = new StreamWriter(client, new UTF8Encoding(false), 1024, leaveOpen: true) { AutoFlush = true };

        var payload = JsonSerializer.Serialize(new ServiceApplyRequest { Request = request }, AppJson.CompactOptions);
        await writer.WriteLineAsync(payload);

        SetStatus("Aguardando resposta do serviço...");
        var responseLine = await reader.ReadLineAsync();
        if (string.IsNullOrWhiteSpace(responseLine))
        {
            throw new InvalidOperationException("O serviço encerrou a conexão sem retornar resposta.");
        }

        var response = JsonSerializer.Deserialize<ServiceApplyResponse>(responseLine, AppJson.CompactOptions)
            ?? throw new InvalidOperationException("A resposta do serviço não pôde ser interpretada.");

        if (!response.Success || response.Result is null)
        {
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(response.ErrorMessage)
                ? "O serviço não conseguiu aplicar a configuração."
                : response.ErrorMessage);
        }

        return response.Result;
    }

    private async Task RunPingAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe um host ou IP.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- Ping {host} ---", Color.Cyan);

        try
        {
            using var ping = new Ping();
            var succeeded = 0;
            var totalMs = 0L;
            const int count = 4;

            for (var i = 0; i < count; i++)
            {
                if (_diagCts.Token.IsCancellationRequested) break;

                try
                {
                    var reply = await ping.SendPingAsync(host, TimeSpan.FromSeconds(3), null, null, _diagCts.Token);
                    if (reply.Status == IPStatus.Success)
                    {
                        succeeded++;
                        totalMs += reply.RoundtripTime;
                        AppendDiagLine($"Resposta de {reply.Address}: bytes=32 tempo={reply.RoundtripTime}ms TTL={reply.Options?.Ttl ?? 0}", Color.LightGreen);
                    }
                    else
                    {
                        AppendDiagLine($"Falha ({reply.Status})", Color.OrangeRed);
                    }
                }
                catch (PingException ex)
                {
                    AppendDiagLine($"Erro: {ex.InnerException?.Message ?? ex.Message}", Color.OrangeRed);
                }

                if (i < count - 1 && !_diagCts.Token.IsCancellationRequested)
                    await Task.Delay(1000, _diagCts.Token);
            }

            AppendDiagLine(
                $"Estatísticas: {succeeded}/{count} recebidos" +
                (succeeded > 0 ? $", média {totalMs / succeeded}ms" : "") + ".",
                Color.White);
        }
        catch (OperationCanceledException)
        {
            AppendDiagLine("Cancelado.", Color.Yellow);
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed);
        }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private async Task RunTracerouteAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe um host ou IP.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- Traceroute {host} (máx 30 saltos) ---", Color.Cyan);

        try
        {
            using var ping = new Ping();
            const int maxHops = 30;

            for (var ttl = 1; ttl <= maxHops; ttl++)
            {
                if (_diagCts.Token.IsCancellationRequested) break;

                PingReply reply;
                try
                {
                    var options = new PingOptions(ttl, dontFragment: true);
                    reply = await ping.SendPingAsync(host, TimeSpan.FromSeconds(3), new byte[32], options, _diagCts.Token);
                }
                catch (PingException ex)
                {
                    AppendDiagLine($"{ttl,3}.  *  Erro: {ex.InnerException?.Message ?? ex.Message}", Color.OrangeRed);
                    continue;
                }

                var hopAddress = reply.Address?.ToString() ?? "*";
                var hopTime = reply.Status is IPStatus.Success or IPStatus.TtlExpired
                    ? $"{reply.RoundtripTime}ms"
                    : "*";
                var lineColor = reply.Status == IPStatus.Success ? Color.LightGreen
                    : reply.Status == IPStatus.TtlExpired ? Color.FromArgb(160, 210, 160)
                    : Color.Gray;

                AppendDiagLine($"{ttl,3}.  {hopAddress,-46} {hopTime}", lineColor);

                if (reply.Status == IPStatus.Success) break;

                if (!_diagCts.Token.IsCancellationRequested)
                    await Task.Delay(150, _diagCts.Token);
            }

            AppendDiagLine("Traceroute concluído.", Color.White);
        }
        catch (OperationCanceledException)
        {
            AppendDiagLine("Cancelado.", Color.Yellow);
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed);
        }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private async Task RunDnsLookupAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe um host para a consulta DNS.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- DNS Lookup: {host} ---", Color.Cyan);

        try
        {
            var entry = await Dns.GetHostEntryAsync(host, _diagCts.Token);
            AppendDiagLine($"Nome canônico : {entry.HostName}", Color.White);

            if (entry.Aliases.Length > 0)
                AppendDiagLine($"Aliases       : {string.Join(", ", entry.Aliases)}", Color.White);

            foreach (var addr in entry.AddressList)
            {
                var family = addr.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6 ? "IPv6" : "IPv4";
                AppendDiagLine($"  {family,-6}: {addr}", Color.LightGreen);
            }

            AppendDiagLine($"Total: {entry.AddressList.Length} endereço(s).", Color.White);
        }
        catch (OperationCanceledException)
        {
            AppendDiagLine("Cancelado.", Color.Yellow);
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed);
        }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private async Task RunFqdnAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            // Se vazio, usa o hostname local
            host = Dns.GetHostName();
        }
        AddToHostHistory(host);

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- FQDN / rDNS: {host} ---", Color.Cyan);

        try
        {
            // Resolve para IP primeiro (caso seja nome)
            IPAddress[] addresses;
            try
            {
                addresses = await Dns.GetHostAddressesAsync(host, _diagCts.Token);
            }
            catch
            {
                addresses = Array.Empty<IPAddress>();
            }

            // Reverse lookup de cada IP encontrado (ou do próprio valor se for IP)
            IPAddress[] targets;
            if (addresses.Length > 0)
            {
                targets = addresses;
            }
            else if (IPAddress.TryParse(host, out var parsed))
            {
                targets = new[] { parsed };
            }
            else
            {
                targets = Array.Empty<IPAddress>();
            }

            if (targets.Length == 0)
            {
                AppendDiagLine($"Não foi possível resolver '{host}'.", Color.OrangeRed);
                return;
            }

            foreach (var ip in targets)
            {
                try
                {
                    var entry = await Dns.GetHostEntryAsync(ip.ToString(), _diagCts.Token);
                    var family = ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6 ? "IPv6" : "IPv4";
                    AppendDiagLine($"  {family,-6} {ip,-46} -> {entry.HostName}", Color.LightGreen);
                }
                catch
                {
                    AppendDiagLine($"  {ip,-46} -> (sem registro PTR)", Color.Gray);
                }

                if (_diagCts.Token.IsCancellationRequested) break;
            }
        }
        catch (OperationCanceledException)
        {
            AppendDiagLine("Cancelado.", Color.Yellow);
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed);
        }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private async Task RunArpTableAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Cache ARP (IP \u2194 MAC) ---", Color.Cyan);

        try
        {
            var entries = await Task.Run(GetArpEntries, _diagCts.Token);

            if (entries.Count == 0)
            {
                AppendDiagLine("Cache ARP vazio.", Color.Gray);
                return;
            }

            string? currentIface = null;
            foreach (var entry in entries)
            {
                if (_diagCts.Token.IsCancellationRequested) break;

                // Imprimir cabe\u00e7alho de interface quando mudar
                if (!string.Equals(currentIface, entry.Interface, StringComparison.OrdinalIgnoreCase))
                {
                    currentIface = entry.Interface;
                    AppendDiagLine(string.Empty, Color.White);
                    AppendDiagLine($"\u25ba Interface: {entry.Interface}", Color.FromArgb(140, 180, 255));
                    AppendDiagLine(
                        string.Format("  {0,-20} {1,-20} {2}", "IP", "MAC", "Tipo"),
                        Color.FromArgb(140, 180, 255));
                    AppendDiagLine(new string('-', 55), Color.FromArgb(40, 55, 45));
                }

                var color = entry.Type.Equals("dynamic", StringComparison.OrdinalIgnoreCase)
                    ? Color.FromArgb(180, 230, 180)
                    : entry.Type.Equals("static", StringComparison.OrdinalIgnoreCase)
                        ? Color.FromArgb(255, 220, 140)
                        : Color.Gray;

                AppendDiagLine(
                    string.Format("  {0,-20} {1,-20} {2}", entry.IPAddress, entry.MacAddress, entry.Type),
                    color);
            }

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine($"Total: {entries.Count} entrada(s).", Color.White);
        }
        catch (OperationCanceledException)
        {
            AppendDiagLine("Cancelado.", Color.Yellow);
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed);
        }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private static List<ArpEntry> GetArpEntries()
    {
        var result = new List<ArpEntry>();

        // Executa `arp -a` e parseia a sa\u00edda
        using var process = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = "arp",
            Arguments = "-a",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.OEMCodePage)
        });

        if (process is null) return result;

        string? currentIface = null;
        while (!process.StandardOutput.EndOfStream)
        {
            var line = process.StandardOutput.ReadLine();
            if (line is null) continue;

            // Linha de interface: "Interface: 192.168.1.5 --- 0x5"
            if (line.TrimStart().StartsWith("Interface:", StringComparison.OrdinalIgnoreCase))
            {
                var parts = line.Split(':', 2);
                currentIface = parts.Length > 1 ? parts[1].Trim().Split(' ')[0] : line.Trim();
                continue;
            }

            // Linha de entrada: "  192.168.1.1         aa-bb-cc-dd-ee-ff     dynamic"
            var cols = line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (cols.Length < 2) continue;
            if (!IPAddress.TryParse(cols[0], out _)) continue;

            result.Add(new ArpEntry(
                currentIface ?? "?",
                cols[0],
                cols.Length > 1 ? cols[1] : "?",
                cols.Length > 2 ? cols[2] : "?"));
        }

        process.WaitForExit();
        return result;
    }

    private async Task RunPortScanAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe um host ou IP para o scan.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        var portRangeText = _portRangeTextBox.Text.Trim();
        int[] ports;
        try
        {
            ports = ParsePortRange(string.IsNullOrWhiteSpace(portRangeText) ? "1-1024" : portRangeText);
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Intervalo de portas inválido: {ex.Message}", Color.OrangeRed);
            return;
        }

        if (ports.Length == 0)
        {
            AppendDiagLine("Nenhuma porta no intervalo informado.", Color.Yellow);
            return;
        }

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- Port Scan: {host} ({ports.Length} porta(s)) ---", Color.Cyan);
        AppendDiagLine("Apenas portas abertas serão exibidas.", Color.Gray);

        var openCount = 0;
        var scannedCount = 0;
        const int maxConcurrency = 64;
        const int timeoutMs = 800;

        try
        {
            // Resolve o host uma vez antes de disparar as tasks paralelas
            IPAddress[] resolvedAddresses;
            try
            {
                resolvedAddresses = await Dns.GetHostAddressesAsync(host, _diagCts.Token);
            }
            catch
            {
                AppendDiagLine($"Não foi possível resolver '{host}'.", Color.OrangeRed);
                return;
            }

            var targetIp = resolvedAddresses
                .FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                ?? resolvedAddresses.FirstOrDefault();

            if (targetIp is null)
            {
                AppendDiagLine("Nenhum endereço IP encontrado.", Color.OrangeRed);
                return;
            }

            AppendDiagLine($"Destino resolvido: {targetIp}", Color.White);

            var sem = new SemaphoreSlim(maxConcurrency, maxConcurrency);
            var tasks = ports.Select(async port =>
            {
                await sem.WaitAsync(_diagCts.Token);
                try
                {
                    if (_diagCts.Token.IsCancellationRequested) return;

                    using var cts = CancellationTokenSource.CreateLinkedTokenSource(_diagCts.Token);
                    cts.CancelAfter(timeoutMs);

                    using var tcp = new System.Net.Sockets.TcpClient();
                    try
                    {
                        await tcp.ConnectAsync(targetIp, port, cts.Token);
                        Interlocked.Increment(ref openCount);
                        var service = WellKnownService(port);
                        AppendDiagLine(
                            string.Format("  {0,-6} aberta{1}", port, service is null ? "" : $"  ({service})"),
                            Color.LightGreen);
                    }
                    catch { /* fechada ou filtrada — ignorar */ }
                    finally
                    {
                        Interlocked.Increment(ref scannedCount);
                    }
                }
                finally
                {
                    sem.Release();
                }
            }).ToArray();

            await Task.WhenAll(tasks);

            AppendDiagLine(
                _diagCts.Token.IsCancellationRequested
                    ? $"Scan cancelado. {scannedCount}/{ports.Length} portas verificadas, {openCount} abertas."
                    : $"Scan concluído. {ports.Length} portas verificadas, {openCount} abertas.",
                Color.White);
        }
        catch (OperationCanceledException)
        {
            AppendDiagLine($"Cancelado. {openCount} porta(s) abertas encontradas até o momento.", Color.Yellow);
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed);
        }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private static int[] ParsePortRange(string input)
    {
        var result = new SortedSet<int>();
        foreach (var segment in input.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (segment.Contains('-'))
            {
                var parts = segment.Split('-');
                if (parts.Length != 2 ||
                    !int.TryParse(parts[0].Trim(), out var from) ||
                    !int.TryParse(parts[1].Trim(), out var to) ||
                    from < 1 || to > 65535 || from > to)
                {
                    throw new ArgumentException($"Intervalo inválido: '{segment}'. Use o formato 1-1024.");
                }
                for (var p = from; p <= to; p++) result.Add(p);
            }
            else
            {
                if (!int.TryParse(segment, out var port) || port < 1 || port > 65535)
                    throw new ArgumentException($"Porta inválida: '{segment}'.");
                result.Add(port);
            }
        }
        return result.ToArray();
    }

    private static string? WellKnownService(int port) => port switch
    {
        20 => "FTP-data",
        21 => "FTP",
        22 => "SSH",
        23 => "Telnet",
        25 => "SMTP",
        53 => "DNS",
        67 => "DHCP",
        80 => "HTTP",
        110 => "POP3",
        119 => "NNTP",
        123 => "NTP",
        135 => "RPC",
        139 => "NetBIOS",
        143 => "IMAP",
        161 => "SNMP",
        389 => "LDAP",
        443 => "HTTPS",
        445 => "SMB",
        465 => "SMTPS",
        514 => "Syslog",
        587 => "SMTP-TLS",
        636 => "LDAPS",
        993 => "IMAPS",
        995 => "POP3S",
        1433 => "MSSQL",
        1521 => "Oracle",
        1723 => "PPTP",
        3306 => "MySQL",
        3389 => "RDP",
        5432 => "PostgreSQL",
        5900 => "VNC",
        6379 => "Redis",
        8080 => "HTTP-alt",
        8443 => "HTTPS-alt",
        8888 => "HTTP-alt2",
        27017 => "MongoDB",
        _ => null
    };

    private async Task RunDot1xAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Autenticação 802.1X / Segurança de Rede ---", Color.Cyan);

        try
        {
            // ── LAN (Cabeado) ────────────────────────────────────────────
            AppendDiagLine("► Interfaces LAN (cabeado):", Color.FromArgb(140, 180, 255));

            var lanOutput = await RunCaptureAsync("netsh", "lan show interfaces", _diagCts.Token);
            if (string.IsNullOrWhiteSpace(lanOutput) ||
                lanOutput.Contains("There are no wired", StringComparison.OrdinalIgnoreCase) ||
                lanOutput.Contains("não há", StringComparison.OrdinalIgnoreCase) ||
                lanOutput.Contains("not installed", StringComparison.OrdinalIgnoreCase))
            {
                AppendDiagLine("  Nenhuma interface LAN encontrada ou serviço Wired AutoConfig inativo.", Color.Gray);
            }
            else
            {
                ParseAndDisplayNetshBlock(lanOutput, new[]
                {
                    ("Name", "Nome"),
                    ("Description", "Descrição"),
                    ("State", "Estado"),
                    ("802.1X Status", "802.1X Status"),
                    ("Authentication Mode", "Modo Auth"),
                    ("EAP Type", "EAP Type"),
                });
            }

            if (_diagCts.Token.IsCancellationRequested) return;

            // ── WLAN (Wi-Fi) ─────────────────────────────────────────────
            AppendDiagLine("► Interfaces WLAN (Wi-Fi):", Color.FromArgb(140, 180, 255));

            var wlanOutput = await RunCaptureAsync("netsh", "wlan show interfaces", _diagCts.Token);
            if (string.IsNullOrWhiteSpace(wlanOutput) ||
                wlanOutput.Contains("There is no wireless interface", StringComparison.OrdinalIgnoreCase) ||
                wlanOutput.Contains("não há interface", StringComparison.OrdinalIgnoreCase))
            {
                AppendDiagLine("  Nenhuma interface Wi-Fi encontrada.", Color.Gray);
            }
            else
            {
                ParseAndDisplayNetshBlock(wlanOutput, new[]
                {
                    ("Name", "Nome"),
                    ("Description", "Descrição"),
                    ("SSID", "SSID"),
                    ("BSSID", "BSSID"),
                    ("Network type", "Tipo"),
                    ("Radio type", "Rádio"),
                    ("Authentication", "Autenticação"),
                    ("Cipher", "Cifração"),
                    ("802.1X enabled", "802.1X"),
                    ("State", "Estado"),
                    ("Signal", "Sinal"),
                    ("Receive rate (Mbps)", "Rx (Mbps)"),
                    ("Transmit rate (Mbps)", "Tx (Mbps)"),
                });
            }

            if (_diagCts.Token.IsCancellationRequested) return;

            // ── Perfis LAN 802.1X detalhados ─────────────────────────────
            if (!string.IsNullOrWhiteSpace(lanOutput) && lanOutput.Contains("Enabled", StringComparison.OrdinalIgnoreCase))
            {
                AppendDiagLine("► Perfis 802.1X LAN:", Color.FromArgb(140, 180, 255));
                var lanProfileOutput = await RunCaptureAsync("netsh", "lan show profiles", _diagCts.Token);
                if (!string.IsNullOrWhiteSpace(lanProfileOutput))
                {
                    foreach (var line in lanProfileOutput.Split('\n'))
                    {
                        var trimmed = line.Trim('\r', ' ');
                        if (string.IsNullOrWhiteSpace(trimmed)) continue;
                        AppendDiagLine("  " + trimmed, Color.FromArgb(180, 230, 180));
                    }
                }
            }
        }
        catch (OperationCanceledException)
        {
            AppendDiagLine("Cancelado.", Color.Yellow);
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed);
        }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private void ParseAndDisplayNetshBlock(string output, (string Key, string Label)[] fields)
    {
        // netsh separa blocos de interfaces por linhas em branco duplas
        var blocks = output.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var block in blocks)
        {
            var lines = block.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var found = false;
            foreach (var (key, label) in fields)
            {
                foreach (var line in lines)
                {
                    var colon = line.IndexOf(':');
                    if (colon < 0) continue;
                    var lineKey = line[..colon].Trim();
                    if (!string.Equals(lineKey, key, StringComparison.OrdinalIgnoreCase)) continue;

                    var value = line[(colon + 1)..].Trim();
                    var color = key is "802.1X Status" or "Authentication" or "Authentication Mode" or "802.1X enabled"
                        ? (value.Contains("Enabled", StringComparison.OrdinalIgnoreCase) ||
                           value.Contains("Habilitado", StringComparison.OrdinalIgnoreCase) ||
                           value.Contains("Yes", StringComparison.OrdinalIgnoreCase)
                            ? Color.LightGreen
                            : Color.FromArgb(200, 200, 200))
                        : Color.FromArgb(180, 230, 180);

                    AppendDiagLine($"  {label,-20}: {value}", color);
                    found = true;
                    break;
                }
            }
            if (found) AppendDiagLine(new string('·', 50), Color.FromArgb(40, 55, 45));
        }
    }

    private static async Task<string> RunCaptureAsync(string fileName, string arguments, CancellationToken cancellationToken)
    {
        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.OEMCodePage),
            StandardErrorEncoding  = Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.OEMCodePage),
        };

        using var process = System.Diagnostics.Process.Start(startInfo)
            ?? throw new InvalidOperationException($"Não foi possível iniciar '{fileName}'.");

        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync(cancellationToken);

        var stdout = await stdoutTask;
        var stderr = await stderrTask;
        return string.IsNullOrWhiteSpace(stdout) ? stderr : stdout;
    }

    private async Task RunRouteTableAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Tabela de Rotas IPv4 ---", Color.Cyan);

        try
        {
            var routes = await Task.Run(GetRouteTable, _diagCts.Token);

            if (routes.Count == 0)
            {
                AppendDiagLine("Nenhuma rota encontrada.", Color.Gray);
                return;
            }

            AppendDiagLine(
                string.Format("  {0,-20} {1,-20} {2,-18} {3,-18} {4}", "Destino", "Máscara", "Gateway", "Interface", "Métrica"),
                Color.FromArgb(140, 180, 255));
            AppendDiagLine(new string('-', 90), Color.FromArgb(60, 80, 100));

            foreach (var r in routes)
            {
                if (_diagCts.Token.IsCancellationRequested) break;
                var gw = r.Gateway == "0.0.0.0" ? "on-link" : r.Gateway;
                AppendDiagLine(
                    $"  {r.Destination,-20} {r.Mask,-20} {gw,-18} {r.Interface,-18} {r.Metric}",
                    r.Destination == "0.0.0.0" ? Color.FromArgb(255, 220, 100) : Color.FromArgb(180, 230, 180));
            }

            AppendDiagLine($"Total: {routes.Count} rota(s).", Color.White);
        }
        catch (OperationCanceledException)
        {
            AppendDiagLine("Cancelado.", Color.Yellow);
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed);
        }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private static List<RouteEntry> GetRouteTable()
    {
        var result = new List<RouteEntry>();

        foreach (var iface in NetworkInterface.GetAllNetworkInterfaces())
        {
            if (iface.NetworkInterfaceType == NetworkInterfaceType.Loopback) continue;

            IPInterfaceProperties props;
            IPv4InterfaceProperties? ipv4Props;
            try
            {
                props = iface.GetIPProperties();
                ipv4Props = props.GetIPv4Properties();
            }
            catch { continue; }

            if (ipv4Props is null) continue;

            // Rotas de gateway
            foreach (var gw in props.GatewayAddresses)
            {
                if (gw.Address.AddressFamily != System.Net.Sockets.AddressFamily.InterNetwork) continue;
                result.Add(new RouteEntry("0.0.0.0", "0.0.0.0", gw.Address.ToString(), iface.Name, 0));
            }

            // Sub-redes diretamente conectadas
            foreach (var uni in props.UnicastAddresses)
            {
                if (uni.Address.AddressFamily != System.Net.Sockets.AddressFamily.InterNetwork) continue;

                var mask = PrefixLengthToMask(uni.PrefixLength);
                var network = NetworkAddress(uni.Address, mask);
                result.Add(new RouteEntry(network, mask, "0.0.0.0", iface.Name, 0));
            }
        }

        return result
            .OrderBy(r => r.Destination == "0.0.0.0" ? 0 : 1)
            .ThenBy(r => r.Destination)
            .ToList();
    }

    private static string PrefixLengthToMask(int prefix)
    {
        var mask = prefix == 0 ? 0u : uint.MaxValue << (32 - prefix);
        var bytes = BitConverter.GetBytes((uint)IPAddress.HostToNetworkOrder((int)mask));
        return new IPAddress(bytes).ToString();
    }

    private static string NetworkAddress(IPAddress address, string mask)
    {
        var addrBytes = address.GetAddressBytes();
        var maskBytes = IPAddress.Parse(mask).GetAddressBytes();
        return new IPAddress(new byte[]
        {
            (byte)(addrBytes[0] & maskBytes[0]),
            (byte)(addrBytes[1] & maskBytes[1]),
            (byte)(addrBytes[2] & maskBytes[2]),
            (byte)(addrBytes[3] & maskBytes[3])
        }).ToString();
    }

    private void AppendDiagLine(string text, Color color)
    {
        if (_diagOutputBox.InvokeRequired)
        {
            _diagOutputBox.Invoke(() => AppendDiagLine(text, color));
            return;
        }

        _diagOutputBox.SelectionStart = _diagOutputBox.TextLength;
        _diagOutputBox.SelectionLength = 0;
        _diagOutputBox.SelectionColor = color;
        _diagOutputBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {text}{Environment.NewLine}");
        _diagOutputBox.ScrollToCaret();
    }

    private void SetDiagBusy(bool busy)
    {
        if (_diagOutputBox.InvokeRequired)
        {
            _diagOutputBox.Invoke(() => SetDiagBusy(busy));
            return;
        }

        _pingButton.Enabled = !busy;
        _tracerouteButton.Enabled = !busy;
        _dnsLookupButton.Enabled = !busy;
        _fqdnButton.Enabled = !busy;
        _routeButton.Enabled = !busy;
        _dot1xButton.Enabled = !busy;
        _portScanButton.Enabled = !busy;
        _arpButton.Enabled = !busy;
        _httpButton.Enabled = !busy;
        _sslButton.Enabled = !busy;
        _whoisButton.Enabled = !busy;
        _bandwidthButton.Enabled = !busy;
        _viewLogButton.Enabled = !busy;
        _netstatButton.Enabled = !busy;
        _wifiScanButton.Enabled = !busy;
        _geoIpButton.Enabled = !busy;
        _ntpButton.Enabled = !busy;
        _dnsFlushButton.Enabled = !busy;
        _firewallButton.Enabled = !busy;
        _proxyButton.Enabled = !busy;
        _fwProfileButton.Enabled = !busy;
        _tcpTestButton.Enabled = !busy;
        _smbButton.Enabled = !busy;
        _certButton.Enabled = !busy;
        _mtuButton.Enabled = !busy;
        _ipconfigButton.Enabled = !busy;
        _vpnButton.Enabled = !busy;
        _subnetScanButton.Enabled = !busy;
        _smtpTestButton.Enabled = !busy;
        _ifStatsButton.Enabled = !busy;
        _exportDiagButton.Enabled = !busy;
        _copyOutputButton.Enabled = !busy;
        _cancelDiagButton.Enabled = busy;
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
            _diagCts?.Dispose();
            _connectivityTimer?.Dispose();
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
            SaveWindowSettings();
            return;
        }

        e.Cancel = true;
        HideToTray();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: HTTP Headers
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunHttpHeadersAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe um host ou URL.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();

        var uri = host.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? host : $"https://{host}";
        AppendDiagLine($"--- HTTP Headers: {uri} ---", Color.Cyan);

        try
        {
            using var handler = new System.Net.Http.HttpClientHandler
            {
                AllowAutoRedirect = true,
                ServerCertificateCustomValidationCallback = (_, _, _, _) => true
            };
            using var http = new System.Net.Http.HttpClient(handler) { Timeout = TimeSpan.FromSeconds(10) };
            http.DefaultRequestHeaders.UserAgent.ParseAdd("NetworkConfigurator/1.0");

            using var response = await http.GetAsync(uri, System.Net.Http.HttpCompletionOption.ResponseHeadersRead, _diagCts.Token);

            AppendDiagLine(string.Format("  {0,-30} {1}", "Status:", $"{(int)response.StatusCode} {response.ReasonPhrase}"),
                response.IsSuccessStatusCode ? Color.LightGreen : Color.OrangeRed);
            AppendDiagLine(string.Format("  {0,-30} {1}", "Versão HTTP:", response.Version), Color.White);
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("  Response Headers:", Color.FromArgb(140, 180, 255));

            foreach (var header in response.Headers)
                AppendDiagLine(string.Format("    {0,-40} {1}", header.Key + ":", string.Join(", ", header.Value)), Color.FromArgb(180, 230, 180));

            AppendDiagLine("  Content Headers:", Color.FromArgb(140, 180, 255));
            foreach (var header in response.Content.Headers)
                AppendDiagLine(string.Format("    {0,-40} {1}", header.Key + ":", string.Join(", ", header.Value)), Color.FromArgb(220, 220, 180));
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: SSL/TLS Inspector
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunSslInspectorAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe um host para inspecionar o certificado SSL.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        // Strip scheme if provided
        if (host.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            host = host[8..];
        else if (host.StartsWith("http://", StringComparison.OrdinalIgnoreCase))
            host = host[7..];
        var colonIdx = host.IndexOf(':');
        var port = 443;
        if (colonIdx > 0 && int.TryParse(host[(colonIdx + 1)..], out var p)) { port = p; host = host[..colonIdx]; }

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- SSL/TLS Inspector: {host}:{port} ---", Color.Cyan);

        try
        {
            using var tcp = new System.Net.Sockets.TcpClient();
            await tcp.ConnectAsync(host, port, _diagCts.Token);

            X509Certificate2? cert = null;
            SslProtocols negotiatedProtocol = default;
            string? cipherAlgorithm = null;

            using var sslStream = new SslStream(tcp.GetStream(), false, (_, _, _, _) => true);
            await sslStream.AuthenticateAsClientAsync(new SslClientAuthenticationOptions
            {
                TargetHost = host,
                RemoteCertificateValidationCallback = (_, _, _, _) => true
            }, _diagCts.Token);

            negotiatedProtocol = sslStream.SslProtocol;
            cipherAlgorithm = sslStream.CipherAlgorithm.ToString();

            if (sslStream.RemoteCertificate is X509Certificate2 c2)
                cert = c2;
            else if (sslStream.RemoteCertificate is not null)
                cert = new X509Certificate2(sslStream.RemoteCertificate);

            AppendDiagLine(string.Format("  {0,-28} {1}", "Protocolo:", negotiatedProtocol), Color.LightGreen);
            AppendDiagLine(string.Format("  {0,-28} {1}", "Cipher:", cipherAlgorithm), Color.LightGreen);

            if (cert is not null)
            {
                AppendDiagLine(string.Empty, Color.White);
                AppendDiagLine("  Certificado:", Color.FromArgb(140, 180, 255));
                AppendDiagLine(string.Format("  {0,-28} {1}", "Subject:", cert.Subject), Color.White);
                AppendDiagLine(string.Format("  {0,-28} {1}", "Issuer:", cert.Issuer), Color.White);
                AppendDiagLine(string.Format("  {0,-28} {1}", "Serial:", cert.SerialNumber), Color.White);
                AppendDiagLine(string.Format("  {0,-28} {1}", "Válido de:", cert.NotBefore.ToString("yyyy-MM-dd HH:mm:ss")), Color.White);

                var expired = cert.NotAfter < DateTime.Now;
                var expiringSoon = !expired && cert.NotAfter < DateTime.Now.AddDays(30);
                AppendDiagLine(string.Format("  {0,-28} {1}", "Válido até:",
                    cert.NotAfter.ToString("yyyy-MM-dd HH:mm:ss") + (expired ? " [EXPIRADO]" : expiringSoon ? " [EXPIRA EM BREVE]" : "")),
                    expired ? Color.OrangeRed : expiringSoon ? Color.Yellow : Color.LightGreen);

                AppendDiagLine(string.Format("  {0,-28} {1}", "Thumbprint:", cert.Thumbprint), Color.Gray);
                AppendDiagLine(string.Format("  {0,-28} {1}", "Algoritmo de assinatura:", cert.SignatureAlgorithm.FriendlyName), Color.White);

                // SANs
                foreach (var ext in cert.Extensions)
                {
                    if (ext.Oid?.Value == "2.5.29.17") // Subject Alternative Name
                    {
                        AppendDiagLine(string.Format("  {0,-28} {1}", "SANs:", ext.Format(false)), Color.FromArgb(180, 230, 180));
                    }
                }
            }
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: WHOIS
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunWhoisAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe um domínio ou IP.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- WHOIS: {host} ---", Color.Cyan);

        try
        {
            var result = await QueryWhoisAsync(host, "whois.iana.org", _diagCts.Token);

            // Extract referred whois server and do a second query if available
            var referLine = result.Split('\n').FirstOrDefault(l =>
                l.TrimStart().StartsWith("refer:", StringComparison.OrdinalIgnoreCase) ||
                l.TrimStart().StartsWith("whois:", StringComparison.OrdinalIgnoreCase));

            if (referLine is not null)
            {
                var referred = referLine.Split(':', 2).LastOrDefault()?.Trim();
                if (!string.IsNullOrWhiteSpace(referred) &&
                    !string.Equals(referred, "whois.iana.org", StringComparison.OrdinalIgnoreCase))
                {
                    AppendDiagLine($"  Consultando servidor referenciado: {referred}", Color.Gray);
                    var detailed = await QueryWhoisAsync(host, referred, _diagCts.Token);
                    foreach (var line in detailed.Split('\n'))
                    {
                        var trimmed = line.TrimEnd('\r');
                        if (string.IsNullOrWhiteSpace(trimmed)) continue;
                        if (trimmed.StartsWith('%') || trimmed.StartsWith('#')) continue;
                        AppendDiagLine("  " + trimmed, Color.FromArgb(180, 230, 180));
                    }
                    return;
                }
            }

            foreach (var line in result.Split('\n'))
            {
                var trimmed = line.TrimEnd('\r');
                if (string.IsNullOrWhiteSpace(trimmed)) continue;
                if (trimmed.StartsWith('%') || trimmed.StartsWith('#')) continue;
                AppendDiagLine("  " + trimmed, Color.FromArgb(180, 230, 180));
            }
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private static async Task<string> QueryWhoisAsync(string query, string server, CancellationToken ct)
    {
        using var tcp = new System.Net.Sockets.TcpClient();
        await tcp.ConnectAsync(server, 43, ct);
        using var stream = tcp.GetStream();
        using var writer = new StreamWriter(stream, Encoding.ASCII, leaveOpen: true) { AutoFlush = true };
        using var reader = new StreamReader(stream, Encoding.ASCII, leaveOpen: true);
        await writer.WriteLineAsync(query.AsMemory(), ct);
        return await reader.ReadToEndAsync(ct);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Bandwidth Monitor
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunBandwidthAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Monitor de Largura de Banda (Cancelar para parar) ---", Color.Cyan);
        AppendDiagLine("  Coletando linha de base...", Color.Gray);

        try
        {
            var lastStats = NetworkInterface.GetAllNetworkInterfaces()
                .Where(i => i.OperationalStatus == OperationalStatus.Up &&
                            i.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                .ToDictionary(i => i.Id, i =>
                {
                    var s = i.GetIPStatistics();
                    return (rx: s.BytesReceived, tx: s.BytesSent, name: i.Name);
                });

            await Task.Delay(1000, _diagCts.Token);

            while (!_diagCts.Token.IsCancellationRequested)
            {
                var ifaces = NetworkInterface.GetAllNetworkInterfaces()
                    .Where(i => i.OperationalStatus == OperationalStatus.Up &&
                                i.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                    .ToList();

                var lines = new List<(string text, Color color)>();
                foreach (var iface in ifaces)
                {
                    var stats = iface.GetIPStatistics();
                    if (!lastStats.TryGetValue(iface.Id, out var prev)) continue;

                    var rxDelta = stats.BytesReceived - prev.rx;
                    var txDelta = stats.BytesSent - prev.tx;
                    lastStats[iface.Id] = (stats.BytesReceived, stats.BytesSent, iface.Name);

                    var highlight = rxDelta > 100_000 || txDelta > 100_000;
                    lines.Add((string.Format("  {0,-35} ↓ {1,-14} ↑ {2}",
                        iface.Name.Length > 33 ? iface.Name[..33] + ".." : iface.Name,
                        FormatBytes(rxDelta) + "/s",
                        FormatBytes(txDelta) + "/s"),
                        highlight ? Color.Yellow : Color.FromArgb(180, 230, 180)));
                }

                foreach (var (text, color) in lines)
                    AppendDiagLine(text, color);
                AppendDiagLine(new string('─', 65), Color.FromArgb(40, 55, 45));

                await Task.Delay(1000, _diagCts.Token);
            }
        }
        catch (OperationCanceledException) { AppendDiagLine("Monitor encerrado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private static string FormatBytes(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
        return $"{bytes / (1024.0 * 1024):F2} MB";
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Ver Log
    // ─────────────────────────────────────────────────────────────────────────

    private void ShowLog()
    {
        _diagOutputBox.Clear();
        AppendDiagLine("--- Protocolo de alterações ---", Color.Cyan);

        try
        {
            var logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "network-configurator", "secure-logs");

            if (!Directory.Exists(logDir))
            {
                AppendDiagLine("Nenhum log encontrado.", Color.Gray);
                return;
            }

            var files = Directory.GetFiles(logDir, "*.txt")
                .OrderByDescending(f => f)
                .Take(2)
                .ToArray();

            if (files.Length == 0)
            {
                AppendDiagLine("Nenhum arquivo de log encontrado.", Color.Gray);
                return;
            }

            foreach (var file in files)
            {
                AppendDiagLine($"  Arquivo: {Path.GetFileName(file)}", Color.FromArgb(140, 180, 255));
                var lines = File.ReadAllLines(file);
                foreach (var line in lines.TakeLast(200))
                {
                    var color = line.Contains("[ERROR]") ? Color.OrangeRed
                        : line.Contains("[INFO]") ? Color.LightGreen
                        : Color.FromArgb(180, 230, 180);
                    AppendDiagLine("  " + line, color);
                }
            }
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Erro ao ler log: {ex.Message}", Color.OrangeRed);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Ferramentas: Painéis
    // ─────────────────────────────────────────────────────────────────────────

    private static (TableLayoutPanel outer, RichTextBox output) BuildToolSubpanel(string title, string subtitle)
    {
        var outer = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(0)
        };
        outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        outer.Controls.Add(new Label
        {
            Text = title,
            Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
            ForeColor = Color.FromArgb(24, 39, 52),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 2)
        }, 0, 0);

        outer.Controls.Add(new Label
        {
            Text = subtitle,
            ForeColor = Color.FromArgb(92, 104, 115),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 8)
        }, 0, 1);

        var output = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BackColor = Color.FromArgb(15, 17, 23),
            ForeColor = Color.FromArgb(180, 230, 180),
            Font = new Font("Consolas", 10F),
            BorderStyle = BorderStyle.None,
            ScrollBars = RichTextBoxScrollBars.Vertical
        };
        outer.Controls.Add(output, 0, 2);
        return (outer, output);
    }

    private static void AppendToolLine(RichTextBox box, string text, Color color)
    {
        box.SelectionStart = box.TextLength;
        box.SelectionLength = 0;
        box.SelectionColor = color;
        box.AppendText(text + Environment.NewLine);
        box.ScrollToCaret();
    }

    private Control BuildMacVendorPanel()
    {
        var (outer, outputBox) = BuildToolSubpanel(
            "MAC Vendor Lookup",
            "Identifica o fabricante pelo endereço MAC (OUI). Tenta base local e consulta online.");

        var inputRow = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Margin = new Padding(0, 0, 0, 8)
        };

        _macInputTextBox.Width = 220;
        _macInputTextBox.PlaceholderText = "AA:BB:CC:DD:EE:FF";
        _macInputTextBox.Margin = new Padding(0, 4, 8, 0);

        var lookupBtn = MakePrimaryButton("Consultar");
        var clearBtn = MakeSecondaryButton("Limpar");
        clearBtn.Click += (_, _) => outputBox.Clear();

        lookupBtn.Click += async (_, _) =>
        {
            var mac = _macInputTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(mac)) return;
            outputBox.Clear();
            AppendToolLine(outputBox, $"Consultando: {mac}", Color.Cyan);
            var vendor = await LookupMacVendorAsync(mac);
            AppendToolLine(outputBox, string.Format("  {0,-20} {1}", "Fabricante:", vendor), Color.LightGreen);
            AppendToolLine(outputBox, string.Format("  {0,-20} {1}", "OUI:", mac.Replace("-", "").Replace(":", "").Replace(".", "").ToUpperInvariant().Take(6).Select(c => c.ToString()).Aggregate((a, b) => a + b)), Color.White);
        };

        inputRow.Controls.Add(new Label { Text = "MAC:", AutoSize = true, Margin = new Padding(0, 7, 6, 0) });
        inputRow.Controls.Add(_macInputTextBox);
        inputRow.Controls.Add(lookupBtn);
        inputRow.Controls.Add(clearBtn);

        outer.Controls.Remove(outer.GetControlFromPosition(0, 1));
        outer.RowStyles[1] = new RowStyle(SizeType.AutoSize);
        outer.Controls.Add(new Label
        {
            Text = "Identifica o fabricante pelo endereço MAC (OUI). Tenta base local e consulta online.",
            ForeColor = Color.FromArgb(92, 104, 115),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 8)
        }, 0, 1);

        // Insert inputRow between subtitle and output
        outer.RowCount = 4;
        outer.RowStyles.Insert(2, new RowStyle(SizeType.AutoSize));
        outer.Controls.Add(inputRow, 0, 2);
        outer.Controls.Remove(outputBox);
        outer.RowStyles[3] = new RowStyle(SizeType.Percent, 100F);
        outer.Controls.Add(outputBox, 0, 3);

        return outer;
    }

    private static async Task<string> LookupMacVendorAsync(string mac)
    {
        var normalized = mac.Replace("-", "").Replace(":", "").Replace(".", "").Replace(" ", "").ToUpperInvariant();
        if (normalized.Length < 6) return "MAC inválido";
        var oui = normalized[..6];

        // Local mini-database (top vendors)
        if (LocalOui.TryGetValue(oui, out var localVendor))
            return localVendor;

        // API fallback
        try
        {
            using var http = new System.Net.Http.HttpClient { Timeout = TimeSpan.FromSeconds(5) };
            http.DefaultRequestHeaders.UserAgent.ParseAdd("NetworkConfigurator/1.0");
            var formatted = $"{oui[..2]}:{oui[2..4]}:{oui[4..6]}";
            var response = await http.GetAsync($"https://api.macvendors.com/{formatted}");
            if (response.IsSuccessStatusCode)
                return (await response.Content.ReadAsStringAsync()).Trim();
            if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                return "Fabricante desconhecido";
            return "Fabricante desconhecido";
        }
        catch
        {
            return "Desconhecido (sem internet)";
        }
    }

    private static readonly Dictionary<string, string> LocalOui = new(StringComparer.OrdinalIgnoreCase)
    {
        ["000C29"] = "VMware", ["000569"] = "VMware", ["001C14"] = "VMware", ["005056"] = "VMware",
        ["080027"] = "Oracle VirtualBox", ["0A0027"] = "Oracle VirtualBox", ["525400"] = "QEMU/KVM",
        ["001451"] = "Apple", ["001B63"] = "Apple", ["001E52"] = "Apple", ["001F5B"] = "Apple",
        ["0021E9"] = "Apple", ["002332"] = "Apple", ["ACDE48"] = "Apple", ["A4C3F0"] = "Apple",
        ["686AB8"] = "Apple", ["F0D1A9"] = "Apple", ["3C0754"] = "Apple", ["DC2B2A"] = "Apple",
        ["000393"] = "Apple", ["0017F2"] = "Apple",
        ["001316"] = "Samsung", ["0016DB"] = "Samsung", ["001A8A"] = "Samsung", ["002567"] = "Samsung",
        ["00265D"] = "Samsung", ["001EE1"] = "Samsung", ["A8F274"] = "Samsung",
        ["001B21"] = "Intel", ["001E67"] = "Intel", ["001F3B"] = "Intel", ["002155"] = "Intel",
        ["00249B"] = "Intel", ["0022FA"] = "Intel", ["8086F2"] = "Intel",
        ["000E7F"] = "Cisco", ["001120"] = "Cisco", ["0012D9"] = "Cisco", ["001801"] = "Cisco",
        ["001A30"] = "Cisco", ["001B8F"] = "Cisco", ["001E14"] = "Cisco", ["002155"] = "Intel",
        ["0050C2"] = "Cisco", ["000142"] = "Cisco", ["000143"] = "Cisco",
        ["001350"] = "Dell", ["001A4B"] = "Dell", ["001E4F"] = "Dell", ["001FFF"] = "Dell",
        ["002564"] = "Dell", ["002590"] = "Dell", ["1866DA"] = "Dell", ["D067E5"] = "Dell",
        ["001083"] = "HP", ["001185"] = "HP", ["001321"] = "HP", ["001A4E"] = "HP",
        ["001B78"] = "HP", ["1CC1DE"] = "HP", ["9CB654"] = "HP", ["F0921C"] = "HP",
        ["001C7E"] = "Lenovo", ["001E4F"] = "Dell", ["00212B"] = "Lenovo", ["286ED4"] = "Lenovo",
        ["4CEA"] = "Lenovo",
        ["001DD8"] = "Microsoft", ["00155D"] = "Microsoft", ["3085A9"] = "Microsoft",
        ["7C1E52"] = "Microsoft", ["DC0700"] = "Microsoft",
        ["001CF0"] = "Huawei", ["001E10"] = "Huawei", ["283BDB"] = "Huawei", ["54899A"] = "Huawei",
        ["6C756E"] = "Huawei",
        ["001A11"] = "Google", ["3C5AB4"] = "Google", ["48D705"] = "Google", ["94EB2C"] = "Google",
        ["001422"] = "TP-Link", ["001D0F"] = "TP-Link", ["14CF92"] = "TP-Link", ["5403B0"] = "TP-Link",
        ["001801"] = "Cisco", ["000D97"] = "Comtrend", ["FCEC DA"] = "Netgear",
        ["001018"] = "Broadcom", ["002452"] = "Broadcom",
        ["000C6E"] = "ASUS", ["0016D4"] = "COMPAL", ["001E4C"] = "Ralink",
        ["00265A"] = "Netgear", ["001EE5"] = "Netgear", ["00146C"] = "Netgear", ["206A8A"] = "Netgear",
        ["784476"] = "TP-Link", ["E894F6"] = "TP-Link",
    };

    private Control BuildWakeOnLanPanel()
    {
        var (outer, outputBox) = BuildToolSubpanel(
            "Wake-on-LAN",
            "Envia um magic packet UDP para acordar uma máquina remota na rede local.");

        var grid = new TableLayoutPanel
        {
            AutoSize = true,
            ColumnCount = 2,
            Margin = new Padding(0, 0, 0, 8)
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        _wolMacTextBox.PlaceholderText = "AA:BB:CC:DD:EE:FF";
        _wolMacTextBox.Dock = DockStyle.Fill;
        _wolIpTextBox.PlaceholderText = "255.255.255.255 (broadcast)";
        _wolIpTextBox.Dock = DockStyle.Fill;

        grid.Controls.Add(new Label { Text = "MAC destino:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 7, 8, 0) }, 0, 0);
        grid.Controls.Add(_wolMacTextBox, 1, 0);
        grid.Controls.Add(new Label { Text = "IP destino:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 7, 8, 0) }, 0, 1);
        grid.Controls.Add(_wolIpTextBox, 1, 1);

        var btnRow = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 4, 0, 0) };
        var sendBtn = MakePrimaryButton("Enviar WoL");
        var clearBtn = MakeSecondaryButton("Limpar");
        clearBtn.Click += (_, _) => outputBox.Clear();
        sendBtn.Click += (_, _) => SendWakeOnLan(outputBox);
        btnRow.Controls.Add(sendBtn);
        btnRow.Controls.Add(clearBtn);
        grid.Controls.Add(new Label(), 0, 2);
        grid.Controls.Add(btnRow, 1, 2);

        // Rebuild layout
        outer.RowCount = 4;
        while (outer.RowStyles.Count < 4) outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        outer.RowStyles[2] = new RowStyle(SizeType.AutoSize);
        outer.RowStyles[3] = new RowStyle(SizeType.Percent, 100F);
        outer.Controls.Remove(outputBox);
        outer.Controls.Add(grid, 0, 2);
        outer.Controls.Add(outputBox, 0, 3);
        return outer;
    }

    private void SendWakeOnLan(RichTextBox outputBox)
    {
        var macText = _wolMacTextBox.Text.Trim();
        var normalized = macText.Replace(":", "").Replace("-", "").Replace(".", "").ToUpperInvariant();
        if (normalized.Length != 12)
        {
            AppendToolLine(outputBox, "MAC inválido. Use o formato AA:BB:CC:DD:EE:FF.", Color.OrangeRed);
            return;
        }

        try
        {
            var macBytes = Enumerable.Range(0, 6)
                .Select(i => Convert.ToByte(normalized.Substring(i * 2, 2), 16))
                .ToArray();

            var packet = new byte[17 * 6];
            for (var i = 0; i < 6; i++) packet[i] = 0xFF;
            for (var i = 1; i <= 16; i++)
                Array.Copy(macBytes, 0, packet, i * 6, 6);

            var ip = string.IsNullOrWhiteSpace(_wolIpTextBox.Text)
                ? IPAddress.Broadcast
                : IPAddress.Parse(_wolIpTextBox.Text.Trim());

            using var udp = new System.Net.Sockets.UdpClient();
            udp.EnableBroadcast = true;
            udp.Send(packet, packet.Length, new IPEndPoint(ip, 9));

            AppendToolLine(outputBox, $"Magic packet enviado para {macText} via {ip}:9", Color.LightGreen);
            AppendToolLine(outputBox, "Note: a máquina precisa ter Wake-on-LAN habilitado na BIOS.", Color.Gray);
        }
        catch (Exception ex)
        {
            AppendToolLine(outputBox, $"Erro: {ex.Message}", Color.OrangeRed);
        }
    }

    private Control BuildCidrRangePanel()
    {
        var (outer, outputBox) = BuildToolSubpanel(
            "CIDR Range",
            "Lista todos os endereços IP de um bloco CIDR. Limitado a 65536 para desempenho.");

        var inputRow = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = false,
            Margin = new Padding(0, 0, 0, 8)
        };

        _cidrInputTextBox.Width = 200;
        _cidrInputTextBox.PlaceholderText = "192.168.1.0/24";
        _cidrInputTextBox.Margin = new Padding(0, 4, 8, 0);

        var expandBtn = MakePrimaryButton("Expandir");
        var clearBtn = MakeSecondaryButton("Limpar");
        clearBtn.Click += (_, _) => outputBox.Clear();

        expandBtn.Click += (_, _) =>
        {
            var input = _cidrInputTextBox.Text.Trim();
            outputBox.Clear();
            if (!input.Contains('/'))
            {
                AppendToolLine(outputBox, "Use o formato CIDR: ex. 192.168.1.0/24", Color.OrangeRed);
                return;
            }

            var parts = input.Split('/', 2);
            if (!IPAddress.TryParse(parts[0], out var baseIp) || baseIp.AddressFamily != System.Net.Sockets.AddressFamily.InterNetwork ||
                !int.TryParse(parts[1], out var prefix) || prefix < 0 || prefix > 32)
            {
                AppendToolLine(outputBox, "CIDR inválido.", Color.OrangeRed);
                return;
            }

            var total = (long)Math.Pow(2, 32 - prefix);
            var limit = Math.Min(total, 65536);
            AppendToolLine(outputBox, $"  Rede: {baseIp}/{prefix}  |  Total: {total:N0} endereços{(limit < total ? $"  (exibindo {limit:N0})" : "")}", Color.Cyan);
            AppendToolLine(outputBox, string.Empty, Color.White);

            var addrBytes = baseIp.GetAddressBytes();
            var maskUint = prefix == 0 ? 0u : uint.MaxValue << (32 - prefix);
            var networkUint = (uint)((addrBytes[0] << 24) | (addrBytes[1] << 16) | (addrBytes[2] << 8) | addrBytes[3]) & maskUint;

            for (var i = 0L; i < limit; i++)
            {
                var addrUint = networkUint + (uint)i;
                var b = new byte[] { (byte)(addrUint >> 24), (byte)(addrUint >> 16), (byte)(addrUint >> 8), (byte)addrUint };
                var isNetwork = i == 0;
                var isBroadcast = i == total - 1 && prefix < 31;
                var label = isNetwork ? " (rede)" : isBroadcast ? " (broadcast)" : "";
                AppendToolLine(outputBox, $"  {new IPAddress(b)}{label}",
                    isNetwork || isBroadcast ? Color.FromArgb(255, 220, 140) : Color.FromArgb(180, 230, 180));
            }
        };

        inputRow.Controls.Add(new Label { Text = "CIDR:", AutoSize = true, Margin = new Padding(0, 7, 6, 0) });
        inputRow.Controls.Add(_cidrInputTextBox);
        inputRow.Controls.Add(expandBtn);
        inputRow.Controls.Add(clearBtn);

        outer.RowCount = 4;
        while (outer.RowStyles.Count < 4) outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        outer.RowStyles[2] = new RowStyle(SizeType.AutoSize);
        outer.RowStyles[3] = new RowStyle(SizeType.Percent, 100F);
        outer.Controls.Remove(outputBox);
        outer.Controls.Add(inputRow, 0, 2);
        outer.Controls.Add(outputBox, 0, 3);
        return outer;
    }

    private Control BuildIpConverterPanel()
    {
        var (outer, outputBox) = BuildToolSubpanel(
            "IP Converter",
            "Converte um endereço IPv4 entre decimal, hexadecimal, binário e inteiro de 32 bits.");

        var inputRow = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = false,
            Margin = new Padding(0, 0, 0, 8)
        };

        _ipConvTextBox.Width = 200;
        _ipConvTextBox.PlaceholderText = "192.168.1.1 ou 3232235777";
        _ipConvTextBox.Margin = new Padding(0, 4, 8, 0);

        var convertBtn = MakePrimaryButton("Converter");
        var clearBtn = MakeSecondaryButton("Limpar");
        clearBtn.Click += (_, _) => outputBox.Clear();

        convertBtn.Click += (_, _) =>
        {
            var input = _ipConvTextBox.Text.Trim();
            outputBox.Clear();

            IPAddress? addr = null;
            if (IPAddress.TryParse(input, out var parsed) &&
                parsed.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
            {
                addr = parsed;
            }
            else if (uint.TryParse(input, out var intVal))
            {
                addr = new IPAddress(BitConverter.GetBytes(IPAddress.HostToNetworkOrder((int)intVal)));
            }

            if (addr is null)
            {
                AppendToolLine(outputBox, "Endereço inválido. Informe um IPv4 ou inteiro de 32 bits.", Color.OrangeRed);
                return;
            }

            var bytes = addr.GetAddressBytes();
            var uint32 = (uint)((bytes[0] << 24) | (bytes[1] << 16) | (bytes[2] << 8) | bytes[3]);

            AppendToolLine(outputBox, $"  Endereço IPv4   : {addr}", Color.White);
            AppendToolLine(outputBox, $"  Decimal (uint32): {uint32}", Color.FromArgb(180, 230, 180));
            AppendToolLine(outputBox, $"  Hexadecimal     : 0x{uint32:X8}  ({bytes[0]:X2}.{bytes[1]:X2}.{bytes[2]:X2}.{bytes[3]:X2})", Color.FromArgb(140, 200, 255));
            AppendToolLine(outputBox, $"  Binário         : {Convert.ToString(bytes[0], 2).PadLeft(8, '0')}.{Convert.ToString(bytes[1], 2).PadLeft(8, '0')}.{Convert.ToString(bytes[2], 2).PadLeft(8, '0')}.{Convert.ToString(bytes[3], 2).PadLeft(8, '0')}", Color.FromArgb(255, 220, 140));
            AppendToolLine(outputBox, $"  Octetos         : {bytes[0]} / {bytes[1]} / {bytes[2]} / {bytes[3]}", Color.White);

            // RFC classification
            var cls = bytes[0] < 128 ? "A" : bytes[0] < 192 ? "B" : bytes[0] < 224 ? "C" : bytes[0] < 240 ? "D (Multicast)" : "E (Reservado)";
            var isPrivate = bytes[0] == 10 || (bytes[0] == 172 && bytes[1] >= 16 && bytes[1] <= 31) || (bytes[0] == 192 && bytes[1] == 168);
            var isLoopback = bytes[0] == 127;
            var isLinkLocal = bytes[0] == 169 && bytes[1] == 254;
            var type = isLoopback ? "Loopback" : isLinkLocal ? "Link-local" : isPrivate ? "Privado (RFC 1918)" : "Público";

            AppendToolLine(outputBox, $"  Classe          : {cls}", Color.White);
            AppendToolLine(outputBox, $"  Tipo            : {type}", Color.FromArgb(180, 230, 180));
        };

        inputRow.Controls.Add(new Label { Text = "IP:", AutoSize = true, Margin = new Padding(0, 7, 6, 0) });
        inputRow.Controls.Add(_ipConvTextBox);
        inputRow.Controls.Add(convertBtn);
        inputRow.Controls.Add(clearBtn);

        outer.RowCount = 4;
        while (outer.RowStyles.Count < 4) outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        outer.RowStyles[2] = new RowStyle(SizeType.AutoSize);
        outer.RowStyles[3] = new RowStyle(SizeType.Percent, 100F);
        outer.Controls.Remove(outputBox);
        outer.Controls.Add(inputRow, 0, 2);
        outer.Controls.Add(outputBox, 0, 3);
        return outer;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Perfis de rede
    // ─────────────────────────────────────────────────────────────────────────

    private string ProfilesFilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "NetworkConfigurator", "profiles.json");

    private Control BuildProfilesPanel()
    {
        var outer = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(6)
        };
        outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var card = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(20),
            BorderStyle = BorderStyle.FixedSingle
        };

        var titleLbl = new Label
        {
            Text = "Perfis de Rede",
            Font = new Font("Segoe UI Semibold", 12.5F, FontStyle.Bold),
            ForeColor = Color.FromArgb(24, 39, 52),
            AutoSize = true,
            Location = new Point(20, 20)
        };

        var subLbl = new Label
        {
            Text = "Salve configurações frequentes e carregue-as rapidamente.",
            ForeColor = Color.FromArgb(92, 104, 115),
            AutoSize = true
        };

        // Save section
        var saveTitleLbl = new Label
        {
            Text = "Salvar configuração atual como perfil:",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            AutoSize = true,
            ForeColor = Color.FromArgb(40, 60, 80),
            Margin = new Padding(0, 12, 0, 4)
        };

        var saveRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        _profileNameTextBox.Width = 200;
        _profileNameTextBox.PlaceholderText = "Nome do perfil";
        _profileNameTextBox.Margin = new Padding(0, 2, 8, 0);
        var saveBtn = MakePrimaryButton("Salvar Perfil");
        saveBtn.Click += async (_, _) => await SaveProfileAsync();
        saveRow.Controls.Add(_profileNameTextBox);
        saveRow.Controls.Add(saveBtn);

        // Load section
        var loadTitleLbl = new Label
        {
            Text = "Perfis salvos:",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            AutoSize = true,
            ForeColor = Color.FromArgb(40, 60, 80),
            Margin = new Padding(0, 14, 0, 4)
        };

        var loadRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        _profileComboBox.Width = 200;
        _profileComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
        _profileComboBox.Margin = new Padding(0, 2, 8, 0);

        var loadBtn = MakePrimaryButton("Carregar");
        var deleteBtn = new Button
        {
            Text = "Excluir",
            AutoSize = true,
            Padding = new Padding(10, 5, 10, 5),
            FlatStyle = FlatStyle.Flat,
            FlatAppearance = { BorderColor = Color.FromArgb(200, 60, 60), BorderSize = 1 },
            ForeColor = Color.FromArgb(180, 40, 40),
            BackColor = Color.White,
            Margin = new Padding(4, 2, 0, 0)
        };

        loadBtn.Click += (_, _) => LoadSelectedProfile();
        deleteBtn.Click += async (_, _) => await DeleteSelectedProfileAsync();

        loadRow.Controls.Add(_profileComboBox);
        loadRow.Controls.Add(loadBtn);
        loadRow.Controls.Add(deleteBtn);

        // Layout card content
        var vLayout = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            AutoSize = false,
            WrapContents = false,
            Padding = new Padding(0)
        };
        vLayout.Controls.Add(titleLbl);
        vLayout.Controls.Add(subLbl);
        vLayout.Controls.Add(saveTitleLbl);
        vLayout.Controls.Add(saveRow);
        vLayout.Controls.Add(loadTitleLbl);
        vLayout.Controls.Add(loadRow);

        card.Controls.Add(vLayout);
        outer.Controls.Add(card, 0, 1);
        return outer;
    }

    private async Task LoadProfilesAsync()
    {
        try
        {
            if (!File.Exists(ProfilesFilePath)) return;
            var json = await File.ReadAllTextAsync(ProfilesFilePath);
            _profiles = JsonSerializer.Deserialize<List<NetworkProfile>>(json, AppJson.CompactOptions) ?? new();
            RefreshProfileCombo();
        }
        catch { /* ignore corrupt file */ }
    }

    private void RefreshProfileCombo()
    {
        _profileComboBox.DataSource = null;
        _profileComboBox.DataSource = _profiles.Select(p => p.Name).ToList();
    }

    private async Task SaveProfileAsync()
    {
        var name = _profileNameTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show("Informe um nome para o perfil.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var adapter = _adapterComboBox.SelectedItem as AdapterInfo;
        var profile = new NetworkProfile
        {
            Name = name,
            AdapterName = adapter?.Name ?? string.Empty,
            UseDhcp = _dhcpModeRadioButton.Checked,
            IPAddress = _ipAddressTextBox.Text.Trim(),
            PrefixLength = (int)_prefixLengthUpDown.Value,
            DefaultGateway = _defaultGatewayTextBox.Text.Trim(),
            DnsServers = _dnsServersTextBox.Text.Trim()
        };

        _profiles.RemoveAll(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
        _profiles.Add(profile);

        Directory.CreateDirectory(Path.GetDirectoryName(ProfilesFilePath)!);
        await File.WriteAllTextAsync(ProfilesFilePath, JsonSerializer.Serialize(_profiles, AppJson.CompactOptions));

        RefreshProfileCombo();
        SetStatus($"Perfil '{name}' salvo.");
        _profileNameTextBox.Clear();
    }

    private void LoadSelectedProfile()
    {
        if (_profileComboBox.SelectedItem is not string name) return;
        var profile = _profiles.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
        if (profile is null) return;

        // Select matching adapter
        for (var i = 0; i < _adapterComboBox.Items.Count; i++)
        {
            if (_adapterComboBox.Items[i] is AdapterInfo ai &&
                string.Equals(ai.Name, profile.AdapterName, StringComparison.OrdinalIgnoreCase))
            {
                _adapterComboBox.SelectedIndex = i;
                break;
            }
        }

        if (profile.UseDhcp)
        {
            _dhcpModeRadioButton.Checked = true;
        }
        else
        {
            _staticModeRadioButton.Checked = true;
            _ipAddressTextBox.Text = profile.IPAddress;
            _prefixLengthUpDown.Value = profile.PrefixLength;
            _defaultGatewayTextBox.Text = profile.DefaultGateway;
            _dnsServersTextBox.Text = profile.DnsServers;
        }

        SetStatus($"Perfil '{name}' carregado.");
    }

    private async Task DeleteSelectedProfileAsync()
    {
        if (_profileComboBox.SelectedItem is not string name) return;
        if (MessageBox.Show($"Excluir perfil '{name}'?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        _profiles.RemoveAll(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));

        Directory.CreateDirectory(Path.GetDirectoryName(ProfilesFilePath)!);
        await File.WriteAllTextAsync(ProfilesFilePath, JsonSerializer.Serialize(_profiles, AppJson.CompactOptions));

        RefreshProfileCombo();
        SetStatus($"Perfil '{name}' excluído.");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Helpers de UI
    // ─────────────────────────────────────────────────────────────────────────

    private static Button MakePrimaryButton(string text) => new()
    {
        Text = text,
        AutoSize = true,
        Padding = new Padding(12, 5, 12, 5),
        FlatStyle = FlatStyle.Flat,
        FlatAppearance = { BorderSize = 0 },
        BackColor = Color.FromArgb(15, 114, 121),
        ForeColor = Color.White,
        Margin = new Padding(0, 2, 6, 0)
    };

    private static Button MakeSecondaryButton(string text) => new()
    {
        Text = text,
        AutoSize = true,
        Padding = new Padding(10, 5, 10, 5),
        FlatStyle = FlatStyle.Flat,
        FlatAppearance = { BorderColor = Color.FromArgb(180, 190, 197) },
        BackColor = Color.White,
        Margin = new Padding(0, 2, 6, 0)
    };

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: NetStat
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunNetStatAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Conexões TCP/UDP Ativas (NetStat) ---", Color.Cyan);

        try
        {
            var output = await RunCaptureAsync("netstat", "-ano", _diagCts.Token);
            var lines = output.Split('\n', StringSplitOptions.RemoveEmptyEntries);

            AppendDiagLine(string.Format("  {0,-8} {1,-28} {2,-28} {3,-14} {4}",
                "Proto", "Endereço Local", "Endereço Remoto", "Estado", "PID"), Color.FromArgb(140, 180, 255));
            AppendDiagLine(new string('-', 95), Color.FromArgb(40, 55, 45));

            var count = 0;
            foreach (var line in lines)
            {
                if (_diagCts.Token.IsCancellationRequested) break;
                var cols = line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                if (cols.Length < 4) continue;
                if (!cols[0].StartsWith("TCP", StringComparison.OrdinalIgnoreCase) &&
                    !cols[0].StartsWith("UDP", StringComparison.OrdinalIgnoreCase)) continue;

                var proto = cols[0];
                var local = cols[1];
                var remote = cols.Length > 2 ? cols[2] : "-";
                var stateOrPid = cols.Length > 3 ? cols[3] : "-";
                var pid = cols.Length > 4 ? cols[4] : stateOrPid;
                var state = cols.Length > 4 ? stateOrPid : "-";

                var color = state == "ESTABLISHED" ? Color.LightGreen
                    : state == "LISTEN" || state == "LISTENING" ? Color.FromArgb(140, 200, 255)
                    : state == "TIME_WAIT" ? Color.Yellow
                    : Color.FromArgb(200, 200, 200);

                AppendDiagLine(string.Format("  {0,-8} {1,-28} {2,-28} {3,-14} {4}",
                    proto, local, remote, state, pid), color);
                count++;
            }

            AppendDiagLine($"Total: {count} conexão(ões).", Color.White);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Wi-Fi Scanner
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunWifiScanAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Wi-Fi Scanner (redes disponíveis) ---", Color.Cyan);

        try
        {
            var output = await RunCaptureAsync("netsh", "wlan show networks mode=bssid", _diagCts.Token);

            if (string.IsNullOrWhiteSpace(output) ||
                output.Contains("There is no wireless interface", StringComparison.OrdinalIgnoreCase) ||
                output.Contains("not found", StringComparison.OrdinalIgnoreCase))
            {
                AppendDiagLine("Nenhuma interface Wi-Fi encontrada ou driver não instalado.", Color.Gray);
                return;
            }

            // Parse blocks separated by empty lines
            var blocks = output.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries);
            var networkCount = 0;

            foreach (var block in blocks)
            {
                if (_diagCts.Token.IsCancellationRequested) break;
                var lines = block.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var ssid = "";
                var bssid = "";
                var signal = "";
                var auth = "";
                var cipher = "";
                var channel = "";
                var radioType = "";

                foreach (var line in lines)
                {
                    var colon = line.IndexOf(':');
                    if (colon < 0) continue;
                    var key = line[..colon].Trim();
                    var val = line[(colon + 1)..].Trim();

                    if (key.Equals("SSID", StringComparison.OrdinalIgnoreCase) && !key.Contains("BSSID")) ssid = val;
                    else if (key.StartsWith("BSSID", StringComparison.OrdinalIgnoreCase)) bssid = val;
                    else if (key.Equals("Signal", StringComparison.OrdinalIgnoreCase)) signal = val;
                    else if (key.Equals("Authentication", StringComparison.OrdinalIgnoreCase)) auth = val;
                    else if (key.Equals("Cipher", StringComparison.OrdinalIgnoreCase)) cipher = val;
                    else if (key.Equals("Channel", StringComparison.OrdinalIgnoreCase)) channel = val;
                    else if (key.Equals("Radio type", StringComparison.OrdinalIgnoreCase)) radioType = val;
                }

                if (string.IsNullOrWhiteSpace(ssid) || ssid == "1") continue;

                networkCount++;
                var signalInt = 0;
                if (signal.EndsWith('%')) int.TryParse(signal.TrimEnd('%'), out signalInt);
                var signalColor = signalInt >= 70 ? Color.LightGreen : signalInt >= 40 ? Color.Yellow : Color.OrangeRed;
                var bars = signalInt >= 80 ? "████" : signalInt >= 60 ? "███░" : signalInt >= 40 ? "██░░" : signalInt >= 20 ? "█░░░" : "░░░░";

                AppendDiagLine(string.Empty, Color.White);
                AppendDiagLine($"  ► {ssid}", Color.Cyan);
                AppendDiagLine(string.Format("    {0,-18} {1}", "BSSID:", bssid), Color.White);
                AppendDiagLine(string.Format("    {0,-18} {1} {2}", "Sinal:", bars, signal), signalColor);
                AppendDiagLine(string.Format("    {0,-18} {1}", "Segurança:", auth), Color.FromArgb(180, 230, 180));
                AppendDiagLine(string.Format("    {0,-18} {1}", "Cifração:", cipher), Color.FromArgb(180, 230, 180));
                if (!string.IsNullOrWhiteSpace(channel))
                    AppendDiagLine(string.Format("    {0,-18} {1}", "Canal:", channel), Color.White);
                if (!string.IsNullOrWhiteSpace(radioType))
                    AppendDiagLine(string.Format("    {0,-18} {1}", "Padrão:", radioType), Color.White);
            }

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine($"Total: {networkCount} rede(s) encontrada(s).", Color.White);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Geo IP
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunGeoIpAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe um IP ou hostname.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- Geo IP: {host} ---", Color.Cyan);

        try
        {
            // Resolve to IP if hostname
            if (!IPAddress.TryParse(host, out _))
            {
                var resolved = await Dns.GetHostAddressesAsync(host, _diagCts.Token);
                var ipv4 = resolved.FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    ?? resolved.FirstOrDefault();
                if (ipv4 is null)
                {
                    AppendDiagLine("Não foi possível resolver o hostname.", Color.OrangeRed);
                    return;
                }
                host = ipv4.ToString();
                AppendDiagLine($"  Resolvido para: {host}", Color.Gray);
            }

            using var http = new System.Net.Http.HttpClient { Timeout = TimeSpan.FromSeconds(8) };
            http.DefaultRequestHeaders.UserAgent.ParseAdd("NetworkConfigurator/1.0");
            var response = await http.GetAsync($"https://ipapi.co/{host}/json/", _diagCts.Token);
            var json = await response.Content.ReadAsStringAsync(_diagCts.Token);

            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (root.TryGetProperty("error", out var err) && err.GetBoolean())
            {
                AppendDiagLine($"  Erro da API: {(root.TryGetProperty("reason", out var r) ? r.GetString() : "desconhecido")}", Color.OrangeRed);
                return;
            }

            void PrintField(string label, string propName, Color color)
            {
                if (root.TryGetProperty(propName, out var val) && val.ValueKind != JsonValueKind.Null)
                    AppendDiagLine(string.Format("  {0,-26} {1}", label + ":", val.ToString()), color);
            }

            AppendDiagLine(string.Empty, Color.White);
            PrintField("IP", "ip", Color.White);
            PrintField("Hostname", "hostname", Color.White);
            PrintField("País", "country_name", Color.LightGreen);
            PrintField("Código do país", "country_code", Color.LightGreen);
            PrintField("Região", "region", Color.FromArgb(180, 230, 180));
            PrintField("Cidade", "city", Color.FromArgb(180, 230, 180));
            PrintField("CEP/Postal Code", "postal", Color.White);
            PrintField("Latitude", "latitude", Color.Gray);
            PrintField("Longitude", "longitude", Color.Gray);
            PrintField("Fuso horário", "timezone", Color.White);
            PrintField("ASN", "asn", Color.FromArgb(140, 200, 255));
            PrintField("Organização/ISP", "org", Color.FromArgb(140, 200, 255));
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: NTP Check
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunNtpCheckAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host)) host = "pool.ntp.org";
        AddToHostHistory(host);

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- NTP Check: {host} ---", Color.Cyan);

        try
        {
            var addresses = await Dns.GetHostAddressesAsync(host, _diagCts.Token);
            var server = addresses.FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                ?? addresses.FirstOrDefault();
            if (server is null) { AppendDiagLine("Não foi possível resolver o servidor NTP.", Color.OrangeRed); return; }

            AppendDiagLine($"  Servidor resolvido: {server}", Color.Gray);

            // NTP protocol via UDP port 123
            var ntpData = new byte[48];
            ntpData[0] = 0x1B; // LI=0, VN=3, Mode=3 (client)

            var t1 = DateTime.UtcNow;
            using var udp = new System.Net.Sockets.UdpClient();
            udp.Client.ReceiveTimeout = 5000;
            await udp.SendAsync(ntpData, ntpData.Length, new IPEndPoint(server, 123));

            var receiveTask = udp.ReceiveAsync();
            if (await Task.WhenAny(receiveTask, Task.Delay(5000, _diagCts.Token)) != receiveTask)
            {
                AppendDiagLine("Tempo esgotado aguardando resposta NTP.", Color.OrangeRed);
                return;
            }

            var t4 = DateTime.UtcNow;
            var response = receiveTask.Result.Buffer;

            if (response.Length < 48) { AppendDiagLine("Resposta NTP inválida.", Color.OrangeRed); return; }

            // Extract transmit timestamp (bytes 40-47, big-endian 64-bit fixed-point)
            var intPart = ((ulong)response[40] << 24) | ((ulong)response[41] << 16) |
                          ((ulong)response[42] << 8) | response[43];
            var fracPart = ((ulong)response[44] << 24) | ((ulong)response[45] << 16) |
                           ((ulong)response[46] << 8) | response[47];

            var ntpEpoch = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var serverTime = ntpEpoch.AddSeconds(intPart + fracPart / 4294967296.0);

            var stratum = response[1];
            var offset = (serverTime - t1).TotalMilliseconds / 2.0;
            var rtt = (t4 - t1).TotalMilliseconds;

            AppendDiagLine(string.Format("  {0,-26} {1}", "Hora do servidor (UTC):", serverTime.ToString("yyyy-MM-dd HH:mm:ss.fff")), Color.LightGreen);
            AppendDiagLine(string.Format("  {0,-26} {1}", "Hora local (UTC):", DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff")), Color.White);
            AppendDiagLine(string.Format("  {0,-26} {1:F1} ms", "Offset estimado:", offset),
                Math.Abs(offset) < 500 ? Color.LightGreen : Color.OrangeRed);
            AppendDiagLine(string.Format("  {0,-26} {1:F1} ms", "RTT:", rtt), Color.White);
            AppendDiagLine(string.Format("  {0,-26} {1}", "Stratum:", stratum), Color.FromArgb(180, 230, 180));
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: DNS Flush
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunDnsFlushAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- DNS Flush ---", Color.Cyan);

        try
        {
            var output = await RunCaptureAsync("ipconfig", "/flushdns", _diagCts.Token);
            foreach (var line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = line.TrimEnd('\r');
                if (string.IsNullOrWhiteSpace(trimmed)) continue;
                AppendDiagLine("  " + trimmed, Color.LightGreen);
            }
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Firewall Rules
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunFirewallRulesAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Regras do Windows Firewall (habilitadas) ---", Color.Cyan);

        try
        {
            var output = await RunCaptureAsync("netsh", "advfirewall firewall show rule name=all", _diagCts.Token);

            if (string.IsNullOrWhiteSpace(output))
            {
                AppendDiagLine("Nenhuma regra encontrada ou acesso negado.", Color.Gray);
                return;
            }

            var blocks = output.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries);
            var count = 0;

            foreach (var block in blocks)
            {
                if (_diagCts.Token.IsCancellationRequested) break;

                var lines = block.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var fields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (var line in lines)
                {
                    var colon = line.IndexOf(':');
                    if (colon < 0) continue;
                    fields[line[..colon].Trim()] = line[(colon + 1)..].Trim();
                }

                if (!fields.TryGetValue("Rule Name", out var ruleName) && !fields.TryGetValue("Nome da regra", out ruleName))
                    continue;

                var enabled = fields.GetValueOrDefault("Enabled", fields.GetValueOrDefault("Habilitado", ""));
                if (!enabled.Equals("Yes", StringComparison.OrdinalIgnoreCase) &&
                    !enabled.Equals("Sim", StringComparison.OrdinalIgnoreCase)) continue;

                count++;
                var direction = fields.GetValueOrDefault("Direction", fields.GetValueOrDefault("Direção", "-"));
                var action = fields.GetValueOrDefault("Action", fields.GetValueOrDefault("Ação", "-"));
                var protocol = fields.GetValueOrDefault("Protocol", fields.GetValueOrDefault("Protocolo", "-"));
                var localPort = fields.GetValueOrDefault("LocalPort", fields.GetValueOrDefault("Porta local", "-"));

                var actionColor = action.Equals("Allow", StringComparison.OrdinalIgnoreCase) ||
                                  action.Equals("Permitir", StringComparison.OrdinalIgnoreCase)
                    ? Color.LightGreen : Color.OrangeRed;

                AppendDiagLine(string.Format("  {0,-45} {1,-5} {2,-6} {3,-8} Porta: {4}",
                    ruleName.Length > 43 ? ruleName[..43] + ".." : ruleName,
                    direction, protocol, action, localPort), actionColor);
            }

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine($"Total: {count} regra(s) habilitada(s).", Color.White);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Exportar saída
    // ─────────────────────────────────────────────────────────────────────────

    private void ExportDiagnosticOutput()
    {
        if (string.IsNullOrWhiteSpace(_diagOutputBox.Text))
        {
            AppendDiagLine("Nenhum conteúdo para exportar.", Color.Yellow);
            return;
        }

        try
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var fileName = $"network-diagnostico-{DateTime.Now:yyyyMMdd-HHmmss}.txt";
            var path = Path.Combine(desktop, fileName);
            File.WriteAllText(path, _diagOutputBox.Text, Encoding.UTF8);
            AppendDiagLine($"Exportado para: {path}", Color.LightGreen);
        }
        catch (Exception ex)
        {
            AppendDiagLine($"Erro ao exportar: {ex.Message}", Color.OrangeRed);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Monitor de Conectividade
    // ─────────────────────────────────────────────────────────────────────────

    private void StartConnectivityMonitor()
    {
        _connectivityTimer = new System.Threading.Timer(_ =>
        {
            var isUp = NetworkInterface.GetIsNetworkAvailable();
            if (isUp == _lastConnectivityState) return;
            _lastConnectivityState = isUp;

            if (!IsHandleCreated) return;
            Invoke(() =>
            {
                _notifyIcon.ShowBalloonTip(3000,
                    "Network Configurator",
                    isUp ? "✔ Conectividade restaurada." : "⚠ Sem conectividade de rede.",
                    isUp ? ToolTipIcon.Info : ToolTipIcon.Warning);
                SetStatus(isUp ? "Rede disponível." : "⚠ Sem conectividade de rede.");
            });
        }, null, TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(10));
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Dark Mode
    // ─────────────────────────────────────────────────────────────────────────

    private void ToggleDarkMode()
    {
        _darkMode = !_darkMode;
        _darkModeButton.Text = _darkMode ? "☀ Light" : "🌙 Dark";
        ApplyThemeToControl(this, _darkMode);
    }

    private void ApplyThemeToControl(Control control, bool dark)
    {
        var bg = dark ? Color.FromArgb(24, 27, 35) : Color.FromArgb(242, 245, 247);
        var fg = dark ? Color.FromArgb(220, 222, 230) : Color.FromArgb(24, 39, 52);
        var cardBg = dark ? Color.FromArgb(32, 36, 47) : Color.White;
        var inputBg = dark ? Color.FromArgb(40, 44, 58) : Color.White;
        var inputFg = dark ? Color.FromArgb(220, 222, 230) : Color.FromArgb(24, 39, 52);
        var borderFg = dark ? Color.FromArgb(60, 70, 90) : Color.FromArgb(180, 190, 197);

        void ApplyRecursive(Control c)
        {
            switch (c)
            {
                case Form f:
                    f.BackColor = bg;
                    break;
                case TabControl tc:
                    tc.BackColor = bg;
                    break;
                case TabPage tp:
                    tp.BackColor = dark ? Color.FromArgb(28, 32, 42) : SystemColors.Control;
                    break;
                case TableLayoutPanel tlp:
                    tlp.BackColor = bg;
                    break;
                case FlowLayoutPanel flp:
                    flp.BackColor = bg;
                    break;
                case Panel p:
                    if (p.BackColor == Color.FromArgb(22, 74, 92)) break; // keep header
                    p.BackColor = p.BorderStyle == BorderStyle.FixedSingle ? cardBg : bg;
                    break;
                case TextBox tb:
                    tb.BackColor = inputBg;
                    tb.ForeColor = inputFg;
                    break;
                case ComboBox cb:
                    cb.BackColor = inputBg;
                    cb.ForeColor = inputFg;
                    break;
                case NumericUpDown nud:
                    nud.BackColor = inputBg;
                    nud.ForeColor = inputFg;
                    break;
                case Label lbl:
                    if (lbl.ForeColor == Color.White) break; // keep light text
                    lbl.ForeColor = fg;
                    break;
                case Button btn:
                    if (btn.BackColor == Color.FromArgb(15, 114, 121)) break; // keep primary
                    if (btn.BackColor == Color.FromArgb(22, 74, 92)) break; // keep dark mode btn
                    if (btn.BackColor == Color.FromArgb(180, 40, 40)) break;  // keep cancel
                    btn.BackColor = dark ? Color.FromArgb(44, 50, 65) : Color.White;
                    btn.ForeColor = dark ? Color.FromArgb(200, 210, 230) : Color.FromArgb(40, 50, 60);
                    btn.FlatAppearance.BorderColor = borderFg;
                    break;
                case RichTextBox rtb:
                    if (rtb.BackColor == Color.FromArgb(15, 17, 23)) break; // keep terminal boxes
                    rtb.BackColor = inputBg;
                    rtb.ForeColor = inputFg;
                    break;
            }

            foreach (Control child in c.Controls)
                ApplyRecursive(child);
        }

        ApplyRecursive(control);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Copiar resultado do diagnóstico
    // ─────────────────────────────────────────────────────────────────────────

    private void CopyDiagnosticOutput()
    {
        if (!string.IsNullOrWhiteSpace(_diagOutputBox.Text))
            Clipboard.SetText(_diagOutputBox.Text);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Histórico de comandos — seta ↑/↓ no campo Host
    // ─────────────────────────────────────────────────────────────────────────

    private void DiagHostTextBox_KeyDown(object? sender, KeyEventArgs e)
    {
        if (_hostHistory.Count == 0) return;

        if (e.KeyCode == Keys.Up)
        {
            _hostHistoryCursor = Math.Min(_hostHistoryCursor + 1, _hostHistory.Count - 1);
            _diagHostTextBox.Text = _hostHistory[_hostHistoryCursor];
            _diagHostTextBox.SelectionStart = _diagHostTextBox.Text.Length;
            e.Handled = true;
            e.SuppressKeyPress = true;
        }
        else if (e.KeyCode == Keys.Down)
        {
            _hostHistoryCursor--;
            if (_hostHistoryCursor < 0)
            {
                _hostHistoryCursor = -1;
                _diagHostTextBox.Text = string.Empty;
            }
            else
            {
                _diagHostTextBox.Text = _hostHistory[_hostHistoryCursor];
            }
            _diagHostTextBox.SelectionStart = _diagHostTextBox.Text.Length;
            e.Handled = true;
            e.SuppressKeyPress = true;
        }
    }

    private void AddToHostHistory(string host)
    {
        if (string.IsNullOrWhiteSpace(host)) return;
        _hostHistory.RemoveAll(h => string.Equals(h, host, StringComparison.OrdinalIgnoreCase));
        _hostHistory.Insert(0, host);
        if (_hostHistory.Count > 50) _hostHistory.RemoveRange(50, _hostHistory.Count - 50);
        _hostHistoryCursor = -1;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Persistência de posição e tamanho da janela
    // ─────────────────────────────────────────────────────────────────────────

    private sealed record WindowSettings(int X, int Y, int Width, int Height, string State);

    private void LoadWindowSettings()
    {
        try
        {
            if (!File.Exists(WindowSettingsPath)) return;
            var json = File.ReadAllText(WindowSettingsPath);
            var s = JsonSerializer.Deserialize<WindowSettings>(json, AppJson.CompactOptions);
            if (s is null) return;

            var screen = Screen.FromControl(this).WorkingArea;
            if (s.Width > 400 && s.Height > 300)
            {
                Width = s.Width;
                Height = s.Height;
            }
            if (s.X >= screen.Left && s.X + s.Width <= screen.Right + 100)
                Left = s.X;
            if (s.Y >= screen.Top && s.Y + s.Height <= screen.Bottom + 100)
                Top = s.Y;
            if (s.State == nameof(FormWindowState.Maximized))
                WindowState = FormWindowState.Maximized;
        }
        catch { /* ignorar */ }
    }

    private void SaveWindowSettings()
    {
        try
        {
            // Salvar dimensões da janela restaurada, não do estado maximizado
            var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
            var s = new WindowSettings(bounds.Left, bounds.Top, bounds.Width, bounds.Height,
                WindowState.ToString());
            Directory.CreateDirectory(Path.GetDirectoryName(WindowSettingsPath)!);
            File.WriteAllText(WindowSettingsPath, JsonSerializer.Serialize(s, AppJson.CompactOptions));
        }
        catch { /* ignorar */ }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Proxy
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunProxyDiagAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Configuração de Proxy ---", Color.Cyan);

        try
        {
            await Task.Run(() =>
            {
                // --- Proxy do sistema (Internet Explorer / WinInet) ---
                AppendDiagLine(string.Empty, Color.White);
                AppendDiagLine("● Proxy do Sistema (HKCU Internet Settings):", Color.SteelBlue);
                try
                {
                    using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\Windows\CurrentVersion\Internet Settings");
                    if (key is not null)
                    {
                        var enabled = Convert.ToInt32(key.GetValue("ProxyEnable", 0));
                        var server  = key.GetValue("ProxyServer") as string ?? "(não configurado)";
                        var bypass  = key.GetValue("ProxyOverride") as string ?? "(nenhum)";
                        AppendDiagLine($"  Habilitado : {(enabled == 1 ? "Sim" : "Não")}", enabled == 1 ? Color.LightGreen : Color.Gray);
                        AppendDiagLine($"  Servidor   : {server}", Color.White);
                        AppendDiagLine($"  Bypass     : {bypass}", Color.Gray);
                        var pac = key.GetValue("AutoConfigURL") as string;
                        if (!string.IsNullOrWhiteSpace(pac))
                            AppendDiagLine($"  PAC URL    : {pac}", Color.White);
                    }
                }
                catch (Exception ex) { AppendDiagLine($"  Erro ao ler registro: {ex.Message}", Color.OrangeRed); }

                // --- WinHTTP ---
                AppendDiagLine(string.Empty, Color.White);
                AppendDiagLine("● WinHTTP:", Color.SteelBlue);
            }, _diagCts.Token);

            var winhttp = await RunCaptureAsync("netsh", "winhttp show proxy", _diagCts.Token);
            foreach (var line in winhttp.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                AppendDiagLine("  " + line.TrimEnd(), Color.White);

            await Task.Run(() =>
            {
                // --- Variáveis de ambiente ---
                AppendDiagLine(string.Empty, Color.White);
                AppendDiagLine("● Variáveis de Ambiente:", Color.SteelBlue);
                var vars = new[] { "HTTP_PROXY", "HTTPS_PROXY", "NO_PROXY", "ALL_PROXY",
                                   "http_proxy", "https_proxy", "no_proxy" };
                var found = false;
                foreach (var v in vars)
                {
                    var val = Environment.GetEnvironmentVariable(v);
                    if (!string.IsNullOrWhiteSpace(val))
                    {
                        AppendDiagLine($"  {v} = {val}", Color.LightGreen);
                        found = true;
                    }
                }
                if (!found) AppendDiagLine("  (nenhuma variável proxy definida)", Color.Gray);
            }, _diagCts.Token);

            // --- Teste de conectividade ---
            var host = _diagHostTextBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(host))
            {
                AppendDiagLine(string.Empty, Color.White);
                AppendDiagLine($"● Teste de conectividade via proxy para {host}:", Color.SteelBlue);
                try
                {
                    var webProxy = System.Net.WebRequest.GetSystemWebProxy();
                    webProxy.Credentials = System.Net.CredentialCache.DefaultCredentials;
                    using var handler = new System.Net.Http.HttpClientHandler { Proxy = webProxy, UseProxy = true };
                    using var client  = new System.Net.Http.HttpClient(handler) { Timeout = TimeSpan.FromSeconds(10) };
                    var url = host.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? host : $"http://{host}";
                    var sw  = System.Diagnostics.Stopwatch.StartNew();
                    var resp = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, _diagCts.Token);
                    sw.Stop();
                    AppendDiagLine($"  HTTP {(int)resp.StatusCode} {resp.ReasonPhrase}  ({sw.ElapsedMilliseconds} ms)",
                        resp.IsSuccessStatusCode ? Color.LightGreen : Color.Orange);
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) { AppendDiagLine($"  Falha: {ex.Message}", Color.OrangeRed); }
            }

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("Diagnóstico de Proxy concluído.", Color.LightGreen);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Firewall Profile
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunFirewallProfileAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Status dos Perfis do Firewall ---", Color.Cyan);

        try
        {
            // Perfil atual
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("● Perfil atual:", Color.SteelBlue);
            var current = await RunCaptureAsync("netsh", "advfirewall show currentprofile", _diagCts.Token);
            ParseAndDisplayFwProfile(current);

            if (_diagCts.Token.IsCancellationRequested) return;

            // Todos os perfis
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("● Todos os perfis:", Color.SteelBlue);
            var all = await RunCaptureAsync("netsh", "advfirewall show allprofiles state", _diagCts.Token);
            foreach (var line in all.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var t = line.TrimEnd();
                if (string.IsNullOrWhiteSpace(t)) continue;
                var color = t.Contains("ON",  StringComparison.OrdinalIgnoreCase) ? Color.LightGreen
                          : t.Contains("OFF", StringComparison.OrdinalIgnoreCase) ? Color.OrangeRed
                          : Color.White;
                AppendDiagLine("  " + t, color);
            }

            if (_diagCts.Token.IsCancellationRequested) return;

            // Conexões de entrada/saída padrão
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("● Políticas padrão do perfil ativo:", Color.SteelBlue);
            var policy = await RunCaptureAsync("netsh", "advfirewall show currentprofile firewallpolicy", _diagCts.Token);
            foreach (var line in policy.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var t = line.TrimEnd();
                if (string.IsNullOrWhiteSpace(t)) continue;
                AppendDiagLine("  " + t, Color.White);
            }

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("Verificação de perfis concluída.", Color.LightGreen);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    private void ParseAndDisplayFwProfile(string output)
    {
        foreach (var line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var t = line.TrimEnd();
            if (string.IsNullOrWhiteSpace(t)) continue;
            var isOn   = t.Contains("ON",  StringComparison.OrdinalIgnoreCase) && t.Contains("State");
            var isOff  = t.Contains("OFF", StringComparison.OrdinalIgnoreCase) && t.Contains("State");
            AppendDiagLine("  " + t, isOn ? Color.LightGreen : isOff ? Color.OrangeRed : Color.White);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: TCP Test (conexão raw a host:port)
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunTcpTestAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe um host ou IP.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        var portText = _portRangeTextBox.Text.Trim();
        if (!int.TryParse(portText, out var port) || port < 1 || port > 65535)
        {
            AppendDiagLine("Informe uma porta válida (1-65535) no campo Portas.", Color.Yellow);
            return;
        }

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- TCP Test: {host}:{port} ---", Color.Cyan);

        try
        {
            const int attempts = 4;
            var ok = 0;
            var totalMs = 0L;

            for (var i = 0; i < attempts; i++)
            {
                if (_diagCts.Token.IsCancellationRequested) break;
                var sw = System.Diagnostics.Stopwatch.StartNew();
                try
                {
                    using var cts2 = CancellationTokenSource.CreateLinkedTokenSource(_diagCts.Token);
                    cts2.CancelAfter(3000);
                    using var tcp = new System.Net.Sockets.TcpClient();
                    await tcp.ConnectAsync(host, port, cts2.Token);
                    sw.Stop();
                    ok++;
                    totalMs += sw.ElapsedMilliseconds;
                    AppendDiagLine($"  Tentativa {i + 1}: Conectado em {sw.ElapsedMilliseconds} ms", Color.LightGreen);
                    tcp.Close();
                }
                catch (OperationCanceledException) when (!_diagCts.Token.IsCancellationRequested)
                {
                    sw.Stop();
                    AppendDiagLine($"  Tentativa {i + 1}: Timeout (>3000 ms)", Color.OrangeRed);
                }
                catch (Exception ex)
                {
                    sw.Stop();
                    AppendDiagLine($"  Tentativa {i + 1}: Falhou — {ex.Message}", Color.OrangeRed);
                }

                if (i < attempts - 1)
                    await Task.Delay(500, _diagCts.Token);
            }

            AppendDiagLine(string.Empty, Color.White);
            if (ok > 0)
                AppendDiagLine($"Resultado: {ok}/{attempts} conexões bem-sucedidas. Média: {totalMs / ok} ms.", Color.LightGreen);
            else
                AppendDiagLine($"Resultado: porta {port} FECHADA ou inacessível em {host}.", Color.OrangeRed);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: SMB / Compartilhamentos de rede
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunSmbDiagAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- SMB / Compartilhamentos de Rede ---", Color.Cyan);

        try
        {
            // Compartilhamentos locais
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("● Compartilhamentos locais (net share):", Color.SteelBlue);
            var shares = await RunCaptureAsync("net", "share", _diagCts.Token);
            foreach (var line in shares.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var t = line.TrimEnd();
                if (string.IsNullOrWhiteSpace(t)) continue;
                AppendDiagLine("  " + t, Color.White);
            }

            if (_diagCts.Token.IsCancellationRequested) return;

            // Sessões abertas
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("● Sessões SMB ativas (net session):", Color.SteelBlue);
            try
            {
                var sessions = await RunCaptureAsync("net", "session", _diagCts.Token);
                foreach (var line in sessions.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                {
                    var t = line.TrimEnd();
                    if (string.IsNullOrWhiteSpace(t)) continue;
                    AppendDiagLine("  " + t, Color.White);
                }
            }
            catch { AppendDiagLine("  (privilégio de administrador necessário)", Color.Gray); }

            if (_diagCts.Token.IsCancellationRequested) return;

            // Drives de rede mapeados
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("● Drives de rede mapeados (net use):", Color.SteelBlue);
            var uses = await RunCaptureAsync("net", "use", _diagCts.Token);
            foreach (var line in uses.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var t = line.TrimEnd();
                if (string.IsNullOrWhiteSpace(t)) continue;
                AppendDiagLine("  " + t, Color.White);
            }

            if (_diagCts.Token.IsCancellationRequested) return;

            // Versão do protocolo SMB
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("● Configuração SMB (Get-SmbServerConfiguration):", Color.SteelBlue);
            try
            {
                var smbConf = await RunCaptureAsync("powershell",
                    "-NoProfile -Command \"Get-SmbServerConfiguration | Select EnableSMB1Protocol,EnableSMB2Protocol,RequireSecuritySignature,EncryptData | Format-List\"",
                    _diagCts.Token);
                foreach (var line in smbConf.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                {
                    var t = line.TrimEnd();
                    if (string.IsNullOrWhiteSpace(t)) continue;
                    var isTrue  = t.Contains(": True",  StringComparison.OrdinalIgnoreCase);
                    var isFalse = t.Contains(": False", StringComparison.OrdinalIgnoreCase);
                    // SMB1 habilitado é alerta de segurança
                    var isSMB1 = t.Contains("SMB1", StringComparison.OrdinalIgnoreCase);
                    var color = isSMB1 && isTrue ? Color.OrangeRed
                              : isTrue  ? Color.LightGreen
                              : isFalse ? Color.Gray
                              : Color.White;
                    AppendDiagLine("  " + t, color);
                }
            }
            catch { AppendDiagLine("  (PowerShell indisponível ou sem privilégios)", Color.Gray); }

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("Diagnóstico SMB concluído.", Color.LightGreen);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Certificate Store
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunCertStoreDiagAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Certificados Instalados ---", Color.Cyan);

        try
        {
            await Task.Run(() =>
            {
                var now = DateTime.UtcNow;

                void DumpStore(System.Security.Cryptography.X509Certificates.StoreName storeName,
                               System.Security.Cryptography.X509Certificates.StoreLocation storeLocation,
                               string label)
                {
                    if (_diagCts.Token.IsCancellationRequested) return;
                    AppendDiagLine(string.Empty, Color.White);
                    AppendDiagLine($"● {label}:", Color.SteelBlue);

                    try
                    {
                        using var store = new System.Security.Cryptography.X509Certificates.X509Store(storeName, storeLocation);
                        store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.ReadOnly);

                        var certs = store.Certificates
                            .Cast<System.Security.Cryptography.X509Certificates.X509Certificate2>()
                            .OrderBy(c => c.GetExpirationDateString())
                            .ToList();

                        if (certs.Count == 0)
                        {
                            AppendDiagLine("  (vazio)", Color.Gray);
                            return;
                        }

                        AppendDiagLine(string.Format("  {0,-55} {1,-12} {2}", "Assunto", "Expira em", "Estado"),
                            Color.Gray);
                        AppendDiagLine("  " + new string('-', 85), Color.Gray);

                        foreach (var cert in certs)
                        {
                            if (_diagCts.Token.IsCancellationRequested) break;
                            var expiry = cert.NotAfter;
                            var daysLeft = (expiry - now).TotalDays;
                            var subject = cert.GetNameInfo(
                                System.Security.Cryptography.X509Certificates.X509NameType.SimpleName, false);
                            if (string.IsNullOrWhiteSpace(subject)) subject = cert.Subject;
                            if (subject.Length > 53) subject = subject[..53] + "..";

                            var expiryStr = expiry.ToString("yyyy-MM-dd");
                            string estado;
                            Color color;
                            if (daysLeft < 0)       { estado = "EXPIRADO";      color = Color.OrangeRed; }
                            else if (daysLeft < 30) { estado = $"⚠ {(int)daysLeft}d";  color = Color.Orange; }
                            else if (daysLeft < 90) { estado = $"{(int)daysLeft}d";     color = Color.Yellow; }
                            else                    { estado = "OK";             color = Color.LightGreen; }

                            AppendDiagLine(string.Format("  {0,-55} {1,-12} {2}", subject, expiryStr, estado), color);
                        }

                        AppendDiagLine($"  Total: {certs.Count} certificado(s)", Color.Gray);
                    }
                    catch (Exception ex) { AppendDiagLine($"  Erro: {ex.Message}", Color.OrangeRed); }
                }

                DumpStore(System.Security.Cryptography.X509Certificates.StoreName.My,
                          System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser,
                          "Pessoal (CurrentUser\\My)");

                DumpStore(System.Security.Cryptography.X509Certificates.StoreName.Root,
                          System.Security.Cryptography.X509Certificates.StoreLocation.LocalMachine,
                          "Autoridades Raiz Confiáveis (LocalMachine\\Root)");

                DumpStore(System.Security.Cryptography.X509Certificates.StoreName.CertificateAuthority,
                          System.Security.Cryptography.X509Certificates.StoreLocation.LocalMachine,
                          "Autoridades Intermediárias (LocalMachine\\CA)");

            }, _diagCts.Token);

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("Inspeção de certificados concluída.", Color.LightGreen);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: MTU Discovery
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunMtuDiscoveryAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe um host ou IP.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- MTU Discovery para {host} ---", Color.Cyan);
        AppendDiagLine("Técnica: binary search com ping DF-bit (tamanhos de pacote diminuindo até passar)", Color.Gray);
        AppendDiagLine(string.Empty, Color.White);

        try
        {
            // Resolve destino
            System.Net.IPAddress? target = null;
            try
            {
                var addrs = await Dns.GetHostAddressesAsync(host, _diagCts.Token);
                target = addrs.FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                      ?? addrs.FirstOrDefault();
            }
            catch
            {
                AppendDiagLine("Não foi possível resolver o host.", Color.OrangeRed);
                return;
            }

            AppendDiagLine($"Destino: {target}", Color.Gray);
            AppendDiagLine(string.Empty, Color.White);

            // Binary search entre 512 e 1500
            int lo = 512, hi = 1500, mtu = 0;
            using var ping = new Ping();

            while (lo <= hi)
            {
                if (_diagCts.Token.IsCancellationRequested) break;
                var mid = (lo + hi) / 2;
                var opts = new PingOptions(128, true); // DontFragment = true
                try
                {
                    var buf = new byte[mid];
                    Array.Fill(buf, (byte)0x41);
                    var reply = await ping.SendPingAsync(target!, 2000, buf, opts);
                    if (reply.Status == IPStatus.Success)
                    {
                        mtu = mid;
                        AppendDiagLine($"  {mid,5} bytes → OK ({reply.RoundtripTime} ms)", Color.LightGreen);
                        lo = mid + 1;
                    }
                    else if (reply.Status == IPStatus.PacketTooBig ||
                             reply.Status == IPStatus.TimedOut)
                    {
                        AppendDiagLine($"  {mid,5} bytes → {reply.Status}", Color.Orange);
                        hi = mid - 1;
                    }
                    else
                    {
                        AppendDiagLine($"  {mid,5} bytes → {reply.Status}", Color.Gray);
                        hi = mid - 1;
                    }
                }
                catch { hi = mid - 1; }
            }

            AppendDiagLine(string.Empty, Color.White);
            if (mtu > 0)
            {
                // MTU de caminho = tamanho do payload + 28 bytes (20 IP + 8 ICMP)
                var pathMtu = mtu + 28;
                AppendDiagLine($"MTU de caminho estimado : {pathMtu} bytes (payload {mtu} + header 28)", Color.LightGreen);
                if (pathMtu == 1500)
                    AppendDiagLine("Ethernet standard — sem fragmentação esperada.", Color.Gray);
                else if (pathMtu < 1500)
                    AppendDiagLine("MTU reduzido detectado (provável VPN, PPPoE ou QoS no caminho).", Color.Yellow);
            }
            else
                AppendDiagLine("Não foi possível determinar o MTU (ICMP pode estar bloqueado).", Color.OrangeRed);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: IPConfig Detalhado
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunIpConfigAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- IPConfig Detalhado ---", Color.Cyan);

        try
        {
            var output = await RunCaptureAsync("ipconfig", "/all", _diagCts.Token);
            if (string.IsNullOrWhiteSpace(output))
            {
                AppendDiagLine("Sem saída.", Color.Gray);
                return;
            }

            foreach (var rawLine in output.Split('\n'))
            {
                if (_diagCts.Token.IsCancellationRequested) break;
                var line = rawLine.TrimEnd();
                if (string.IsNullOrWhiteSpace(line)) { AppendDiagLine(string.Empty, Color.White); continue; }

                Color color;
                if (!line.StartsWith(' ') && line.EndsWith(':'))
                    color = Color.SteelBlue;           // cabeçalho de adaptador
                else if (line.Contains("IPv4", StringComparison.OrdinalIgnoreCase) ||
                         line.Contains("IPv6", StringComparison.OrdinalIgnoreCase))
                    color = Color.LightGreen;
                else if (line.Contains("Gateway", StringComparison.OrdinalIgnoreCase))
                    color = Color.Yellow;
                else if (line.Contains("DNS", StringComparison.OrdinalIgnoreCase))
                    color = Color.Cyan;
                else if (line.Contains("DHCP", StringComparison.OrdinalIgnoreCase))
                    color = Color.Orange;
                else if (line.Contains("MAC", StringComparison.OrdinalIgnoreCase) ||
                         line.Contains("Physical", StringComparison.OrdinalIgnoreCase) ||
                         line.Contains("Física", StringComparison.OrdinalIgnoreCase))
                    color = Color.Plum;
                else
                    color = Color.White;

                AppendDiagLine(line, color);
            }
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: VPN Status
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunVpnStatusAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Status de VPN ---", Color.Cyan);

        try
        {
            // Conexões RAS/VPN cadastradas
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("● Conexões VPN configuradas (rasdial /?):", Color.SteelBlue);
            var entries = await RunCaptureAsync("rasdial", string.Empty, _diagCts.Token);
            foreach (var line in entries.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var t = line.TrimEnd();
                if (string.IsNullOrWhiteSpace(t) || t.Contains("O comando foi concluído") ||
                    t.Contains("Command completed")) continue;
                AppendDiagLine("  " + t, Color.White);
            }

            if (_diagCts.Token.IsCancellationRequested) return;

            // Interfaces PPP/VPN no netsh
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("● Interfaces VPN/PPP ativas (netsh interface show interface):", Color.SteelBlue);
            var ifaces = await RunCaptureAsync("netsh", "interface show interface", _diagCts.Token);
            foreach (var line in ifaces.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var t = line.TrimEnd();
                if (string.IsNullOrWhiteSpace(t)) continue;
                var isVpn = t.Contains("VPN", StringComparison.OrdinalIgnoreCase) ||
                            t.Contains("PPP", StringComparison.OrdinalIgnoreCase) ||
                            t.Contains("WAN", StringComparison.OrdinalIgnoreCase) ||
                            t.Contains("Loopback", StringComparison.OrdinalIgnoreCase);
                var isConn = t.Contains("Connected", StringComparison.OrdinalIgnoreCase) ||
                             t.Contains("Conectado", StringComparison.OrdinalIgnoreCase);
                var color = (isVpn && isConn) ? Color.LightGreen
                          : isVpn             ? Color.Orange
                          : Color.Gray;
                AppendDiagLine("  " + t, color);
            }

            if (_diagCts.Token.IsCancellationRequested) return;

            // Via PowerShell: Get-VpnConnection
            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("● Conexões VPN (Get-VpnConnection):", Color.SteelBlue);
            try
            {
                var ps = await RunCaptureAsync("powershell",
                    "-NoProfile -Command \"Get-VpnConnection | Select Name,ServerAddress,ConnectionStatus,TunnelType,UseWinlogonCredential | Format-List\"",
                    _diagCts.Token);
                if (string.IsNullOrWhiteSpace(ps.Trim()))
                    AppendDiagLine("  (nenhuma VPN configurada via Windows)", Color.Gray);
                else
                    foreach (var line in ps.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                    {
                        var t = line.TrimEnd();
                        if (string.IsNullOrWhiteSpace(t)) continue;
                        var isConn = t.Contains("Connected", StringComparison.OrdinalIgnoreCase) ||
                                     t.Contains("Conectado", StringComparison.OrdinalIgnoreCase);
                        AppendDiagLine("  " + t, isConn ? Color.LightGreen : Color.White);
                    }
            }
            catch { AppendDiagLine("  (PowerShell indisponível ou módulo RRAS não instalado)", Color.Gray); }

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("Diagnóstico VPN concluído.", Color.LightGreen);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Subnet Scan (hosts ativos na sub-rede local)
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunSubnetScanAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();

        try
        {
            // Detecta a sub-rede local a partir do adaptador activo
            var gateway = System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                .Where(n => n.OperationalStatus == System.Net.NetworkInformation.OperationalStatus.Up &&
                            n.NetworkInterfaceType != System.Net.NetworkInformation.NetworkInterfaceType.Loopback)
                .SelectMany(n => n.GetIPProperties().UnicastAddresses)
                .FirstOrDefault(a => a.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork &&
                                     !IPAddress.IsLoopback(a.Address));

            string baseIp;
            int prefix;
            if (gateway is not null)
            {
                var addr = gateway.Address.GetAddressBytes();
                var mask = gateway.IPv4Mask.GetAddressBytes();
                var net  = new byte[4];
                for (var i = 0; i < 4; i++) net[i] = (byte)(addr[i] & mask[i]);
                baseIp = $"{net[0]}.{net[1]}.{net[2]}";
                prefix = gateway.PrefixLength;
            }
            else
            {
                AppendDiagLine("Não foi possível detectar a sub-rede local.", Color.Orange);
                return;
            }

            // Limita a /24 para não demorar demais
            var hostField = _diagHostTextBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(hostField) && System.Net.IPAddress.TryParse(hostField, out var parsed))
            {
                var b = parsed.GetAddressBytes();
                baseIp = $"{b[0]}.{b[1]}.{b[2]}";
            }

            AppendDiagLine($"--- Subnet Scan: {baseIp}.0/24 ---", Color.Cyan);
            AppendDiagLine("Enviando ping para 254 hosts em paralelo (timeout 600 ms)...", Color.Gray);

            var found = new System.Collections.Concurrent.ConcurrentBag<(string ip, long ms)>();
            var sem = new System.Threading.SemaphoreSlim(32);
            using var ping = new Ping();

            var tasks = Enumerable.Range(1, 254).Select(async i =>
            {
                await sem.WaitAsync(_diagCts.Token);
                try
                {
                    var ip  = $"{baseIp}.{i}";
                    var buf = new byte[16];
                    using var p = new Ping();
                    var reply = await p.SendPingAsync(ip, 600, buf, new PingOptions(128, true));
                    if (reply.Status == IPStatus.Success)
                        found.Add((ip, reply.RoundtripTime));
                }
                catch { }
                finally { sem.Release(); }
            }).ToList();

            await Task.WhenAll(tasks);

            var sorted = found.OrderBy(x =>
            {
                var parts = x.ip.Split('.');
                return int.TryParse(parts.LastOrDefault(), out var n) ? n : 0;
            }).ToList();

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine($"Hosts ativos encontrados: {sorted.Count}", Color.White);
            AppendDiagLine(string.Format("  {0,-18} {1,-8} {2}", "IP", "RTT (ms)", "Hostname"), Color.Gray);
            AppendDiagLine("  " + new string('-', 50), Color.Gray);

            foreach (var (ip, ms) in sorted)
            {
                if (_diagCts.Token.IsCancellationRequested) break;
                string hostname;
                try
                {
                    var entry = await Dns.GetHostEntryAsync(ip);
                    hostname = entry.HostName;
                }
                catch { hostname = "(sem rDNS)"; }
                AppendDiagLine(string.Format("  {0,-18} {1,-8} {2}", ip, ms, hostname), Color.LightGreen);
            }

            if (sorted.Count == 0)
                AppendDiagLine("  Nenhum host respondeu ao ping (ICMP pode estar bloqueado).", Color.Orange);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: SMTP Test
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunSmtpTestAsync()
    {
        var host = _diagHostTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(host))
        {
            AppendDiagLine("Informe o servidor SMTP no campo Host.", Color.Yellow);
            return;
        }
        AddToHostHistory(host);

        var portText = _portRangeTextBox.Text.Trim();
        var port = int.TryParse(portText, out var p) && p > 0 && p <= 65535 ? p : 25;

        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine($"--- SMTP Test: {host}:{port} ---", Color.Cyan);

        try
        {
            AppendDiagLine($"Conectando a {host}:{port}...", Color.Gray);
            using var cts2 = CancellationTokenSource.CreateLinkedTokenSource(_diagCts.Token);
            cts2.CancelAfter(10_000);

            using var tcp  = new System.Net.Sockets.TcpClient();
            await tcp.ConnectAsync(host, port, cts2.Token);
            AppendDiagLine("TCP conectado.", Color.LightGreen);

            await using var stream = tcp.GetStream();
            using var reader = new System.IO.StreamReader(stream, System.Text.Encoding.ASCII, leaveOpen: true);
            await using var writer = new System.IO.StreamWriter(stream, System.Text.Encoding.ASCII, leaveOpen: true) { AutoFlush = true };

            async Task<string> ReadLineAsync()
            {
                using var lcts = CancellationTokenSource.CreateLinkedTokenSource(_diagCts.Token);
                lcts.CancelAfter(5000);
                return await reader.ReadLineAsync(lcts.Token) ?? string.Empty;
            }

            Color SmtpColor(string l) =>
                l.StartsWith("2") ? Color.LightGreen :
                l.StartsWith("3") ? Color.Yellow :
                l.StartsWith("4") || l.StartsWith("5") ? Color.OrangeRed : Color.White;

            // Banner
            var banner = await ReadLineAsync();
            AppendDiagLine($"← {banner}", SmtpColor(banner));

            // EHLO
            await writer.WriteLineAsync("EHLO networkconfigurator.local");
            AppendDiagLine("→ EHLO networkconfigurator.local", Color.Cyan);
            string line;
            var caps = new List<string>();
            while ((line = await ReadLineAsync()) != string.Empty)
            {
                AppendDiagLine($"← {line}", SmtpColor(line));
                if (line.Length > 4) caps.Add(line[4..].Trim());
                if (line[3] == ' ') break; // última linha do bloco EHLO
            }

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("Recursos anunciados:", Color.SteelBlue);
            foreach (var cap in caps)
            {
                var isSec = cap.StartsWith("STARTTLS", StringComparison.OrdinalIgnoreCase) ||
                            cap.StartsWith("AUTH", StringComparison.OrdinalIgnoreCase);
                AppendDiagLine($"  {cap}", isSec ? Color.LightGreen : Color.White);
            }

            if (!caps.Any(c => c.StartsWith("STARTTLS", StringComparison.OrdinalIgnoreCase)))
                AppendDiagLine("⚠  STARTTLS não anunciado — conexão sem criptografia!", Color.Orange);

            // QUIT
            await writer.WriteLineAsync("QUIT");
            AppendDiagLine("→ QUIT", Color.Cyan);
            var bye = await ReadLineAsync();
            AppendDiagLine($"← {bye}", SmtpColor(bye));

            AppendDiagLine(string.Empty, Color.White);
            AppendDiagLine("Teste SMTP concluído.", Color.LightGreen);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado ou timeout.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally { SetDiagBusy(false); }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Diagnósticos: Interface Statistics
    // ─────────────────────────────────────────────────────────────────────────

    private async Task RunIfStatsAsync()
    {
        SetDiagBusy(true);
        _diagCts = new CancellationTokenSource();
        AppendDiagLine("--- Estatísticas de Interfaces de Rede ---", Color.Cyan);

        try
        {
            await Task.Run(() =>
            {
                var ifaces = System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                    .Where(n => n.NetworkInterfaceType != System.Net.NetworkInformation.NetworkInterfaceType.Loopback)
                    .OrderByDescending(n => n.OperationalStatus)
                    .ToList();

                foreach (var ni in ifaces)
                {
                    if (_diagCts.Token.IsCancellationRequested) break;

                    var stat = ni.GetIPStatistics();
                    var isUp = ni.OperationalStatus == System.Net.NetworkInformation.OperationalStatus.Up;
                    var statusColor = isUp ? Color.LightGreen : Color.Gray;

                    AppendDiagLine(string.Empty, Color.White);
                    AppendDiagLine($"● {ni.Name}  [{ni.NetworkInterfaceType}]  {ni.OperationalStatus}", statusColor);
                    AppendDiagLine($"  Descrição  : {ni.Description}", Color.Gray);
                    AppendDiagLine($"  MAC        : {ni.GetPhysicalAddress()}", Color.Plum);
                    AppendDiagLine($"  Velocidade : {(ni.Speed > 0 ? $"{ni.Speed / 1_000_000} Mbps" : "N/A")}", Color.White);

                    // Endereços IP
                    var ips = ni.GetIPProperties().UnicastAddresses
                        .Select(a => $"{a.Address}/{a.PrefixLength}")
                        .ToList();
                    if (ips.Count > 0)
                        AppendDiagLine($"  Endereços  : {string.Join("  ", ips)}", Color.LightGreen);

                    // Tráfego
                    var rx = stat.BytesReceived;
                    var tx = stat.BytesSent;
                    AppendDiagLine($"  RX         : {FormatBytes(rx)}  ({stat.UnicastPacketsReceived:N0} pkts," +
                                   $" {stat.IncomingPacketsDiscarded:N0} dropped, {stat.IncomingPacketsWithErrors:N0} erros)",
                                   rx > 0 ? Color.Cyan : Color.Gray);
                    AppendDiagLine($"  TX         : {FormatBytes(tx)}  ({stat.UnicastPacketsSent:N0} pkts," +
                                   $" {stat.OutgoingPacketsDiscarded:N0} dropped, {stat.OutgoingPacketsWithErrors:N0} erros)",
                                   tx > 0 ? Color.Yellow : Color.Gray);
                }

                AppendDiagLine(string.Empty, Color.White);
                AppendDiagLine($"Total de interfaces: {ifaces.Count}  |  Ativas: {ifaces.Count(n => n.OperationalStatus == System.Net.NetworkInformation.OperationalStatus.Up)}",
                    Color.White);
            }, _diagCts.Token);
        }
        catch (OperationCanceledException) { AppendDiagLine("Cancelado.", Color.Yellow); }
        catch (Exception ex) { AppendDiagLine($"Erro: {ex.Message}", Color.OrangeRed); }
        finally
        {
            _diagCts?.Dispose();
            _diagCts = null;
            SetDiagBusy(false);
        }
    }

}
