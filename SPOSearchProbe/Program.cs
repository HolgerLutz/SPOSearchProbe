using System.Reflection;
using System.Text.Json;

namespace SPOSearchProbe;

/// <summary>
/// Application entry point. Handles command-line argument parsing, mode selection
/// (Admin vs. End-User), single-instance enforcement via named mutexes, config file
/// resolution, and launches the appropriate WinForms UI.
/// </summary>
static class Program
{
    /// <summary>
    /// Returns the display version string, e.g. "v1.26.220.1".
    /// Prefers the <see cref="AssemblyInformationalVersionAttribute"/> set by the build,
    /// then falls back to the assembly's four-part file version, and finally "v0.0.0"
    /// if neither is available (e.g. during development without a proper build).
    /// </summary>
    internal static string AppVersion
    {
        get
        {
            // InformationalVersion can contain semver + build metadata (set in .csproj)
            var v = Assembly.GetExecutingAssembly()
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            if (!string.IsNullOrEmpty(v)) return v;
            // Fallback to the classic Major.Minor.Build.Revision version
            var fv = Assembly.GetExecutingAssembly().GetName().Version;
            return fv != null ? $"v{fv}" : "v0.0.0";
        }
    }

    /// <summary>Author attribution shown in title bars and reports.</summary>
    internal const string AppAuthor = "created by Holger Lutz (holger.lutz@microsoft.com)";

    /// <summary>Combined version + author string used in every window title bar.</summary>
    internal static string AppTitleSuffix => $"{AppVersion} - {AppAuthor}";

    /// <summary>
    /// Checks whether a command-line argument looks like a switch/flag (starts with '-' or '/').
    /// Supports both Unix-style (-flag) and Windows-style (/flag) conventions.
    /// </summary>
    static bool IsSwitch(string arg) =>
        arg.StartsWith('-') || arg.StartsWith('/');

    /// <summary>
    /// Case-insensitive check whether <paramref name="arg"/> matches any of the given
    /// <paramref name="names"/>. Used to recognize the same flag regardless of prefix style
    /// (e.g. "-admin", "--admin", "/admin" all match).
    /// </summary>
    static bool IsFlag(string arg, params string[] names) =>
        names.Any(n => arg.Equals(n, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Application entry point. Must be [STAThread] because WinForms requires
    /// a single-threaded apartment for COM interop (clipboard, file dialogs, etc.).
    /// </summary>
    [STAThread]
    static void Main(string[] args)
    {
        // --- Help ---
        // Show a modal help dialog if any help flag is detected.
        // We check before ApplicationConfiguration.Initialize() so we can show
        // the help even if the app config is broken.
        if (args.Any(a => IsFlag(a, "-help", "--help", "/help", "-?", "/?")))
        {
            var help = $"""
                SPO Search Probe Tool {AppVersion}
                {AppAuthor}

                Usage:  SPOSearchProbe.exe [options] [config-file]

                Options:
                  -admin          Launch in Admin mode (multi-user,
                                  full config editor).
                                  Default is End-User mode (simple single-user GUI).

                  -config <path>  Path to search-config.json.
                                  Default: search-config.json next to the executable.

                  -help           Show this help message.

                Examples:
                  SPOSearchProbe.exe                          End-User mode, default config
                  SPOSearchProbe.exe -admin                   Admin mode, default config
                  SPOSearchProbe.exe -admin -config C:\cfg.json   Admin mode, custom config
                  SPOSearchProbe.exe -config .\my-config.json     End-User mode, custom config
                """;
            MessageBox.Show(help, $"SPO Search Probe - Help ({AppVersion})", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Initialize WinForms (DPI awareness, visual styles, default font, etc.)
        ApplicationConfiguration.Initialize();

        // --- Parse flags ---
        // Determine whether the user requested Admin mode (multi-user config editor)
        // vs. the default End-User mode (single-user search probe GUI).
        bool adminMode = args.Any(a => IsFlag(a, "-admin", "--admin", "/admin"));

        // --- Duplicate instance check ---
        // A named system mutex prevents two instances of the same mode from running
        // simultaneously. This avoids conflicts with shared token cache files and
        // overlapping log file writes.
        //
        // IMPORTANT: We do NOT use 'using' here because the mutex must remain alive
        // for the entire application lifetime. If the GC collected the Mutex object,
        // the OS would release the named mutex and another instance could start.
        // The OS automatically releases it when the process exits.
        var mutexName = adminMode
            ? "SPOSearchProbe_Admin_SingleInstance"
            : "SPOSearchProbe_EndUser_SingleInstance";

        var mutex = new Mutex(true, mutexName, out bool createdNew);
        GC.KeepAlive(mutex); // prevent GC from collecting (and releasing) the mutex

        if (!createdNew)
        {
            // Another instance already holds the mutex — warn the user but allow
            // override. Two instances could corrupt shared state (token files, logs).
            var modeLabel = adminMode ? "Admin" : "End-User";
            var result = MessageBox.Show(
                $"An SPO Search Probe instance is already running in {modeLabel} mode.\n\n" +
                "Running multiple instances of the same mode may cause conflicts\n" +
                "with token cache files and log files.\n\n" +
                "Start a second instance anyway?",
                "SPO Search Probe - Already Running",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result != DialogResult.Yes)
                return;
        }

        // --- Resolve config path ---
        // Scan args for an explicit -config <path> flag. The value follows the flag
        // as the next argument (args[i+1]).
        string? configPath = null;
        for (int i = 0; i < args.Length; i++)
        {
            if (IsFlag(args[i], "-config", "--config", "/config") && i + 1 < args.Length)
            {
                configPath = args[++i]; // consume the next arg as the path
                break;
            }
        }

        // Fallback: look for search-config.json next to the executable.
        // Environment.ProcessPath gives the .exe path; AppContext.BaseDirectory
        // is the fallback when running as a single-file publish or under a debugger.
        var exeDir = Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory;
        configPath ??= Path.Combine(exeDir, "search-config.json");

        if (!File.Exists(configPath))
        {
            if (adminMode)
            {
                // Admin mode auto-creates a default config file so the admin can
                // edit it via the built-in config editor UI.
                var defaultCfg = new SearchConfig();
                var json = JsonSerializer.Serialize(defaultCfg, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(configPath, json);
            }
            else
            {
                // End-User mode requires a pre-existing config — the user can't edit
                // it in this mode, so there's no point launching without one.
                MessageBox.Show($"Configuration file not found:\n{configPath}\n\nPlace search-config.json next to the executable.",
                    "SPO Search Probe", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        // --- Load and parse the JSON config ---
        SearchConfig config;
        try
        {
            var json = File.ReadAllText(configPath);
            config = JsonSerializer.Deserialize<SearchConfig>(json) ?? new SearchConfig();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load config:\n{ex.Message}", "SPO Search Probe", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // --- End-User mode guard rails ---
        // Validate that essential config values are set before launching the UI.
        // The placeholder URL is the default from a freshly created config file.
        if (!adminMode && (string.IsNullOrEmpty(config.SiteUrl) || config.SiteUrl == "https://yourtenant.sharepoint.com"))
        {
            MessageBox.Show("Please configure a valid SiteUrl in search-config.json.",
                "SPO Search Probe", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // ClientId is the Azure AD app registration used for OAuth2 login.
        // Without it, interactive login is impossible.
        if (!adminMode && string.IsNullOrEmpty(config.ClientId))
        {
            MessageBox.Show("Please configure a valid ClientId in search-config.json.\n\n" +
                "Example: 9bc3ab49-b65d-410a-85ad-de819febfddc (PnP Management Shell)",
                "SPO Search Probe", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // --- Launch the appropriate form ---
        // AdminForm: multi-user management, full config editor, advanced features.
        // MainForm: single-user search probe with login, scheduler, and live chart.
        if (adminMode)
            Application.Run(new AdminForm(config, configPath, exeDir));
        else
            Application.Run(new MainForm(config, exeDir));
    }
}