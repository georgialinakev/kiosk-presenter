using System.ComponentModel;
using System.Windows;
using System.Windows.Interop;
using KioskPresenter.Core;
using Microsoft.Web.WebView2.Core;

namespace KioskPresenter.Host;

/// <summary>
/// Fullscreen window hosting the Claude Design UI in WebView2. Owns the bridge plumbing only;
/// <see cref="KioskController"/> decides what every message means.
/// </summary>
public partial class MainWindow : Window
{
    public const string AppHost = "app.kiosk.local";
    public const string BrandHost = "brand.kiosk.local";

    private string? _configScriptId;

    public MainWindow()
    {
        InitializeComponent();
        Web.DefaultBackgroundColor = System.Drawing.Color.FromArgb(0x1A, 0x1A, 0x2E);
    }

    /// <summary>A command the UI posted (open, home, pinSubmit, ...).</summary>
    public event Action<UiCommand>? CommandReceived;

    /// <summary>Set by the controller; anything else trying to close the window is refused.</summary>
    public bool AllowClose { get; set; }

    public IntPtr Handle => new WindowInteropHelper(this).EnsureHandle();

    public bool IsReady => Web.CoreWebView2 is not null;

    /// <summary>Dev mode: a normal resizable window that doesn't sit on top.</summary>
    public void UseWindowedMode()
    {
        Topmost = false;
        WindowState = WindowState.Normal;
        WindowStyle = WindowStyle.SingleBorderWindow;
        ResizeMode = ResizeMode.CanResize;
        Width = 1366;
        Height = 768;
    }

    public async Task InitializeWebViewAsync(KioskPaths paths, string uiFolder, string configScript, bool devTools)
    {
        var options = new CoreWebView2EnvironmentOptions("--autoplay-policy=no-user-gesture-required");
        var environment = await CoreWebView2Environment.CreateAsync(null, paths.WebViewDataFolder, options);
        await Web.EnsureCoreWebView2Async(environment);
        var core = Web.CoreWebView2;

        var settings = core.Settings;
        settings.AreDevToolsEnabled = devTools;
        settings.AreDefaultContextMenusEnabled = devTools;
        settings.AreBrowserAcceleratorKeysEnabled = devTools;
        settings.AreDefaultScriptDialogsEnabled = false;
        settings.IsStatusBarEnabled = false;
        settings.IsZoomControlEnabled = false;
        settings.IsPinchZoomEnabled = false;
        settings.IsSwipeNavigationEnabled = false;
        settings.IsGeneralAutofillEnabled = false;
        settings.IsPasswordAutosaveEnabled = false;

        core.SetVirtualHostNameToFolderMapping(AppHost, uiFolder, CoreWebView2HostResourceAccessKind.Allow);
        core.SetVirtualHostNameToFolderMapping(LibraryService.CacheHost, paths.CacheFolder, CoreWebView2HostResourceAccessKind.Allow);
        core.SetVirtualHostNameToFolderMapping(BrandHost, paths.BrandFolder, CoreWebView2HostResourceAccessKind.Allow);

        // The kiosk page is the only page: no popups, no navigating away.
        core.NewWindowRequested += (_, e) => e.Handled = true;
        core.NavigationStarting += (_, e) =>
        {
            if (!e.Uri.StartsWith($"https://{AppHost}/", StringComparison.OrdinalIgnoreCase))
            {
                Log.Warn($"Blocked navigation to {e.Uri}");
                e.Cancel = true;
            }
        };
        core.WebMessageReceived += (_, e) =>
        {
            var command = Bridge.ParseCommand(e.WebMessageAsJson);
            if (command is not null) CommandReceived?.Invoke(command);
        };
        core.ProcessFailed += (_, e) =>
        {
            Log.Error($"WebView2 process failed: {e.ProcessFailedKind} ({e.Reason}). Reloading.");
            Dispatcher.InvokeAsync(ReloadUi);
        };

        _configScriptId = await core.AddScriptToExecuteOnDocumentCreatedAsync(configScript);
        core.Navigate($"https://{AppHost}/index.html");
    }

    /// <summary>Swaps the injected KIOSK_HOST_CONFIG (after the PIN is set, say) for the next page load.</summary>
    public async Task UpdateConfigScriptAsync(string configScript)
    {
        if (Web.CoreWebView2 is null) return;
        if (_configScriptId is not null) Web.CoreWebView2.RemoveScriptToExecuteOnDocumentCreated(_configScriptId);
        _configScriptId = await Web.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(configScript);
    }

    public void ReloadUi()
    {
        try
        {
            Web.CoreWebView2?.Navigate($"https://{AppHost}/index.html");
        }
        catch (Exception ex)
        {
            Log.Error("UI reload failed.", ex);
        }
    }

    /// <summary>Sends a host message to the UI (see <see cref="Bridge"/>).</summary>
    public void Post(object message)
    {
        if (Web.CoreWebView2 is null) return;
        try
        {
            Web.CoreWebView2.PostWebMessageAsJson(Bridge.Serialize(message));
        }
        catch (Exception ex)
        {
            Log.Error("Posting to the UI failed.", ex);
        }
    }

    /// <summary>Puts keyboard focus inside the page so its own key handling works.</summary>
    public void FocusUi()
    {
        Activate();
        Web.Focus();
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (!AllowClose)
        {
            e.Cancel = true;
            return;
        }
        base.OnClosing(e);
    }
}
