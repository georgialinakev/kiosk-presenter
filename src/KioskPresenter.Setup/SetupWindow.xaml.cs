using System.Windows;
using KioskPresenter.Core;
using KioskPresenter.PowerPoint;

namespace KioskPresenter.Setup;

public partial class SetupWindow : Window
{
    private bool _busy;

    public SetupWindow(bool startWithUninstall)
    {
        InitializeComponent();
        Loaded += (_, _) => Prepare(startWithUninstall);
    }

    private void Prepare(bool uninstallMode)
    {
        var installed = Installer.IsInstalled;
        var installedVersion = Installer.InstalledVersion();
        SubHeading.Text = installed
            ? $"Version {installedVersion ?? "?"} is installed. This is version {Installer.Version}."
            : $"Version {Installer.Version}. Installs to {Installer.InstallDir}.";
        InstallButton.Content = installed ? "Upgrade" : "Install";
        UninstallButton.Visibility = installed ? Visibility.Visible : Visibility.Collapsed;
        PurgeCheck.Visibility = installed ? Visibility.Visible : Visibility.Collapsed;

        var ppt = PowerPointInfo.Describe();
        PowerPointStatus.Text = ppt is null
            ? "⚠ PowerPoint is not installed. Install and license it before the kiosk can play presentations."
            : $"✓ {ppt}";
        WebViewStatus.Text = Installer.WebView2Installed()
            ? "✓ WebView2 runtime installed"
            : "• WebView2 runtime will be installed (needs internet)";

        var config = ConfigStore.Load(KioskPaths.Default.ConfigFile, out _);
        TitleBox.Text = config.Brand.Title;

        var users = WindowsUsers.WithProfiles();
        UserCombo.ItemsSource = users;
        var console = WindowsUsers.ConsoleUser();
        var preferred = users.FirstOrDefault(u => string.Equals(u.Name, console, StringComparison.OrdinalIgnoreCase));
        if (preferred is not null) UserCombo.SelectedItem = preferred;
        else if (console is not null) UserCombo.Text = console;

        if (!Installer.HasPayload)
        {
            Append("This Setup.exe has no app inside it (a local build). Use the release build from GitHub to install.");
            InstallButton.IsEnabled = false;
        }

        if (uninstallMode && installed)
        {
            InstallOptionsPanel.Visibility = Visibility.Collapsed;
            InstallButton.Visibility = Visibility.Collapsed;
            Heading.Text = "Uninstall Kiosk Presenter";
        }
    }

    private async void Install_Click(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        WindowsUser? user = null;
        if (AutoStartCheck.IsChecked == true)
        {
            user = UserCombo.SelectedItem as WindowsUser ?? WindowsUsers.Find(UserCombo.Text);
            if (user is null)
            {
                MessageBox.Show(this, $"Windows doesn't recognise the account \"{UserCombo.Text}\". Pick one from the list or type DOMAIN\\name.",
                    "Kiosk Presenter Setup", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
        var options = new InstallOptions(user, DemoCheck.IsChecked == true, TitleBox.Text);
        await RunAsync(installer => installer.Install(options));
    }

    private async void Uninstall_Click(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        var purge = PurgeCheck.IsChecked == true;
        var question = purge
            ? "Remove Kiosk Presenter and delete all presentations, settings and logs?"
            : "Remove Kiosk Presenter? Presentations and settings are kept.";
        if (MessageBox.Show(this, question, "Kiosk Presenter Setup", MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK) return;
        await RunAsync(installer => installer.Uninstall(purge));
    }

    private async Task RunAsync(Action<Installer> work)
    {
        _busy = true;
        InstallButton.IsEnabled = UninstallButton.IsEnabled = false;
        var installer = new Installer(message => Dispatcher.InvokeAsync(() => Append(message)));
        try
        {
            await Task.Run(() => work(installer));
        }
        catch (Exception ex)
        {
            Log.Error("Setup failed.", ex);
            Append($"FAILED: {ex.Message}");
            Append($"Details are in {KioskPaths.Default.LogFolder}");
        }
        finally
        {
            _busy = false;
            CloseButton.Focus();
        }
    }

    private void Append(string line)
    {
        LogBox.AppendText(line + Environment.NewLine);
        LogBox.ScrollToEnd();
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
