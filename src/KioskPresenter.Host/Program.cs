using System.IO;
using System.Windows;
using System.Windows.Threading;
using KioskPresenter.Core;

namespace KioskPresenter.Host;

internal static class Program
{
    public const string HostMutexName = @"Local\KioskPresenter.Host";

    [STAThread]
    private static int Main(string[] args)
    {
        var options = HostOptions.Parse(args);
        var paths = options.Paths;

        if (options.WatchdogMode)
        {
            Log.Init(paths.LogFolder, "watchdog");
            return Watchdog.Run(options, paths);
        }

        Log.Init(paths.LogFolder, "host");
        using var mutex = new Mutex(initiallyOwned: true, HostMutexName, out var createdNew);
        if (!createdNew)
        {
            Log.Info("Another kiosk instance is already running in this session; exiting.");
            return 0;
        }

        Log.Info($"Starting Kiosk Presenter {AppVersion.Current} as {Environment.UserDomainName}\\{Environment.UserName}" +
                 (options.Windowed ? " (windowed dev mode)" : string.Empty));

        var app = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        app.DispatcherUnhandledException += (_, e) =>
        {
            // Log and exit with an error code: the watchdog restarts a clean instance.
            Log.Error("Unhandled UI exception; exiting for the watchdog to restart.", e.Exception);
            e.Handled = true;
            app.Shutdown(1);
        };
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
            Log.Error("Unhandled exception.", e.ExceptionObject as Exception);
        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            Log.Error("Unobserved task exception.", e.Exception);
            e.SetObserved();
        };

        var window = new MainWindow();
        if (options.Windowed) window.UseWindowedMode();
        var controller = new KioskController(options, paths, window);
        app.SessionEnding += (_, _) => controller.OnSessionEnding();
        app.Exit += (_, _) => controller.Dispose();

        app.Startup += async (_, _) =>
        {
            window.Show();
            try
            {
                await controller.StartAsync();
            }
            catch (Exception ex)
            {
                Log.Error("Kiosk failed to start.", ex);
                app.Shutdown(2);
            }
        };

        var code = app.Run();
        Log.Info($"Kiosk Presenter exited with code {code}.");
        return code;
    }
}

internal static class AppVersion
{
    public static string Current { get; } =
        typeof(AppVersion).Assembly
            .GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
            .OfType<System.Reflection.AssemblyInformationalVersionAttribute>()
            .FirstOrDefault()?.InformationalVersion?.Split('+')[0]
        ?? "0.0.0";
}
