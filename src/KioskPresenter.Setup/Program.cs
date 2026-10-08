using System.Windows;
using KioskPresenter.Core;

namespace KioskPresenter.Setup;

/// <summary>
/// KioskPresenter-Setup.exe
///   (no args)                  Wizard: install, upgrade or uninstall.
///   /quiet [/user:NAME] [/demo] [/title:"Hall of Discovery"]
///                              Silent install. NAME is the account the kiosk starts for; omit for no auto-start.
///   /uninstall [/quiet] [/purge]
///                              Remove the app; /purge also deletes presentations, settings and logs.
/// Exit codes: 0 success, 1 failure, 2 bad arguments.
/// </summary>
internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        var quiet = Has(args, "/quiet");
        var uninstall = Has(args, "/uninstall");

        if (quiet)
        {
            var installer = new Installer(Console.WriteLine);
            try
            {
                if (uninstall)
                {
                    installer.Uninstall(removeData: Has(args, "/purge"));
                    return 0;
                }

                WindowsUser? user = null;
                var userArg = Value(args, "/user:");
                if (!string.IsNullOrWhiteSpace(userArg))
                {
                    user = WindowsUsers.Find(userArg);
                    if (user is null)
                    {
                        Log.Error($"Setup: user '{userArg}' not found.");
                        return 2;
                    }
                }
                installer.Install(new InstallOptions(user, Has(args, "/demo"), Value(args, "/title:")));
                return 0;
            }
            catch (Exception ex)
            {
                Log.Error("Setup failed.", ex);
                return 1;
            }
        }

        var app = new Application();
        var window = new SetupWindow(startWithUninstall: uninstall);
        return app.Run(window);
    }

    private static bool Has(string[] args, string flag) =>
        args.Any(a => string.Equals(a, flag, StringComparison.OrdinalIgnoreCase));

    private static string? Value(string[] args, string prefix)
    {
        var arg = args.FirstOrDefault(a => a.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        return arg?[prefix.Length..].Trim('"');
    }
}
