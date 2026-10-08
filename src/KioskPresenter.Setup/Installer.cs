using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using KioskPresenter.Core;
using KioskPresenter.PowerPoint;
using Microsoft.Win32;

namespace KioskPresenter.Setup;

internal sealed record InstallOptions(WindowsUser? AutoStartUser, bool InstallDemos, string? KioskTitle);

/// <summary>
/// Everything Setup does to a PC, and its exact reverse. Every step logs what it did, so a failed
/// install can be read back from C:\ProgramData\KioskPresenter\Logs.
/// </summary>
internal sealed class Installer(Action<string> report)
{
    public const string ProductName = "Kiosk Presenter";
    public const string TaskName = "Kiosk Presenter";
    public const string HostExe = "KioskPresenter.exe";
    public const string SetupExe = "KioskPresenter-Setup.exe";
    private const string ProductKey = @"SOFTWARE\KioskPresenter";
    private const string UninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\KioskPresenter";
    private const string WebView2ClientId = "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}";

    public static string InstallDir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "KioskPresenter");

    public static string ShortcutPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms), ProductName + ".lnk");

    public static string Version => typeof(Installer).Assembly
        .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion?.Split('+')[0] ?? "0.0.0";

    public static bool IsInstalled => File.Exists(Path.Combine(InstallDir, HostExe));

    public static string? InstalledVersion()
    {
        using var key = Registry.LocalMachine.OpenSubKey(ProductKey);
        return key?.GetValue("Version") as string;
    }

    public static bool HasPayload => Assembly.GetExecutingAssembly().GetManifestResourceInfo("payload.zip") is not null;

    public static bool WebView2Installed()
    {
        foreach (var (hive, path) in new[]
        {
            (Registry.LocalMachine, $@"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{WebView2ClientId}"),
            (Registry.LocalMachine, $@"SOFTWARE\Microsoft\EdgeUpdate\Clients\{WebView2ClientId}"),
            (Registry.CurrentUser, $@"SOFTWARE\Microsoft\EdgeUpdate\Clients\{WebView2ClientId}"),
        })
        {
            using var key = hive.OpenSubKey(path);
            if (key?.GetValue("pv") is string version && version.Length > 0 && version != "0.0.0.0") return true;
        }
        return false;
    }

    // ------------------------------------------------------------------ install

    public void Install(InstallOptions options)
    {
        var paths = KioskPaths.Default;
        Log.Init(paths.LogFolder, "setup");
        Step($"Installing {ProductName} {Version}");
        if (!HasPayload) throw new InvalidOperationException("This Setup.exe was built without the app inside it. Use the build from the GitHub release.");

        if (!PowerPointInfo.IsInstalled) Step("WARNING: PowerPoint is not installed. The kiosk will list decks but can't play them until it is.");

        StopRunningKiosk(paths);

        using var payload = OpenPayload();
        ExtractApp(payload);
        CopySelf();
        EnsureWebView2(payload);
        PrepareDataFolder(paths, options);
        if (options.InstallDemos) CopyDemos(payload, paths);
        CreateShortcut();

        if (options.AutoStartUser is { } user) RegisterAutoStart(user);
        else RemoveAutoStart();

        WriteRegistry(options.AutoStartUser);
        Step("Done.");
    }

    private void StopRunningKiosk(KioskPaths paths)
    {
        // The flag stops the watchdog from restarting the host we're about to replace.
        try
        {
            Directory.CreateDirectory(paths.StateFolder);
            File.WriteAllText(paths.AdminExitFlag, "setup " + DateTime.Now.ToString("O"));
        }
        catch (Exception) { /* folder may not exist yet */ }

        var stopped = 0;
        foreach (var p in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(HostExe)))
        {
            using (p)
            {
                try
                {
                    p.Kill(entireProcessTree: true);
                    p.WaitForExit(5000);
                    stopped++;
                }
                catch (Exception ex)
                {
                    Step($"Could not stop process {p.Id}: {ex.Message}");
                }
            }
        }
        if (stopped > 0) Step($"Stopped {stopped} running kiosk process(es).");
    }

    private static ZipArchive OpenPayload()
    {
        var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("payload.zip")
            ?? throw new InvalidOperationException("Payload missing.");
        return new ZipArchive(stream, ZipArchiveMode.Read);
    }

    private void ExtractApp(ZipArchive payload)
    {
        var dir = InstallDir;
        Directory.CreateDirectory(dir);

        // Clear the previous version's files, keeping a running copy of Setup itself.
        foreach (var file in Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories))
        {
            if (string.Equals(file, Environment.ProcessPath, StringComparison.OrdinalIgnoreCase)) continue;
            try { File.Delete(file); }
            catch (Exception ex) { Step($"Could not remove old file {file}: {ex.Message}"); }
        }

        var count = 0;
        var root = Path.GetFullPath(dir) + Path.DirectorySeparatorChar;
        foreach (var entry in payload.Entries)
        {
            var name = entry.FullName.Replace('\\', '/');
            if (!name.StartsWith("app/", StringComparison.Ordinal) || entry.Name.Length == 0) continue;
            var target = Path.GetFullPath(Path.Combine(dir, name["app/".Length..]));
            if (!target.StartsWith(root, StringComparison.OrdinalIgnoreCase)) continue; // zip-slip guard
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            entry.ExtractToFile(target, overwrite: true);
            count++;
        }
        Step($"Copied {count} file(s) to {dir}");
    }

    private void CopySelf()
    {
        var self = Environment.ProcessPath;
        var target = Path.Combine(InstallDir, SetupExe);
        if (self is null || string.Equals(self, target, StringComparison.OrdinalIgnoreCase)) return;
        File.Copy(self, target, overwrite: true);
        Step("Kept a copy of Setup for uninstall.");
    }

    private void EnsureWebView2(ZipArchive payload)
    {
        if (WebView2Installed())
        {
            Step("WebView2 runtime is already installed.");
            return;
        }
        var entry = payload.GetEntry("MicrosoftEdgeWebview2Setup.exe");
        if (entry is null)
        {
            Step("WARNING: WebView2 runtime missing and no installer bundled. Install it from Microsoft before starting the kiosk.");
            return;
        }
        var temp = Path.Combine(Path.GetTempPath(), "MicrosoftEdgeWebview2Setup.exe");
        entry.ExtractToFile(temp, overwrite: true);
        Step("Installing the WebView2 runtime (needs internet)...");
        using var process = Process.Start(new ProcessStartInfo(temp, "/silent /install") { UseShellExecute = false });
        if (process is null || !process.WaitForExit(TimeSpan.FromMinutes(5)))
        {
            Step("WARNING: the WebView2 installer did not finish within 5 minutes.");
            return;
        }
        Step(process.ExitCode == 0 ? "WebView2 runtime installed." : $"WARNING: WebView2 installer exit code {process.ExitCode}.");
    }

    private void PrepareDataFolder(KioskPaths paths, InstallOptions options)
    {
        paths.EnsureCreated();

        // The kiosk runs as a standard user and must write its cache, logs and PIN.
        var info = new DirectoryInfo(paths.DataRoot);
        var security = info.GetAccessControl();
        var users = new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null);
        security.AddAccessRule(new FileSystemAccessRule(
            users,
            FileSystemRights.Modify,
            InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
            PropagationFlags.None,
            AccessControlType.Allow));
        info.SetAccessControl(security);
        Step($"Prepared {paths.DataRoot} (Users can modify).");

        var config = ConfigStore.Load(paths.ConfigFile, out var error);
        if (error is not null) Step($"WARNING: existing config.json unreadable, keeping it untouched: {error}");
        if (error is null)
        {
            if (!string.IsNullOrWhiteSpace(options.KioskTitle)) config.Brand.Title = options.KioskTitle.Trim();
            var existed = File.Exists(paths.ConfigFile);
            ConfigStore.Save(paths.ConfigFile, config);
            Step(existed ? "Updated config.json (existing PIN and content folder kept)." : "Wrote a new config.json.");
        }

        try
        {
            File.Delete(paths.AdminExitFlag);
        }
        catch (Exception) { /* not there */ }
    }

    private void CopyDemos(ZipArchive payload, KioskPaths paths)
    {
        var content = paths.DefaultContentFolder;
        Directory.CreateDirectory(content);
        var copied = 0;
        foreach (var entry in payload.Entries.Where(e => e.FullName.Replace('\\', '/').StartsWith("demo/", StringComparison.Ordinal) && e.Name.Length > 0))
        {
            var target = Path.Combine(content, entry.Name);
            if (File.Exists(target)) continue;
            entry.ExtractToFile(target);
            copied++;
        }
        Step(copied > 0 ? $"Added {copied} demo presentation(s) to {content}" : "Demo presentations already present.");
    }

    private void CreateShortcut()
    {
        try
        {
            var shellType = Type.GetTypeFromProgID("WScript.Shell") ?? throw new InvalidOperationException("WScript.Shell unavailable.");
            dynamic shell = Activator.CreateInstance(shellType)!;
            dynamic link = shell.CreateShortcut(ShortcutPath);
            link.TargetPath = Path.Combine(InstallDir, HostExe);
            link.WorkingDirectory = InstallDir;
            link.Description = "Start the kiosk";
            link.Save();
            Step("Added a Start menu shortcut.");
        }
        catch (Exception ex)
        {
            Step($"WARNING: could not create the Start menu shortcut: {ex.Message}");
        }
    }

    private void RegisterAutoStart(WindowsUser user)
    {
        var exe = SecurityElement.Escape(Path.Combine(InstallDir, HostExe));
        var sid = SecurityElement.Escape(user.Sid);
        var xml = $"""
            <?xml version="1.0" encoding="UTF-16"?>
            <Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
              <RegistrationInfo>
                <Description>Starts Kiosk Presenter when {SecurityElement.Escape(user.Name)} signs in.</Description>
              </RegistrationInfo>
              <Triggers>
                <LogonTrigger>
                  <Enabled>true</Enabled>
                  <UserId>{sid}</UserId>
                  <Delay>PT3S</Delay>
                </LogonTrigger>
              </Triggers>
              <Principals>
                <Principal id="Author">
                  <UserId>{sid}</UserId>
                  <LogonType>InteractiveToken</LogonType>
                  <RunLevel>LeastPrivilege</RunLevel>
                </Principal>
              </Principals>
              <Settings>
                <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
                <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
                <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
                <AllowHardTerminate>true</AllowHardTerminate>
                <StartWhenAvailable>false</StartWhenAvailable>
                <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
                <Priority>4</Priority>
                <Enabled>true</Enabled>
              </Settings>
              <Actions Context="Author">
                <Exec>
                  <Command>"{exe}"</Command>
                </Exec>
              </Actions>
            </Task>
            """;
        var file = Path.Combine(Path.GetTempPath(), "KioskPresenterTask.xml");
        File.WriteAllText(file, xml, Encoding.Unicode);
        try
        {
            var (code, output) = RunSchtasks($"/Create /TN \"{TaskName}\" /XML \"{file}\" /F");
            if (code != 0) throw new InvalidOperationException($"schtasks failed ({code}): {output.Trim()}");
            Step($"The kiosk will start when {user.Name} signs in.");
        }
        finally
        {
            File.Delete(file);
        }
    }

    private void RemoveAutoStart()
    {
        var (code, _) = RunSchtasks($"/Delete /TN \"{TaskName}\" /F");
        if (code == 0) Step("Removed the auto-start task.");
    }

    private static (int Code, string Output) RunSchtasks(string arguments)
    {
        var info = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "schtasks.exe"), arguments)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        using var process = Process.Start(info)!;
        var output = process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
        process.WaitForExit();
        return (process.ExitCode, output);
    }

    private void WriteRegistry(WindowsUser? autoStartUser)
    {
        using (var product = Registry.LocalMachine.CreateSubKey(ProductKey))
        {
            product.SetValue("InstallDir", InstallDir);
            product.SetValue("Version", Version);
            if (autoStartUser is not null)
            {
                product.SetValue("AutoStartSid", autoStartUser.Sid);
                product.SetValue("AutoStartUser", autoStartUser.Name);
            }
            else
            {
                product.DeleteValue("AutoStartSid", throwOnMissingValue: false);
                product.DeleteValue("AutoStartUser", throwOnMissingValue: false);
            }
        }

        var setup = Path.Combine(InstallDir, SetupExe);
        using var uninstall = Registry.LocalMachine.CreateSubKey(UninstallKey);
        uninstall.SetValue("DisplayName", ProductName);
        uninstall.SetValue("DisplayVersion", Version);
        uninstall.SetValue("Publisher", "Georgialina Networks");
        uninstall.SetValue("InstallLocation", InstallDir);
        uninstall.SetValue("DisplayIcon", Path.Combine(InstallDir, HostExe));
        uninstall.SetValue("UninstallString", $"\"{setup}\" /uninstall");
        uninstall.SetValue("QuietUninstallString", $"\"{setup}\" /uninstall /quiet");
        uninstall.SetValue("NoModify", 1, RegistryValueKind.DWord);
        uninstall.SetValue("NoRepair", 1, RegistryValueKind.DWord);
        uninstall.SetValue("EstimatedSize", (int)(DirectorySize(InstallDir) / 1024), RegistryValueKind.DWord);
        Step("Registered in Settings > Apps.");
    }

    // ------------------------------------------------------------------ uninstall

    public void Uninstall(bool removeData)
    {
        var paths = KioskPaths.Default;
        Log.Init(paths.LogFolder, "setup");
        Step($"Uninstalling {ProductName}");

        StopRunningKiosk(paths);
        RemoveAutoStart();
        try
        {
            if (File.Exists(ShortcutPath)) File.Delete(ShortcutPath);
            Step("Removed the Start menu shortcut.");
        }
        catch (Exception ex) { Step($"Could not remove the shortcut: {ex.Message}"); }

        Registry.LocalMachine.DeleteSubKeyTree(UninstallKey, throwOnMissingSubKey: false);
        Registry.LocalMachine.DeleteSubKeyTree(ProductKey, throwOnMissingSubKey: false);
        Step("Removed registry entries.");

        if (removeData)
        {
            try
            {
                Directory.Delete(paths.DataRoot, recursive: true);
                Step($"Removed {paths.DataRoot} (presentations, settings, logs).");
            }
            catch (Exception ex) { Step($"Could not remove {paths.DataRoot}: {ex.Message}"); }
        }
        else
        {
            try { File.Delete(paths.AdminExitFlag); } catch (Exception) { /* fine */ }
            Step($"Kept presentations and settings in {paths.DataRoot}.");
        }

        RemoveInstallDir();
        Step("Done.");
    }

    private void RemoveInstallDir()
    {
        var dir = InstallDir;
        if (!Directory.Exists(dir)) return;
        var runningFromInstallDir = Environment.ProcessPath?.StartsWith(dir, StringComparison.OrdinalIgnoreCase) == true;
        if (!runningFromInstallDir)
        {
            Directory.Delete(dir, recursive: true);
            Step($"Removed {dir}");
            return;
        }
        // Setup can't delete its own exe while running: let cmd finish the job once we've exited.
        Process.Start(new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cmd.exe"),
            $"/c timeout /t 3 /nobreak >nul & rmdir /s /q \"{dir}\"")
        {
            UseShellExecute = false,
            CreateNoWindow = true,
        });
        Step($"{dir} will be removed a few seconds after Setup closes.");
    }

    private static long DirectorySize(string dir)
    {
        try
        {
            return Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories).Sum(f => new FileInfo(f).Length);
        }
        catch (Exception)
        {
            return 0;
        }
    }

    private void Step(string message)
    {
        if (message.StartsWith("WARNING", StringComparison.Ordinal)) Log.Warn(message);
        else Log.Info(message);
        report(message);
    }
}
