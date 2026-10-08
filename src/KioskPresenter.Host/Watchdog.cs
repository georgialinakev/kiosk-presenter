using System.Diagnostics;
using System.IO;
using KioskPresenter.Core;

namespace KioskPresenter.Host;

/// <summary>
/// The same exe started with --watchdog. Waits for the host to exit and starts it again,
/// unless the exit was an admin exit (flag file) or a sign-out. One per session.
/// </summary>
internal static class Watchdog
{
    public const string MutexName = @"Local\KioskPresenter.Watchdog";
    private static readonly TimeSpan RestartDelay = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan CrashWindow = TimeSpan.FromMinutes(2);
    private const int MaxRestartsInWindow = 5;

    public static int Run(HostOptions options, KioskPaths paths)
    {
        using var mutex = new Mutex(initiallyOwned: true, MutexName, out var createdNew);
        if (!createdNew) return 0;

        var exe = Environment.ProcessPath;
        if (exe is null)
        {
            Log.Error("Watchdog can't find its own exe; exiting.");
            return 1;
        }

        Log.Info($"Watchdog started for host pid {options.WatchPid}.");
        var restarts = new Queue<DateTime>();
        var pid = options.WatchPid;

        while (true)
        {
            using (var host = FindHost(pid))
            {
                if (host is not null)
                {
                    host.WaitForExit();
                    Log.Info($"Host {host.Id} exited with code {SafeExitCode(host)}.");
                }
            }

            if (File.Exists(paths.AdminExitFlag))
            {
                Log.Info("Admin exit or sign-out recorded; watchdog standing down.");
                return 0;
            }

            restarts.Enqueue(DateTime.UtcNow);
            while (restarts.Count > 0 && DateTime.UtcNow - restarts.Peek() > CrashWindow) restarts.Dequeue();
            if (restarts.Count > MaxRestartsInWindow)
            {
                Log.Error($"Host restarted {restarts.Count} times in {CrashWindow.TotalMinutes:0} minutes; backing off 60 s.");
                Thread.Sleep(TimeSpan.FromSeconds(60));
                restarts.Clear();
            }
            else
            {
                Thread.Sleep(RestartDelay);
            }

            if (File.Exists(paths.AdminExitFlag)) return 0;

            try
            {
                var args = options.DataRoot is null ? string.Empty : $"--data \"{options.DataRoot}\"";
                using var started = Process.Start(new ProcessStartInfo(exe, args) { UseShellExecute = false });
                pid = started?.Id ?? 0;
                Log.Warn($"Restarted the host (pid {pid}).");
            }
            catch (Exception ex)
            {
                Log.Error("Watchdog could not restart the host.", ex);
                pid = 0;
                Thread.Sleep(TimeSpan.FromSeconds(10));
            }
        }
    }

    /// <summary>The host by pid, or any other Kiosk Presenter process in this session.</summary>
    private static Process? FindHost(int pid)
    {
        if (pid != 0)
        {
            try
            {
                var p = Process.GetProcessById(pid);
                if (!p.HasExited) return p;
                p.Dispose();
            }
            catch (ArgumentException)
            {
                // Gone already.
            }
        }

        var self = Environment.ProcessId;
        var session = Process.GetCurrentProcess().SessionId;
        var name = Path.GetFileNameWithoutExtension(Environment.ProcessPath) ?? "KioskPresenter";
        foreach (var p in Process.GetProcessesByName(name))
        {
            if (p.Id != self && p.SessionId == session) return p;
            p.Dispose();
        }
        return null;
    }

    private static string SafeExitCode(Process p)
    {
        try { return p.ExitCode.ToString(); }
        catch (Exception) { return "?"; }
    }

    /// <summary>Starts the watchdog for this host if none is running in the session.</summary>
    public static void EnsureRunning(HostOptions options)
    {
        if (Mutex.TryOpenExisting(MutexName, out var existing))
        {
            existing.Dispose();
            return;
        }
        var exe = Environment.ProcessPath;
        if (exe is null) return;
        try
        {
            var data = options.DataRoot is null ? string.Empty : $" --data \"{options.DataRoot}\"";
            using var _ = Process.Start(new ProcessStartInfo(exe, $"--watchdog --pid {Environment.ProcessId}{data}") { UseShellExecute = false });
            Log.Info("Watchdog started.");
        }
        catch (Exception ex)
        {
            Log.Error("Could not start the watchdog.", ex);
        }
    }
}
