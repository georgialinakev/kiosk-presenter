using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32;

namespace KioskPresenter.Setup;

internal sealed record WindowsUser(string Name, string Sid)
{
    public override string ToString() => Name;
}

/// <summary>Finds the accounts the kiosk could start for, without any extra libraries.</summary>
internal static class WindowsUsers
{
    /// <summary>Every local, domain or Entra ID account that has signed in to this PC (has a profile).</summary>
    public static IReadOnlyList<WindowsUser> WithProfiles()
    {
        var users = new List<WindowsUser>();
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList");
            if (key is null) return users;
            foreach (var sid in key.GetSubKeyNames())
            {
                // S-1-5-21: local and domain accounts. S-1-12-1: Entra ID accounts.
                if (!sid.StartsWith("S-1-5-21-", StringComparison.Ordinal) && !sid.StartsWith("S-1-12-1-", StringComparison.Ordinal)) continue;
                var name = Resolve(sid);
                if (name is null)
                {
                    using var profile = key.OpenSubKey(sid);
                    var folder = profile?.GetValue("ProfileImagePath") as string;
                    name = folder is null ? sid : Path.GetFileName(folder);
                }
                users.Add(new WindowsUser(name, sid));
            }
        }
        catch (Exception)
        {
            // An empty list still lets the admin type a name.
        }
        return users.OrderBy(u => u.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    /// <summary>"PC\kiosk", "kiosk", "DOMAIN\user" or "AzureAD\name" to an account, or null.</summary>
    public static WindowsUser? Find(string nameOrSid)
    {
        var text = nameOrSid.Trim();
        if (text.Length == 0) return null;
        try
        {
            if (text.StartsWith("S-1-", StringComparison.OrdinalIgnoreCase))
            {
                return new WindowsUser(Resolve(text) ?? text, new SecurityIdentifier(text).Value);
            }
            var account = new NTAccount(text);
            var sid = (SecurityIdentifier)account.Translate(typeof(SecurityIdentifier));
            return new WindowsUser(Resolve(sid.Value) ?? text, sid.Value);
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>The user signed in at the physical console, the usual kiosk account candidate.</summary>
    public static string? ConsoleUser()
    {
        var session = WTSGetActiveConsoleSessionId();
        if (session == 0xFFFFFFFF) return null;
        var user = QuerySession(session, WtsUserName);
        if (string.IsNullOrEmpty(user)) return null;
        var domain = QuerySession(session, WtsDomainName);
        return string.IsNullOrEmpty(domain) ? user : $"{domain}\\{user}";
    }

    private static string? Resolve(string sid)
    {
        try
        {
            return new SecurityIdentifier(sid).Translate(typeof(NTAccount)).Value;
        }
        catch (Exception)
        {
            return null;
        }
    }

    private const int WtsUserName = 5;
    private const int WtsDomainName = 7;

    [DllImport("kernel32.dll")]
    private static extern uint WTSGetActiveConsoleSessionId();

    [DllImport("wtsapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool WTSQuerySessionInformation(IntPtr server, uint sessionId, int infoClass, out IntPtr buffer, out uint bytes);

    [DllImport("wtsapi32.dll")]
    private static extern void WTSFreeMemory(IntPtr memory);

    private static string? QuerySession(uint session, int infoClass)
    {
        if (!WTSQuerySessionInformation(IntPtr.Zero, session, infoClass, out var buffer, out _)) return null;
        try
        {
            return Marshal.PtrToStringUni(buffer);
        }
        finally
        {
            WTSFreeMemory(buffer);
        }
    }
}
