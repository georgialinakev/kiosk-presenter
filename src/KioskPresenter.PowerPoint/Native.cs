using System.Runtime.InteropServices;

namespace KioskPresenter.PowerPoint;

internal static class Native
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("kernel32.dll", EntryPoint = "DeleteFileW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DeleteFile(string path);

    public static int ProcessIdOf(IntPtr hwnd)
    {
        GetWindowThreadProcessId(hwnd, out var pid);
        return (int)pid;
    }
}
