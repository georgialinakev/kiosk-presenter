using System.Runtime.InteropServices;
using static KioskPresenter.Host.Interop.NativeMethods;

namespace KioskPresenter.Host.Interop;

/// <summary>
/// Session-wide low-level keyboard hook. Swallows the shortcuts that would take a visitor out of
/// the kiosk (Win, Alt+Tab, Ctrl+Esc, Alt+F4, Task Manager, F1, PowerPoint's Ctrl+T) and raises
/// <see cref="AdminChord"/> for Ctrl+Shift+Q. Ctrl+Alt+Del can't be hooked by design.
/// </summary>
internal sealed class KeyboardHook : IDisposable
{
    private const int VK_TAB = 0x09;
    private const int VK_SHIFT = 0x10;
    private const int VK_CONTROL = 0x11;
    private const int VK_ESCAPE = 0x1B;
    private const int VK_LWIN = 0x5B;
    private const int VK_RWIN = 0x5C;
    private const int VK_APPS = 0x5D;
    private const int VK_F1 = 0x70;
    private const int VK_F4 = 0x73;
    private const int VK_Q = 0x51;
    private const int VK_T = 0x54;

    // Held in a field: if the GC collected the delegate, Windows would call freed memory.
    private readonly LowLevelKeyboardProc _proc;
    private readonly Func<bool> _showActive;
    private IntPtr _hook;

    public KeyboardHook(Func<bool> showActive)
    {
        _showActive = showActive;
        _proc = Callback;
    }

    /// <summary>Ctrl+Shift+Q pressed anywhere.</summary>
    public event Action? AdminChord;

    public bool IsInstalled => _hook != IntPtr.Zero;

    public void Install()
    {
        if (_hook != IntPtr.Zero) return;
        _hook = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(null), 0);
        if (_hook == IntPtr.Zero)
        {
            Core.Log.Error($"Keyboard hook failed (Win32 error {Marshal.GetLastWin32Error()}).");
        }
    }

    public void Uninstall()
    {
        if (_hook == IntPtr.Zero) return;
        UnhookWindowsHookEx(_hook);
        _hook = IntPtr.Zero;
    }

    private IntPtr Callback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            var data = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
            var message = wParam.ToInt32();
            var isDown = message is WM_KEYDOWN or WM_SYSKEYDOWN;
            if (ShouldBlock(data, isDown)) return new IntPtr(1);
        }
        return CallNextHookEx(_hook, nCode, wParam, lParam);
    }

    private bool ShouldBlock(KBDLLHOOKSTRUCT key, bool isDown)
    {
        var vk = (int)key.vkCode;
        var alt = (key.flags & LLKHF_ALTDOWN) != 0;
        var ctrl = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
        var shift = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;

        switch (vk)
        {
            case VK_LWIN or VK_RWIN or VK_APPS or VK_F1:
                return true;
            case VK_TAB when alt:
            case VK_ESCAPE when alt || ctrl:
            case VK_F4 when alt:
                return true;
            case VK_Q when ctrl && shift:
                if (isDown) AdminChord?.Invoke();
                return true;
            case VK_T when ctrl && _showActive():
                return true;
            default:
                return false;
        }
    }

    public void Dispose() => Uninstall();
}
