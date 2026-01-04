using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WindowsDesktop; // Slions.VirtualDesktop

internal static class Program
{
    private const int WM_HOTKEY = 0x0312;

    private const uint MOD_ALT = 0x0001;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;
    private const uint MOD_WIN = 0x0008;
    private const uint MOD_NOREPEAT = 0x4000;

    private const int MoveHotkeyOffset = 100;
    private const uint GA_ROOT = 2;
    private const string TrayWindowClass = "Shell_TrayWnd";
    private const uint ATTACH_INPUT = 0x0001;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

    [DllImport("user32.dll")]
    private static extern IntPtr GetShellWindow();

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr FindWindow(string lpClassName, string? lpWindowName);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

    [DllImport("user32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    [DllImport("user32.dll")]
    private static extern IntPtr GetActiveWindow();

    [DllImport("user32.dll")]
    private static extern IntPtr GetFocus();

    [StructLayout(LayoutKind.Sequential)]
    private struct GUITHREADINFO
    {
        public int cbSize;
        public int flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public Rectangle rcCaret;
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        var icon = LoadEmbeddedIcon("icon.ico") ?? SystemIcons.Application;

        using var hotkeyWindow = new HotkeyWindow();
        using var tray = CreateTrayIcon(icon);

        // Hotkeys: Ctrl+Alt+Win+1..9 (switch), Shift+Ctrl+Alt+Win+1..9 (move active window)
        uint switchMods = MOD_CONTROL | MOD_ALT | MOD_WIN | MOD_NOREPEAT;
        uint moveMods = MOD_SHIFT | MOD_CONTROL | MOD_ALT | MOD_WIN | MOD_NOREPEAT;
        for (int i = 1; i <= 9; i++)
        {
            uint vk = (uint)('0' + i);
            RegisterHotKey(hotkeyWindow.Handle, i, switchMods, vk);
            RegisterHotKey(hotkeyWindow.Handle, MoveHotkeyOffset + i, moveMods, vk);
        }

        Application.ApplicationExit += (_, _) =>
        {
            for (int i = 1; i <= 9; i++)
            {
                UnregisterHotKey(hotkeyWindow.Handle, i);
                UnregisterHotKey(hotkeyWindow.Handle, MoveHotkeyOffset + i);
            }
        };

        // Handle WM_HOTKEY from the hidden window
        hotkeyWindow.HotkeyPressed += id =>
        {
            var desktops = VirtualDesktop.GetDesktops();
            if (id is >= 1 and <= 9)
            {
                int index = id - 1;
                if (index >= 0 && index < desktops.Length)
                    desktops[index].Switch();
                return;
            }

            if (id >= MoveHotkeyOffset + 1 && id <= MoveHotkeyOffset + 9)
            {
                int index = id - MoveHotkeyOffset - 1;
                if (index < 0 || index >= desktops.Length)
                    return;

                if (TryMoveForegroundWindow(desktops[index]))
                    desktops[index].Switch();
            }
        };

        // ????????? WinForms message pump (??? ? ????? ??????????? ????)
        Application.Run();
    }

    private static NotifyIcon CreateTrayIcon(Icon icon)
    {
        var menu = new ContextMenuStrip();

        var exitItem = new ToolStripMenuItem("Exit");
        exitItem.Click += (_, _) => Application.Exit();
        menu.Items.Add(exitItem);

        var tray = new NotifyIcon
        {
            Text = "wdhotkeys (Ctrl+Alt+Win+1..9 switch, Shift+Ctrl+Alt+Win+1..9 move)",
            Icon = icon,
            Visible = true,
            ContextMenuStrip = menu
        };

        // ??????? ?????? "??????? ???? = ?????"
        // (?????? ?? ??????????? ?? DoubleClick)

        return tray;
    }

    private static bool TryMoveForegroundWindow(VirtualDesktop targetDesktop)
    {
        var hWnd = GetActiveTopLevelWindow();
        if (hWnd == IntPtr.Zero)
            return false;

        try
        {
            VirtualDesktop.MoveToDesktop(hWnd, targetDesktop);
            return true;
        }
        catch
        {
            return false;
        }
    }

    // Try to get a real user window (not the shell/tray/child window) even while a hotkey is pressed.
    private static IntPtr GetActiveTopLevelWindow()
    {
        var shell = GetShellWindow();
        var tray = FindWindow(TrayWindowClass, null);

        IntPtr Normalize(IntPtr handle) => handle == IntPtr.Zero ? IntPtr.Zero : GetAncestor(handle, GA_ROOT);
        bool IsUsable(IntPtr handle) =>
            handle != IntPtr.Zero &&
            handle != shell &&
            handle != tray &&
            IsWindowVisible(handle);

        IntPtr PickFromThread(uint threadId)
        {
            if (threadId == 0)
                return IntPtr.Zero;

            uint current = GetCurrentThreadId();
            bool attached = false;
            try
            {
                if (current != threadId)
                {
                    attached = AttachThreadInput(current, threadId, true);
                }

                IntPtr[] threadCandidates =
                {
                    Normalize(GetActiveWindow()),
                    Normalize(GetFocus()),
                };

                foreach (var c in threadCandidates)
                {
                    if (IsUsable(c))
                        return c;
                }
            }
            finally
            {
                if (attached)
                    AttachThreadInput(current, threadId, false);
            }

            return IntPtr.Zero;
        }

        var fg = Normalize(GetForegroundWindow());
        if (IsUsable(fg))
            return fg;

        // If foreground is shell/tray, try the GUI thread info of that thread.
        uint tid = GetWindowThreadProcessId(GetForegroundWindow(), out _);
        if (tid != 0)
        {
            var info = new GUITHREADINFO { cbSize = Marshal.SizeOf<GUITHREADINFO>() };
            if (GetGUIThreadInfo(tid, ref info))
            {
                IntPtr[] candidates =
                {
                    Normalize(info.hwndActive),
                    Normalize(info.hwndFocus),
                    Normalize(info.hwndCapture),
                    Normalize(info.hwndCaret),
                };

                foreach (var candidate in candidates)
                {
                    if (IsUsable(candidate))
                        return candidate;
                }
            }

            var attached = PickFromThread(tid);
            if (IsUsable(attached))
                return attached;
        }

        return IntPtr.Zero;
    }

    private static Icon? LoadEmbeddedIcon(string fileName)
    {
        // ???? ??????, ??????????????? ?? ".icon.ico" ??? "icon.ico"
        var asm = Assembly.GetExecutingAssembly();
        var resName = Array.Find(asm.GetManifestResourceNames(),
            n => n.EndsWith("." + fileName, StringComparison.OrdinalIgnoreCase) ||
                 n.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));

        if (resName is null) return null;

        using var stream = asm.GetManifestResourceStream(resName);
        if (stream is null) return null;

        return new Icon(stream);
    }

    private sealed class HotkeyWindow : NativeWindow, IDisposable
    {
        public event Action<int>? HotkeyPressed;

        public HotkeyWindow()
        {
            // ??????? ????????? ????, ?? ??????? ????? ????????? WM_HOTKEY
            CreateHandle(new CreateParams
            {
                Caption = "wdhotkeys-hotkey-window",
                X = 0, Y = 0, Height = 0, Width = 0,
                Style = 0
            });
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                int id = m.WParam.ToInt32();
                HotkeyPressed?.Invoke(id);
            }

            base.WndProc(ref m);
        }

        public void Dispose()
        {
            DestroyHandle();
        }
    }
}
