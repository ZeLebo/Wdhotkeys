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
    private const uint MOD_WIN = 0x0008;
    private const uint MOD_NOREPEAT = 0x4000;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        var icon = LoadEmbeddedIcon("icon.ico") ?? SystemIcons.Application;

        using var hotkeyWindow = new HotkeyWindow();
        using var tray = CreateTrayIcon(icon);

        // Хоткеи: Ctrl+Alt+Win+1..9
        uint mods = MOD_CONTROL | MOD_ALT | MOD_WIN | MOD_NOREPEAT;
        for (int i = 1; i <= 9; i++)
        {
            uint vk = (uint)('0' + i);
            if (!RegisterHotKey(hotkeyWindow.Handle, i, mods, vk))
            {
                // если нужно — можно писать в файл/ивентлог
                // int err = Marshal.GetLastWin32Error();
            }
        }

        Application.ApplicationExit += (_, _) =>
        {
            for (int i = 1; i <= 9; i++)
                UnregisterHotKey(hotkeyWindow.Handle, i);
        };

        // Обработка WM_HOTKEY
        hotkeyWindow.HotkeyPressed += id =>
        {
            int index = id - 1;
            var desktops = VirtualDesktop.GetDesktops();
            if (index >= 0 && index < desktops.Length)
                desktops[index].Switch();
        };

        // Запускаем WinForms message pump (это и чинит контекстное меню)
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
            Text = "wdhotkeys (Ctrl+Alt+Win+1..9)",
            Icon = icon,
            Visible = true,
            ContextMenuStrip = menu
        };

        // Убираем логику “двойной клик = выход”
        // (ничего не подписываем на DoubleClick)

        return tray;
    }

    private static Icon? LoadEmbeddedIcon(string fileName)
    {
        // Ищем ресурс, заканчивающийся на ".icon.ico" или "icon.ico"
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
            // Создаём невидимое окно, на которое будут приходить WM_HOTKEY
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
