using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using WindowsDesktop; // Slions.VirtualDesktop
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

internal static class Program
{
    private const int WM_HOTKEY = 0x0312;

    private const uint MOD_ALT = 0x0001;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;
    private const uint MOD_WIN = 0x0008;
    private const uint MOD_NOREPEAT = 0x4000;

    private const uint GA_ROOT = 2;
    private const string TrayWindowClass = "Shell_TrayWnd";
    private const string ConfigFileName = "wdhotkeys.yaml";
    private const bool DefaultHardMode = false;

    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_SYSKEYDOWN = 0x0104;
    private const int WM_KEYUP = 0x0101;
    private const int WM_SYSKEYUP = 0x0105;

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

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public int vkCode;
        public int scanCode;
        public int flags;
        public int time;
        public IntPtr dwExtraInfo;
    }

    private static readonly Dictionary<Guid, IntPtr> LastDesktopFocus = new();
    private static string LogPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wdhotkeys.log");

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
        try
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (_, e) => Log($"ThreadException: {e.Exception}");
            AppDomain.CurrentDomain.UnhandledException += (_, e) => Log($"UnhandledException: {e.ExceptionObject}");

        using var mutex = new Mutex(true, "Global\\wdhotkeys-single-instance", out bool createdNew);
        if (!createdNew)
            return;

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        var uiContext = new WindowsFormsSynchronizationContext();
        SynchronizationContext.SetSynchronizationContext(uiContext);

        var icon = LoadEmbeddedIcon("icon.ico") ?? SystemIcons.Application;

        using var hotkeyWindow = new HotkeyWindow();
        using var hotkeys = new HotkeyManager(hotkeyWindow.Handle, uiContext, HandleHotkey);
        using var tray = CreateTrayIcon(
            icon,
            reloadConfig: () => hotkeys.Reload(LoadConfig()),
            openConfig: OpenConfigFile);

        EnsureConfigFileExists();
        hotkeys.Reload(LoadConfig());

        hotkeyWindow.HotkeyPressed += hotkeys.HandleRegisteredHotkey;

        Application.Run();
    }
        catch (Exception ex)
        {
            Log($"Fatal: {ex}");
            MessageBox.Show($"wdhotkeys crashed:\n{ex.Message}", "wdhotkeys", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static void HandleHotkey(HotkeyAction action)
    {
        try
        {
            var desktops = VirtualDesktop.GetDesktops();
            int targetIdx = action.Desktop - 1;
            if (targetIdx < 0 || targetIdx >= desktops.Length)
                return;

            if (action.Kind == HotkeyActionKind.Switch)
            {
                TrySaveCurrentFocus();
                desktops[targetIdx].Switch();
                RestoreDesktopFocus(desktops[targetIdx]);
            }
            else if (action.Kind == HotkeyActionKind.Move)
            {
                if (TryMoveForegroundWindow(desktops[targetIdx], out var movedWindow))
                {
                    desktops[targetIdx].Switch();
                    if (movedWindow != IntPtr.Zero)
                        SetForegroundWindow(movedWindow);
                }
            }
            Log($"Handled {action.Kind} -> desktop {action.Desktop}");
        }
        catch (Exception ex)
        {
            Log($"HandleHotkey error: {ex}");
        }
    }

    private static NotifyIcon CreateTrayIcon(Icon icon, Action reloadConfig, Action openConfig)
    {
        var menu = new ContextMenuStrip();

        var openItem = new ToolStripMenuItem("Open config");
        openItem.Click += (_, _) => openConfig();
        menu.Items.Add(openItem);

        var reloadItem = new ToolStripMenuItem("Reload config");
        reloadItem.Click += (_, _) => reloadConfig();
        menu.Items.Add(reloadItem);

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

        return tray;
    }

    private static ConfigModel LoadConfig()
    {
        try
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigFileName);
            if (!File.Exists(path))
            {
                var def = DefaultConfig();
                Log($"Config missing, using defaults. HardMode={def.HardMode}, desktops={def.Desktops.Count}");
                return def;
            }

            var yaml = File.ReadAllText(path);
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            var model = deserializer.Deserialize<ConfigModel?>(yaml);
            var result = model is null || model.Desktops.Count == 0 ? DefaultConfig() : model;
            Log($"Config loaded. HardMode={result.HardMode}, desktops={result.Desktops.Count}");
            return result;
        }
        catch (Exception ex)
        {
            Log($"LoadConfig failed: {ex}");
            return DefaultConfig();
        }
    }

    private static ConfigModel DefaultConfig()
    {
        var desktops = new List<DesktopHotkeys>();
        for (int i = 1; i <= 5; i++)
        {
            desktops.Add(new DesktopHotkeys
            {
                Desktop = i,
                Switch = new List<string> { $"Ctrl+Alt+Win+{i}" },
                Move = new List<string> { $"Shift+Ctrl+Alt+Win+{i}" }
            });
        }

        return new ConfigModel { Desktops = desktops };
    }

    private static void EnsureConfigFileExists()
    {
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigFileName);
        if (File.Exists(path))
            return;

        try
        {
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            var yaml = serializer.Serialize(DefaultConfig());
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? AppDomain.CurrentDomain.BaseDirectory);
            File.WriteAllText(path, yaml);
        }
        catch
        {
            // Ignore I/O errors; app will fall back to defaults in-memory.
        }
    }

    private static void OpenConfigFile()
    {
        EnsureConfigFileExists();
        try
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigFileName);
            var psi = new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            };
            Process.Start(psi);
        }
        catch
        {
            // ignore failures to launch editor
        }
    }

    private static bool TryMoveForegroundWindow(VirtualDesktop targetDesktop, out IntPtr movedWindow)
    {
        movedWindow = IntPtr.Zero;
        var hWnd = GetActiveTopLevelWindow();
        if (hWnd == IntPtr.Zero)
            return false;

        try
        {
            VirtualDesktop.MoveToDesktop(hWnd, targetDesktop);
            movedWindow = hWnd;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void SaveCurrentDesktopFocus(VirtualDesktop? desktop)
    {
        if (desktop is null)
            return;

        var hWnd = GetActiveTopLevelWindow();
        if (hWnd == IntPtr.Zero)
            return;

        LastDesktopFocus[desktop.Id] = hWnd;
    }

    private static void TrySaveCurrentFocus()
    {
        try
        {
            SaveCurrentDesktopFocus(VirtualDesktop.Current);
        }
        catch (Exception ex)
        {
            Log($"Save focus failed: {ex}");
        }
    }

    private static void RestoreDesktopFocus(VirtualDesktop desktop)
    {
        if (!LastDesktopFocus.TryGetValue(desktop.Id, out var hWnd))
            return;

        if (hWnd == IntPtr.Zero || !IsWindowVisible(hWnd))
            return;

        // Ensure the window is on the current desktop before focusing.
        if (!VirtualDesktop.IsCurrentVirtualDesktop(hWnd))
            return;

        SetForegroundWindow(hWnd);
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
        // ищем ресурс, инкапсулированный как ".icon.ico" или "icon.ico"
        var asm = Assembly.GetExecutingAssembly();
        var resName = Array.Find(asm.GetManifestResourceNames(),
            n => n.EndsWith("." + fileName, StringComparison.OrdinalIgnoreCase) ||
                 n.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));

        if (resName is null) return null;

        using var stream = asm.GetManifestResourceStream(resName);
        if (stream is null) return null;

        return new Icon(stream);
    }

    private sealed class HotkeyManager : IDisposable
    {
        private readonly IntPtr _windowHandle;
        private readonly SynchronizationContext _uiContext;
        private readonly Action<HotkeyAction> _onHotkey;
        private readonly Dictionary<int, HotkeyAction> _actions = new();
        private readonly List<HookBinding> _hookBindings = new();
        private IntPtr _hookHandle = IntPtr.Zero;
        private LowLevelKeyboardProc? _hookProc;
        private int _nextId = 1;
        private bool _hardMode;

        public HotkeyManager(IntPtr windowHandle, SynchronizationContext uiContext, Action<HotkeyAction> onHotkey)
        {
            _windowHandle = windowHandle;
            _uiContext = uiContext;
            _onHotkey = onHotkey;
        }

        public void Reload(ConfigModel config)
        {
            UnregisterAll();
            _nextId = 1;
            _hardMode = config.HardMode;
            _hookBindings.Clear();

            foreach (var desktop in config.Desktops)
            {
                Register(desktop.Desktop, HotkeyActionKind.Switch, desktop.Switch);
                Register(desktop.Desktop, HotkeyActionKind.Move, desktop.Move);
            }

            if (_hardMode && _hookBindings.Count > 0 && !EnsureHook())
            {
                // Hook failed; fall back to safe mode for everything
                Log("Hook failed, falling back to safe mode");
                _hardMode = false;
                _hookBindings.Clear();
                UnregisterAll();
                foreach (var desktop in config.Desktops)
                {
                    Register(desktop.Desktop, HotkeyActionKind.Switch, desktop.Switch);
                    Register(desktop.Desktop, HotkeyActionKind.Move, desktop.Move);
                }
                ReleaseHook();
            }
            else if (!_hardMode || _hookBindings.Count == 0)
            {
                ReleaseHook();
            }

            Log($"Reload: HardMode={_hardMode}, actions={_actions.Count}, hookBindings={_hookBindings.Count}, hookHandle={_hookHandle}");
        }

        public bool TryGetAction(int id, out HotkeyAction action) => _actions.TryGetValue(id, out action);

        public void HandleRegisteredHotkey(int id) => DispatchIfKnown(id);

        private void Register(int desktop, HotkeyActionKind kind, List<string> hotkeys)
        {
            foreach (var hotkey in hotkeys)
            {
                if (!TryParseHotkey(hotkey, out var mods, out var vk))
                    continue;

                if (_hardMode)
                {
                    uint hookMods = mods & ~MOD_NOREPEAT; // low-level hook doesn't use MOD_NOREPEAT
                    bool registered = RegisterHotKey(_windowHandle, _nextId, mods, vk);
                    if (registered)
                    {
                        _actions[_nextId] = new HotkeyAction(kind, desktop);
                        _nextId++;
                        Log($"Registered via RegisterHotKey: {hotkey} -> desktop {desktop} ({kind})");
                    }
                    else
                    {
                        _hookBindings.Add(new HookBinding(new HotkeyAction(kind, desktop), hookMods, (int)vk));
                        Log($"Registered via hook: {hotkey} -> desktop {desktop} ({kind})");
                    }
                }
                else
                {
                    int id = _nextId++;
                    if (RegisterHotKey(_windowHandle, id, mods, vk))
                    {
                        _actions[id] = new HotkeyAction(kind, desktop);
                        Log($"Registered (safe): {hotkey} -> desktop {desktop} ({kind})");
                    }
                }
            }
        }

        private void UnregisterAll()
        {
            foreach (var id in _actions.Keys)
                UnregisterHotKey(_windowHandle, id);
            _actions.Clear();
            _hookBindings.Clear();
        }

        public void Dispose()
        {
            UnregisterAll();
            ReleaseHook();
        }

        private bool EnsureHook()
        {
            if (_hookHandle != IntPtr.Zero)
                return true;

            _hookProc = HookCallback;
            IntPtr hModule = IntPtr.Zero;
            try
            {
                using var curProcess = Process.GetCurrentProcess();
                using var curModule = curProcess.MainModule;
                if (curModule != null)
                    hModule = GetModuleHandle(curModule.ModuleName);
            }
            catch
            {
                hModule = IntPtr.Zero;
            }

            _hookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, hModule, 0);
            if (_hookHandle == IntPtr.Zero)
                _hookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, IntPtr.Zero, 0);

            if (_hookHandle != IntPtr.Zero)
                Log("Hook installed");
            else
                Log("Hook install failed");

            return _hookHandle != IntPtr.Zero;
        }

        private void ReleaseHook()
        {
            if (_hookHandle != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookHandle);
                _hookHandle = IntPtr.Zero;
            }
            _hookProc = null;
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                bool isKeyDown = wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN;
                if (nCode >= 0 && isKeyDown)
                {
                    var info = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                    int vkCode = info.vkCode;

                    uint currentMods = GetCurrentModifiersSnapshot();

                    foreach (var binding in _hookBindings)
                    {
                        if (binding.VirtualKey == vkCode && currentMods == binding.Modifiers)
                        {
                            Log($"Hook matched: vk={vkCode}, mods={currentMods} -> {binding.Action.Kind} desktop {binding.Action.Desktop}");
                            Dispatch(binding.Action);
                            return (IntPtr)1; // suppress
                        }
                    }
                }
            }
            catch
            {
                // swallow to avoid crashing the hook
            }

            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }

        private static uint GetCurrentModifiersSnapshot()
        {
            uint mods = 0;
            if ((GetAsyncKeyState((int)Keys.ControlKey) & 0x8000) != 0)
                mods |= MOD_CONTROL;
            if ((GetAsyncKeyState((int)Keys.Menu) & 0x8000) != 0)
                mods |= MOD_ALT;
            if ((GetAsyncKeyState((int)Keys.ShiftKey) & 0x8000) != 0)
                mods |= MOD_SHIFT;
            if ((GetAsyncKeyState((int)Keys.LWin) & 0x8000) != 0 || (GetAsyncKeyState((int)Keys.RWin) & 0x8000) != 0)
                mods |= MOD_WIN;
            return mods;
        }

        private void DispatchIfKnown(int id)
        {
            if (TryGetAction(id, out var action))
                Dispatch(action);
        }

        private void Dispatch(HotkeyAction action)
        {
            _uiContext.Post(_ => _onHotkey(action), null);
        }
    }

    private enum HotkeyActionKind
    {
        Switch,
        Move
    }

    private readonly record struct HotkeyAction(HotkeyActionKind Kind, int Desktop);

    private sealed record HookBinding(HotkeyAction Action, uint Modifiers, int VirtualKey);

    private static bool TryParseHotkey(string hotkey, out uint modifiers, out uint vk)
    {
        modifiers = 0;
        vk = 0;

        var parts = hotkey.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
            return false;

        bool keySet = false;
        foreach (var part in parts)
        {
            var token = part.ToUpperInvariant();
            switch (token)
            {
                case "CTRL":
                case "CONTROL":
                    modifiers |= MOD_CONTROL;
                    break;
                case "ALT":
                    modifiers |= MOD_ALT;
                    break;
                case "SHIFT":
                    modifiers |= MOD_SHIFT;
                    break;
                case "WIN":
                case "META":
                    modifiers |= MOD_WIN;
                    break;
                default:
                    if (keySet)
                        return false;

                    if (token.Length == 1 && char.IsLetterOrDigit(token[0]))
                    {
                        vk = (uint)char.ToUpperInvariant(token[0]);
                        keySet = true;
                        break;
                    }

                    if (token.StartsWith("F", StringComparison.OrdinalIgnoreCase) &&
                        int.TryParse(token.AsSpan(1), out int fn) && fn is >= 1 and <= 24)
                    {
                        vk = (uint)(Keys.F1 + fn - 1);
                        keySet = true;
                        break;
                    }

                    return false;
            }
        }

        if (!keySet)
            return false;

        modifiers |= MOD_NOREPEAT;
        return true;
    }

    private sealed class ConfigModel
    {
        public bool HardMode { get; set; } = DefaultHardMode;
        public List<DesktopHotkeys> Desktops { get; set; } = new();
    }

    private sealed class DesktopHotkeys
    {
        public int Desktop { get; set; }
        public List<string> Switch { get; set; } = new();
        public List<string> Move { get; set; } = new();
    }

    private static void Log(string message)
    {
        try
        {
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
            File.AppendAllText(LogPath, line);
        }
        catch
        {
            // ignore logging failures
        }
    }

    private sealed class HotkeyWindow : NativeWindow, IDisposable
    {
        public event Action<int>? HotkeyPressed;

        public HotkeyWindow()
        {
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
