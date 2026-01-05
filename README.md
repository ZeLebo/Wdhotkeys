# wdhotkeys

`wdhotkeys` is a small Windows utility to switch between virtual desktops and move windows across them. It runs in the background and lives in the system tray.

## Features
- Default shortcuts:  
  - `Ctrl + Alt + Win + 1..9` — switch to desktop 1..9  
  - `Shift + Ctrl + Alt + Win + 1..9` — move the active window to desktop 1..9 (and follow it)
- Uses Windows 10/11 Virtual Desktops
- Tray menu: Reload config, Open config, Exit
- Auto-creates `wdhotkeys.yaml` on first run
- Single-instance guard (won’t start a second copy)
- Can be published as single-file exe
- Optional “hard mode” to grab even shell shortcuts (Win+1..9 etc.) via low-level keyboard hook

## How it works
It uses **Slions.VirtualDesktop** to talk to the Windows virtual desktops API. Hotkeys are registered via `RegisterHotKey`, and actions are driven by the YAML config.

## Config (YAML)
Place `wdhotkeys.yaml` next to the exe (or next to the project when running from IDE). Each desktop can have multiple hotkeys for switching and moving.

```yaml
desktops:
  - desktop: 1
    switch: ["Ctrl+Alt+Win+1"]
    move:   ["Shift+Ctrl+Alt+Win+1"]
  - desktop: 2
    switch: ["Ctrl+Alt+Win+2"]
    move:   ["Shift+Ctrl+Alt+Win+2"]
```

- If the file is missing, it’s created with defaults. If it’s invalid, in-memory defaults are used.
- After editing, click **Reload config** in the tray. **Open config** opens the file with the default editor.
- To enable aggressive capture of system shortcuts, add `hardMode: true` at the root of the YAML. This enables a low-level keyboard hook and will attempt to intercept shell shortcuts like Win+1..9. Leave it `false` (default) for safer behavior using `RegisterHotKey`.

## Requirements
- Windows 10 / 11
- .NET SDK 8.0+
- x64

### Build single-file exe
```powershell
dotnet publish -c Release -r win-x64 `
  /p:PublishSingleFile=true `
  /p:SelfContained=true
```
