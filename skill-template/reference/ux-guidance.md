# UX Guidance — Ask Before Building

When creating or modifying an Excel add-in, ask these questions BEFORE writing XAML, ViewModel, or AddInEntry code.

## Question 1: Task pane hosting

| Option | Description | Best for |
|---|---|---|
| **Floating WPF Window** | Separate window, always on top, draggable/resizable | Prototypes, multi-monitor, large UIs |
| **Docked CTP** | Inside Excel (left/right/top/bottom), resizes with Excel. WinForms UserControl + ElementHost bridge | Production add-ins, native feel |
| **Ribbon-only** | No pane, all via ribbon buttons + dialogs | Simple tools, macro launchers |

Ask: "How should the pane appear — **floating window**, **docked panel**, or **ribbon-only**?"
If docked: "Which side — **right** (most common), **left**, **bottom**, or **top**?"

## Question 2: Visual theme

| Option | Description |
|---|---|
| **Dark** (#1a1a2e bg, light text) | Modern, premium look |
| **Light** (white/gray bg, dark text) | Matches Excel default |
| **Match Excel** (detect at runtime) | Most polished, complex to implement |

## Question 3: Pane dimensions

Ask: "Width — **narrow** (320px), **standard** (420px), or **wide** (560px)?"
For floating windows, also ask height.

## Question 4: Control framework

| Option | Description |
|---|---|
| **WPF** (default) | MVVM, data binding, rich styling. Full XRai Hooks support |
| **WinForms** | Simpler, native Windows look. XRai Hooks support |

## Question 5: Architecture

| Option | Description |
|---|---|
| **MVVM** (default for WPF) | ViewModel + INotifyPropertyChanged + ICommand. XRai `model` reads full ViewModel |
| **Code-behind** | Event handlers in .xaml.cs. XRai drives via `pane.*` commands |

## Question 6: Ribbon integration

Ask: "Custom ribbon tab? If yes — tab name and button names?"

## Store preferences in CLAUDE.md

After the user answers, save to the project's `CLAUDE.md`:

```markdown
## Add-in UX Preferences
- Hosting: Docked CTP, right side
- Theme: Dark (#1a1a2e)
- Width: 420px
- Framework: WPF
- Architecture: MVVM
- Ribbon: Custom tab "MyTool" with buttons: Scan, Settings, About
```

If preferences already exist in `CLAUDE.md`, use them without re-asking.
