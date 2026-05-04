# Teardown, shutdown & dialog handling — what XRai does, what it can't, what it works around

## TL;DR

In your `IExcelAddIn.AutoClose()`:
```csharp
public void AutoClose() => Pilot.Shutdown();
```
That's the entire consumer-side recommendation.

In your CLI flow: just keep using `xrai rebuild` — it auto-scrubs blocker
state from the previous session, so every launch starts clean.

If you launch Excel any other way and hit a startup blocker:
```json
{"cmd":"connect","auto_dismiss":true}
```
will detect and dismiss "Excel didn't shut down properly — Safe Mode?"
and similar prompts.

---

## The .NET 8 + Excel-DNA + WPF teardown crash

### What actually happens

Excel-DNA addins built against **.NET 8** that load **WPF** assemblies hit
a CoreCLR strict-shutdown FailFast (`0x80131506`, `COR_E_EXECUTIONENGINE`)
when Excel exits. This is a documented runtime-level interaction between:

- CoreCLR's strict AssemblyLoadContext-unload sweep on .NET 8+
- WPF's process-wide static state (KeyboardDevice, automation peers,
  Application shadow state, etc.)
- Excel-DNA's collectible AssemblyLoadContext model

**XRai cannot fix this from inside the addin's load context.** Verified
empirically: even with an entirely empty `AutoClose()` method, the crash
reproduces 100% with the same fault offset every time. The problem is in
CoreCLR + WPF + Excel-DNA, not in XRai.

The same crash hits Bloomberg, FactSet, and most large Excel-DNA + WPF
addins on .NET 8. Microsoft and Excel-DNA both have open tracking issues.

### What XRai DOES fix (visibility + recovery)

The crash itself is a single Event Viewer entry — cosmetic. The real
problems were the *workflow blockers* that Excel layered on top:

1. **"Excel didn't shut down properly — Safe Mode?"** prompt on next launch
2. **Document Recovery** task pane
3. **Addin auto-disabled** in `DisabledItems`

XRai now handles all three:

#### 1. Targeted Resiliency scrub on `xrai rebuild`

Every `xrai rebuild` does a per-addin scrub of the user's Office
`Resiliency` registry:

```
HKCU\Software\Microsoft\Office\{ver}\{Excel|Word|PowerPoint}\Resiliency\*
```

Only entries whose binary value contains the .xll filename of the addin
being built are deleted. **Office's protection of every other addin on
the system is preserved.** If you have other Excel addins (anything),
their `DisabledItems`/`StartupItems` entries are untouched.

The scrub appears as a step in the `rebuild` response when it removes
something:

```json
"steps": [..., "resiliency-scrub: ok (0 ms) — cleared 2 stale crash entries for this addin", ...]
```

#### 2. Per-addin scrub in `Pilot.Start`

When the addin loads, `Pilot.Start()` also scrubs Resiliency for the
specific .xll currently hosting it (via `ExcelDnaUtil.XllPath`). Same
targeted match. This catches the case where Excel was launched outside
of `xrai rebuild` and the addin happens to load.

#### 3. Startup-blocker detection in `connect`

`xrai connect` now probes Excel for blocking dialogs BEFORE COM attach:

```json
{"cmd":"connect"}
```

If Excel is showing the Recovery prompt or any other blocker dialog, the
response is structured:

```json
{
  "ok": false,
  "code": "XRAI_STARTUP_BLOCKED",
  "error": "Excel has 1 blocking dialog(s) preventing attach...",
  "data": {
    "blockers": [{
      "title": "Microsoft Excel",
      "hwnd": 1234567,
      "messageExcerpt": "Excel didn't shut down properly. Safe Mode could help...",
      "autoDismissAction": "no"
    }]
  }
}
```

To auto-dismiss in a single call:

```json
{"cmd":"connect","auto_dismiss":true}
```

XRai clicks the safe button (`No` for Safe Mode prompts, `Cancel` /
`Close` for recovery) and re-probes. Workflow continues.

#### 4. `Pilot.Shutdown()` clean-teardown helper

Single-call cleanup for consumer addins. `public void AutoClose() =>
Pilot.Shutdown();` is the entire AutoClose. Internally:

- Drops `AppDomain.UnhandledException` subscription
- Removes `Trace.Listeners`
- Detaches `ProcessExit` / `AssemblyLoadContext.Unloading` safety hooks
- Clears `static event` subscribers (`PipeServer.OnEventEmitted`,
  `ControlAdapter.OnControlChanged`, `ModelAdapter.OnModelChanged`)
- Disposes every `ControlAdapter` (releases DPD subscriptions)
- Force-closes the active pipe, joins the worker thread (8s)
- `InvokeShutdown` + `Join` on every WPF dispatcher passed to `Expose`
- `GC.Collect()` + `GC.WaitForPendingFinalizers()` to drain finalizers
- (Office hosts only) `TerminateProcess(0)` to skip the strict sweep

The TerminateProcess only fires from *explicit* `Pilot.Shutdown()` calls
inside `EXCEL.EXE`/`WINWORD.EXE`/`POWERPNT.EXE`. The safety-net hooks
(`ProcessExit`, `ALC.Unloading`) call `Stop()` only — they will NEVER
TerminateProcess mid-session.

Disable the bypass entirely if undesirable for your host:
```csharp
Pilot.DisableStrictShutdownBypass();
```

---

## The canonical Excel-DNA + WPF addin pattern

The demo addin (`demo/XRai.Demo.PortfolioAddin`) uses the documented
Excel-DNA pattern that works WITH XRai:

1. **`.dna` file**: `LoadFromBytes="false"` — load into the default
   (non-collectible) AssemblyLoadContext. Avoids the strict ALC unload
   sweep where possible.

2. **WPF inside a CTP via ElementHost**: `CustomTaskPaneFactory.CreateCustomTaskPane`
   on Excel's main thread. No dedicated `Dispatcher.Run()` STA worker.

3. **`AutoClose() => Pilot.Shutdown()`**: single line.

4. **Pilot.Expose called from `Pane.Loaded`**: ensures the WPF visual
   tree is fully realized before walking it. Walking too early misses
   collapsed/unrealized subtrees.

If you need to use a side STA thread for some other reason, just expect
the runtime crash on close — it's unavoidable. XRai's mitigations still
prevent it from blocking your workflow.

---

## Verified guarantees

| Path | Guaranteed clean? |
|---|---|
| Cell automation, `read`/`type`/`format` | ✅ |
| Sheets, charts, pivots, tables, ribbon | ✅ |
| Task pane controls (`pane.click`, `pane.type`, ...) | ✅ |
| ViewModel binding (`model`, `model.set`) | ✅ |
| Hooks pipe lifecycle (no thread leak) | ✅ |
| Static event lifecycle (no rooted-delegate leak) | ✅ |
| DPD subscriptions (auto-drop on Unloaded) | ✅ |
| Recovery dialog after close | ✅ — auto-scrubbed on next rebuild / detected by connect |
| Addin auto-disabled by Office | ✅ — same scrub clears DisabledItems entry |
| `xrai rebuild` round-trip | ✅ — every cycle |
| Underlying CoreCLR crash log | ❌ — runtime issue, awaits Microsoft/Excel-DNA fix |

---

## What an agent should do when starting fresh

1. `xrai rebuild --project=...` — kills any zombie Excel, scrubs
   per-addin Resiliency, builds, launches, attaches.
2. If the rebuild reports `attach-com: ok`, the addin is loaded.
3. If for any reason the agent connects to a pre-existing Excel:
   `{"cmd":"connect","auto_dismiss":true}` handles any blocker dialog.
4. Drive normally with `read`, `type`, `pane.click`, `model`, etc.
5. `Pilot.Shutdown()` is automatic via the addin's `AutoClose`. No
   special handling needed on the agent side.

That's the entire workflow. No registry surgery, no manual recovery, no
half-broken state across sessions.
