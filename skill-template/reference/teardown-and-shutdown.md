# Teardown & shutdown — what XRai does for you, and what it can't

## TL;DR

In your `IExcelAddIn.AutoClose()`:
```csharp
public void AutoClose() => Pilot.Shutdown();
```
That's the entire recommendation. `Pilot.Shutdown()` does every cleanup XRai
needs: drops static event subscriptions, joins the pipe-server thread, shuts
down every WPF dispatcher it has been exposed to, joins those threads, drains
finalizers, and then (when running inside Excel/Word/PowerPoint) calls
Win32 `TerminateProcess` to skip CoreCLR's strict ALC-unload sweep.

## What `Pilot.Shutdown()` actually does

1. **`Pilot.Stop()`** — the strict cleanup pass:
   - Removes the `AppDomain.UnhandledException` subscription installed by
     `ErrorCapture` (rooted delegate that would otherwise pin the addin's
     load context).
   - Removes the `Trace.Listeners` entry installed by `LogCapture`.
   - Detaches the `ProcessExit` and `AssemblyLoadContext.Unloading` safety-net
     hooks Pilot installs in `Start()`.
   - Clears the static events `PipeServer.OnEventEmitted`,
     `ControlAdapter.OnControlChanged`, `ModelAdapter.OnModelChanged` so any
     subscriber delegates from consumer code release.
   - Disposes every `ControlAdapter` (each one removes its
     `DependencyPropertyDescriptor.AddValueChanged` callbacks and its
     `Unloaded` / `Dispatcher.ShutdownStarted` handlers).
   - Force-closes the active `NamedPipeServerStream` so the pipe-server
     thread's blocking `ReadLine()` returns immediately, then `Join`s the
     thread (8s budget) so it actually exits before we proceed.

2. **WPF dispatcher shutdown.** Every dispatcher Pilot saw via `Pilot.Expose`
   is tracked by weak reference. `Shutdown()` calls `InvokeShutdown()` on each
   one and `Join`s its thread (5s budget) so the WPF STA threads actually
   terminate before AutoClose returns.

3. **Finalizer drain.** `GC.Collect()` + `GC.WaitForPendingFinalizers()` x 2.

4. **Strict-shutdown bypass.** `Pilot.Shutdown()` ends with a call to Win32
   `TerminateProcess(GetCurrentProcess(), 0)` — but ONLY when the host
   process is `EXCEL.EXE`, `WINWORD.EXE`, or `POWERPNT.EXE` AND the bypass
   hasn't been explicitly disabled via `Pilot.DisableStrictShutdownBypass()`.

   **Why:** .NET 8 CoreCLR's collectible-AssemblyLoadContext-unload sweep
   is fundamentally incompatible with WPF static state. WPF's process-wide
   statics (e.g. `KeyboardDevice`, `MouseDevice`, automation peers,
   `Application` shadow state) root types in the unloading context. CoreCLR
   cannot drain them and FailFasts with `0x80131506` (`COR_E_EXECUTIONENGINE`)
   on Excel exit. We bypass the sweep by terminating the host cleanly *just*
   before the sweep would run.

## When the crash still happens

There is one scenario XRai cannot fully prevent on its own:

**Excel-DNA does not always call `IExcelAddIn.AutoClose()`** when Excel is
closing. If Excel exits via process termination rather than addin unload,
neither `AutoClose`, nor `AppDomain.ProcessExit`, nor
`AssemblyLoadContext.Unloading` fires before CoreCLR's strict sweep. None
of XRai's hooks have a chance to run.

**Mitigations (in order of effectiveness):**

1. **Use `LoadFromBytes="false"` in your `.dna` file** — loads the addin
   into the default (non-collectible) load context. The strict sweep
   doesn't run on the default context, so most teardown crashes disappear.
   Trade-off: addin can no longer hot-reload via Excel-DNA's normal path.

   Custom `.dna` template:
   ```xml
   <?xml version="1.0" encoding="utf-8"?>
   <DnaLibrary RollForward="LatestMinor" Name="My Addin" RuntimeVersion="v8.0"
               xmlns="http://schemas.excel-dna.net/addin/2020/07/dnalibrary">
     <ExternalLibrary Path="MyAddin.dll" ExplicitExports="false"
                      LoadFromBytes="false" Pack="true" IncludePdb="false" />
   </DnaLibrary>
   ```

2. **Don't host WPF on a dedicated `Dispatcher.Run()` thread.** If the
   addin only uses WinForms (or hosts WPF inside an `ElementHost` on
   Excel's main thread), the strict-sweep crash class becomes much rarer.
   Excel-DNA's official guidance is the same.

3. **Suppress the Windows Error Reporting dialog.** The crash log appears
   in Event Viewer regardless, but you can stop the "Excel didn't shut
   down properly — Safe Mode?" prompt by clearing the Resiliency keys
   under `HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency` after
   each session. (Belt-and-braces, not a real fix.)

## What XRai 1.0+ guarantees

**On clean teardown paths (AutoClose / ProcessExit / ALC.Unloading fires):**

- ✅ No stuck modifier keys (PostMessage instead of SendKeys.Send)
- ✅ No leaked DPD subscriptions (auto-detach on Unloaded)
- ✅ No leaked `AppDomain.UnhandledException` handlers (Uninstall in Stop)
- ✅ No leaked static-event delegates (atomic clear in Stop)
- ✅ No leaked pipe-server threads (force-close + Join)
- ✅ No leaked dispatcher threads (InvokeShutdown + Join)
- ✅ Exit via TerminateProcess instead of CLR strict sweep on Office hosts
- ✅ Single-call `Pilot.Shutdown()` — consumers don't have to orchestrate

**On unclean paths (AutoClose skipped, sweep runs first):**

- ❌ XRai cannot prevent the runtime FailFast — by the time any code
  could run, the sweep is already in progress.
- ✅ Mitigation: use `LoadFromBytes="false"` in `.dna` (see above).

This is a known Excel-DNA + .NET 8 + WPF runtime limitation, not an XRai
bug. Bloomberg, FactSet, and most other large Excel-DNA + WPF addins on
.NET 8 hit the same crash. Microsoft and Excel-DNA both have open issues
tracking it.
