# XRai Security Model

This document describes the security architecture of XRai's named-pipe interface,
the threats it defends against, and the known limitations of the current design.

## What the pipes expose

XRai exposes a JSON-over-named-pipe protocol that gives callers full automation
authority over a running Excel instance. A subset of the commands an authenticated
client can invoke:

- **Read and write any workbook cell, formula, format, table, pivot, chart, or
  slicer.** Clients can exfiltrate the contents of every open workbook or mutate
  them arbitrarily.
- **Execute VBA** via the `vba.*` command family (`vba.import`, `vba.run`, etc.).
- **Execute Excel macros** via `macro.run`. This is a remote-code-execution
  primitive against whatever VBA is loaded in Excel.
- **Send arbitrary keystrokes** via `keys.send` and friends. On a focused Excel
  window this can invoke any feature of Excel, the Office shell, or a modal
  dialog.
- **Start local processes** via `process.start`.
- **Drive WPF task pane controls** exposed by an add-in using `XRai.Hooks`
  (`pane.click`, `pane.type`, `pane.grid.*`).
- **Mutate ViewModel state** on any exposed MVVM model (`model.set`).

Anyone who can connect to an XRai pipe and write a valid command can trivially
execute code inside the target Excel process.

## Threat model

**Threats in scope:**

1. **Cross-user attacker on the same machine.** A low-privilege account on a
   shared workstation (Remote Desktop, Terminal Server, Citrix) attempting to
   read or manipulate another user's Excel session through the pipe.
2. **Same-user attacker via an unrelated process.** A process running as the
   same Windows user (e.g. a browser extension, a malicious NPM post-install
   script) attempting to enumerate XRai pipes and inject commands.
3. **Accidental local cross-talk.** Two unrelated tools on the same machine
   enumerating `\\.\pipe\xrai_*` and colliding.

**Threats out of scope:**

- An attacker who already has full control of the target Windows user session
  (debugger, kernel driver, physical access). At that point the attacker can
  read the token file, inject into Excel directly, or replace the XRai binary.
  There is nothing cryptographically meaningful we can do from user-mode.
- Malicious Excel add-ins loaded into the same Excel process as XRai.Hooks.
  They share the same address space and can bypass any checks.
- Network attackers. XRai pipes are local only; there is no remote transport.

## Defense layers

### 1. Per-user pipe ACL (primary)

Every XRai pipe — the hooks pipe (`xrai_{pid}`) and the daemon pipe
(`xrai_daemon_{user}`) — is created using
`NamedPipeServerStreamAcl.Create` with a `PipeSecurity` that grants access to
**only two principals**:

- The current Windows user (`WindowsIdentity.GetCurrent().User`) with
  `ReadWrite | CreateNewInstance | Synchronize`.
- `NT AUTHORITY\SYSTEM` with `FullControl` (so system services can participate
  if needed; does not grant access to other interactive users).

Any other local account that calls `NamedPipeClientStream.Connect` on the pipe
gets `ERROR_ACCESS_DENIED`. The code lives in
`XRai.Core.PipeAuth.CreateRestrictedServerPipe`.

### 2. Token handshake (secondary / belt-and-braces)

Even for an attacker running as the same user, XRai requires a handshake at
connect time:

1. On server startup, the server generates a cryptographically random 128-bit
   token via `RandomNumberGenerator.GetBytes(16)`, base64-encodes it, and
   persists it to a file at
   `%LOCALAPPDATA%\XRai\tokens\{pipe_name}.token`.
2. The token file's NTFS ACL is tightened to grant Read+Write+Delete to the
   creating user only; inherited rules are removed.
3. Clients read the token from the same file path and, as the first line after
   `Connect()`, send `{"auth_token":"BASE64..."}`.
4. The server validates the token using `CryptographicOperations.FixedTimeEquals`
   (constant-time comparison, no timing side-channel).
5. If the token is missing or wrong, the server writes
   `{"ok":false,"error":"Authentication failed","code":"XRAI_AUTH_FAILED"}` and
   closes the pipe. The rejection is logged.

On graceful shutdown the server deletes its token file. Daemon startup also
runs `PipeAuth.CleanupOrphanedTokens` to delete token files whose pipes are no
longer alive, so crashes don't leave stale tokens.

The daemon's control messages `__daemon_ping__` and `__daemon_stop__` are
**exempt from token auth** — they carry no payload, exist so a client can
probe the daemon before it knows the token, and are protected by the pipe ACL
alone.

### 3. Backwards-compatibility escape hatch

Setting the environment variable `XRAI_ALLOW_UNAUTH=1` on the server side
allows unauthenticated clients through, logging a loud warning on every
connection. The default is **strict** (no env var = enforce auth). This exists
only to unblock migration for legacy clients; it should never be enabled in
production.

## Handshake protocol (wire format)

```
C→S: {"auth_token":"<base64>"}\n
S→C (success): {"ok":true,"auth":"ok"}\n
S→C (failure): {"ok":false,"error":"Authentication failed","code":"XRAI_AUTH_FAILED"}\n
```

After a successful handshake the pipe behaves exactly as before — newline-
delimited JSON commands in, newline-delimited JSON responses out.

## Known limitations

- **Same-user token exfiltration.** An attacker with read access to the
  current user's `%LOCALAPPDATA%` can read the token file. This is accepted:
  at that privilege level the attacker already controls the Windows session
  and could inject into Excel via other mechanisms (DLL injection, WinAPI
  automation, Windows UI Automation). Defending against this from within a
  user-mode process is not possible.
- **Race on first connect.** There is a narrow window between
  `GenerateAndStoreToken` writing the file and the subsequent ACL tightening
  during which the file is world-readable by inheritance. The window is a
  few milliseconds on a typical machine. If this matters to your deployment,
  pre-provision `%LOCALAPPDATA%\XRai\tokens\` with an explicit ACL before
  starting any XRai process.
- **ACL fallback.** If `NamedPipeServerStreamAcl.Create` throws (very rare;
  SID resolution failures or running under an exotic identity), XRai falls
  back to the unsecured constructor and logs a warning. The `security.status`
  command reports `pipe_acl_restricted: false` in that case.
- **Malicious add-ins in the same Excel process.** XRai.Hooks runs in the
  Excel add-in sandbox, which is to say: no sandbox. Any other add-in in the
  same Excel instance can read the token from memory, patch the validation,
  or call the exposed commands directly without going through the pipe.

## Auditing the ACL

You can verify the applied ACL with PowerShell (Sysinternals
[`pipelist`](https://learn.microsoft.com/en-us/sysinternals/downloads/pipelist)
is also useful):

```powershell
# List XRai pipes
Get-ChildItem \\.\pipe\ | Where-Object Name -like "xrai*"

# Inspect the ACL (icacls doesn't work on pipes, use accesschk from Sysinternals)
accesschk.exe -p xrai_12345
```

A correctly-secured pipe should list:

- `DOMAIN\CurrentUser`: ReadWrite, CreateNewInstance
- `NT AUTHORITY\SYSTEM`: FullControl
- No other principals.

You can also ask the server directly:

```json
{"cmd":"security.status"}
```

which returns:

```json
{
  "ok": true,
  "pipe_acl_restricted": true,
  "token_auth_enabled": true,
  "token_file_exists": true,
  "token_file_path": "C:\\Users\\...\\AppData\\Local\\XRai\\tokens\\xrai_daemon_X.token",
  "hooks_pipe_name": "xrai_12345",
  "daemon_pipe_name": "xrai_daemon_DOMAIN_user",
  "current_user": "DOMAIN\\user",
  "allow_unauthenticated": false
}
```

## Migration notes

- Clients shipped before the token handshake are rejected by default on updated
  servers. Upgrade the client, or temporarily set `XRAI_ALLOW_UNAUTH=1` on the
  server side during migration (and audit the logs for the unauthenticated
  warnings).
- Old servers don't create token files. Updated clients check for the file's
  existence and surface a clear error pointing at the expected path.
- Both server and client read the same token path. Deleting
  `%LOCALAPPDATA%\XRai\tokens\` is equivalent to rotating every in-flight
  token; next server start re-provisions fresh tokens.
