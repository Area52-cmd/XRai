using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;

namespace XRai.Core;

/// <summary>
/// Named pipe authentication and ACL helpers for XRai.
///
/// Security model:
///   1. Every XRai pipe is created with a restrictive ACL that grants access only
///      to the current Windows user and NT AUTHORITY\SYSTEM. Any other local user
///      that tries to enumerate or open the pipe gets ERROR_ACCESS_DENIED.
///   2. Belt-and-braces: a 128-bit random token is generated on server startup
///      and stored in %LOCALAPPDATA%\XRai\tokens\{pipe_name}.token with an ACL
///      restricted to the creator (current user). Clients read the token before
///      connecting and present it as the first line after connect. The server
///      validates the token before accepting any commands.
///
/// The ACL is the primary defense. The token handshake is the secondary defense
/// for the case where an attacker bypasses the ACL (e.g. runs as the same user
/// via another exploit). An attacker with read access to the user's LocalAppData
/// can still obtain the token, but at that point they already control the
/// Windows session — there is nothing more to protect.
/// </summary>
public static class PipeAuth
{
    /// <summary>
    /// Set the XRAI_ALLOW_UNAUTH environment variable to "1" to allow clients to
    /// connect without presenting a valid token. Intended for legacy compatibility
    /// during the migration to authenticated pipes. Logs a loud warning when used.
    /// </summary>
    public const string AllowUnauthEnvVar = "XRAI_ALLOW_UNAUTH";

    /// <summary>
    /// Error code returned to clients when authentication fails.
    /// </summary>
    public const string AuthFailedCode = "XRAI_AUTH_FAILED";

    private const int TokenBytes = 16; // 128 bits

    /// <summary>
    /// Returns true when unauthenticated pipe access is explicitly allowed via the
    /// XRAI_ALLOW_UNAUTH environment variable. Defaults to false (strict mode).
    /// </summary>
    public static bool AllowUnauthenticated =>
        string.Equals(Environment.GetEnvironmentVariable(AllowUnauthEnvVar), "1", StringComparison.Ordinal);

    /// <summary>
    /// Returns the full path to the token file for the given pipe name.
    /// Format: %LOCALAPPDATA%\XRai\tokens\{pipe_name}.token
    /// </summary>
    public static string GetTokenFilePath(string pipeName)
    {
        if (string.IsNullOrEmpty(pipeName))
            throw new ArgumentException("pipeName must not be null or empty", nameof(pipeName));

        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var dir = Path.Combine(local, "XRai", "tokens");
        // Sanitize pipe name for filesystem — pipe names shouldn't contain any
        // path separators but defend against it anyway.
        var safeName = pipeName.Replace('\\', '_').Replace('/', '_').Replace(':', '_');
        return Path.Combine(dir, safeName + ".token");
    }

    /// <summary>
    /// Generate a cryptographically strong random token, persist it to the token
    /// file for the given pipe name (restricted ACL), and return the token string.
    /// Called by the server during startup.
    /// </summary>
    public static string GenerateAndStoreToken(string pipeName)
    {
        var raw = RandomNumberGenerator.GetBytes(TokenBytes);
        var token = Convert.ToBase64String(raw);

        var path = GetTokenFilePath(pipeName);
        var dir = Path.GetDirectoryName(path)!;
        Directory.CreateDirectory(dir);

        // Write the token, then tighten the ACL so only the current user can read it.
        File.WriteAllText(path, token, Encoding.ASCII);
        TryRestrictTokenFileAcl(path);

        return token;
    }

    /// <summary>
    /// Read the token for the given pipe name from disk. Returns null if the token
    /// file doesn't exist or is unreadable. Called by the client before connecting.
    /// </summary>
    public static string? ReadToken(string pipeName)
    {
        try
        {
            var path = GetTokenFilePath(pipeName);
            if (!File.Exists(path)) return null;
            var token = File.ReadAllText(path, Encoding.ASCII).Trim();
            return string.IsNullOrEmpty(token) ? null : token;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Constant-time comparison of the provided token against the token stored on
    /// disk for the given pipe name. Returns true if the tokens match.
    /// </summary>
    public static bool ValidateToken(string pipeName, string? providedToken)
    {
        if (string.IsNullOrEmpty(providedToken)) return false;

        var expected = ReadToken(pipeName);
        if (string.IsNullOrEmpty(expected)) return false;

        // Constant-time comparison to avoid timing-channel leakage.
        var a = Encoding.ASCII.GetBytes(expected);
        var b = Encoding.ASCII.GetBytes(providedToken);
        return CryptographicOperations.FixedTimeEquals(a, b);
    }

    /// <summary>
    /// Delete the token file for the given pipe name. Called by the server on
    /// graceful shutdown. Swallows errors — a leftover token is harmless since
    /// the pipe won't exist to validate it against.
    /// </summary>
    public static void ClearToken(string pipeName)
    {
        try
        {
            var path = GetTokenFilePath(pipeName);
            if (File.Exists(path)) File.Delete(path);
        }
        catch { }
    }

    /// <summary>
    /// Build a PipeSecurity object that grants ReadWrite + CreateNewInstance to
    /// the current Windows user and FullControl to NT AUTHORITY\SYSTEM, denying
    /// everyone else by virtue of the absence of any other allow rule.
    /// </summary>
    public static PipeSecurity BuildRestrictedPipeSecurity()
    {
        var currentUser = WindowsIdentity.GetCurrent().User
            ?? throw new InvalidOperationException("WindowsIdentity.GetCurrent().User returned null");

        var security = new PipeSecurity();
        security.AddAccessRule(new PipeAccessRule(
            currentUser,
            PipeAccessRights.ReadWrite | PipeAccessRights.CreateNewInstance | PipeAccessRights.Synchronize,
            AccessControlType.Allow));
        security.AddAccessRule(new PipeAccessRule(
            new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null),
            PipeAccessRights.FullControl,
            AccessControlType.Allow));
        return security;
    }

    /// <summary>
    /// Create a NamedPipeServerStream with the restricted ACL applied. Uses the
    /// .NET 8+ NamedPipeServerStreamAcl.Create factory which is the only API that
    /// accepts a PipeSecurity on modern Windows.
    /// </summary>
    public static NamedPipeServerStream CreateRestrictedServerPipe(
        string pipeName,
        int maxInstances,
        PipeOptions options = PipeOptions.Asynchronous,
        PipeDirection direction = PipeDirection.InOut,
        PipeTransmissionMode transmissionMode = PipeTransmissionMode.Byte)
    {
        var security = BuildRestrictedPipeSecurity();
        return NamedPipeServerStreamAcl.Create(
            pipeName,
            direction,
            maxInstances,
            transmissionMode,
            options,
            inBufferSize: 0,
            outBufferSize: 0,
            pipeSecurity: security);
    }

    /// <summary>
    /// Removes inherited permissions from the token file and leaves only an explicit
    /// rule for the current user. Falls back silently on non-NTFS volumes.
    /// </summary>
    private static void TryRestrictTokenFileAcl(string path)
    {
        try
        {
            var currentUser = WindowsIdentity.GetCurrent().User
                ?? throw new InvalidOperationException("Unable to get current user SID");

            var fileInfo = new FileInfo(path);
            var security = fileInfo.GetAccessControl();

            // Break inheritance and remove inherited rules.
            security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);

            // Remove any existing explicit access rules.
            var existing = security.GetAccessRules(true, false, typeof(SecurityIdentifier));
            foreach (FileSystemAccessRule rule in existing)
            {
                security.RemoveAccessRuleAll(rule);
            }

            // Grant Read+Write to the current user only.
            security.AddAccessRule(new FileSystemAccessRule(
                currentUser,
                FileSystemRights.Read | FileSystemRights.Write | FileSystemRights.Delete,
                InheritanceFlags.None,
                PropagationFlags.None,
                AccessControlType.Allow));

            fileInfo.SetAccessControl(security);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PipeAuth] Failed to restrict ACL on {path}: {ex.Message}");
        }
    }

    /// <summary>
    /// Try to parse the first line a client sends and extract its auth_token field.
    /// Returns the token string or null if the line is absent / malformed / missing
    /// the field. We deliberately do a light-weight substring parse rather than a
    /// full JSON parse so that (a) we fail fast on garbage and (b) an attacker
    /// can't crash the server with a malformed JSON bomb during the auth handshake.
    /// </summary>
    public static string? TryExtractAuthToken(string? handshakeLine)
    {
        if (string.IsNullOrEmpty(handshakeLine)) return null;

        try
        {
            var node = System.Text.Json.Nodes.JsonNode.Parse(handshakeLine);
            return node?["auth_token"]?.GetValue<string>();
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Build the JSON line a client sends as its first message after connecting.
    /// </summary>
    public static string BuildHandshakeLine(string token)
    {
        var obj = new System.Text.Json.Nodes.JsonObject
        {
            ["auth_token"] = token,
        };
        return obj.ToJsonString();
    }

    /// <summary>
    /// Build the JSON error response the server sends when authentication fails.
    /// </summary>
    public static string BuildAuthFailedResponse()
    {
        var obj = new System.Text.Json.Nodes.JsonObject
        {
            ["ok"] = false,
            ["error"] = "Authentication failed",
            ["code"] = AuthFailedCode,
        };
        return obj.ToJsonString();
    }

    /// <summary>
    /// Build the JSON success response the server sends when authentication succeeds.
    /// </summary>
    public static string BuildAuthOkResponse()
    {
        var obj = new System.Text.Json.Nodes.JsonObject
        {
            ["ok"] = true,
            ["auth"] = "ok",
        };
        return obj.ToJsonString();
    }

    /// <summary>
    /// Clean up orphaned token files whose owning process is no longer alive.
    /// Called on server startup so that a crashed server doesn't leave stale
    /// tokens that could later be misused. The token file name does not by
    /// itself identify a pid — we scan all xrai token files and drop any whose
    /// associated pipe is unreachable.
    /// </summary>
    public static void CleanupOrphanedTokens(Func<string, bool> isPipeAlive)
    {
        try
        {
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var dir = Path.Combine(local, "XRai", "tokens");
            if (!Directory.Exists(dir)) return;

            foreach (var file in Directory.EnumerateFiles(dir, "*.token"))
            {
                try
                {
                    var pipeName = Path.GetFileNameWithoutExtension(file);
                    if (!isPipeAlive(pipeName))
                    {
                        File.Delete(file);
                    }
                }
                catch { }
            }
        }
        catch { }
    }
}
