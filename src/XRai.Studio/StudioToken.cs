using System.Security.Cryptography;
using System.Text;

namespace XRai.Studio;

/// <summary>
/// Token-based auth for the Studio localhost web server. Rotates on every
/// daemon start. Same pattern as Jupyter: URL contains the token on first
/// load, cookie is set, subsequent requests use the cookie.
///
/// The token is generated in memory (not persisted between runs) and also
/// written to %LOCALAPPDATA%\XRai\studio\token.txt so an external consumer
/// (VS Code extension, VS 2022 extension, scripts) can read it without
/// having to intercept stdout.
/// </summary>
public static class StudioToken
{
    private const int TokenBytes = 32;

    public static string GetTokenFilePath()
    {
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(local, "XRai", "studio", "token.txt");
    }

    /// <summary>
    /// Generate a cryptographically-random base64url token and persist it.
    /// Returns the token string. Called once per daemon start.
    /// </summary>
    public static string GenerateAndStore()
    {
        var raw = RandomNumberGenerator.GetBytes(TokenBytes);
        var token = Base64UrlEncode(raw);

        var path = GetTokenFilePath();
        var dir = Path.GetDirectoryName(path)!;
        Directory.CreateDirectory(dir);

        File.WriteAllText(path, token, Encoding.ASCII);
        TryRestrictAcl(path);

        return token;
    }

    /// <summary>
    /// Read the token that was stored by the currently-running daemon.
    /// Returns null if no daemon has run yet or the file isn't readable.
    /// Used by external clients (VS Code extension, etc.) to authenticate.
    /// </summary>
    public static string? TryReadStoredToken()
    {
        try
        {
            var path = GetTokenFilePath();
            if (!File.Exists(path)) return null;
            var token = File.ReadAllText(path, Encoding.ASCII).Trim();
            return string.IsNullOrEmpty(token) ? null : token;
        }
        catch { return null; }
    }

    /// <summary>
    /// Delete the token file on daemon shutdown. Best-effort.
    /// </summary>
    public static void ClearStoredToken()
    {
        try
        {
            var path = GetTokenFilePath();
            if (File.Exists(path)) File.Delete(path);
        }
        catch { }
    }

    /// <summary>
    /// Constant-time compare of two token strings. Prevents timing attacks
    /// on the HMAC-free token match.
    /// </summary>
    public static bool ValidateToken(string? expected, string? provided)
    {
        if (string.IsNullOrEmpty(expected) || string.IsNullOrEmpty(provided)) return false;
        if (expected.Length != provided.Length) return false;
        var a = Encoding.ASCII.GetBytes(expected);
        var b = Encoding.ASCII.GetBytes(provided);
        return CryptographicOperations.FixedTimeEquals(a, b);
    }

    private static string Base64UrlEncode(byte[] bytes)
    {
        var s = Convert.ToBase64String(bytes);
        return s.TrimEnd('=').Replace('+', '-').Replace('/', '_');
    }

    private static void TryRestrictAcl(string path)
    {
        // Best-effort ACL tightening — same as PipeAuth. Failures are logged
        // but not fatal (home directories are already per-user on Windows).
        try
        {
            if (!OperatingSystem.IsWindows()) return;
            var fi = new FileInfo(path);
            var security = fi.GetAccessControl();
            security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
            var rules = security.GetAccessRules(true, true, typeof(System.Security.Principal.NTAccount));
            foreach (System.Security.AccessControl.FileSystemAccessRule rule in rules)
            {
                security.RemoveAccessRule(rule);
            }
            var user = new System.Security.Principal.NTAccount(Environment.UserDomainName, Environment.UserName);
            security.AddAccessRule(new System.Security.AccessControl.FileSystemAccessRule(
                user,
                System.Security.AccessControl.FileSystemRights.FullControl,
                System.Security.AccessControl.AccessControlType.Allow));
            fi.SetAccessControl(security);
        }
        catch { /* ACL tightening is defense in depth; the file is already per-user */ }
    }
}
