using System.Diagnostics;
using System.Net.WebSockets;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using XRai.Core;

namespace XRai.Studio;

/// <summary>
/// The Studio web host. Owns the Kestrel server, the event bus, the state
/// provider, and the embedded dashboard resources. Started from
/// XRai.Tool.DaemonServer when --studio is passed (or when {"cmd":"studio"}
/// is sent to an already-running daemon).
///
/// Design goals:
///   1. Zero config — picks an ephemeral port, opens the browser itself.
///   2. Single-user — binds to 127.0.0.1 only, one-time token handshake.
///   3. Additive — does not modify the existing XRai command router or pipe.
///   4. Self-contained — HTML/JS/CSS are embedded resources, no wwwroot dir.
/// </summary>
public sealed class StudioHost : IDisposable
{
    public EventBus Bus { get; } = new();
    public string Token { get; private set; } = "";
    public string Url { get; private set; } = "";
    public int Port { get; private set; }

    private IHost? _host;
    private readonly Func<JsonObject> _stateProvider;
    private readonly Func<string, JsonObject, string>? _commandDispatcher;
    private readonly List<IDisposable> _disposables = new();

    /// <param name="stateProvider">Called on every GET /state request to
    ///     snapshot the current state (attach, pane, model, screenshot URL).</param>
    /// <param name="commandDispatcher">Optional. POST /command forwards here
    ///     (cmd name + args object → JSON response). If null, /command returns
    ///     501 Not Implemented.</param>
    public StudioHost(
        Func<JsonObject> stateProvider,
        Func<string, JsonObject, string>? commandDispatcher = null)
    {
        _stateProvider = stateProvider;
        _commandDispatcher = commandDispatcher;
    }

    /// <summary>
    /// Register a disposable whose lifetime is tied to the host (sources
    /// like CaptureLoop, FileWatcherSource, PipeEventSource). They get
    /// disposed automatically when the host stops.
    /// </summary>
    public void RegisterDisposable(IDisposable d) => _disposables.Add(d);

    /// <summary>
    /// Start Kestrel on an ephemeral 127.0.0.1 port, mint a fresh token,
    /// and optionally launch the default browser at the authenticated URL.
    /// Returns the full launch URL (with ?t=token query string).
    /// </summary>
    public string Start(bool launchBrowser = true)
    {
        Token = StudioToken.GenerateAndStore();

        var builder = Host.CreateDefaultBuilder()
            .ConfigureWebHostDefaults(web =>
            {
                web.UseUrls("http://127.0.0.1:0");
                web.Configure(Configure);
            })
            .ConfigureLogging(logging =>
            {
                // Keep Kestrel quiet — the daemon's log file is the source of
                // truth. We only want errors in case something goes wrong.
                logging.ClearProviders();
                logging.AddFilter("Microsoft.AspNetCore", LogLevel.Warning);
                logging.AddFilter("Microsoft.Hosting", LogLevel.Warning);
            });

        _host = builder.Build();
        _host.Start();

        // Resolve the actual bound port from the server addresses feature
        var server = _host.Services.GetRequiredService<Microsoft.AspNetCore.Hosting.Server.IServer>();
        var addresses = server.Features.Get<IServerAddressesFeature>();
        var addr = addresses?.Addresses.FirstOrDefault() ?? "http://127.0.0.1:0";
        // Parse port out of "http://127.0.0.1:54321"
        var uri = new Uri(addr);
        Port = uri.Port;
        Url = $"http://127.0.0.1:{Port}/?t={Token}";

        if (launchBrowser)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = Url,
                    UseShellExecute = true,
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to launch browser: {ex.Message}");
                Console.Error.WriteLine($"Studio URL: {Url}");
            }
        }

        return Url;
    }

    private void Configure(IApplicationBuilder app)
    {
        app.UseWebSockets(new WebSocketOptions { KeepAliveInterval = TimeSpan.FromSeconds(30) });

        app.Use(async (ctx, next) =>
        {
            // Token handshake:
            //  - Cookie is the normal long-lived auth for browser tabs.
            //  - ?t= query param is accepted on every request (API clients,
            //    WebSocket upgrades, and the first HTML load).
            //  - ONLY the browser HTML entry point (/ or /index.html) gets the
            //    promote-to-cookie-then-redirect treatment — for API endpoints
            //    the redirect strips the token and breaks curl-style callers.
            var path = ctx.Request.Path.Value ?? "/";
            var queryToken = ctx.Request.Query["t"].ToString();
            var cookieToken = ctx.Request.Cookies["xrai-studio"] ?? "";

            bool queryAuthed = StudioToken.ValidateToken(Token, queryToken);
            bool cookieAuthed = StudioToken.ValidateToken(Token, cookieToken);
            bool authed = queryAuthed || cookieAuthed;

            // HTML entry point: promote ?t= to a cookie on first load so the
            // browser doesn't keep the token in the URL bar / history.
            bool isEntryPoint = path == "/" || path == "/index.html";
            if (isEntryPoint && queryAuthed && !cookieAuthed)
            {
                ctx.Response.Cookies.Append("xrai-studio", Token, new CookieOptions
                {
                    HttpOnly = true,
                    Secure = false, // localhost, no TLS
                    SameSite = SameSiteMode.Strict,
                    Path = "/",
                });
                ctx.Response.Redirect(ctx.Request.Path);
                return;
            }

            if (!authed)
            {
                ctx.Response.StatusCode = 401;
                await ctx.Response.WriteAsync("Unauthorized. Use the URL printed when the daemon started (includes a one-time token).");
                return;
            }

            await next();
        });

        app.Use(async (ctx, next) =>
        {
            var path = ctx.Request.Path.Value ?? "/";

            if (path == "/events" && ctx.WebSockets.IsWebSocketRequest)
            {
                await HandleEventsWebSocket(ctx);
                return;
            }

            if (path == "/state")
            {
                var json = _stateProvider().ToJsonString();
                ctx.Response.ContentType = "application/json";
                await ctx.Response.WriteAsync(json);
                return;
            }

            if (path == "/ides")
            {
                var ides = IdeLauncher.DetectAll();
                var arr = new System.Text.Json.Nodes.JsonArray();
                foreach (var i in ides) arr.Add(i.ToJson());
                ctx.Response.ContentType = "application/json";
                await ctx.Response.WriteAsync(arr.ToJsonString());
                return;
            }

            if (path == "/preferences" && ctx.Request.Method == "GET")
            {
                var prefs = StudioPreferences.Load();
                ctx.Response.ContentType = "application/json";
                await ctx.Response.WriteAsync(prefs.ToJson().ToJsonString());
                return;
            }

            if (path == "/preferences" && ctx.Request.Method == "POST")
            {
                using var reader = new StreamReader(ctx.Request.Body);
                var body = await reader.ReadToEndAsync();
                try
                {
                    var node = System.Text.Json.Nodes.JsonNode.Parse(body);
                    var prefs = StudioPreferences.FromJson(node);
                    prefs.Save();
                    ctx.Response.ContentType = "application/json";
                    await ctx.Response.WriteAsync(prefs.ToJson().ToJsonString());
                }
                catch (Exception ex)
                {
                    ctx.Response.StatusCode = 400;
                    ctx.Response.ContentType = "application/json";
                    await ctx.Response.WriteAsync(Response.ErrorFromException(ex, "studio.preferences.save"));
                }
                return;
            }

            if (path == "/studio/open-logs" && ctx.Request.Method == "POST")
            {
                try
                {
                    var logsDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "XRai", "logs");
                    Directory.CreateDirectory(logsDir);
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = $"\"{logsDir}\"",
                        UseShellExecute = true,
                    });
                    ctx.Response.ContentType = "application/json";
                    await ctx.Response.WriteAsync(Response.Ok(new { path = logsDir }));
                }
                catch (Exception ex)
                {
                    ctx.Response.StatusCode = 500;
                    ctx.Response.ContentType = "application/json";
                    await ctx.Response.WriteAsync(Response.ErrorFromException(ex, "studio.open-logs"));
                }
                return;
            }

            if (path == "/ide/open" && ctx.Request.Method == "POST")
            {
                using var reader = new StreamReader(ctx.Request.Body);
                var body = await reader.ReadToEndAsync();
                try
                {
                    var node = System.Text.Json.Nodes.JsonNode.Parse(body) as System.Text.Json.Nodes.JsonObject;
                    var filePath = node?["filePath"]?.GetValue<string>();
                    var line = node?["line"]?.GetValue<int?>();
                    var kindStr = node?["kind"]?.GetValue<string?>();

                    // If no explicit line but a searchText is provided, find the
                    // line by searching the file for the first occurrence. This is
                    // how Studio targets the exact edit location from an Edit tool
                    // call — the old_string is passed as searchText.
                    var searchText = node?["searchText"]?.GetValue<string?>();
                    if (!line.HasValue && !string.IsNullOrEmpty(searchText) && !string.IsNullOrEmpty(filePath))
                    {
                        try
                        {
                            var fullPath = Path.GetFullPath(filePath);
                            if (File.Exists(fullPath))
                            {
                                var content = File.ReadAllText(fullPath);
                                var idx = content.IndexOf(searchText, StringComparison.Ordinal);
                                if (idx < 0)
                                {
                                    // Exact match failed — try first line of searchText
                                    var firstLine = searchText.Split('\n')[0].Trim();
                                    if (firstLine.Length > 10)
                                        idx = content.IndexOf(firstLine, StringComparison.Ordinal);
                                }
                                if (idx >= 0)
                                {
                                    // Count newlines up to the match position
                                    line = 1;
                                    for (int ci = 0; ci < idx; ci++)
                                    {
                                        if (content[ci] == '\n') line++;
                                    }
                                }
                            }
                        }
                        catch { /* search failed — open at top, still better than nothing */ }
                    }
                    IdeLauncher.IdeKind? kind = null;
                    if (!string.IsNullOrEmpty(kindStr) &&
                        Enum.TryParse<IdeLauncher.IdeKind>(kindStr, out var parsed))
                    {
                        kind = parsed;
                    }

                    System.Text.Json.Nodes.JsonObject result;
                    if (string.IsNullOrEmpty(filePath))
                    {
                        if (kind.HasValue)
                        {
                            result = IdeLauncher.LaunchBlank(kind.Value);
                        }
                        else
                        {
                            ctx.Response.StatusCode = 400;
                            ctx.Response.ContentType = "application/json";
                            await ctx.Response.WriteAsync(Response.Error("filePath or kind is required", code: ErrorCodes.MissingArgument));
                            return;
                        }
                    }
                    else
                    {
                        result = IdeLauncher.Open(filePath, line, null, kind);
                    }

                    ctx.Response.ContentType = "application/json";
                    await ctx.Response.WriteAsync(result.ToJsonString());
                }
                catch (Exception ex)
                {
                    ctx.Response.StatusCode = 500;
                    ctx.Response.ContentType = "application/json";
                    await ctx.Response.WriteAsync(Response.ErrorFromException(ex, "studio.ide.open"));
                }
                return;
            }

            if (path == "/command" && ctx.Request.Method == "POST")
            {
                await HandleCommand(ctx);
                return;
            }

            // Static assets — serve from embedded resources
            if (path == "/" || path == "/index.html")
            {
                await ServeEmbedded(ctx, "index.html", "text/html; charset=utf-8");
                return;
            }

            if (path == "/studio.js")
            {
                await ServeEmbedded(ctx, "studio.js", "application/javascript; charset=utf-8");
                return;
            }

            if (path == "/studio.css")
            {
                await ServeEmbedded(ctx, "studio.css", "text/css; charset=utf-8");
                return;
            }

            await next();
        });
    }

    private async Task HandleEventsWebSocket(HttpContext ctx)
    {
        using var ws = await ctx.WebSockets.AcceptWebSocketAsync();
        var (id, reader) = Bus.Subscribe();
        var cancel = ctx.RequestAborted;

        try
        {
            await foreach (var evt in reader.ReadAllAsync(cancel))
            {
                if (ws.State != WebSocketState.Open) break;
                var json = evt.ToJson().ToJsonString();
                var bytes = Encoding.UTF8.GetBytes(json);
                await ws.SendAsync(bytes, WebSocketMessageType.Text, endOfMessage: true, cancel);
            }
        }
        catch (OperationCanceledException) { /* client went away */ }
        catch (WebSocketException) { /* client went away hard */ }
        finally
        {
            Bus.Unsubscribe(id);
            try { await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "bye", CancellationToken.None); } catch { }
        }
    }

    private async Task HandleCommand(HttpContext ctx)
    {
        ctx.Response.ContentType = "application/json";

        if (_commandDispatcher == null)
        {
            ctx.Response.StatusCode = 501;
            await ctx.Response.WriteAsync(Response.Error(
                "/command is disabled on this Studio instance",
                code: ErrorCodes.NotImplemented));
            return;
        }

        using var reader = new StreamReader(ctx.Request.Body);
        var body = await reader.ReadToEndAsync();
        JsonObject? parsed;
        try { parsed = JsonNode.Parse(body) as JsonObject; }
        catch (Exception parseEx)
        {
            ctx.Response.StatusCode = 400;
            await ctx.Response.WriteAsync(Response.ErrorFromException(parseEx, "studio.command.parse"));
            return;
        }

        var cmd = parsed?["cmd"]?.GetValue<string>();
        if (string.IsNullOrEmpty(cmd))
        {
            ctx.Response.StatusCode = 400;
            await ctx.Response.WriteAsync(Response.Error("Missing 'cmd' field", code: ErrorCodes.MissingArgument));
            return;
        }

        try
        {
            var result = _commandDispatcher(cmd, parsed!);
            await ctx.Response.WriteAsync(result);
        }
        catch (Exception ex)
        {
            ctx.Response.StatusCode = 500;
            await ctx.Response.WriteAsync(Response.ErrorFromException(ex, $"studio.command.{cmd}"));
        }
    }

    private async Task ServeEmbedded(HttpContext ctx, string name, string contentType)
    {
        var asm = typeof(StudioHost).Assembly;
        // Embedded resource naming convention: <default_namespace>.<folder>.<name>
        var resourceName = $"XRai.Studio.wwwroot.{name}";
        using var stream = asm.GetManifestResourceStream(resourceName);
        if (stream == null)
        {
            ctx.Response.StatusCode = 404;
            await ctx.Response.WriteAsync($"Resource not found: {resourceName}");
            return;
        }

        ctx.Response.ContentType = contentType;
        ctx.Response.Headers["Cache-Control"] = "no-cache";
        await stream.CopyToAsync(ctx.Response.Body);
    }

    public void Dispose()
    {
        foreach (var d in _disposables)
        {
            try { d.Dispose(); } catch { }
        }
        _disposables.Clear();

        try { _host?.StopAsync(TimeSpan.FromSeconds(2)).GetAwaiter().GetResult(); } catch { }
        try { _host?.Dispose(); } catch { }
        _host = null;

        StudioToken.ClearStoredToken();
    }
}
