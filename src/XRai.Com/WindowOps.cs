using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class WindowOps
{
    private readonly ExcelSession _session;
    public WindowOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("window.zoom", HandleZoom);
        router.Register("window.scroll", HandleScroll);
        router.Register("window.split", HandleSplit);
        router.Register("window.view", HandleView);
        router.Register("window.gridlines", HandleGridlines);
        router.Register("window.headings", HandleHeadings);
        router.Register("window.statusbar", HandleStatusbar);
        router.Register("window.fullscreen", HandleFullscreen);
    }

    private string HandleZoom(JsonObject args)
    {
        using var guard = new ComGuard();
        var win = guard.Track(_session.App.ActiveWindow);
        if (args["level"] != null)
        {
            var level = args["level"]!.GetValue<int>();
            win.Zoom = level;
            return Response.Ok(new { zoom = level });
        }
        return Response.Ok(new { zoom = (int)win.Zoom });
    }

    private string HandleScroll(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>() ?? throw new ArgumentException("window.scroll requires 'ref'");
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        _session.App.Goto(range, true);
        return Response.Ok(new { scrolled_to = refStr });
    }

    private string HandleSplit(JsonObject args)
    {
        using var guard = new ComGuard();
        var win = guard.Track(_session.App.ActiveWindow);
        if (args["ref"] != null)
        {
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[args["ref"]!.GetValue<string>()]);
            range.Select();
            win.Split = true;
            return Response.Ok(new { split = true, at = args["ref"]!.GetValue<string>() });
        }
        win.Split = false;
        return Response.Ok(new { split = false });
    }

    private string HandleView(JsonObject args)
    {
        using var guard = new ComGuard();
        var win = guard.Track(_session.App.ActiveWindow);
        if (args["mode"] != null)
        {
            var mode = args["mode"]!.GetValue<string>();
            win.View = mode.ToLowerInvariant() switch
            {
                "normal" => Excel.XlWindowView.xlNormalView,
                "pagebreak" => Excel.XlWindowView.xlPageBreakPreview,
                "pagelayout" => Excel.XlWindowView.xlPageLayoutView,
                _ => throw new ArgumentException($"Unknown view mode: {mode}")
            };
            return Response.Ok(new { view = mode });
        }
        var current = win.View switch
        {
            Excel.XlWindowView.xlNormalView => "normal",
            Excel.XlWindowView.xlPageBreakPreview => "pagebreak",
            Excel.XlWindowView.xlPageLayoutView => "pagelayout",
            _ => "unknown"
        };
        return Response.Ok(new { view = current });
    }

    private string HandleGridlines(JsonObject args)
    {
        using var guard = new ComGuard();
        var win = guard.Track(_session.App.ActiveWindow);
        if (args["show"] != null)
        {
            win.DisplayGridlines = args["show"]!.GetValue<bool>();
            return Response.Ok(new { gridlines = win.DisplayGridlines });
        }
        return Response.Ok(new { gridlines = win.DisplayGridlines });
    }

    private string HandleHeadings(JsonObject args)
    {
        using var guard = new ComGuard();
        var win = guard.Track(_session.App.ActiveWindow);
        if (args["show"] != null)
        {
            win.DisplayHeadings = args["show"]!.GetValue<bool>();
            return Response.Ok(new { headings = win.DisplayHeadings });
        }
        return Response.Ok(new { headings = win.DisplayHeadings });
    }

    private string HandleStatusbar(JsonObject args)
    {
        var text = _session.App.StatusBar;
        return Response.Ok(new { statusbar = text?.ToString() ?? "" });
    }

    private string HandleFullscreen(JsonObject args)
    {
        if (args["on"] != null)
        {
            _session.App.DisplayFullScreen = args["on"]!.GetValue<bool>();
            return Response.Ok(new { fullscreen = _session.App.DisplayFullScreen });
        }
        return Response.Ok(new { fullscreen = _session.App.DisplayFullScreen });
    }
}
