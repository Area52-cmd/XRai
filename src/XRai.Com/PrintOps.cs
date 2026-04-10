using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class PrintOps
{
    private readonly ExcelSession _session;

    public PrintOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("print.setup", HandlePrintSetup);
        router.Register("print.margins", HandlePrintMargins);
        router.Register("print.area", HandlePrintArea);
        router.Register("print.area.clear", HandlePrintAreaClear);
        router.Register("print.titles", HandlePrintTitles);
        router.Register("print.headers", HandlePrintHeaders);
        router.Register("print.gridlines", HandlePrintGridlines);
        router.Register("print.breaks", HandlePrintBreaks);
        router.Register("print.preview", HandlePrintPreview);
    }

    private string HandlePrintSetup(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var ps = guard.Track(sheet.PageSetup);

        if (args["orientation"] != null)
        {
            ps.Orientation = args["orientation"]!.GetValue<string>().ToLowerInvariant() switch
            {
                "landscape" => Excel.XlPageOrientation.xlLandscape,
                _ => Excel.XlPageOrientation.xlPortrait,
            };
        }

        if (args["paper_size"] != null)
        {
            ps.PaperSize = args["paper_size"]!.GetValue<string>().ToLowerInvariant() switch
            {
                "a4" => Excel.XlPaperSize.xlPaperA4,
                "legal" => Excel.XlPaperSize.xlPaperLegal,
                _ => Excel.XlPaperSize.xlPaperLetter,
            };
        }

        if (args["scale"] != null)
        {
            var scale = args["scale"]!.GetValue<int>();
            if (scale < 1 || scale > 400)
                throw new ArgumentException("scale must be between 1 and 400");
            ps.Zoom = scale;
        }

        if (args["fit_wide"] != null)
        {
            ps.Zoom = false;
            ps.FitToPagesWide = args["fit_wide"]!.GetValue<int>();
        }

        if (args["fit_tall"] != null)
        {
            ps.Zoom = false;
            ps.FitToPagesTall = args["fit_tall"]!.GetValue<int>();
        }

        return Response.Ok(new
        {
            orientation = ps.Orientation == Excel.XlPageOrientation.xlLandscape ? "landscape" : "portrait",
            zoom = ps.Zoom is false ? (object?)null : Convert.ToInt32(ps.Zoom),
            fit_wide = Convert.ToInt32(ps.FitToPagesWide),
            fit_tall = Convert.ToInt32(ps.FitToPagesTall),
            setup = true,
        });
    }

    private string HandlePrintMargins(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var ps = guard.Track(sheet.PageSetup);

        // Excel PageSetup margins are in points (1 inch = 72 points)
        const double pointsPerInch = 72.0;

        if (args["top"] != null)
            ps.TopMargin = args["top"]!.GetValue<double>() * pointsPerInch;
        if (args["bottom"] != null)
            ps.BottomMargin = args["bottom"]!.GetValue<double>() * pointsPerInch;
        if (args["left"] != null)
            ps.LeftMargin = args["left"]!.GetValue<double>() * pointsPerInch;
        if (args["right"] != null)
            ps.RightMargin = args["right"]!.GetValue<double>() * pointsPerInch;
        if (args["header"] != null)
            ps.HeaderMargin = args["header"]!.GetValue<double>() * pointsPerInch;
        if (args["footer"] != null)
            ps.FooterMargin = args["footer"]!.GetValue<double>() * pointsPerInch;

        return Response.Ok(new
        {
            top = Math.Round(ps.TopMargin / pointsPerInch, 3),
            bottom = Math.Round(ps.BottomMargin / pointsPerInch, 3),
            left = Math.Round(ps.LeftMargin / pointsPerInch, 3),
            right = Math.Round(ps.RightMargin / pointsPerInch, 3),
            header = Math.Round(ps.HeaderMargin / pointsPerInch, 3),
            footer = Math.Round(ps.FooterMargin / pointsPerInch, 3),
            margins_set = true,
        });
    }

    private string HandlePrintArea(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var ps = guard.Track(sheet.PageSetup);

        if (args["ref"] != null)
        {
            var refStr = args["ref"]!.GetValue<string>();
            ps.PrintArea = refStr;
            return Response.Ok(new { print_area = refStr, set = true });
        }

        // Get current print area
        var current = ps.PrintArea;
        if (string.IsNullOrEmpty(current))
            return Response.Ok(new { print_area = (string?)null });

        return Response.Ok(new { print_area = current });
    }

    private string HandlePrintAreaClear(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var ps = guard.Track(sheet.PageSetup);
        ps.PrintArea = "";

        return Response.Ok(new { print_area_cleared = true });
    }

    private string HandlePrintTitles(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var ps = guard.Track(sheet.PageSetup);

        if (args["rows"] != null)
            ps.PrintTitleRows = "$" + args["rows"]!.GetValue<string>().Replace(":", ":$");
        if (args["columns"] != null)
            ps.PrintTitleColumns = "$" + args["columns"]!.GetValue<string>().Replace(":", ":$");

        return Response.Ok(new
        {
            print_title_rows = ps.PrintTitleRows,
            print_title_columns = ps.PrintTitleColumns,
            titles_set = true,
        });
    }

    private string HandlePrintHeaders(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var ps = guard.Track(sheet.PageSetup);

        // Header sections
        if (args["left"] != null)
            ps.LeftHeader = args["left"]!.GetValue<string>();
        if (args["center"] != null)
            ps.CenterHeader = args["center"]!.GetValue<string>();
        if (args["right"] != null)
            ps.RightHeader = args["right"]!.GetValue<string>();

        // Footer sections
        if (args["footer_left"] != null)
            ps.LeftFooter = args["footer_left"]!.GetValue<string>();
        if (args["footer_center"] != null)
            ps.CenterFooter = args["footer_center"]!.GetValue<string>();
        if (args["footer_right"] != null)
            ps.RightFooter = args["footer_right"]!.GetValue<string>();

        return Response.Ok(new
        {
            header_left = ps.LeftHeader,
            header_center = ps.CenterHeader,
            header_right = ps.RightHeader,
            footer_left = ps.LeftFooter,
            footer_center = ps.CenterFooter,
            footer_right = ps.RightFooter,
            headers_set = true,
        });
    }

    private string HandlePrintGridlines(JsonObject args)
    {
        var show = args["show"]?.GetValue<bool>()
            ?? throw new ArgumentException("print.gridlines requires 'show' (bool)");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var ps = guard.Track(sheet.PageSetup);
        ps.PrintGridlines = show;

        return Response.Ok(new { gridlines = show });
    }

    private string HandlePrintBreaks(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());

        var action = args["action"]?.GetValue<string>();

        if (action != null && action.Equals("clear", StringComparison.OrdinalIgnoreCase))
        {
            sheet.ResetAllPageBreaks();
            return Response.Ok(new { page_breaks_cleared = true });
        }

        if (action != null && action.Equals("list", StringComparison.OrdinalIgnoreCase))
        {
            var hBreaks = guard.Track(sheet.HPageBreaks);
            var vBreaks = guard.Track(sheet.VPageBreaks);

            var breaks = new JsonArray();
            for (int i = 1; i <= hBreaks.Count; i++)
            {
                var hb = hBreaks[i];
                var loc = hb.Location;
                breaks.Add(new JsonObject
                {
                    ["type"] = "row",
                    ["location"] = loc.Address[false, false],
                });
                Marshal.ReleaseComObject(loc);
                Marshal.ReleaseComObject(hb);
            }
            for (int i = 1; i <= vBreaks.Count; i++)
            {
                var vb = vBreaks[i];
                var loc = vb.Location;
                breaks.Add(new JsonObject
                {
                    ["type"] = "column",
                    ["location"] = loc.Address[false, false],
                });
                Marshal.ReleaseComObject(loc);
                Marshal.ReleaseComObject(vb);
            }

            return Response.Ok(new { page_breaks = breaks });
        }

        // Insert a page break
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("print.breaks requires 'ref' or 'action'");
        var type = args["type"]?.GetValue<string>() ?? "row";

        var range = guard.Track(sheet.Range[refStr]);

        if (type.Equals("column", StringComparison.OrdinalIgnoreCase) ||
            type.Equals("col", StringComparison.OrdinalIgnoreCase))
        {
            var vBreaksCol = guard.Track(sheet.VPageBreaks);
            vBreaksCol.Add(range);
        }
        else
        {
            var hBreaksRow = guard.Track(sheet.HPageBreaks);
            hBreaksRow.Add(range);
        }

        return Response.Ok(new { @ref = refStr, page_break = type, inserted = true });
    }

    private string HandlePrintPreview(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var ps = guard.Track(sheet.PageSetup);
        var hBreaks = guard.Track(sheet.HPageBreaks);
        var vBreaks = guard.Track(sheet.VPageBreaks);

        // Estimate page count by triggering a calculation
        // Accessing HPageBreaks.Count after setting print view gives page info
        var pageCount = (hBreaks.Count + 1) * (vBreaks.Count + 1);

        var orientation = ps.Orientation == Excel.XlPageOrientation.xlLandscape
            ? "landscape" : "portrait";

        var printArea = ps.PrintArea;

        return Response.Ok(new
        {
            estimated_pages = pageCount,
            orientation,
            print_area = string.IsNullOrEmpty(printArea) ? null : printArea,
            zoom = ps.Zoom is false ? (object?)null : Convert.ToInt32(ps.Zoom),
            fit_wide = Convert.ToInt32(ps.FitToPagesWide),
            fit_tall = Convert.ToInt32(ps.FitToPagesTall),
            gridlines = ps.PrintGridlines,
            print_title_rows = string.IsNullOrEmpty(ps.PrintTitleRows) ? null : ps.PrintTitleRows,
            print_title_columns = string.IsNullOrEmpty(ps.PrintTitleColumns) ? null : ps.PrintTitleColumns,
            horizontal_breaks = hBreaks.Count,
            vertical_breaks = vBreaks.Count,
        });
    }
}
