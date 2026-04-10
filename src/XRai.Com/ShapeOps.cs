using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class ShapeOps
{
    private readonly ExcelSession _session;
    public ShapeOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("shape.list", HandleList);
        router.Register("shape.add", HandleAdd);
        router.Register("shape.delete", HandleDelete);
        router.Register("shape.text", HandleText);
        router.Register("shape.move", HandleMove);
        router.Register("shape.resize", HandleResize);
        router.Register("shape.format", HandleFormat);
        router.Register("image.insert", HandleImageInsert);
        router.Register("image.delete", HandleImageDelete);
    }

    private string HandleList(JsonObject args)
    {
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var shapes = guard.Track(sheet.Shapes);
        var result = new JsonArray();
        for (int i = 1; i <= shapes.Count; i++)
        {
            var s = shapes.Item(i);
            result.Add(new JsonObject
            {
                ["name"] = s.Name,
                ["type"] = s.Type.ToString(),
                ["left"] = s.Left,
                ["top"] = s.Top,
                ["width"] = s.Width,
                ["height"] = s.Height,
            });
            Marshal.ReleaseComObject(s);
        }
        return Response.Ok(new { shapes = result, count = result.Count });
    }

    private string HandleAdd(JsonObject args)
    {
        var type = args["type"]?.GetValue<string>() ?? "rectangle";
        var left = args["left"]?.GetValue<float>() ?? 100;
        var top = args["top"]?.GetValue<float>() ?? 100;
        var width = args["width"]?.GetValue<float>() ?? 200;
        var height = args["height"]?.GetValue<float>() ?? 100;
        var text = args["text"]?.GetValue<string>();

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var shapes = guard.Track(sheet.Shapes);

        var autoType = type.ToLowerInvariant() switch
        {
            "rectangle" or "rect" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
            "oval" or "ellipse" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
            "rounded" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle,
            "callout" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangularCallout,
            "diamond" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond,
            "arrow" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightArrow,
            _ => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
        };

        dynamic shape;
        if (type == "textbox")
        {
            shape = shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
        }
        else if (type == "line")
        {
            shape = shapes.AddLine(left, top, left + width, top + height);
        }
        else
        {
            shape = shapes.AddShape(autoType, left, top, width, height);
        }

        if (text != null)
        {
            try { shape.TextFrame.Characters().Text = text; } catch { }
        }

        string name = shape.Name;
        Marshal.ReleaseComObject(shape);
        return Response.Ok(new { name, type, left, top, width, height, added = true });
    }

    private string HandleDelete(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>() ?? throw new ArgumentException("shape.delete requires 'name'");
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var shapes = guard.Track(sheet.Shapes);
        var shape = shapes.Item(name);
        shape.Delete();
        Marshal.ReleaseComObject(shape);
        return Response.Ok(new { name, deleted = true });
    }

    private string HandleText(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>() ?? throw new ArgumentException("shape.text requires 'name'");
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var shapes = guard.Track(sheet.Shapes);
        var shape = shapes.Item(name);

        if (args["text"] != null)
        {
            shape.TextFrame.Characters().Text = args["text"]!.GetValue<string>();
            var result = shape.TextFrame.Characters().Text;
            Marshal.ReleaseComObject(shape);
            return Response.Ok(new { name, text = result });
        }

        string currentText;
        try { currentText = shape.TextFrame.Characters().Text; }
        catch { currentText = ""; }
        Marshal.ReleaseComObject(shape);
        return Response.Ok(new { name, text = currentText });
    }

    private string HandleMove(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>() ?? throw new ArgumentException("shape.move requires 'name'");
        var left = args["left"]?.GetValue<float>() ?? 0;
        var top = args["top"]?.GetValue<float>() ?? 0;
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var shapes = guard.Track(sheet.Shapes);
        var shape = shapes.Item(name);
        shape.Left = left;
        shape.Top = top;
        Marshal.ReleaseComObject(shape);
        return Response.Ok(new { name, left, top, moved = true });
    }

    private string HandleResize(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>() ?? throw new ArgumentException("shape.resize requires 'name'");
        var width = args["width"]?.GetValue<float>() ?? 0;
        var height = args["height"]?.GetValue<float>() ?? 0;
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var shapes = guard.Track(sheet.Shapes);
        var shape = shapes.Item(name);
        if (width > 0) shape.Width = width;
        if (height > 0) shape.Height = height;
        Marshal.ReleaseComObject(shape);
        return Response.Ok(new { name, width, height, resized = true });
    }

    private string HandleFormat(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>() ?? throw new ArgumentException("shape.format requires 'name'");
        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var shapes = guard.Track(sheet.Shapes);
        var shape = shapes.Item(name);

        if (args["fill_color"] != null)
        {
            var color = ParseColor(args["fill_color"]!.GetValue<string>());
            shape.Fill.ForeColor.RGB = color;
        }
        if (args["line_color"] != null)
        {
            var color = ParseColor(args["line_color"]!.GetValue<string>());
            shape.Line.ForeColor.RGB = color;
        }
        if (args["line_weight"] != null)
        {
            shape.Line.Weight = args["line_weight"]!.GetValue<float>();
        }

        Marshal.ReleaseComObject(shape);
        return Response.Ok(new { name, formatted = true });
    }

    private string HandleImageInsert(JsonObject args)
    {
        var path = args["path"]?.GetValue<string>() ?? throw new ArgumentException("image.insert requires 'path'");
        var left = args["left"]?.GetValue<float>() ?? 100;
        var top = args["top"]?.GetValue<float>() ?? 100;
        var width = args["width"]?.GetValue<float>() ?? -1;
        var height = args["height"]?.GetValue<float>() ?? -1;

        if (!File.Exists(path))
            return Response.Error($"File not found: {path}");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var shapes = guard.Track(sheet.Shapes);
        var pic = shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoFalse,
            Microsoft.Office.Core.MsoTriState.msoTrue, left, top,
            width > 0 ? width : -1, height > 0 ? height : -1);

        string name = pic.Name;
        Marshal.ReleaseComObject(pic);
        return Response.Ok(new { name, path, inserted = true });
    }

    private string HandleImageDelete(JsonObject args)
    {
        return HandleDelete(args); // Same as shape.delete
    }

    private static int ParseColor(string color)
    {
        if (color.StartsWith("#") && color.Length == 7)
        {
            int r = Convert.ToInt32(color[1..3], 16);
            int g = Convert.ToInt32(color[3..5], 16);
            int b = Convert.ToInt32(color[5..7], 16);
            return r | (g << 8) | (b << 16);
        }
        return 0;
    }
}
