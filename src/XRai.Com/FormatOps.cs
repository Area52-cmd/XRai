using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class FormatOps
{
    private readonly ExcelSession _session;

    public FormatOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("format.border", HandleBorder);
        router.Register("format.align", HandleAlign);
        router.Register("format.font", HandleFont);
        router.Register("format.style", HandleStyle);
        router.Register("format.conditional", HandleConditional);
        router.Register("format.conditional.read", HandleConditionalRead);
        router.Register("format.conditional.clear", HandleConditionalClear);
        router.Register("format.style.list", HandleStyleList);
    }

    private string HandleBorder(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("format.border requires 'ref'");
            var side = args["side"]?.GetValue<string>() ?? "all";
            var weight = args["weight"]?.GetValue<string>() ?? "thin";
            var color = args["color"]?.GetValue<string>();
            var style = args["style"]?.GetValue<string>() ?? "continuous";

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);

            var xlWeight = ParseBorderWeight(weight);
            var xlStyle = ParseLineStyle(style);
            var borders = GetBorderIndices(side);

            foreach (var borderIndex in borders)
            {
                var border = guard.Track(range.Borders[borderIndex]);
                border.LineStyle = xlStyle;
                if (xlStyle != Excel.XlLineStyle.xlLineStyleNone)
                {
                    border.Weight = xlWeight;
                    if (color != null)
                        border.Color = ParseColor(color);
                }
            }

            return Response.Ok(new { @ref = refStr, side, weight, style, bordered = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"format.border: {ex.Message}");
        }
    }

    private string HandleAlign(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("format.align requires 'ref'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);

            if (args["horizontal"] is JsonNode hNode)
            {
                range.HorizontalAlignment = hNode.GetValue<string>().ToLowerInvariant() switch
                {
                    "left" => Excel.XlHAlign.xlHAlignLeft,
                    "center" => Excel.XlHAlign.xlHAlignCenter,
                    "right" => Excel.XlHAlign.xlHAlignRight,
                    "justify" => Excel.XlHAlign.xlHAlignJustify,
                    _ => Excel.XlHAlign.xlHAlignGeneral,
                };
            }

            if (args["vertical"] is JsonNode vNode)
            {
                range.VerticalAlignment = vNode.GetValue<string>().ToLowerInvariant() switch
                {
                    "top" => Excel.XlVAlign.xlVAlignTop,
                    "center" => Excel.XlVAlign.xlVAlignCenter,
                    "bottom" => Excel.XlVAlign.xlVAlignBottom,
                    _ => Excel.XlVAlign.xlVAlignBottom,
                };
            }

            if (args["wrap_text"] is JsonNode wrapNode)
                range.WrapText = wrapNode.GetValue<bool>();

            if (args["shrink_to_fit"] is JsonNode shrinkNode)
                range.ShrinkToFit = shrinkNode.GetValue<bool>();

            if (args["text_rotation"] is JsonNode rotNode)
                range.Orientation = rotNode.GetValue<int>();

            if (args["indent"] is JsonNode indentNode)
                range.IndentLevel = indentNode.GetValue<int>();

            return Response.Ok(new { @ref = refStr, aligned = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"format.align: {ex.Message}");
        }
    }

    private string HandleFont(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("format.font requires 'ref'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);
            var font = guard.Track(range.Font);

            if (args["italic"] is JsonNode italicNode)
                font.Italic = italicNode.GetValue<bool>();

            if (args["underline"] is JsonNode underlineNode)
            {
                // Support bool or string (single/double)
                var ulStr = underlineNode.ToString().ToLowerInvariant();
                font.Underline = ulStr switch
                {
                    "true" or "single" => Excel.XlUnderlineStyle.xlUnderlineStyleSingle,
                    "double" => Excel.XlUnderlineStyle.xlUnderlineStyleDouble,
                    _ => Excel.XlUnderlineStyle.xlUnderlineStyleNone,
                };
            }

            if (args["strikethrough"] is JsonNode strikeNode)
                font.Strikethrough = strikeNode.GetValue<bool>();

            if (args["color"] is JsonNode colorNode)
                font.Color = ParseColor(colorNode.GetValue<string>());

            if (args["subscript"] is JsonNode subNode)
                font.Subscript = subNode.GetValue<bool>();

            if (args["superscript"] is JsonNode supNode)
                font.Superscript = supNode.GetValue<bool>();

            return Response.Ok(new { @ref = refStr, font_set = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"format.font: {ex.Message}");
        }
    }

    private string HandleStyle(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("format.style requires 'ref'");
            var styleName = args["style"]?.GetValue<string>()
                ?? throw new ArgumentException("format.style requires 'style'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);
            range.Style = styleName;

            return Response.Ok(new { @ref = refStr, style = styleName, applied = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"format.style: {ex.Message}");
        }
    }

    private string HandleConditional(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("format.conditional requires 'ref'");
            var type = args["type"]?.GetValue<string>() ?? "cell_value";

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);
            var fcs = guard.Track(range.FormatConditions);

            switch (type.ToLowerInvariant())
            {
                case "cell_value":
                {
                    var op = args["operator"]?.GetValue<string>() ?? ">";
                    var value = args["value"]?.ToString()
                        ?? throw new ArgumentException("format.conditional requires 'value' for cell_value type");

                    var xlOp = ParseConditionOperator(op);
                    var value2 = args["value2"]?.ToString();

                    Excel.FormatCondition fc;
                    if (xlOp == Excel.XlFormatConditionOperator.xlBetween && value2 != null)
                        fc = (Excel.FormatCondition)fcs.Add(
                            Excel.XlFormatConditionType.xlCellValue, xlOp, value, value2);
                    else
                        fc = (Excel.FormatCondition)fcs.Add(
                            Excel.XlFormatConditionType.xlCellValue, xlOp, value);

                    ApplyConditionalFormat(fc, args, guard);
                    Marshal.ReleaseComObject(fc);
                    break;
                }
                case "text_contains":
                {
                    var value = args["value"]?.GetValue<string>()
                        ?? throw new ArgumentException("format.conditional requires 'value' for text_contains type");

                    // Use xlTextString with the formula approach
                    var formula = $"=NOT(ISERROR(SEARCH(\"{value}\",{range.Cells[1, 1].Address[false, false]})))";
                    var fc = (Excel.FormatCondition)fcs.Add(
                        Excel.XlFormatConditionType.xlExpression, Formula1: formula);
                    ApplyConditionalFormat(fc, args, guard);
                    Marshal.ReleaseComObject(fc);
                    break;
                }
                case "top_n":
                {
                    var value = args["value"]?.GetValue<int>() ?? 10;
                    dynamic fc = fcs.AddTop10();
                    fc.TopBottom = Excel.XlTopBottom.xlTop10Top;
                    fc.Rank = value;
                    if (args["format"] is JsonNode fmtNode)
                    {
                        fc.Interior.Color = ParseColor(fmtNode.GetValue<string>());
                    }
                    Marshal.ReleaseComObject((object)fc);
                    break;
                }
                case "color_scale":
                {
                    var fc = (Excel.ColorScale)fcs.AddColorScale(2);
                    // Default: red to green
                    fc.ColorScaleCriteria[1].FormatColor.Color = ParseColor("#F8696B");
                    fc.ColorScaleCriteria[2].FormatColor.Color = ParseColor("#63BE7B");
                    Marshal.ReleaseComObject(fc);
                    break;
                }
                case "data_bar":
                {
                    var fc = (Excel.Databar)fcs.AddDatabar();
                    if (args["format"] is JsonNode barColor)
                        fc.BarColor.Color = ParseColor(barColor.GetValue<string>());
                    Marshal.ReleaseComObject(fc);
                    break;
                }
                default:
                    return Response.Error($"format.conditional: unsupported type '{type}'");
            }

            return Response.Ok(new { @ref = refStr, type, conditional_added = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"format.conditional: {ex.Message}");
        }
    }

    private string HandleConditionalClear(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("format.conditional.clear requires 'ref'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);
            var fcs = guard.Track(range.FormatConditions);
            fcs.Delete();

            return Response.Ok(new { @ref = refStr, conditional_cleared = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"format.conditional.clear: {ex.Message}");
        }
    }

    private string HandleConditionalRead(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("format.conditional.read requires 'ref'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);
            var fcs = guard.Track(range.FormatConditions);

            var rules = new JsonArray();
            for (int i = 1; i <= fcs.Count; i++)
            {
                try
                {
                    dynamic fc = fcs[i];
                    var rule = new JsonObject();

                    try { rule["priority"] = (int)fc.Priority; } catch { }
                    try { rule["applies_to"] = ((Excel.Range)fc.AppliesTo).Address[false, false]; } catch { }

                    var fcType = (Excel.XlFormatConditionType)fc.Type;
                    rule["type"] = fcType.ToString();

                    switch (fcType)
                    {
                        case Excel.XlFormatConditionType.xlCellValue:
                        case Excel.XlFormatConditionType.xlExpression:
                        {
                            try { rule["formula1"] = (string)fc.Formula1; } catch { }
                            try { rule["formula2"] = (string)fc.Formula2; } catch { }
                            try { rule["operator"] = ((Excel.XlFormatConditionOperator)fc.Operator).ToString(); } catch { }
                            try
                            {
                                var fmt = new JsonObject();
                                try { fmt["bold"] = (bool)fc.Font.Bold; } catch { }
                                try { fmt["font_color"] = ColorToHex((int)fc.Font.Color); } catch { }
                                try { fmt["bg_color"] = ColorToHex((int)fc.Interior.Color); } catch { }
                                rule["format"] = fmt;
                            }
                            catch { }
                            break;
                        }
                        case Excel.XlFormatConditionType.xlColorScale:
                        {
                            rule["type"] = "ColorScale";
                            break;
                        }
                        case Excel.XlFormatConditionType.xlDatabar:
                        {
                            rule["type"] = "DataBar";
                            break;
                        }
                        case (Excel.XlFormatConditionType)6: // xlIconSet
                        {
                            rule["type"] = "IconSet";
                            break;
                        }
                        case Excel.XlFormatConditionType.xlTop10:
                        {
                            rule["type"] = "Top10";
                            try { rule["rank"] = (int)fc.Rank; } catch { }
                            break;
                        }
                    }

                    rules.Add(rule);
                    Marshal.ReleaseComObject((object)fc);
                }
                catch
                {
                    // Skip rules that cannot be read
                }
            }

            return Response.Ok(new { @ref = refStr, rules, count = rules.Count });
        }
        catch (Exception ex)
        {
            return Response.Error($"format.conditional.read: {ex.Message}");
        }
    }

    private string HandleStyleList(JsonObject args)
    {
        try
        {
            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());
            var styles = guard.Track(wb.Styles);

            var result = new JsonArray();
            int limit = 200;
            foreach (Excel.Style style in styles)
            {
                if (result.Count >= limit)
                {
                    Marshal.ReleaseComObject(style);
                    break;
                }

                var entry = new JsonObject
                {
                    ["name"] = style.Name,
                    ["built_in"] = style.BuiltIn,
                };
                try { entry["name_local"] = style.NameLocal; } catch { }
                try { entry["number_format"] = style.NumberFormat; } catch { }

                result.Add(entry);
                Marshal.ReleaseComObject(style);
            }

            return Response.Ok(new { styles = result, count = result.Count });
        }
        catch (Exception ex)
        {
            return Response.Error($"format.style.list: {ex.Message}");
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static void ApplyConditionalFormat(Excel.FormatCondition fc, JsonObject args, ComGuard guard)
    {
        if (args["format"] is JsonNode fmtNode)
        {
            var interior = guard.Track(fc.Interior);
            interior.Color = ParseColor(fmtNode.GetValue<string>());
        }
    }

    private static Excel.XlBordersIndex[] GetBorderIndices(string side)
    {
        return side.ToLowerInvariant() switch
        {
            "top" => [Excel.XlBordersIndex.xlEdgeTop],
            "bottom" => [Excel.XlBordersIndex.xlEdgeBottom],
            "left" => [Excel.XlBordersIndex.xlEdgeLeft],
            "right" => [Excel.XlBordersIndex.xlEdgeRight],
            _ => [
                Excel.XlBordersIndex.xlEdgeTop,
                Excel.XlBordersIndex.xlEdgeBottom,
                Excel.XlBordersIndex.xlEdgeLeft,
                Excel.XlBordersIndex.xlEdgeRight,
            ],
        };
    }

    private static Excel.XlBorderWeight ParseBorderWeight(string weight)
    {
        return weight.ToLowerInvariant() switch
        {
            "thin" => Excel.XlBorderWeight.xlThin,
            "medium" => Excel.XlBorderWeight.xlMedium,
            "thick" => Excel.XlBorderWeight.xlThick,
            "hairline" => Excel.XlBorderWeight.xlHairline,
            _ => Excel.XlBorderWeight.xlThin,
        };
    }

    private static Excel.XlLineStyle ParseLineStyle(string style)
    {
        return style.ToLowerInvariant() switch
        {
            "continuous" => Excel.XlLineStyle.xlContinuous,
            "dash" => Excel.XlLineStyle.xlDash,
            "dot" => Excel.XlLineStyle.xlDot,
            "none" => Excel.XlLineStyle.xlLineStyleNone,
            _ => Excel.XlLineStyle.xlContinuous,
        };
    }

    private static Excel.XlFormatConditionOperator ParseConditionOperator(string op)
    {
        return op switch
        {
            ">" or "greater" => Excel.XlFormatConditionOperator.xlGreater,
            "<" or "less" => Excel.XlFormatConditionOperator.xlLess,
            "=" or "equal" => Excel.XlFormatConditionOperator.xlEqual,
            ">=" or "greater_equal" => Excel.XlFormatConditionOperator.xlGreaterEqual,
            "<=" or "less_equal" => Excel.XlFormatConditionOperator.xlLessEqual,
            "!=" or "not_equal" => Excel.XlFormatConditionOperator.xlNotEqual,
            "between" => Excel.XlFormatConditionOperator.xlBetween,
            _ => Excel.XlFormatConditionOperator.xlGreater,
        };
    }

    private static int ParseColor(string color)
    {
        if (color.StartsWith("#") && color.Length == 7)
        {
            int r = Convert.ToInt32(color[1..3], 16);
            int g = Convert.ToInt32(color[3..5], 16);
            int b = Convert.ToInt32(color[5..7], 16);
            return r | (g << 8) | (b << 16); // OLE color is BGR
        }

        return color.ToLowerInvariant() switch
        {
            "yellow" => 65535,
            "red" => 255,
            "green" => 65280,
            "blue" => 16711680,
            "white" => 16777215,
            _ => 0,
        };
    }

    private static string ColorToHex(int oleColor)
    {
        int r = oleColor & 0xFF;
        int g = (oleColor >> 8) & 0xFF;
        int b = (oleColor >> 16) & 0xFF;
        return $"#{r:X2}{g:X2}{b:X2}";
    }
}
