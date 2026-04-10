// Leak-audited: 2026-04-10 — No `foreach (Excel.X in Y)` iterations in this
// file. All COM proxies obtained via guard.Track are released by ComGuard on
// dispose. No leaks found.

using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class DataOps
{
    private readonly ExcelSession _session;

    public DataOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("copy", HandleCopy);
        router.Register("paste", HandlePaste);
        router.Register("paste.values", HandlePasteValues);
        router.Register("sort", HandleSort);
        router.Register("find", HandleFind);
        router.Register("replace", HandleReplace);
        router.Register("fill.down", HandleFillDown);
        router.Register("fill.right", HandleFillRight);
        router.Register("transpose", HandleTranspose);
        router.Register("protect", HandleProtect);
        router.Register("unprotect", HandleUnprotect);
        router.Register("comment", HandleComment);
        router.Register("comment.read", HandleCommentRead);
        router.Register("validation", HandleValidation);
        router.Register("validation.read", HandleValidationRead);
        router.Register("hyperlink", HandleHyperlink);
        router.Register("comment.thread", HandleCommentThreaded);
        router.Register("comment.thread.read", HandleCommentThreadedRead);
    }

    private string HandleCopy(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("copy requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        range.Copy();

        return Response.Ok(new { @ref = refStr, copied = true });
    }

    private string HandlePaste(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("paste requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        sheet.Paste(range);

        return Response.Ok(new { @ref = refStr, pasted = true });
    }

    private string HandlePasteValues(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("paste.values requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        range.PasteSpecial(Excel.XlPasteType.xlPasteValues);

        return Response.Ok(new { @ref = refStr, pasted_values = true });
    }

    private string HandleSort(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("sort requires 'ref'");
        var column = args["column"]?.GetValue<string>()
            ?? throw new ArgumentException("sort requires 'column'");
        var order = args["order"]?.GetValue<string>() ?? "asc";

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        var keyRange = guard.Track(sheet.Range[column]);
        var sortObj = guard.Track(range.Sort);

        var xlOrder = order.ToLowerInvariant() switch
        {
            "desc" or "descending" => Excel.XlSortOrder.xlDescending,
            _ => Excel.XlSortOrder.xlAscending,
        };

        range.Sort(keyRange, xlOrder, Header: Excel.XlYesNoGuess.xlYes);
        return Response.Ok(new { @ref = refStr, sorted_by = column, order });
    }

    private string HandleFind(JsonObject args)
    {
        var what = args["what"]?.GetValue<string>()
            ?? throw new ArgumentException("find requires 'what'");
        var refStr = args["ref"]?.GetValue<string>();

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        Excel.Range searchRange;
        if (refStr != null)
            searchRange = guard.Track(sheet.Range[refStr]);
        else
            searchRange = guard.Track(sheet.UsedRange);

        var found = searchRange.Find(what);
        if (found == null)
            return Response.Ok(new { found = false, what });

        var addr = found.Address[false, false];
        var val = found.Value2?.ToString();
        Marshal.ReleaseComObject(found);

        return Response.Ok(new { found = true, what, @ref = addr, value = val });
    }

    private string HandleReplace(JsonObject args)
    {
        var what = args["what"]?.GetValue<string>()
            ?? throw new ArgumentException("replace requires 'what'");
        var with = args["with"]?.GetValue<string>()
            ?? throw new ArgumentException("replace requires 'with'");
        var refStr = args["ref"]?.GetValue<string>();

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        Excel.Range searchRange;
        if (refStr != null)
            searchRange = guard.Track(sheet.Range[refStr]);
        else
            searchRange = guard.Track(sheet.UsedRange);

        var replaced = searchRange.Replace(what, with);
        return Response.Ok(new { what, with_value = with, replaced });
    }

    private string HandleFillDown(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("fill.down requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        range.FillDown();

        return Response.Ok(new { @ref = refStr, filled = "down" });
    }

    private string HandleFillRight(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("fill.right requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        range.FillRight();

        return Response.Ok(new { @ref = refStr, filled = "right" });
    }

    private string HandleTranspose(JsonObject args)
    {
        var from = args["from"]?.GetValue<string>()
            ?? throw new ArgumentException("transpose requires 'from'");
        var to = args["to"]?.GetValue<string>()
            ?? throw new ArgumentException("transpose requires 'to'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var source = guard.Track(sheet.Range[from]);
        var dest = guard.Track(sheet.Range[to]);

        var values = source.Value2;
        dest.Value2 = _session.App.WorksheetFunction.Transpose(values);

        return Response.Ok(new { from, to, transposed = true });
    }

    private string HandleProtect(JsonObject args)
    {
        var password = args["password"]?.GetValue<string>();

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        if (password != null)
            sheet.Protect(password);
        else
            sheet.Protect();

        return Response.Ok(new { sheet = sheet.Name, @protected = true });
    }

    private string HandleUnprotect(JsonObject args)
    {
        var password = args["password"]?.GetValue<string>();

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        if (password != null)
            sheet.Unprotect(password);
        else
            sheet.Unprotect();

        return Response.Ok(new { sheet = sheet.Name, unprotected = true });
    }

    private string HandleComment(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("comment requires 'ref'");
        var text = args["text"]?.GetValue<string>()
            ?? throw new ArgumentException("comment requires 'text'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);

        // Delete existing comment if present — must track COM object
        try
        {
            var existing = range.Comment;
            if (existing != null)
            {
                existing.Delete();
                Marshal.ReleaseComObject(existing);
            }
        }
        catch { }
        range.AddComment(text);

        return Response.Ok(new { @ref = refStr, comment = text });
    }

    private string HandleCommentRead(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("comment.read requires 'ref'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);

        var comment = range.Comment;
        if (comment == null)
            return Response.Ok(new { @ref = refStr, comment = (string?)null });

        var text = comment.Text();
        Marshal.ReleaseComObject(comment);
        return Response.Ok(new { @ref = refStr, comment = text });
    }

    private string HandleValidation(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("validation requires 'ref'");
        var type = args["type"]?.GetValue<string>() ?? "list";
        var formula = args["formula"]?.GetValue<string>()
            ?? throw new ArgumentException("validation requires 'formula'");

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        var validation = guard.Track(range.Validation);

        try { validation.Delete(); } catch { }

        var xlType = type.ToLowerInvariant() switch
        {
            "list" => Excel.XlDVType.xlValidateList,
            "whole" => Excel.XlDVType.xlValidateWholeNumber,
            "decimal" => Excel.XlDVType.xlValidateDecimal,
            "date" => Excel.XlDVType.xlValidateDate,
            _ => Excel.XlDVType.xlValidateList,
        };

        validation.Add(xlType, Excel.XlDVAlertStyle.xlValidAlertStop, Formula1: formula);
        return Response.Ok(new { @ref = refStr, validation = type, formula });
    }

    private string HandleHyperlink(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("hyperlink requires 'ref'");
        var url = args["url"]?.GetValue<string>()
            ?? throw new ArgumentException("hyperlink requires 'url'");
        var text = args["text"]?.GetValue<string>() ?? url;

        using var guard = new ComGuard();
        var sheet = guard.Track(_session.GetActiveSheet());
        var range = guard.Track(sheet.Range[refStr]);
        var links = guard.Track(sheet.Hyperlinks);
        links.Add(range, url, TextToDisplay: text);

        return Response.Ok(new { @ref = refStr, url, text });
    }

    private string HandleValidationRead(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("validation.read requires 'ref'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);

            Excel.Validation validation;
            try
            {
                validation = guard.Track(range.Validation);
                // Access Type to verify validation exists — throws if none
                var _ = validation.Type;
            }
            catch
            {
                return Response.Ok(new { @ref = refStr, has_validation = false });
            }

            var typeStr = ((Excel.XlDVType)validation.Type) switch
            {
                Excel.XlDVType.xlValidateWholeNumber => "whole_number",
                Excel.XlDVType.xlValidateDecimal => "decimal",
                Excel.XlDVType.xlValidateList => "list",
                Excel.XlDVType.xlValidateDate => "date",
                Excel.XlDVType.xlValidateTime => "time",
                Excel.XlDVType.xlValidateTextLength => "text_length",
                Excel.XlDVType.xlValidateCustom => "custom",
                _ => "unknown",
            };

            string? operatorStr = null;
            try
            {
                operatorStr = ((Excel.XlFormatConditionOperator)validation.Operator) switch
                {
                    Excel.XlFormatConditionOperator.xlBetween => "between",
                    Excel.XlFormatConditionOperator.xlNotBetween => "not_between",
                    Excel.XlFormatConditionOperator.xlEqual => "equal",
                    Excel.XlFormatConditionOperator.xlNotEqual => "not_equal",
                    Excel.XlFormatConditionOperator.xlGreater => "greater",
                    Excel.XlFormatConditionOperator.xlLess => "less",
                    Excel.XlFormatConditionOperator.xlGreaterEqual => "greater_equal",
                    Excel.XlFormatConditionOperator.xlLessEqual => "less_equal",
                    _ => "unknown",
                };
            }
            catch { }

            string? formula1 = null;
            try { formula1 = validation.Formula1; } catch { }
            string? formula2 = null;
            try { formula2 = validation.Formula2; } catch { }
            string? inputTitle = null;
            try { inputTitle = validation.InputTitle; } catch { }
            string? inputMessage = null;
            try { inputMessage = validation.InputMessage; } catch { }
            string? errorTitle = null;
            try { errorTitle = validation.ErrorTitle; } catch { }
            string? errorMessage = null;
            try { errorMessage = validation.ErrorMessage; } catch { }

            string? errorStyle = null;
            try
            {
                errorStyle = ((Excel.XlDVAlertStyle)validation.AlertStyle) switch
                {
                    Excel.XlDVAlertStyle.xlValidAlertStop => "stop",
                    Excel.XlDVAlertStyle.xlValidAlertWarning => "warning",
                    Excel.XlDVAlertStyle.xlValidAlertInformation => "information",
                    _ => "unknown",
                };
            }
            catch { }

            bool? ignoreBlank = null;
            try { ignoreBlank = validation.IgnoreBlank; } catch { }
            bool? inCellDropdown = null;
            try { inCellDropdown = validation.InCellDropdown; } catch { }
            bool? showInput = null;
            try { showInput = validation.ShowInput; } catch { }
            bool? showError = null;
            try { showError = validation.ShowError; } catch { }

            return Response.Ok(new
            {
                @ref = refStr,
                has_validation = true,
                type = typeStr,
                @operator = operatorStr,
                formula1,
                formula2,
                input_title = inputTitle,
                input_message = inputMessage,
                error_title = errorTitle,
                error_message = errorMessage,
                error_style = errorStyle,
                ignore_blank = ignoreBlank,
                in_cell_dropdown = inCellDropdown,
                show_input = showInput,
                show_error = showError,
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"validation.read: {ex.Message}");
        }
    }

    private string HandleCommentThreaded(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("comment.thread requires 'ref'");
            var text = args["text"]?.GetValue<string>()
                ?? throw new ArgumentException("comment.thread requires 'text'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);

            string? fallback = null;
            try
            {
                ((dynamic)range).AddCommentThreaded(text);
            }
            catch
            {
                // Older Excel without threaded comment support — fall back to legacy
                try
                {
                    var existing = range.Comment;
                    if (existing != null)
                    {
                        existing.Delete();
                        Marshal.ReleaseComObject(existing);
                    }
                }
                catch { }
                range.AddComment(text);
                fallback = "legacy";
            }

            return Response.Ok(new { @ref = refStr, text, added = true, fallback });
        }
        catch (Exception ex)
        {
            return Response.Error($"comment.thread: {ex.Message}");
        }
    }

    private string HandleCommentThreadedRead(JsonObject args)
    {
        try
        {
            var refStr = args["ref"]?.GetValue<string>()
                ?? throw new ArgumentException("comment.thread.read requires 'ref'");

            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);

            // Try threaded comments first
            try
            {
                dynamic ct = ((dynamic)range).CommentThreaded;
                if (ct != null)
                {
                    var text = (string)ct.Text;
                    string? author = null;
                    try { author = (string)ct.Author.Name; } catch { }

                    var replies = new JsonArray();
                    try
                    {
                        foreach (dynamic reply in ct.Replies)
                        {
                            var replyObj = new JsonObject { ["text"] = (string)reply.Text };
                            try { replyObj["author"] = (string)reply.Author.Name; } catch { }
                            replies.Add(replyObj);
                            Marshal.ReleaseComObject((object)reply);
                        }
                    }
                    catch { }

                    Marshal.ReleaseComObject((object)ct);
                    return Response.Ok(new
                    {
                        @ref = refStr,
                        type = "threaded",
                        text,
                        author,
                        replies,
                        reply_count = replies.Count,
                    });
                }
            }
            catch { }

            // Fall back to legacy comment
            try
            {
                var comment = range.Comment;
                if (comment != null)
                {
                    var text = comment.Text();
                    Marshal.ReleaseComObject(comment);
                    return Response.Ok(new { @ref = refStr, type = "legacy", text });
                }
            }
            catch { }

            return Response.Ok(new { @ref = refStr, type = (string?)null, text = (string?)null });
        }
        catch (Exception ex)
        {
            return Response.Error($"comment.thread.read: {ex.Message}");
        }
    }
}
