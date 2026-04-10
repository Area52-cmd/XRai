using System.IO;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Com;

public class VbaOps
{
    private readonly ExcelSession _session;

    public VbaOps(ExcelSession session) { _session = session; }

    public void Register(CommandRouter router)
    {
        router.Register("vba.list", HandleList);
        router.Register("vba.view", HandleView);
        router.Register("vba.import", HandleImport);
        router.Register("vba.update", HandleUpdate);
        router.Register("vba.delete", HandleDelete);
    }

    private string HandleList(JsonObject args)
    {
        try
        {
            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());

            dynamic vbProject;
            try { vbProject = wb.VBProject; }
            catch
            {
                return Response.Error("VBA project access denied. Enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model", code: ErrorCodes.VbaAccessDenied);
            }

            dynamic components = vbProject.VBComponents;
            int count = components.Count;
            var result = new JsonArray();

            for (int i = 1; i <= count; i++)
            {
                dynamic comp = components.Item(i);
                string name = comp.Name;
                int typeVal = (int)comp.Type;
                string typeName = typeVal switch
                {
                    1 => "Standard",
                    2 => "Class",
                    3 => "UserForm",
                    100 => "Document",
                    _ => $"Unknown({typeVal})",
                };

                int lineCount = 0;
                try { lineCount = comp.CodeModule.CountOfLines; } catch { }

                result.Add(new JsonObject
                {
                    ["name"] = name,
                    ["type"] = typeName,
                    ["type_id"] = typeVal,
                    ["line_count"] = lineCount,
                });
            }

            return Response.Ok(new { count, components = result });
        }
        catch (Exception ex)
        {
            return Response.Error($"vba.list: {ex.Message}");
        }
    }

    private string HandleView(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("vba.view requires 'name'");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());

            dynamic vbProject;
            try { vbProject = wb.VBProject; }
            catch
            {
                return Response.Error("VBA project access denied. Enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model", code: ErrorCodes.VbaAccessDenied);
            }

            dynamic comp = FindComponent(vbProject.VBComponents, name);
            dynamic codeModule = comp.CodeModule;
            int lineCount = codeModule.CountOfLines;
            string code = lineCount > 0 ? codeModule.Lines(1, lineCount) : "";

            return Response.Ok(new { name, line_count = lineCount, code });
        }
        catch (Exception ex)
        {
            return Response.Error($"vba.view: {ex.Message}");
        }
    }

    private string HandleImport(JsonObject args)
    {
        try
        {
            var path = args["path"]?.GetValue<string>()
                ?? throw new ArgumentException("vba.import requires 'path'");

            if (!File.Exists(path))
                return Response.Error($"File not found: {path}");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());

            dynamic vbProject;
            try { vbProject = wb.VBProject; }
            catch
            {
                return Response.Error("VBA project access denied. Enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model", code: ErrorCodes.VbaAccessDenied);
            }

            dynamic imported = vbProject.VBComponents.Import(path);
            string importedName = imported.Name;

            return Response.Ok(new { name = importedName, path, imported = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"vba.import: {ex.Message}");
        }
    }

    private string HandleUpdate(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("vba.update requires 'name'");
            var code = args["code"]?.GetValue<string>()
                ?? throw new ArgumentException("vba.update requires 'code'");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());

            dynamic vbProject;
            try { vbProject = wb.VBProject; }
            catch
            {
                return Response.Error("VBA project access denied. Enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model", code: ErrorCodes.VbaAccessDenied);
            }

            dynamic comp = FindComponent(vbProject.VBComponents, name);
            dynamic codeModule = comp.CodeModule;

            // Delete all existing lines
            int existingLines = codeModule.CountOfLines;
            if (existingLines > 0)
                codeModule.DeleteLines(1, existingLines);

            // Insert new code
            if (!string.IsNullOrEmpty(code))
                codeModule.InsertLines(1, code);

            int newLineCount = codeModule.CountOfLines;
            return Response.Ok(new { name, line_count = newLineCount, updated = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"vba.update: {ex.Message}");
        }
    }

    private string HandleDelete(JsonObject args)
    {
        try
        {
            var name = args["name"]?.GetValue<string>()
                ?? throw new ArgumentException("vba.delete requires 'name'");

            using var guard = new ComGuard();
            var wb = guard.Track(_session.GetActiveWorkbook());

            dynamic vbProject;
            try { vbProject = wb.VBProject; }
            catch
            {
                return Response.Error("VBA project access denied. Enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model", code: ErrorCodes.VbaAccessDenied);
            }

            dynamic comp = FindComponent(vbProject.VBComponents, name);
            vbProject.VBComponents.Remove(comp);

            return Response.Ok(new { name, deleted = true });
        }
        catch (Exception ex)
        {
            return Response.Error($"vba.delete: {ex.Message}");
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static dynamic FindComponent(dynamic components, string name)
    {
        int count = components.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic comp = components.Item(i);
            if (string.Equals((string)comp.Name, name, StringComparison.OrdinalIgnoreCase))
                return comp;
        }
        throw new ArgumentException($"VBA component not found: {name}");
    }
}
