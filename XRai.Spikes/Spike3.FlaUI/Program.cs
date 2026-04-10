using System.Diagnostics;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace Spike3.FlaUITest;

class Program
{
    static readonly Dictionary<string, string> _results = new();

    static void Main()
    {
        try
        {
            // 1. Find Excel process
            var excelProcesses = Process.GetProcessesByName("EXCEL");
            if (excelProcesses.Length == 0)
            {
                Console.WriteLine("ERROR: No Excel process found.");
                return;
            }

            var excelProcess = excelProcesses[0];
            Console.WriteLine($"Found Excel: PID {excelProcess.Id}, Title: {excelProcess.MainWindowTitle}");

            // 2. Attach FlaUI
            using var automation = new UIA3Automation();
            var app = FlaUI.Core.Application.Attach(excelProcess);
            var window = app.GetMainWindow(automation, TimeSpan.FromSeconds(5));
            Console.WriteLine($"Attached FlaUI to: {window.Title}");
            _results["FlaUI attach"] = "OK";

            // 3. Dump top-level UI tree (2 levels deep)
            Console.WriteLine();
            Console.WriteLine("=== UI TREE (2 levels) ===");
            var children = window.FindAllChildren();
            foreach (var child in children)
            {
                PrintElement(child, 1);
                try
                {
                    var grandchildren = child.FindAllChildren();
                    foreach (var gc in grandchildren)
                    {
                        PrintElement(gc, 2);
                    }
                }
                catch { }
            }
            _results["UI tree readable"] = "OK";
            Console.WriteLine();

            // 4. Find the ribbon
            AutomationElement? ribbon = null;
            try
            {
                // Try by ClassName first (Excel ribbon is often "NetUIHWND")
                ribbon = window.FindFirstDescendant(cf => cf.ByControlType(ControlType.ToolBar));
                if (ribbon == null)
                    ribbon = window.FindFirstDescendant(cf => cf.ByClassName("NetUIHWND"));
                if (ribbon == null)
                    ribbon = window.FindFirstDescendant(cf => cf.ByName("Ribbon"));
            }
            catch { }

            if (ribbon != null)
            {
                Console.WriteLine($"Found ribbon element: {ribbon.Name} (Type: {ribbon.ControlType}, Class: {ribbon.ClassName})");
                _results["Ribbon found"] = "OK";
            }
            else
            {
                Console.WriteLine("Ribbon element not found by ToolBar/NetUIHWND/Name search");
                _results["Ribbon found"] = "FAIL";
            }

            // 5. List all ribbon tab names
            Console.WriteLine();
            Console.WriteLine("=== RIBBON TABS ===");
            AutomationElement[] tabs = Array.Empty<AutomationElement>();
            try
            {
                tabs = window.FindAllDescendants(cf => cf.ByControlType(ControlType.TabItem));
                int tabCount = 0;
                foreach (var tab in tabs)
                {
                    if (!string.IsNullOrWhiteSpace(tab.Name))
                    {
                        Console.WriteLine($"Tab: {tab.Name}");
                        tabCount++;
                    }
                }
                Console.WriteLine($"Found {tabCount} ribbon tabs");
                _results[$"Tab names found"] = tabCount > 0 ? $"OK ({tabCount} tabs)" : "FAIL";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error finding tabs: {ex.Message}");
                _results["Tab names found"] = "FAIL";
            }

            // 6. Inspect one tab's contents (Home tab or active tab)
            Console.WriteLine();
            Console.WriteLine("=== TAB CONTENTS ===");
            try
            {
                // Find groups within the ribbon area
                var groups = window.FindAllDescendants(cf => cf.ByControlType(ControlType.Group));
                int groupCount = 0;
                foreach (var group in groups)
                {
                    if (!string.IsNullOrWhiteSpace(group.Name))
                    {
                        Console.WriteLine($"  Group: {group.Name}");
                        groupCount++;
                        try
                        {
                            var buttons = group.FindAllChildren(cf => cf.ByControlType(ControlType.Button));
                            foreach (var btn in buttons.Take(5)) // limit output
                            {
                                if (!string.IsNullOrWhiteSpace(btn.Name))
                                {
                                    Console.WriteLine($"    Button: {btn.Name} [Enabled: {btn.IsEnabled}]");
                                }
                            }
                        }
                        catch { }
                    }
                    if (groupCount >= 10) break; // limit output
                }
                _results["Tab contents"] = groupCount > 0 ? "OK" : "FAIL";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading tab contents: {ex.Message}");
                _results["Tab contents"] = "FAIL";
            }

            // 7. Find the formula bar
            Console.WriteLine();
            Console.WriteLine("=== FORMULA BAR ===");
            try
            {
                var formulaBar = window.FindFirstDescendant(cf => cf.ByAutomationId("FormulaBar"))
                    ?? window.FindFirstDescendant(cf => cf.ByName("Formula Bar"));

                if (formulaBar != null)
                {
                    Console.WriteLine($"Formula bar found: {formulaBar.Name} (Type: {formulaBar.ControlType})");
                    // Try to get text
                    try
                    {
                        var edit = formulaBar.FindFirstDescendant(cf => cf.ByControlType(ControlType.Edit));
                        if (edit != null)
                        {
                            Console.WriteLine($"Formula bar text: {edit.Name}");
                        }
                    }
                    catch { }
                    _results["Formula bar"] = "OK";
                }
                else
                {
                    // Try searching by ClassName
                    var nameBox = window.FindFirstDescendant(cf => cf.ByAutomationId("Name Box"))
                        ?? window.FindFirstDescendant(cf => cf.ByName("Name Box"));
                    if (nameBox != null)
                    {
                        Console.WriteLine($"Name Box found (near formula bar): {nameBox.Name}");
                        _results["Formula bar"] = "OK (Name Box)";
                    }
                    else
                    {
                        Console.WriteLine("Formula bar not found");
                        _results["Formula bar"] = "FAIL";
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error finding formula bar: {ex.Message}");
                _results["Formula bar"] = "FAIL";
            }

            // 8. Find sheet tabs
            Console.WriteLine();
            Console.WriteLine("=== SHEET TABS ===");
            try
            {
                // Sheet tabs are usually TabItem controls at the bottom
                var allTabs = window.FindAllDescendants(cf => cf.ByControlType(ControlType.TabItem));
                var sheetTabs = allTabs.Where(t =>
                    t.Name.StartsWith("Sheet") || t.Name.Contains("Sheet")).ToArray();

                if (sheetTabs.Length == 0)
                {
                    // Try by AutomationId pattern
                    sheetTabs = allTabs.Where(t =>
                    {
                        try { return t.AutomationId?.Contains("Sheet") == true; }
                        catch { return false; }
                    }).ToArray();
                }

                if (sheetTabs.Length > 0)
                {
                    foreach (var st in sheetTabs)
                    {
                        Console.WriteLine($"Sheet tab: {st.Name}");
                    }
                    _results["Sheet tabs"] = "OK";
                }
                else
                {
                    Console.WriteLine("No sheet tabs found via TabItem search. Dumping all tab items:");
                    foreach (var t in allTabs)
                    {
                        Console.WriteLine($"  TabItem: \"{SafeGet(() => t.Name)}\" [AutomationId: {SafeGet(() => t.AutomationId)}]");
                    }
                    _results["Sheet tabs"] = "FAIL";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error finding sheet tabs: {ex.Message}");
                _results["Sheet tabs"] = "FAIL";
            }

            // 9. Screenshot
            Console.WriteLine();
            try
            {
                var capture = FlaUI.Core.Capturing.Capture.Element(window);
                string path = Path.Combine(AppContext.BaseDirectory, "spike3_screenshot.png");
                capture.ToFile(path);
                Console.WriteLine($"Screenshot saved to: {path}");
                _results["Screenshot"] = "OK";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Screenshot failed: {ex.Message}");
                _results["Screenshot"] = "FAIL";
            }

            // 10. Summary
            Console.WriteLine();
            Console.WriteLine("═══════════════════════════════════════");
            Console.WriteLine("SPIKE 3 RESULTS");
            foreach (var kvp in _results)
            {
                Console.WriteLine($"- {kvp.Key + ":",-22} {kvp.Value}");
            }
            Console.WriteLine("═══════════════════════════════════════");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"FATAL ERROR: {ex.GetType().Name}: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    static string SafeGet(Func<string?> getter)
    {
        try { return getter() ?? ""; }
        catch { return ""; }
    }

    static void PrintElement(AutomationElement el, int indent)
    {
        string prefix = new string(' ', indent * 2);
        string ct = SafeGet(() => el.ControlType.ToString());
        string name = SafeGet(() => el.Name);
        string aid = SafeGet(() => el.AutomationId);
        string cls = SafeGet(() => el.ClassName);
        Console.WriteLine($"{prefix}{ct}: \"{name}\" [AutomationId: {aid}, Class: {cls}]");
    }
}
