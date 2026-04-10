using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Spike1.Com;

class Program
{
    [DllImport("oleaut32.dll", PreserveSig = false)]
    static extern void GetActiveObject(
        [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
        nint pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    static object GetActiveObject(string progId)
    {
        Guid clsid = Type.GetTypeFromProgID(progId, true)!.GUID;
        GetActiveObject(clsid, 0, out object obj);
        return obj;
    }

    static void Main()
    {
        Excel.Application? app = null;
        Excel.Workbook? workbook = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? rangeA1 = null;
        Excel.Range? rangeA2 = null;
        Excel.Range? rangeA3 = null;
        Excel.Range? rangeA4 = null;
        Excel.Range? rangeA5 = null;
        Excel.Range? rangeA6 = null;
        Excel.Range? rangeA1A6 = null;
        Excel.Range? rangeB1 = null;
        Excel.Font? fontA1 = null;
        Excel.Interior? interiorA1 = null;
        Excel.Sheets? sheets = null;

        try
        {
            // 1. Attach to running Excel
            object excelObj = GetActiveObject("Excel.Application");
            app = (Excel.Application)excelObj;
            Console.WriteLine($"Attached to Excel: {app.Version}");

            workbook = app.ActiveWorkbook;
            if (workbook == null)
            {
                // No workbook open (Excel may be on start screen) — create one
                Excel.Workbooks workbooks = app.Workbooks;
                workbook = workbooks.Add();
                Marshal.ReleaseComObject(workbooks);
                Console.WriteLine($"Created new workbook: {workbook.Name}");
            }
            else
            {
                Console.WriteLine($"Active workbook: {workbook.Name}");
            }

            // 2. Write a string value
            sheet = (Excel.Worksheet)workbook.ActiveSheet;
            rangeA1 = sheet.Range["A1"];
            rangeA1.Value2 = "Hello from XRai";
            Console.WriteLine("Wrote string to A1");

            // 3. Read it back
            object? valA1 = rangeA1.Value2;
            Console.WriteLine($"Read A1: {valA1}");
            if (valA1?.ToString() != "Hello from XRai")
                throw new Exception("ASSERTION FAILED: A1 value mismatch");

            // 4. Write numeric values to A2:A6
            rangeA2 = sheet.Range["A2"]; rangeA2.Value2 = 10;
            rangeA3 = sheet.Range["A3"]; rangeA3.Value2 = 20;
            rangeA4 = sheet.Range["A4"]; rangeA4.Value2 = 30;
            rangeA5 = sheet.Range["A5"]; rangeA5.Value2 = 40;
            rangeA6 = sheet.Range["A6"]; rangeA6.Value2 = 50;
            Console.WriteLine("Wrote values to A2:A6");

            // 5. Read range A1:A6
            rangeA1A6 = sheet.Range["A1", "A6"];
            foreach (Excel.Range cell in rangeA1A6)
            {
                string addr = cell.Address[false, false];
                object? val = cell.Value2;
                Console.WriteLine($"{addr}: {val}");
                Marshal.ReleaseComObject(cell);
            }

            // 6. Write formula to B1
            rangeB1 = sheet.Range["B1"];
            rangeB1.Formula = "=SUM(A2:A6)";
            Console.WriteLine("Wrote formula to B1");

            // 7. Trigger calculation
            app.Calculate();
            Console.WriteLine("Calculation triggered");

            // 8. Read formula result
            object? b1Value = rangeB1.Value2;
            object? b1Formula = rangeB1.Formula;
            Console.WriteLine($"B1 value: {b1Value}, formula: {b1Formula}");
            if (Convert.ToDouble(b1Value) != 150.0)
                throw new Exception($"ASSERTION FAILED: B1 value expected 150, got {b1Value}");

            // 9. Read cell formatting
            fontA1 = rangeA1.Font;
            string fontName = fontA1.Name?.ToString() ?? "Unknown";
            double fontSize = Convert.ToDouble(fontA1.Size);
            bool isBold = Convert.ToBoolean(fontA1.Bold);
            Console.WriteLine($"A1 font: {fontName}, size: {fontSize}, bold: {isBold}");

            // 10. Write formatting
            fontA1.Bold = true;
            fontA1.Size = 14;
            interiorA1 = rangeA1.Interior;
            // Yellow in OLE color format: RGB(255, 255, 0) = 0x00FFFF = 65535
            interiorA1.Color = 65535;
            Console.WriteLine("Applied formatting to A1");

            // Read back formatting to verify
            string fontName2 = fontA1.Name?.ToString() ?? "Unknown";
            double fontSize2 = Convert.ToDouble(fontA1.Size);
            bool isBold2 = Convert.ToBoolean(fontA1.Bold);
            Console.WriteLine($"A1 font: {fontName2}, size: {fontSize2}, bold: {isBold2}");

            // 11. List all sheets
            sheets = workbook.Sheets;
            int sheetCount = sheets.Count;
            foreach (Excel.Worksheet s in sheets)
            {
                Console.WriteLine($"Sheet: {s.Name}");
                Marshal.ReleaseComObject(s);
            }
            Console.WriteLine($"Found {sheetCount} sheets");

            // 12. Clean up COM objects
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR: {ex.GetType().Name}: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
        finally
        {
            // Release in reverse order
            if (sheets != null) Marshal.ReleaseComObject(sheets);
            if (interiorA1 != null) Marshal.ReleaseComObject(interiorA1);
            if (fontA1 != null) Marshal.ReleaseComObject(fontA1);
            if (rangeB1 != null) Marshal.ReleaseComObject(rangeB1);
            if (rangeA1A6 != null) Marshal.ReleaseComObject(rangeA1A6);
            if (rangeA6 != null) Marshal.ReleaseComObject(rangeA6);
            if (rangeA5 != null) Marshal.ReleaseComObject(rangeA5);
            if (rangeA4 != null) Marshal.ReleaseComObject(rangeA4);
            if (rangeA3 != null) Marshal.ReleaseComObject(rangeA3);
            if (rangeA2 != null) Marshal.ReleaseComObject(rangeA2);
            if (rangeA1 != null) Marshal.ReleaseComObject(rangeA1);
            if (sheet != null) Marshal.ReleaseComObject(sheet);
            if (workbook != null) Marshal.ReleaseComObject(workbook);
            if (app != null) Marshal.ReleaseComObject(app);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.WriteLine("COM cleanup complete");
        }
    }
}
