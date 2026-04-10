using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Services;

public static class SnippetLibrary
{
    public static List<VbaSnippet> GetAllSnippets() => new()
    {
        new VbaSnippet
        {
            Name = "Error Handler Template",
            Category = "Error Handling",
            Description = "Standard error handling pattern with logging",
            Code = @"Sub MySub()
    On Error GoTo ErrHandler
    ' --- your code here ---
    Exit Sub
ErrHandler:
    MsgBox ""Error "" & Err.Number & "": "" & Err.Description, vbCritical
    Resume Next
End Sub"
        },
        new VbaSnippet
        {
            Name = "FileDialog Open",
            Category = "File I/O",
            Description = "Open file dialog to select a file",
            Code = @"Function BrowseForFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = ""Select a file""
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        BrowseForFile = fd.SelectedItems(1)
    End If
End Function"
        },
        new VbaSnippet
        {
            Name = "FileDialog Save",
            Category = "File I/O",
            Description = "Save-as file dialog",
            Code = @"Function BrowseForSave() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    If fd.Show = -1 Then
        BrowseForSave = fd.SelectedItems(1)
    End If
End Function"
        },
        new VbaSnippet
        {
            Name = "Array Sort (Bubble)",
            Category = "Data Structures",
            Description = "Sort a variant array using bubble sort",
            Code = @"Sub BubbleSort(arr() As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub"
        },
        new VbaSnippet
        {
            Name = "Dictionary Usage",
            Category = "Data Structures",
            Description = "Scripting.Dictionary pattern for key-value lookups",
            Code = @"Sub UseDictionary()
    Dim dict As Object
    Set dict = CreateObject(""Scripting.Dictionary"")
    dict.Add ""key1"", ""value1""
    dict.Add ""key2"", ""value2""
    If dict.Exists(""key1"") Then
        Debug.Print dict(""key1"")
    End If
    Dim k As Variant
    For Each k In dict.Keys
        Debug.Print k & "" = "" & dict(k)
    Next k
End Sub"
        },
        new VbaSnippet
        {
            Name = "RegExp Pattern Match",
            Category = "Text Processing",
            Description = "Use VBScript.RegExp for pattern matching",
            Code = @"Function RegExMatch(text As String, pattern As String) As Boolean
    Dim re As Object
    Set re = CreateObject(""VBScript.RegExp"")
    re.Pattern = pattern
    re.IgnoreCase = True
    re.Global = True
    RegExMatch = re.Test(text)
End Function"
        },
        new VbaSnippet
        {
            Name = "Send Email via Outlook",
            Category = "Automation",
            Description = "Create and send an email using Outlook",
            Code = @"Sub SendEmail(toAddr As String, subject As String, body As String)
    Dim olApp As Object, olMail As Object
    Set olApp = CreateObject(""Outlook.Application"")
    Set olMail = olApp.CreateItem(0)
    With olMail
        .To = toAddr
        .Subject = subject
        .Body = body
        .Send
    End With
    Set olMail = Nothing: Set olApp = Nothing
End Sub"
        },
        new VbaSnippet
        {
            Name = "SQL Query via ADO",
            Category = "Database",
            Description = "Execute a SQL query and return a recordset",
            Code = @"Function RunQuery(connStr As String, sql As String) As Object
    Dim conn As Object, rs As Object
    Set conn = CreateObject(""ADODB.Connection"")
    Set rs = CreateObject(""ADODB.Recordset"")
    conn.Open connStr
    rs.Open sql, conn, 3, 1 ' adOpenStatic, adLockReadOnly
    Set RunQuery = rs
    ' Caller must close rs and conn
End Function"
        },
        new VbaSnippet
        {
            Name = "Progress Bar in Status Bar",
            Category = "UI",
            Description = "Show progress in the Excel status bar",
            Code = @"Sub ShowProgress(current As Long, total As Long)
    Dim pct As Long
    pct = Int((current / total) * 100)
    Application.StatusBar = ""Processing: "" & pct & ""%  "" & String(pct \ 2, ""|"") & String(50 - pct \ 2, ""-"")
    If current = total Then Application.StatusBar = False
End Sub"
        },
        new VbaSnippet
        {
            Name = "Worksheet Loop",
            Category = "Excel",
            Description = "Loop through all worksheets in the active workbook",
            Code = @"Sub LoopWorksheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name & "": "" & ws.UsedRange.Rows.Count & "" rows""
    Next ws
End Sub"
        },
        new VbaSnippet
        {
            Name = "Last Row / Column Finder",
            Category = "Excel",
            Description = "Find the last used row and column on a sheet",
            Code = @"Function LastRow(ws As Worksheet, Optional col As Long = 1) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Function LastCol(ws As Worksheet, Optional row As Long = 1) As Long
    LastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
End Function"
        },
        new VbaSnippet
        {
            Name = "Timer Macro",
            Category = "Utility",
            Description = "Schedule a macro to run at a specific time or interval",
            Code = @"Sub StartTimer()
    Application.OnTime Now + TimeValue(""00:00:30""), ""MyRepeatingMacro""
End Sub

Sub MyRepeatingMacro()
    ' --- your code here ---
    Debug.Print ""Tick: "" & Now
    StartTimer ' reschedule
End Sub

Sub StopTimer()
    On Error Resume Next
    Application.OnTime Now + TimeValue(""00:00:30""), ""MyRepeatingMacro"", , False
End Sub"
        },
        new VbaSnippet
        {
            Name = "JSON Parser",
            Category = "Data Processing",
            Description = "Minimal JSON string parser using ScriptControl",
            Code = @"Function ParseJson(jsonStr As String) As Object
    Dim sc As Object
    Set sc = CreateObject(""MSScriptControl.ScriptControl"")
    sc.Language = ""JScript""
    Set ParseJson = sc.Eval(""("" & jsonStr & "")"")
End Function"
        },
        new VbaSnippet
        {
            Name = "HTTP GET Request",
            Category = "Web",
            Description = "Perform an HTTP GET request and return the response body",
            Code = @"Function HttpGet(url As String) As String
    Dim http As Object
    Set http = CreateObject(""MSXML2.XMLHTTP"")
    http.Open ""GET"", url, False
    http.Send
    HttpGet = http.responseText
End Function"
        },
        new VbaSnippet
        {
            Name = "PDF Export",
            Category = "Export",
            Description = "Export the active sheet as a PDF file",
            Code = @"Sub ExportToPdf(filePath As String)
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        OpenAfterPublish:=False
End Sub"
        },
        new VbaSnippet
        {
            Name = "Pivot Table Creation",
            Category = "Excel",
            Description = "Create a pivot table from a data range",
            Code = @"Sub CreatePivot(dataRange As Range, destSheet As String, tableName As String)
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(destSheet)
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range(""A3""), TableName:=tableName)
    ' Add fields as needed:
    ' pt.PivotFields(""Category"").Orientation = xlRowField
    ' pt.AddDataField pt.PivotFields(""Amount""), ""Sum of Amount"", xlSum
End Sub"
        },
    };
}
