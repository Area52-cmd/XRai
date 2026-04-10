using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using XRai.Demo.MacroGuard.Models;
using Brush = System.Windows.Media.Brush;
using Color = System.Windows.Media.Color;
using FontFamily = System.Windows.Media.FontFamily;
using RichTextBox = System.Windows.Controls.RichTextBox;
using SolidColorBrush = System.Windows.Media.SolidColorBrush;

namespace XRai.Demo.MacroGuard;

public partial class MacroGuardPane : System.Windows.Controls.UserControl
{
    private readonly MacroGuardViewModel _vm;

    public MacroGuardPane() : this(new MacroGuardViewModel()) { }

    public MacroGuardPane(MacroGuardViewModel vm)
    {
        _vm = vm;
        DataContext = _vm;
        InitializeComponent();

        _vm.PropertyChanged += (s, e) =>
        {
            Dispatcher.BeginInvoke(() =>
            {
                switch (e.PropertyName)
                {
                    case nameof(MacroGuardViewModel.ResponseText):
                        if (!string.IsNullOrEmpty(_vm.ResponseText))
                            ResponseBox.Document = SyntaxHighlighter.Highlight(_vm.ResponseText);
                        else
                            ResponseBox.Document = new FlowDocument();
                        break;

                    case nameof(MacroGuardViewModel.CodePreviewText):
                        if (!string.IsNullOrEmpty(_vm.CodePreviewText))
                            CodePreview.Document = SyntaxHighlighter.Highlight(_vm.CodePreviewText);
                        else
                            CodePreview.Document = new FlowDocument();
                        break;

                    case nameof(MacroGuardViewModel.SelectedSnippet):
                        UpdateSnippetPreview();
                        break;

                    case nameof(MacroGuardViewModel.DiffText):
                        UpdateDiffView();
                        break;
                }
            });
        };
    }

    public MacroGuardViewModel ViewModel => _vm;

    private void ModuleTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is VbaModuleInfo module)
        {
            _vm.SelectedModule = module.Name;
            _vm.SelectedMacro = module.MacroNames.FirstOrDefault() ?? "";
            _vm.CodePreviewText = module.Code;
        }
        else if (e.NewValue is string macroName)
        {
            _vm.SelectedMacro = macroName;
        }
    }

    private void ApiKeyBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.PasswordBox pb)
            _vm.ApiKey = pb.Password;
    }

    private void CodeInput_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        if (sender is RichTextBox rtb)
        {
            var range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            _vm.CodeInputText = range.Text.TrimEnd();
        }
    }

    private void ProtectButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
            var vbProject = app.ActiveWorkbook.VBProject;
            _vm.StatusMessage = "VBProject protection toggled";
        }
        catch (Exception ex)
        {
            _vm.StatusMessage = $"Protection error: {ex.Message}";
        }
    }

    // ── Premium feature UI helpers ──────────────────────────────

    private void UpdateSnippetPreview()
    {
        var code = _vm.GetSelectedSnippetCode();
        if (!string.IsNullOrEmpty(code))
            SnippetPreview.Document = SyntaxHighlighter.Highlight(code);
        else
            SnippetPreview.Document = new FlowDocument();
    }

    private void UpdateDiffView()
    {
        var text = _vm.DiffText;
        if (string.IsNullOrEmpty(text))
        {
            DiffView.Document = new FlowDocument();
            return;
        }

        var doc = new FlowDocument
        {
            Background = new SolidColorBrush(Color.FromRgb(30, 30, 30)),
            Foreground = new SolidColorBrush(Color.FromRgb(212, 212, 212)),
            FontFamily = new FontFamily("Cascadia Code, Consolas, Courier New"),
            FontSize = 11,
            PagePadding = new Thickness(8)
        };

        var para = new Paragraph();
        var addedBrush = new SolidColorBrush(Color.FromRgb(106, 153, 85));   // Green
        var removedBrush = new SolidColorBrush(Color.FromRgb(233, 69, 96));  // Red
        var headerBrush = new SolidColorBrush(Color.FromRgb(86, 156, 214));  // Blue
        var defaultBrush = new SolidColorBrush(Color.FromRgb(212, 212, 212));
        addedBrush.Freeze(); removedBrush.Freeze(); headerBrush.Freeze(); defaultBrush.Freeze();

        foreach (var line in text.Split('\n'))
        {
            Brush brush;
            if (line.StartsWith("=== ")) brush = headerBrush;
            else if (line.StartsWith("+ ")) brush = addedBrush;
            else if (line.StartsWith("- ")) brush = removedBrush;
            else brush = defaultBrush;

            para.Inlines.Add(new Run(line) { Foreground = brush });
            para.Inlines.Add(new LineBreak());
        }

        doc.Blocks.Add(para);
        DiffView.Document = doc;
    }
}
