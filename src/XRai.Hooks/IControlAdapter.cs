namespace XRai.Hooks;

/// <summary>
/// Common interface for control adapters (WPF and WinForms).
/// Every method that PipeServer calls on a control goes through this interface.
/// </summary>
public interface IControlAdapter
{
    string Name { get; }
    string Type { get; }
    bool IsEnabled { get; }
    bool IsVisible { get; }
    bool HasCommand { get; }

    string? GetValue();
    void SetValue(string value);
    ControlAdapter.ClickResult Click();
    void Toggle();
    void Focus();
    void DoubleClick();
    void RightClick();
    void Hover();
    string Expand(bool open = true);
    ControlAdapter.SendKeysResult SendKeys(string keys);
    (string[] Items, int SelectedIndex) ReadListItems();
    string SelectListItem(int? index, string? text);
    object? GetDetailedInfo();

    // DataGrid operations
    string? GetDataGridCell(int row, int col);
    object?[][] GetDataGridAllData();
    void SelectDataGridRow(int index);

    // Tree operations
    void ExpandTreeNode(string path, bool open = true);

    // Scroll
    void ScrollTo(double offset);

    // Drag-and-drop (WPF-only; WinForms adapters throw NotSupportedException)
    void DragTo(System.Windows.FrameworkElement target);

    // Context menu
    void OpenContextMenu();
    string[] GetContextMenuItems();
    void ClickContextMenuItem(string name);

    // WPF element access (null for WinForms adapters)
    System.Windows.FrameworkElement? Element { get; }
}
