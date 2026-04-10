using System.Collections;
using System.Globalization;
using System.Windows.Data;

namespace XRai.Demo.MacroGuard;

/// <summary>
/// Converts a list of strings to a comma-separated string for display binding.
/// </summary>
public class StringJoinConverter : IValueConverter
{
    public static readonly StringJoinConverter Instance = new();

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is IEnumerable<string> strings)
            return string.Join(", ", strings);
        if (value is IEnumerable list)
            return string.Join(", ", list.Cast<object>().Select(o => o?.ToString() ?? ""));
        return value?.ToString() ?? "";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
