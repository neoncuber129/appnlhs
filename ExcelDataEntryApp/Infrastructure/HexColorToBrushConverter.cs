using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelDataEntryApp.Infrastructure;

public sealed class HexColorToBrushConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is not string s || string.IsNullOrWhiteSpace(s))
            return Brushes.Transparent;

        if (s.Equals("Transparent", StringComparison.OrdinalIgnoreCase))
            return Brushes.Transparent;

        try
        {
            var color = (Color)ColorConverter.ConvertFromString(s);
            return new SolidColorBrush(color);
        }
        catch (FormatException)
        {
            return Brushes.Transparent;
        }
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
