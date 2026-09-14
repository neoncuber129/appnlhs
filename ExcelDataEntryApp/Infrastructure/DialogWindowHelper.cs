using System.Windows;

namespace ExcelDataEntryApp.Infrastructure;

internal static class DialogWindowHelper
{
    public static void ClampToWorkArea(Window window)
    {
        window.SourceInitialized += (_, _) =>
        {
            var area = SystemParameters.WorkArea;
            if (window.Width > area.Width * 0.96)
            {
                window.Width = area.Width * 0.96;
            }

            if (window.Height > area.Height * 0.92)
            {
                window.Height = area.Height * 0.92;
            }
        };
    }
}
