using System.Windows;
using System.Windows.Media.Imaging;

namespace ExcelDataEntryApp;

public partial class App : Application
{
    private static BitmapFrame? _appIcon;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        _appIcon = BitmapFrame.Create(
            new Uri("pack://application:,,,/Assets/app.ico", UriKind.Absolute),
            BitmapCreateOptions.None,
            BitmapCacheOption.OnLoad);

        EventManager.RegisterClassHandler(
            typeof(Window),
            FrameworkElement.LoadedEvent,
            new RoutedEventHandler(OnWindowLoaded));
    }

    private static void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
        if (sender is Window window && window.Icon is null && _appIcon is not null)
        {
            window.Icon = _appIcon;
        }
    }
}
