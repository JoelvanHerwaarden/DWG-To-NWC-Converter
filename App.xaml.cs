using System.Windows;

namespace DWGToNWCConverter;

public partial class App : System.Windows.Application
{
    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        if (e.Args.Any(arg => string.Equals(arg, "--run-scheduled", StringComparison.OrdinalIgnoreCase)))
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;

            var settings = SettingsService.Load();
            var converter = new ConversionService();
            await converter.RunBatchAsync(settings, CancellationToken.None);

            Shutdown();
            return;
        }

        var window = new MainWindow();
        MainWindow = window;
        window.Show();
    }
}
