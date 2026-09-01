using System.Windows;
using System.Windows.Threading;
using OneNoteExporter.Helpers;

namespace OneNoteExporter;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Global unhandled exception handler
        DispatcherUnhandledException += OnDispatcherUnhandledException;
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        MessageBox.Show(
            UserFacingError.Describe(
                e.Exception,
                "An unexpected error occurred. Please restart the app and try again."),
            "OneNote Exporter – Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        e.Handled = true;
    }
}
