using System.Reflection;
using System.Windows;

namespace OneNoteExporter;

public partial class AboutWindow : Window
{
    public AboutWindow()
    {
        InitializeComponent();

        var assembly = Assembly.GetExecutingAssembly();
        var informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        var versionParts = informationalVersion?.Split('+', 2);
        var version = versionParts is { Length: > 0 } && !string.IsNullOrWhiteSpace(versionParts[0])
            ? versionParts[0]
            : assembly.GetName().Version?.ToString(3) ?? "Unknown";
        var configuration = assembly
            .GetCustomAttribute<AssemblyConfigurationAttribute>()?
            .Configuration ?? "Release";

        VersionText.Text = $"Version {version}";
        BuildText.Text = versionParts is { Length: 2 }
            ? $"{configuration} build · {versionParts[1][..Math.Min(7, versionParts[1].Length)]}"
            : $"{configuration} build";
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
}
