using System.Windows;

namespace OneNoteExporter;

public enum RenderedExportChoice
{
    Cancel,
    IndividualSections,
    OneNotePackage,
    WholeNotebook
}

public partial class RenderedExportWarningWindow : Window
{
    public RenderedExportChoice Choice { get; private set; } = RenderedExportChoice.Cancel;

    public RenderedExportWarningWindow(
        string exportFormat,
        bool hasSelectedSections)
    {
        InitializeComponent();

        string formatName = exportFormat.Equals("pdf", StringComparison.OrdinalIgnoreCase)
            ? "PDF"
            : "XPS";
        TitleText.Text = $"Large {formatName} exports may fail";
        ExplanationText.Text =
            "This is a limitation of OneNote's export engine. Choose how you want to export:";
        WholeNotebookButtonText.Text =
            $"Export the whole notebook as one {formatName} file";

        if (hasSelectedSections)
        {
            OneNotePackageButton.IsEnabled = false;
            OneNotePackageButtonText.Text =
                "OneNote package unavailable with selected sections";
            OneNotePackageButton.ToolTip =
                "Available when whole notebooks are selected.";
        }
    }

    private void Complete(RenderedExportChoice choice)
    {
        Choice = choice;
        DialogResult = choice != RenderedExportChoice.Cancel;
        Close();
    }

    private void SectionsButton_Click(object sender, RoutedEventArgs e)
        => Complete(RenderedExportChoice.IndividualSections);

    private void OneNotePackageButton_Click(object sender, RoutedEventArgs e)
        => Complete(RenderedExportChoice.OneNotePackage);

    private void WholeNotebookButton_Click(object sender, RoutedEventArgs e)
        => Complete(RenderedExportChoice.WholeNotebook);

    private void CancelButton_Click(object sender, RoutedEventArgs e)
        => Complete(RenderedExportChoice.Cancel);
}
