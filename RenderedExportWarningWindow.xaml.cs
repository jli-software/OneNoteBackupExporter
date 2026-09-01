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
        int notebookCount,
        bool hasSelectedSections)
    {
        InitializeComponent();

        string formatName = exportFormat.Equals("pdf", StringComparison.OrdinalIgnoreCase)
            ? "PDF"
            : "XPS";
        string notebookLabel = notebookCount == 1
            ? "the selected notebook"
            : $"the {notebookCount} selected notebooks";

        TitleText.Text = $"Whole-notebook {formatName} exports can fail when they are large";
        ExplanationText.Text =
            $"OneNote Desktop must render {notebookLabel} through its own {formatName} engine. " +
            "For large notebooks this engine can time out or disconnect from the app.";
        SectionsDescriptionText.Text =
            $"The app loads all exportable sections automatically and creates one {formatName} " +
            "file per section. Locked sections are not included. " +
            "Smaller render jobs are significantly more reliable.";
        WholeNotebookButton.Content =
            $"Try the whole {formatName} export anyway (may fail for large notebooks)";

        if (hasSelectedSections)
        {
            OneNotePackageButton.IsEnabled = false;
            OneNotePackageButton.Content =
                "OneNote package requires no individually selected sections";
            OneNotePackageButton.ToolTip =
                "Cancel, deselect the individual sections, and try again to create a complete package.";
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
