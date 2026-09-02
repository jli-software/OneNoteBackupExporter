using OneNoteExporter.Models;
using OneNoteExporter.Services;
using System.Diagnostics;
using Xunit;

namespace OneNoteExporter.Tests;

public sealed class HtmlExportTests : IDisposable
{
    private readonly string _testRoot = Path.Combine(
        Path.GetTempPath(), "OneNoteExporterTests", Guid.NewGuid().ToString("N"));

    [Fact]
    public async Task WaitForHtmlPageAsync_WaitsForTheWholePageBundle()
    {
        string pageDirectory = Path.Combine(_testRoot, "staging", "Page");
        string targetPath = Path.Combine(pageDirectory, "index.htm");
        Directory.CreateDirectory(Path.Combine(pageDirectory, "index_files"));
        await File.WriteAllTextAsync(targetPath, "<html><body>Test</body></html>");
        await File.WriteAllBytesAsync(
            Path.Combine(pageDirectory, "index_files", "image.png"),
            [1, 2, 3, 4]);
        var page = new HtmlPageExport(
            "page-id", "Page", "Section", "Section\\Page", targetPath);

        var stopwatch = Stopwatch.StartNew();
        Task lateAsset = Task.Run(async () =>
        {
            await Task.Delay(1500);
            await File.WriteAllTextAsync(
                Path.Combine(pageDirectory, "index_files", "late.css"), "body{}");
        });

        ExportResult result = await OneNoteService.WaitForHtmlPageAsync(
            page, isCloud: false, progress: null, CancellationToken.None);
        Assert.True(lateAsset.IsCompleted);
        await lateAsset;

        Assert.True(result.Success);
        Assert.Equal(targetPath, result.ExportedPath);
        Assert.True(stopwatch.Elapsed >= TimeSpan.FromSeconds(5));
    }

    [Fact]
    public async Task WaitForHtmlPageAsync_HonorsCancellation()
    {
        string targetPath = Path.Combine(_testRoot, "missing", "index.htm");
        var page = new HtmlPageExport(
            "page-id", "Page", "Section", "Page", targetPath);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            OneNoteService.WaitForHtmlPageAsync(
                page, isCloud: false, progress: null, cancellation.Token));
    }

    [Fact]
    public async Task FinalizeHtmlExport_ReplacesTheCompleteBundleAndBuildsIndexes()
    {
        string stagingDirectory = Path.Combine(_testRoot, ".Notebook_html.partial");
        string finalDirectory = Path.Combine(_testRoot, "Notebook_html");
        string pageDirectory = Path.Combine(stagingDirectory, "Section", "Page & Notes");
        string targetPath = Path.Combine(pageDirectory, "index.htm");
        Directory.CreateDirectory(Path.Combine(pageDirectory, "index_files"));
        await File.WriteAllTextAsync(targetPath, "<html><body>New page</body></html>");
        await File.WriteAllTextAsync(
            Path.Combine(pageDirectory, "index_files", "asset.txt"), "asset");
        Directory.CreateDirectory(finalDirectory);
        await File.WriteAllTextAsync(Path.Combine(finalDirectory, "obsolete.txt"), "old");

        var page = new HtmlPageExport(
            "page-id", "Page & Notes", "Section", "Section\\Page & Notes", targetPath);
        var plan = new HtmlExportPlan(
            "Notebook",
            stagingDirectory,
            finalDirectory,
            finalDirectory,
            "Notebook",
            IsCloud: false,
            [page]);
        var result = new ExportResult { Success = true };

        ExportResult finalized = OneNoteService.FinalizeHtmlExport(plan, result);

        Assert.False(File.Exists(Path.Combine(finalDirectory, "obsolete.txt")));
        Assert.True(File.Exists(Path.Combine(
            finalDirectory, "Section", "Page & Notes", "index.htm")));
        Assert.True(File.Exists(Path.Combine(finalDirectory, "index.html")));
        string rootIndex = await File.ReadAllTextAsync(
            Path.Combine(finalDirectory, "index.html"));
        Assert.Contains("Page &amp; Notes", rootIndex);
        Assert.Contains("Page%20%26%20Notes/index.htm", rootIndex);
        Assert.Equal(Path.Combine(finalDirectory, "index.html"), finalized.ExportedPath);
        Assert.Empty(Directory.EnumerateDirectories(
            _testRoot, ".Notebook_html.*.previous"));
    }

    [Fact]
    public async Task FinalizeHtmlExport_SectionExportRebuildsNotebookIndex()
    {
        string collectionDirectory = Path.Combine(_testRoot, "Notebook_html");
        string sectionParent = Path.Combine(collectionDirectory, "Group");
        string stagingDirectory = Path.Combine(sectionParent, ".Section.partial");
        string finalDirectory = Path.Combine(sectionParent, "Section__1234abcd");
        string targetPath = Path.Combine(stagingDirectory, "Page", "index.htm");
        Directory.CreateDirectory(Path.GetDirectoryName(targetPath)!);
        await File.WriteAllTextAsync(targetPath, "<html><body>Page</body></html>");

        string retainedSection = Path.Combine(collectionDirectory, "Retained__deadbeef");
        Directory.CreateDirectory(retainedSection);
        await File.WriteAllTextAsync(
            Path.Combine(retainedSection, "index.html"), "retained");
        await File.WriteAllTextAsync(
            Path.Combine(collectionDirectory, "index.html"), "obsolete link");

        var page = new HtmlPageExport(
            "page-id", "Page", "Section", "Page", targetPath);
        var plan = new HtmlExportPlan(
            "Section",
            stagingDirectory,
            finalDirectory,
            collectionDirectory,
            "Notebook",
            IsCloud: false,
            [page]);

        ExportResult result = OneNoteService.FinalizeHtmlExport(
            plan, new ExportResult { Success = true });

        string collectionIndex = await File.ReadAllTextAsync(
            Path.Combine(collectionDirectory, "index.html"));
        Assert.DoesNotContain("obsolete link", collectionIndex);
        Assert.Contains("Group / Section", collectionIndex);
        Assert.Contains("Retained", collectionIndex);
        Assert.Contains("Group/Section__1234abcd/index.html", collectionIndex);
        Assert.Equal(Path.Combine(collectionDirectory, "index.html"), result.ExportedPath);
    }

    public void Dispose()
    {
        if (Directory.Exists(_testRoot))
            Directory.Delete(_testRoot, recursive: true);
    }
}
