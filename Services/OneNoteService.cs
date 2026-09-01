using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteExporter.Models;

namespace OneNoteExporter.Services;

internal sealed record PendingExport(
    string StagingPath,
    string FinalPath,
    bool IsCloud);

/// <summary>
/// Synchronous OneNote COM API implementation owned by <see cref="OneNoteComWorker"/>.
/// Construction, all calls and disposal must stay on that worker's STA thread.
/// </summary>
internal sealed class OneNoteService : IDisposable
{
    private Application? _oneNote;
    private bool _disposed = false;

    public OneNoteService()
    {
        try
        {
            _oneNote = new Application();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "Failed to initialize OneNote Desktop. " +
                "Please ensure that OneNote 2016 is installed.", ex);
        }
    }

    // ── Public API ───────────────────────────────────────────────────────────

    public VersionInfo GetVersionInfo()
    {
        var info = new VersionInfo { OneNoteInstalled = _oneNote != null };

        if (_oneNote != null)
        {
            try
            {
                _oneNote.GetHierarchy("", HierarchyScope.hsNotebooks, out _);
                info.OneNoteVersion = "OneNote Desktop (COM API available)";
            }
            catch
            {
                info.OneNoteVersion = "Unknown";
            }
        }

        return info;
    }

    public List<NotebookInfo> GetNotebooks()
    {
        if (_oneNote == null)
            throw new InvalidOperationException("OneNote is not initialized.");

        var notebooks = new List<NotebookInfo>();

        _oneNote.GetHierarchy("", HierarchyScope.hsNotebooks, out string xml);

        var xdoc = XDocument.Parse(xml);
        var ns   = xdoc.Root?.Name.Namespace;
        if (ns == null) return notebooks;

        foreach (var nb in xdoc.Descendants(ns + "Notebook"))
        {
            notebooks.Add(new NotebookInfo
            {
                Id               = nb.Attribute("ID")?.Value              ?? "",
                Name             = nb.Attribute("name")?.Value            ?? "Unnamed",
                Path             = nb.Attribute("path")?.Value            ?? "",
                LastModified     = nb.Attribute("lastModifiedTime")?.Value ?? "",
                IsCurrentlyViewed = nb.Attribute("isCurrentlyViewed")?.Value == "true"
            });
        }

        return notebooks;
    }

    public void ValidateConnection()
    {
        if (_oneNote == null)
            throw new InvalidOperationException("OneNote is not initialized.");

        _oneNote.GetHierarchy("", HierarchyScope.hsNotebooks, out _);
    }

    /// <summary>
    /// Starts a notebook export and returns as soon as OneNote accepted the
    /// publish request. File completion is monitored outside the STA thread.
    /// </summary>
    public PendingExport BeginNotebookExport(
        string notebookId,
        string destinationPath,
        string exportFormat = "onepkg",
        IProgress<string>? progress = null)
    {
        if (_oneNote == null)
            throw new InvalidOperationException("OneNote is not initialized.");

        // Get notebook name and path from hierarchy
        _oneNote.GetHierarchy(notebookId, HierarchyScope.hsSelf, out string xml);
        var xdoc         = XDocument.Parse(xml);
        var notebookName = xdoc.Root?.Attribute("name")?.Value ?? "Notebook";
        var notebookPath = xdoc.Root?.Attribute("path")?.Value ?? "";

        bool isCloud = notebookPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                       notebookPath.StartsWith("http://",  StringComparison.OrdinalIgnoreCase);

        progress?.Report($"Exporting: {notebookName} (format: {exportFormat})");

        string fileExtension = exportFormat.ToLowerInvariant() switch
        {
            "xps" => ".xps",
            "pdf" => ".pdf",
            _     => ".onepkg"
        };

        PublishFormat publishFormat = exportFormat.ToLowerInvariant() switch
        {
            "xps" => PublishFormat.pfXPS,
            "pdf" => PublishFormat.pfPDF,
            _     => PublishFormat.pfOneNotePackage
        };

        var sanitizedName = string.Join("_", notebookName.Split(Path.GetInvalidFileNameChars()));
        var finalPath     = Path.Combine(destinationPath, sanitizedName + fileExtension);
        var stagingPath   = CreateStagingPath(destinationPath, sanitizedName, fileExtension);

        Directory.CreateDirectory(destinationPath);
        CleanupStaleStagingFiles(
            destinationPath, sanitizedName, fileExtension, DateTime.UtcNow.AddDays(-2));

        _oneNote.OpenHierarchy(notebookPath, "", out string openedId, CreateFileType.cftNone);

        progress?.Report("OneNote is writing to a protected staging file...");

        try
        {
            _oneNote.Publish(openedId, stagingPath, publishFormat, "");
        }
        catch
        {
            TryDeleteStagingFile(stagingPath);
            throw;
        }

        return new PendingExport(stagingPath, finalPath, isCloud);
    }

    private static string CreateStagingPath(
        string directory,
        string sanitizedName,
        string fileExtension)
        => Path.Combine(
            directory,
            $".{sanitizedName}.{Guid.NewGuid():N}.partial{fileExtension}");

    private static void TryDeleteStagingFile(string stagingPath)
    {
        try
        {
            if (File.Exists(stagingPath))
                File.Delete(stagingPath);
        }
        catch (IOException)
        {
            // OneNote may still hold the file after a broken RPC call. A unique
            // staging path keeps later attempts and the last good export safe.
        }
        catch (UnauthorizedAccessException)
        {
            // Preserve the original export failure; cleanup can happen later.
        }
    }

    private static void CleanupStaleStagingFiles(
        string directory,
        string sanitizedName,
        string fileExtension,
        DateTime olderThanUtc)
    {
        string searchPattern = $".{sanitizedName}.*.partial{fileExtension}";

        try
        {
            foreach (string path in Directory.EnumerateFiles(directory, searchPattern))
            {
                try
                {
                    if (File.GetLastWriteTimeUtc(path) < olderThanUtc)
                        File.Delete(path);
                }
                catch (IOException)
                {
                    // An active or locked OneNote write must never be disturbed.
                }
                catch (UnauthorizedAccessException)
                {
                    // Cleanup is best-effort and must not prevent an export.
                }
            }
        }
        catch (IOException)
        {
            // Directory enumeration is best-effort.
        }
        catch (UnauthorizedAccessException)
        {
            // The export itself will provide the actionable permission error.
        }
    }

    /// <summary>
    /// Returns all exportable sections in a notebook (including those inside section groups).
    /// Skips encrypted and recycle-bin entries. Called by OneNoteComWorker on its STA thread.
    /// </summary>
    public List<SectionInfo> GetSections(string notebookId)
    {
        if (_oneNote == null)
            throw new InvalidOperationException("OneNote is not initialized.");

        var sections = new List<SectionInfo>();

        // Notebook metadata (name, path, cloud flag)
        _oneNote.GetHierarchy(notebookId, HierarchyScope.hsSelf, out string nbXml);
        var nbDoc  = XDocument.Parse(nbXml);
        var nbName = nbDoc.Root?.Attribute("name")?.Value ?? "Notebook";
        var nbPath = nbDoc.Root?.Attribute("path")?.Value ?? "";
        bool isCloud = nbPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                       nbPath.StartsWith("http://",  StringComparison.OrdinalIgnoreCase);

        // Full section hierarchy
        _oneNote.GetHierarchy(notebookId, HierarchyScope.hsSections, out string xml);
        var xdoc = XDocument.Parse(xml);
        var ns   = xdoc.Root?.Name.Namespace;
        if (ns == null) return sections;

        CollectSections(xdoc.Root!, ns, notebookId, nbName, nbPath, isCloud, "", sections);
        return sections;
    }

    /// <summary>
    /// Starts a section export and returns as soon as OneNote accepted the
    /// publish request. File completion is monitored outside the STA thread.
    /// </summary>
    public PendingExport BeginSectionExport(
        SectionInfo section,
        string destinationPath,
        string exportFormat = "onepkg",
        IProgress<string>? progress = null)
    {
        if (_oneNote == null)
            throw new InvalidOperationException("OneNote is not initialized.");

        progress?.Report($"Exporting section: {section.Name} (format: {exportFormat})");

        string fileExtension = exportFormat.ToLowerInvariant() switch
        {
            "xps" => ".xps",
            "pdf" => ".pdf",
            _     => ".onepkg"
        };

        PublishFormat publishFormat = exportFormat.ToLowerInvariant() switch
        {
            "xps" => PublishFormat.pfXPS,
            "pdf" => PublishFormat.pfPDF,
            _     => PublishFormat.pfOneNotePackage
        };

        // Sections go into a notebook-named subfolder to prevent filename clashes
        var sanitizedNb  = string.Join("_", section.NotebookName.Split(Path.GetInvalidFileNameChars()));
        var sanitizedSec = string.Join("_", section.Name.Split(Path.GetInvalidFileNameChars()));
        var notebookDir  = Path.Combine(destinationPath, sanitizedNb);
        var finalPath    = Path.Combine(notebookDir, sanitizedSec + fileExtension);
        var stagingPath  = CreateStagingPath(notebookDir, sanitizedSec, fileExtension);

        Directory.CreateDirectory(notebookDir);
        CleanupStaleStagingFiles(
            notebookDir, sanitizedSec, fileExtension, DateTime.UtcNow.AddDays(-2));

        progress?.Report("Opening parent notebook...");
        _oneNote.OpenHierarchy(section.NotebookPath, "", out _, CreateFileType.cftNone);

        progress?.Report("OneNote is writing to a protected staging file...");

        try
        {
            _oneNote.Publish(section.Id, stagingPath, publishFormat, "");
        }
        catch
        {
            TryDeleteStagingFile(stagingPath);
            throw;
        }

        return new PendingExport(stagingPath, finalPath, section.IsCloud);
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    /// <summary>
    /// Polls until the exported file exists and its size has been stable for 10 seconds.
    /// Cloud notebooks get a 30-minute timeout; local notebooks get 20 minutes.
    /// </summary>
    internal static async Task<ExportResult> WaitForFileAsync(
        string fullPath,
        bool isCloud,
        IProgress<string>? progress,
        CancellationToken ct)
    {
        int  maxAttempts     = isCloud ? 900 : 600; // 30 min : 20 min
        int  checkIntervalMs = 2000;
        long previousSize    = -1;
        int  stableCount     = 0;
        bool fileCreated     = false;
        long fileSize        = 0;

        // Initial wait – give OneNote a moment to start writing
        await Task.Delay(checkIntervalMs, ct).ConfigureAwait(false);

        for (int i = 0; i < maxAttempts; i++)
        {
            ct.ThrowIfCancellationRequested();

            if (File.Exists(fullPath))
            {
                fileCreated = true;
                fileSize    = new FileInfo(fullPath).Length;

                if (fileSize > 0 && fileSize == previousSize)
                {
                    stableCount++;
                    if (stableCount >= 5)
                    {
                        progress?.Report($"✓ File written! Final size: {FormatBytes(fileSize)}");
                        return new ExportResult
                        {
                            Success  = true,
                            Message  = $"Exported successfully ({FormatBytes(fileSize)})",
                            ExportedPath = fullPath
                        };
                    }
                    progress?.Report($"Size stable at {FormatBytes(fileSize)} ({stableCount * 2}s)...");
                }
                else if (fileSize > 0)
                {
                    double mbps = previousSize > 0
                        ? (fileSize - previousSize) / 1024.0 / 1024.0 / (checkIntervalMs / 1000.0)
                        : 0;
                    progress?.Report(mbps > 0
                        ? $"Writing: {FormatBytes(fileSize)} (~{mbps:0.0} MB/s)"
                        : $"Writing: {FormatBytes(fileSize)}");
                    stableCount = 0;
                }
                else
                {
                    progress?.Report("File created, waiting for data...");
                    stableCount = 0;
                }

                previousSize = fileSize;
            }
            else
            {
                if (i % 10 == 0)
                    progress?.Report($"Waiting for file creation... ({i * checkIntervalMs / 1000}s)");
            }

            await Task.Delay(checkIntervalMs, ct).ConfigureAwait(false);
        }

        // Timed out
        int timeoutMin = maxAttempts * checkIntervalMs / 60_000;
        if (fileCreated && fileSize == 0)
            return new ExportResult { Success = false, Message = "File was created but is empty (0 bytes)." };

        return new ExportResult
        {
            Success = false,
            Message = $"Timeout after {timeoutMin} min. OneNote may still be writing in the background. " +
                      $"Check {Path.GetDirectoryName(fullPath)} again in a few minutes."
        };
    }

    internal static ExportResult FinalizeExport(
        PendingExport pending,
        ExportResult result)
    {
        File.Move(pending.StagingPath, pending.FinalPath, overwrite: true);
        result.ExportedPath = pending.FinalPath;
        return result;
    }

    internal static bool IsRecoverableComFailure(Exception exception)
        => exception is COMException ex && ex.HResult is
            unchecked((int)0x800706BA) or // RPC_S_SERVER_UNAVAILABLE
            unchecked((int)0x80010108) or // RPC_E_DISCONNECTED
            unchecked((int)0x800401FD) or // CO_E_OBJNOTCONNECTED
            unchecked((int)0x8001010A) or // RPC_E_SERVERCALL_RETRYLATER
            unchecked((int)0x80010001);   // RPC_E_CALL_REJECTED

    internal static bool RequiresComConnectionReset(Exception exception)
        => exception is COMException ex && ex.HResult is
            unchecked((int)0x800706BA) or
            unchecked((int)0x80010108) or
            unchecked((int)0x800401FD);

    internal static string DescribeComFailure(Exception exception)
    {
        if (exception is not COMException ex)
            return exception.Message;

        return ex.HResult switch
        {
            unchecked((int)0x800706BA) => "RPC connection lost (0x800706BA)",
            unchecked((int)0x80010108) => "COM object disconnected (0x80010108)",
            unchecked((int)0x800401FD) => "COM object is no longer connected (0x800401FD)",
            unchecked((int)0x80010007) => "OneNote COM server stopped (0x80010007)",
            unchecked((int)0x80010012) => "OneNote COM server unavailable (0x80010012)",
            unchecked((int)0x80042023) => "OneNote internal timeout (0x80042023)",
            unchecked((int)0x8001010A) => "OneNote is busy (0x8001010A)",
            unchecked((int)0x80010001) => "OneNote rejected the COM call (0x80010001)",
            _ => $"COM error 0x{ex.HResult:X8}"
        };
    }

    internal static ExportResult CreateConnectionRecoveryFailure(
        Exception originalException,
        Exception recoveryException)
        => new()
        {
            Success = false,
            Message = $"OneNote communication failed: " +
                      $"{DescribeComFailure(originalException)}. " +
                      "The app tried to reconnect automatically, but OneNote Desktop " +
                      $"could not be initialized again: {recoveryException.Message} " +
                      "Open OneNote, wait until synchronization has finished, and retry. " +
                      "The existing backup was not changed."
        };

    internal static ExportResult CreateExportFailure(
        Exception exception,
        int attemptCount,
        string exportFormat,
        bool connectionRecoveryAttempted)
    {
        string attemptSuffix = attemptCount > 1
            ? $" after {attemptCount} attempts"
            : "";
        bool isRenderedFormat = exportFormat.Equals("pdf", StringComparison.OrdinalIgnoreCase) ||
                                exportFormat.Equals("xps", StringComparison.OrdinalIgnoreCase);
        string renderedFormatAdvice = isRenderedFormat
            ? " OneNote's PDF/XPS renderer is less reliable for large notebooks; " +
              "try .onepkg if the problem continues."
            : "";
        string recoveryDetails = connectionRecoveryAttempted
            ? " The app reconnected to OneNote automatically, but the connection failed again."
            : "";

        return exception switch
        {
            UnauthorizedAccessException ex => new ExportResult
            {
                Success = false,
                Message = $"Access denied: {ex.Message} The existing backup was not changed."
            },
            COMException ex when ex.HResult == unchecked((int)0x8004201A) => new ExportResult
            {
                Success = false,
                Message = "OneNote error 0x8004201A: The export file already exists."
            },
            COMException ex when ex.HResult == unchecked((int)0x800706BA) => new ExportResult
            {
                Success = false,
                Message = $"OneNote COM/RPC connection lost (0x800706BA){attemptSuffix}." +
                          recoveryDetails +
                          " This is a OneNote Desktop communication failure, not a " +
                          "destination-folder error. Open OneNote, wait until the notebook " +
                          "is fully synchronized, and retry." + renderedFormatAdvice
                          + " The existing backup was not changed."
            },
            COMException ex when ex.HResult == unchecked((int)0x80042023) => new ExportResult
            {
                Success = false,
                Message = $"OneNote internal timeout (0x80042023){attemptSuffix}. " +
                          "OneNote did not finish the export request in time." +
                          renderedFormatAdvice
            },
            COMException ex when ex.HResult == unchecked((int)0x8004201D) => new ExportResult
            {
                Success = false,
                Message = "OneNote reports that the notebook is not fully synchronized " +
                          "(0x8004201D). Open it in OneNote, wait for synchronization, " +
                          "and retry."
            },
            COMException ex when ex.HResult == unchecked((int)0x80042030) => new ExportResult
            {
                Success = false,
                Message = "A modal OneNote dialog is blocking the export (0x80042030). " +
                          "Close the dialog in OneNote and retry."
            },
            COMException ex when IsRecoverableComFailure(ex) => new ExportResult
            {
                Success = false,
                Message = $"OneNote communication failed: {DescribeComFailure(ex)}" +
                          $"{attemptSuffix}." + recoveryDetails +
                          " Open OneNote and retry after it is responsive and synchronized." +
                          renderedFormatAdvice
            },
            COMException ex => new ExportResult
            {
                Success = false,
                Message = $"OneNote COM error 0x{ex.HResult:X8}: {ex.Message}"
            },
            _ => new ExportResult
            {
                Success = false,
                Message = $"Error during export: {exception.Message} " +
                          "The existing backup was not changed."
            }
        };
    }

    /// <summary>
    /// Recursively walks a hierarchy XML element, collecting Section entries
    /// and descending into SectionGroups. Skips encrypted sections and the recycle bin.
    /// </summary>
    private static void CollectSections(
        XElement parent, XNamespace ns,
        string notebookId, string notebookName, string notebookPath, bool isCloud,
        string groupName, List<SectionInfo> sections)
    {
        foreach (var el in parent.Elements())
        {
            if (el.Name == ns + "Section")
            {
                if (el.Attribute("encrypted")?.Value == "true") continue;

                sections.Add(new SectionInfo
                {
                    Id           = el.Attribute("ID")?.Value   ?? "",
                    Name         = el.Attribute("name")?.Value ?? "Unnamed",
                    NotebookId   = notebookId,
                    NotebookName = notebookName,
                    NotebookPath = notebookPath,
                    GroupName    = groupName,
                    IsCloud      = isCloud
                });
            }
            else if (el.Name == ns + "SectionGroup")
            {
                if (el.Attribute("isRecycleBin")?.Value == "true") continue;

                var gName = el.Attribute("name")?.Value ?? "Group";
                CollectSections(el, ns, notebookId, notebookName, notebookPath, isCloud,
                    string.IsNullOrEmpty(groupName) ? gName : $"{groupName} / {gName}",
                    sections);
            }
        }
    }

    private static string FormatBytes(long bytes)
    {
        string[] units = { "Bytes", "KB", "MB", "GB" };
        double   value = bytes;
        int      order = 0;
        while (value >= 1024 && order < units.Length - 1) { order++; value /= 1024; }
        return $"{value:0.##} {units[order]}";
    }

    // ── IDisposable ──────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;

        if (_oneNote != null)
        {
            try
            {
                Marshal.ReleaseComObject(_oneNote);
                _oneNote = null;
            }
            catch { /* ignore cleanup errors */ }
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
