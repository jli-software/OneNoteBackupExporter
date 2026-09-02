using System.Diagnostics;
using System.IO;
using OneNoteExporter.Models;

namespace OneNoteExporter.Helpers;

public static class FileHelper
{
    /// <summary>Returns the user's Downloads folder path.</summary>
    public static string GetDefaultDownloadsPath()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        return Path.Combine(home, "Downloads");
    }

    /// <summary>Returns the path to OneNote's local backup folder.</summary>
    public static string GetOneNoteBackupPath()
    {
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(local, "Microsoft", "OneNote", "16.0", "Sicherung");
    }

    /// <summary>Checks whether local OneNote backups exist and are readable.</summary>
    public static BackupAvailability CheckLocalBackupAvailable()
    {
        var backupPath = GetOneNoteBackupPath();

        if (string.IsNullOrEmpty(backupPath))
            return new BackupAvailability
            {
                Available = false,
                Reason    = "path_error",
                Message   = "Backup path could not be determined."
            };

        if (!Directory.Exists(backupPath))
            return new BackupAvailability
            {
                Available = false,
                Reason    = "folder_not_found",
                Message   = "Backup folder not found. OneNote may not have created backups yet.",
                Path      = backupPath
            };

        string[] items;
        try
        {
            items = Directory.GetFileSystemEntries(backupPath);
        }
        catch (UnauthorizedAccessException)
        {
            return new BackupAvailability
            {
                Available = false,
                Reason    = "access_error",
                Message   = "Access denied to backup folder.",
                Path      = backupPath
            };
        }
        catch (Exception ex)
        {
            return new BackupAvailability
            {
                Available = false,
                Reason    = "read_error",
                Message   = UserFacingError.Describe(
                    ex,
                    "The OneNote backup folder could not be read."),
                Path      = backupPath
            };
        }

        if (items.Length == 0)
            return new BackupAvailability
            {
                Available = false,
                Reason    = "folder_empty",
                Message   = "Backup folder is empty. Open OneNote and wait for backups to be created.",
                Path      = backupPath
            };

        return new BackupAvailability
        {
            Available     = true,
            Message       = $"Local backup available ({items.Length} notebooks)",
            Path          = backupPath,
            NotebookCount = items.Length
        };
    }

    /// <summary>
    /// Copies the entire OneNote local backup folder to <paramref name="destPath"/>.
    /// </summary>
    public static ExportResult CopyLocalBackup(
        string destPath,
        CancellationToken ct = default)
    {
        var backupPath = GetOneNoteBackupPath();

        if (string.IsNullOrEmpty(backupPath))
            return new ExportResult { Success = false, Message = "Could not determine OneNote backup path." };

        try
        {
            ct.ThrowIfCancellationRequested();
            Directory.CreateDirectory(destPath);
            CopyDirectory(backupPath, destPath, ct);
            return new ExportResult { Success = true, Message = "Local backup copied successfully.", ExportedPath = destPath };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            return new ExportResult
            {
                Success = false,
                Message = UserFacingError.Describe(
                    ex,
                    "The local OneNote backup could not be copied.")
            };
        }
    }

    /// <summary>Opens a folder in Windows Explorer.</summary>
    public static void OpenFolder(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName        = "explorer.exe",
                Arguments       = path,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Debug.WriteLine(
                UserFacingError.Describe(ex, "The destination folder could not be opened."));
        }
    }

    // ── private helpers ──────────────────────────────────────────────────────

    private static void CopyDirectory(string src, string dst, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        Directory.CreateDirectory(dst);

        foreach (var file in Directory.GetFiles(src))
        {
            ct.ThrowIfCancellationRequested();
            var dest = Path.Combine(dst, Path.GetFileName(file));
            File.Copy(file, dest, overwrite: true);
            ct.ThrowIfCancellationRequested();
        }

        foreach (var dir in Directory.GetDirectories(src))
        {
            ct.ThrowIfCancellationRequested();
            var dest = Path.Combine(dst, Path.GetFileName(dir));
            CopyDirectory(dir, dest, ct);
        }
    }
}
