using System.Windows.Threading;
using OneNoteExporter.Models;

namespace OneNoteExporter.Services;

/// <summary>
/// Owns the OneNote COM service on a dedicated STA thread. All operations are
/// dispatched through the thread's message pump and therefore run serially.
/// </summary>
public sealed class OneNoteComWorker : IDisposable
{
    private const int MaxExportAttempts = 2;
    private static readonly TimeSpan ExportRetryDelay = TimeSpan.FromSeconds(5);

    private readonly Thread _thread;
    private readonly SemaphoreSlim _operationGate = new(1, 1);
    private readonly TaskCompletionSource _started =
        new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly TaskCompletionSource _stopped =
        new(TaskCreationOptions.RunContinuationsAsynchronously);

    private Dispatcher? _dispatcher;
    private OneNoteService? _service;
    private int _disposeState;

    private OneNoteComWorker()
    {
        _thread = new Thread(RunMessageLoop)
        {
            IsBackground = true,
            Name = "OneNote COM STA Worker"
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public static async Task<OneNoteComWorker> CreateAsync()
    {
        var worker = new OneNoteComWorker();

        try
        {
            await worker._started.Task.ConfigureAwait(false);
            return worker;
        }
        catch
        {
            worker.Dispose();
            throw;
        }
    }

    public Task<VersionInfo> GetVersionInfoAsync(CancellationToken ct = default)
        => RunSerializedAsync(
            () => InvokeAsync(service => service.GetVersionInfo(), ct), ct);

    public Task<List<NotebookInfo>> GetNotebooksAsync(CancellationToken ct = default)
        => RunSerializedAsync(
            () => InvokeAsync(service => service.GetNotebooks(), ct), ct);

    public Task<List<SectionInfo>> GetSectionsAsync(
        string notebookId,
        CancellationToken ct = default)
        => RunSerializedAsync(
            () => InvokeAsync(service => service.GetSections(notebookId), ct), ct);

    public Task<ExportResult> ExportNotebookAsync(
        string notebookId,
        string destinationPath,
        string exportFormat = "onepkg",
        IProgress<string>? progress = null,
        CancellationToken ct = default)
        => RunSerializedAsync(
            () => ExportAsync(
                service => service.BeginNotebookExport(
                    notebookId, destinationPath, exportFormat, progress),
                exportFormat,
                progress,
                ct),
            ct);

    public Task<ExportResult> ExportSectionAsync(
        SectionInfo section,
        string destinationPath,
        string exportFormat = "onepkg",
        IProgress<string>? progress = null,
        CancellationToken ct = default)
        => RunSerializedAsync(
            () => ExportAsync(
                service => service.BeginSectionExport(
                    section, destinationPath, exportFormat, progress),
                exportFormat,
                progress,
                ct),
            ct);

    private async Task<ExportResult> ExportAsync(
        Func<OneNoteService, PendingExport> beginExport,
        string exportFormat,
        IProgress<string>? progress,
        CancellationToken ct)
    {
        bool connectionRecoveryAttempted = false;

        for (int attempt = 1; attempt <= MaxExportAttempts; attempt++)
        {
            try
            {
                var pending = await InvokeAsync(beginExport, ct).ConfigureAwait(false);
                var result = await OneNoteService.WaitForFileAsync(
                        pending.StagingPath, pending.IsCloud, progress, ct)
                    .ConfigureAwait(false);

                if (!result.Success)
                {
                    result.Message +=
                        " The existing backup was not changed. A temporary staging " +
                        "file may be retained until automatic cleanup.";
                    return result;
                }

                progress?.Report("Finalizing export and replacing the previous backup...");
                ct.ThrowIfCancellationRequested();
                var finalResult = OneNoteService.FinalizeExport(pending, result);
                finalResult.AttemptCount = attempt;
                finalResult.RecoveredAfterRetry = attempt > 1;

                if (finalResult.RecoveredAfterRetry)
                {
                    finalResult.Message +=
                        $" Export succeeded after OneNote communication recovery " +
                        $"(attempt {attempt}/{MaxExportAttempts}).";
                }

                return finalResult;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex) when (
                attempt < MaxExportAttempts &&
                OneNoteService.IsRecoverableComFailure(ex))
            {
                string failure = OneNoteService.DescribeComFailure(ex);
                bool resetConnection = OneNoteService.RequiresComConnectionReset(ex);

                progress?.Report(
                    $"OneNote communication failed: {failure}. " +
                    $"Automatic retry {attempt + 1}/{MaxExportAttempts} in " +
                    $"{ExportRetryDelay.TotalSeconds:0} seconds...");

                await Task.Delay(ExportRetryDelay, ct).ConfigureAwait(false);

                if (resetConnection)
                {
                    progress?.Report("Reconnecting to OneNote Desktop...");

                    try
                    {
                        await RecreateServiceAsync(ct).ConfigureAwait(false);
                        connectionRecoveryAttempted = true;
                        progress?.Report(
                            $"OneNote connection restored. Retrying export " +
                            $"({attempt + 1}/{MaxExportAttempts})...");
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception recoveryException)
                    {
                        return OneNoteService.CreateConnectionRecoveryFailure(
                            ex, recoveryException);
                    }
                }
            }
            catch (Exception ex)
            {
                return OneNoteService.CreateExportFailure(
                    ex, attempt, exportFormat, connectionRecoveryAttempted);
            }
        }

        throw new InvalidOperationException("The export retry loop ended unexpectedly.");
    }

    private async Task RecreateServiceAsync(CancellationToken ct)
        => _ = await InvokeOnWorkerAsync(
            () =>
            {
                var previousService = _service;
                var candidateService = new OneNoteService();

                try
                {
                    candidateService.ValidateConnection();
                }
                catch
                {
                    candidateService.Dispose();
                    throw;
                }

                _service = candidateService;
                previousService?.Dispose();
                return true;
            },
            ct).ConfigureAwait(false);

    private async Task<T> RunSerializedAsync<T>(
        Func<Task<T>> operation,
        CancellationToken ct)
    {
        ObjectDisposedException.ThrowIf(
            Volatile.Read(ref _disposeState) != 0,
            this);

        await _operationGate.WaitAsync(ct).ConfigureAwait(false);

        try
        {
            ObjectDisposedException.ThrowIf(
                Volatile.Read(ref _disposeState) != 0,
                this);
            return await operation().ConfigureAwait(false);
        }
        finally
        {
            _operationGate.Release();
        }
    }

    private async Task<T> InvokeAsync<T>(
        Func<OneNoteService, T> operation,
        CancellationToken ct)
        => await InvokeOnWorkerAsync(
                () => operation(_service
                    ?? throw new InvalidOperationException("OneNote is not initialized.")),
                ct)
            .ConfigureAwait(false);

    private async Task<T> InvokeOnWorkerAsync<T>(
        Func<T> operation,
        CancellationToken ct)
    {
        ObjectDisposedException.ThrowIf(
            Volatile.Read(ref _disposeState) != 0,
            this);

        await _started.Task.ConfigureAwait(false);

        var dispatcher = _dispatcher
            ?? throw new InvalidOperationException("The OneNote COM worker did not start.");

        var dispatcherOperation = dispatcher.InvokeAsync(
            () =>
            {
                ct.ThrowIfCancellationRequested();
                return operation();
            },
            DispatcherPriority.Normal,
            ct);

        return await dispatcherOperation.Task.ConfigureAwait(false);
    }

    private void RunMessageLoop()
    {
        _dispatcher = Dispatcher.CurrentDispatcher;

        try
        {
            _service = new OneNoteService();

            if (Volatile.Read(ref _disposeState) != 0)
            {
                _started.TrySetCanceled();
                return;
            }

            _started.TrySetResult();
            Dispatcher.Run();
        }
        catch (OperationCanceledException) when (Volatile.Read(ref _disposeState) != 0)
        {
            _started.TrySetCanceled();
        }
        catch (Exception ex)
        {
            _started.TrySetException(ex);
        }
        finally
        {
            _service?.Dispose();
            _service = null;
            _stopped.TrySetResult();
        }
    }

    public void Dispose()
    {
        if (Interlocked.CompareExchange(ref _disposeState, 1, 0) != 0)
        {
            _stopped.Task.GetAwaiter().GetResult();
            return;
        }

        _operationGate.Wait();

        try
        {
            var dispatcher = _dispatcher;
            if (dispatcher != null && !dispatcher.HasShutdownStarted)
                dispatcher.BeginInvokeShutdown(DispatcherPriority.Send);

            if (Thread.CurrentThread != _thread)
                _stopped.Task.GetAwaiter().GetResult();
        }
        finally
        {
            _operationGate.Release();
        }
    }
}
