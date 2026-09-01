using System.Windows.Threading;
using OneNoteExporter.Models;

namespace OneNoteExporter.Services;

/// <summary>
/// Owns the OneNote COM service on a dedicated STA thread. All operations are
/// dispatched through the thread's message pump and therefore run serially.
/// </summary>
public sealed class OneNoteComWorker : IDisposable
{
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
                progress,
                ct),
            ct);

    private async Task<ExportResult> ExportAsync(
        Func<OneNoteService, PendingExport> beginExport,
        IProgress<string>? progress,
        CancellationToken ct)
    {
        try
        {
            var pending = await InvokeAsync(beginExport, ct).ConfigureAwait(false);
            return await OneNoteService.WaitForFileAsync(
                    pending.FullPath, pending.IsCloud, progress, ct)
                .ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            return OneNoteService.CreateExportFailure(ex);
        }
    }

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
                return operation(_service
                    ?? throw new InvalidOperationException("OneNote is not initialized."));
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
