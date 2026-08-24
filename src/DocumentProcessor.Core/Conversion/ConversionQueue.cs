using System.Collections.Concurrent;

namespace DocumentProcessor.Core.Conversion;

/// <summary>
/// Coalesces concurrent conversion requests into shared LibreOffice invocations.
/// <para>
/// LibreOffice spends roughly 450 ms starting before it looks at the first document, and nothing
/// after the first costs that again. Under load that startup is pure waste: eight requests arriving
/// together previously meant eight processes each paying it. Here they queue, and whichever
/// dispatcher next frees a slot takes the whole waiting set as one invocation — measured at 715 ms
/// for eight small documents against 3,262 ms for eight separate conversions.
/// </para>
/// <para>
/// The batching is opportunistic, never speculative: a request that can start immediately does,
/// and nothing is ever held back hoping a companion arrives. An idle server therefore behaves
/// exactly as it did before, and only a server that is already queueing gets the benefit — which is
/// the only time it is available to take.
/// </para>
/// </summary>
internal static class ConversionQueue
{
    /// <summary>
    /// A conversion waiting for, or occupying, a LibreOffice invocation.
    /// <para>
    /// <see cref="TryClaim"/> is what keeps the three ways a request can leave the queue — being
    /// dispatched, timing out, and being cancelled — from racing each other. Exactly one of them
    /// wins, so a request cannot both be handed to LibreOffice and reported to its caller as timed
    /// out.
    /// </para>
    /// </summary>
    private sealed class PendingRequest(ConversionItem item, CancellationToken callerToken)
    {
        private int _claimed;

        public ConversionItem Item { get; } = item;
        public CancellationToken CallerToken { get; } = callerToken;

        public TaskCompletionSource Completion { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public CancellationTokenSource? QueueDeadline { get; set; }

        public bool TryClaim() => Interlocked.Exchange(ref _claimed, 1) == 0;

        /// <summary>Starts the clocks that can take this request out of the queue before it runs.</summary>
        public void ArmWhileQueued(TimeSpan queueTimeout)
        {
            // Load-shedding: a request that has waited out its queue budget fails fast rather than
            // holding a connection open behind a saturated host.
            var deadline = new CancellationTokenSource(queueTimeout);
            QueueDeadline = deadline;
            deadline.Token.Register(() =>
            {
                if (TryClaim())
                    Completion.TrySetException(LibreOfficeGate.SaturatedException(queueTimeout));
            });

            CallerToken.Register(() =>
            {
                if (TryClaim())
                    Completion.TrySetCanceled(CallerToken);
            });
        }

        public void Finish(Exception? error)
        {
            QueueDeadline?.Dispose();

            // Honour a cancellation that arrived after dispatch. The document may well have been
            // converted anyway — a batch is not abandoned because one participant walked away — but
            // the caller asked to stop, so that is what they are told.
            if (CallerToken.IsCancellationRequested)
                Completion.TrySetCanceled(CallerToken);
            else if (error is not null)
                Completion.TrySetException(error);
            else
                Completion.TrySetResult();
        }
    }

    /// <summary>Requests that can share an invocation: same LibreOffice configuration, same target format.</summary>
    private sealed class Lane
    {
        public Queue<PendingRequest> Pending { get; } = new();
        public int ActiveDispatchers { get; set; }
    }

    private static readonly ConcurrentDictionary<(LibreOfficeSettings Settings, string Extension), Lane> Lanes = new();

    /// <summary>
    /// Converts one document, sharing an invocation with any others waiting on the same
    /// configuration. Completes when that document's result has been written.
    /// </summary>
    public static Task EnqueueAsync(
        LibreOfficeSettings settings, ConversionItem item, string targetExtension, CancellationToken cancellationToken)
    {
        var lane = Lanes.GetOrAdd((settings, targetExtension), _ => new Lane());
        var request = new PendingRequest(item, cancellationToken);

        bool startDispatcher;
        lock (lane)
        {
            lane.Pending.Enqueue(request);

            // More dispatchers than conversion slots would just queue on the gate, and each one
            // holds a batch hostage while it waits.
            startDispatcher = lane.ActiveDispatchers < LibreOfficeGate.Limit;
            if (startDispatcher)
                lane.ActiveDispatchers++;
        }

        request.ArmWhileQueued(settings.QueueTimeout);

        if (startDispatcher)
            _ = Task.Run(() => DispatchLoopAsync(lane, settings, targetExtension), CancellationToken.None);

        return request.Completion.Task;
    }

    private static async Task DispatchLoopAsync(Lane lane, LibreOfficeSettings settings, string targetExtension)
    {
        try
        {
            while (true)
            {
                // The slot is taken before the queue is drained, not after: everything that arrives
                // while this dispatcher waits for a slot can then join the batch it is about to run.
                var slot = await LibreOfficeGate.TryEnterAsync(settings.QueueTimeout, CancellationToken.None)
                    .ConfigureAwait(false);

                List<PendingRequest> batch;
                lock (lane)
                {
                    batch = DrainLocked(lane, settings.MaxBatchSize);
                    if (batch.Count == 0)
                    {
                        slot?.Dispose();
                        return;
                    }
                }

                if (slot is null)
                {
                    var saturated = LibreOfficeGate.SaturatedException(settings.QueueTimeout);
                    foreach (var request in batch)
                        request.Finish(saturated);
                    continue;
                }

                try
                {
                    await RunBatchAsync(batch, settings, targetExtension).ConfigureAwait(false);
                }
                finally
                {
                    slot.Dispose();
                }
            }
        }
        finally
        {
            // Releasing the dispatcher slot and deciding whether work is left must happen together.
            // Apart, a request enqueued in the gap would see a full dispatcher count, decline to
            // start one, and then wait forever behind a dispatcher that has just exited.
            lock (lane)
            {
                lane.ActiveDispatchers--;

                if (lane.Pending.Count > 0 && lane.ActiveDispatchers < LibreOfficeGate.Limit)
                {
                    lane.ActiveDispatchers++;
                    _ = Task.Run(() => DispatchLoopAsync(lane, settings, targetExtension), CancellationToken.None);
                }
            }
        }
    }

    private static List<PendingRequest> DrainLocked(Lane lane, int maxBatchSize)
    {
        var batch = new List<PendingRequest>();

        while (batch.Count < maxBatchSize && lane.Pending.TryDequeue(out var request))
        {
            // Requests already taken by their own timeout or cancellation are simply dropped here.
            if (request.TryClaim())
                batch.Add(request);
        }

        return batch;
    }

    private static async Task RunBatchAsync(List<PendingRequest> batch, LibreOfficeSettings settings, string targetExtension)
    {
        using var cancellation = new AggregateCancellation(batch);

        try
        {
            var outcomes = await LibreOfficeRunner.ConvertBatchAsync(
                settings, batch.Select(r => r.Item).ToList(), targetExtension, cancellation.Token).ConfigureAwait(false);

            for (var i = 0; i < batch.Count; i++)
                batch[i].Finish(outcomes[i].Error);
        }
        catch (Exception ex)
        {
            // Whole-invocation faults — LibreOffice missing, timed out, killed — belong to everyone
            // in the batch.
            foreach (var request in batch)
                request.Finish(ex);
        }
    }

    /// <summary>
    /// Cancels the LibreOffice process only once every participant has given up.
    /// <para>
    /// One caller disconnecting must not abort a conversion five other tenants are still waiting
    /// on. For the single-document case this collapses to the caller's own token, preserving the
    /// rule that a cancelled request kills its subprocess rather than leaving it running.
    /// </para>
    /// </summary>
    private sealed class AggregateCancellation : IDisposable
    {
        private readonly CancellationTokenSource _source = new();
        private readonly List<CancellationTokenRegistration> _registrations;
        private int _remaining;

        public AggregateCancellation(IReadOnlyList<PendingRequest> batch)
        {
            _registrations = new List<CancellationTokenRegistration>(batch.Count);

            // A participant whose token can never be cancelled is one that never gives up, so it is
            // not counted — and a batch containing one is never cancelled at all.
            var cancellable = batch.Where(r => r.CallerToken.CanBeCanceled).ToList();
            if (cancellable.Count < batch.Count)
                return;

            _remaining = cancellable.Count;
            foreach (var request in cancellable)
                _registrations.Add(request.CallerToken.Register(OnParticipantCancelled));
        }

        public CancellationToken Token => _source.Token;

        private void OnParticipantCancelled()
        {
            if (Interlocked.Decrement(ref _remaining) == 0)
                _source.Cancel();
        }

        public void Dispose()
        {
            foreach (var registration in _registrations)
                registration.Dispose();

            _source.Dispose();
        }
    }
}
