using System;
using System.Threading;

namespace BachFlixNfo
{
    internal static class ApiRetry
    {
        internal static bool Transient(int status) { return status == 408 || status == 429 || status == 500 || status == 502 || status == 503 || status == 504; }
        internal static void Pause(int attempt, TimeSpan? retryAfter, Action<TimeSpan> delay)
        {
            var wait = retryAfter ?? TimeSpan.FromSeconds(attempt + 1);
            // Never retry earlier than a long server-requested wait: fail instead.
            if (wait > TimeSpan.FromSeconds(30)) throw new InvalidOperationException("Service requested a retry delay over 30 seconds; retry the run later.");
            (delay ?? Thread.Sleep)(wait < TimeSpan.Zero ? TimeSpan.Zero : wait);
        }
    }
}
