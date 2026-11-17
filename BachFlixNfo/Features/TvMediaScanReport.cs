using System;
using System.Collections.Generic;

namespace BachFlixNfo.Features
{
    internal sealed class TvMediaScanReport
    {
        internal int Examined, Unchanged, Updated, Skipped, Errors, ScanErrors;
        internal int Discovered, Reused, Probed, New, Modified, Deleted, ResolutionChanges, FormatChanges, Cells;
        internal readonly List<string> ErrorDetails = new List<string>();

        internal void Add(TvMediaScanResult result)
        {
            Discovered += result.Discovered; Reused += result.Reused; Probed += result.Probed;
            New += result.New; Modified += result.Modified; Deleted += result.Deleted;
        }

        internal void Record(string message, TvMediaShow show)
        {
            if (!IsError(message.TrimStart())) return;
            ErrorDetails.Add((show == null ? "Scan" : show.Title + " (DB row " + show.Row + ")") + ": " + message.TrimStart());
        }

        internal void CompleteUnfinished(bool interrupted)
        {
            int unfinished = Math.Max(0, Examined - Unchanged - Updated - Skipped - Errors);
            if (interrupted) Skipped += unfinished;
            else Errors += unfinished;
        }

        private static bool IsError(string text)
        {
            return text.StartsWith("ERROR:") || text.StartsWith("Scan stopped") || text.StartsWith("FFProbe/file metadata failure");
        }

        internal static string Color(string message)
        {
            string text = message.TrimStart();
            if (IsError(text) || text.StartsWith("Errors to review:")) return "Red";
            if (text.StartsWith("TV Shows skipped:")) return text.EndsWith(": 0") ? "DarkGray" : "Yellow";
            if (text.StartsWith("TV Shows with errors:") || text.StartsWith("Scan errors:")) return text.EndsWith(": 0") ? "DarkGray" : "Red";
            if (text.StartsWith("SKIPPED:") || text.StartsWith("Unsupported") || text.StartsWith("Scan interrupted") || text.StartsWith("Unreadable") ||
                text.StartsWith("Cache schema") || text.IndexOf("formula", StringComparison.OrdinalIgnoreCase) >= 0) return "Yellow";
            if (text.StartsWith("Google Sheet: Updated") || text.StartsWith("TV Shows updated:") || text.StartsWith("Google Sheets cells updated:")) return "Green";
            if (text.StartsWith("Google Sheet: Queued")) return "Yellow";
            if (text.StartsWith("Google Sheet: Unchanged") || text.StartsWith("TV Shows unchanged:")) return "DarkGray";
            if (text.StartsWith("Resolution:") || text.StartsWith("Format:") || text.StartsWith("Resolution changes:") || text.StartsWith("Format changes:")) return "Magenta";
            if (text.StartsWith("[") || text.StartsWith("61 -") || text.StartsWith("TV Shows examined:") || text.StartsWith("Total execution time:") || text.StartsWith("Errors-only log:")) return "Cyan";
            if (text.StartsWith("Episode progress:") || text.StartsWith("Cache location:")) return "DarkCyan";
            return "Gray";
        }
    }
}
