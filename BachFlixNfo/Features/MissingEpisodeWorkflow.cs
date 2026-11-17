using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace BachFlixNfo.Features
{
    internal sealed class MissingEpisodeOptions
    {
        internal bool Update;
        internal int Limit = int.MaxValue;
        internal int? Row;
        internal static MissingEpisodeOptions Parse(string[] args)
        {
            var options = new MissingEpisodeOptions();
            if (args.Length == 0 || (args[0] != "--missing-episodes-preview" && args[0] != "--missing-episodes-update"))
                throw new ArgumentException("Unknown missing-episode command.");
            options.Update = args[0] == "--missing-episodes-update";
            int n;
            if (args.Length == 1) return options;
            if (args.Length == 2 && int.TryParse(args[1], out n) && n > 0) { options.Limit = n; return options; }
            if (args.Length == 3 && args[1] == "--row" && int.TryParse(args[2], out n) && n >= 3) { options.Row = n; return options; }
            throw new ArgumentException("Usage: --missing-episodes-preview|--missing-episodes-update [positive show limit | --row DB-row]");
        }
    }

    internal sealed class MissingEpisodeCellChange
    {
        internal MissingEpisodeResult Result;
        internal string Range { get { return Result.Show.OutputCell; } }
        internal string Value { get { return Result.ProposedText; } }
    }

    internal interface IMissingEpisodeSheet
    {
        IList<MissingEpisodeShow> Read();
        void Write(IList<MissingEpisodeCellChange> changes);
    }

    internal sealed class MissingEpisodeRun
    {
        internal bool UpdateMode, Aborted;
        internal string Error;
        internal int TotalShows, SelectedShows;
        internal IList<MissingEpisodeShow> SelectedRows = new List<MissingEpisodeShow>();
        internal readonly List<MissingEpisodeResult> Results = new List<MissingEpisodeResult>();
        internal int SuccessfullyChecked { get { return Results.Count(r => r.ProposedValue.HasValue); } }
        internal int WithMissing { get { return Results.Count(r => r.ProposedValue > 0); } }
        internal int ZeroMissing { get { return Results.Count(r => r.ProposedValue == 0); } }
        internal int TotalMissing { get { return Results.Sum(r => r.ProposedValue ?? 0); } }
        internal int NotInEmby { get { return Results.Count(r => r.NotInEmby); } }
        internal int Failed { get { return Results.Count - SuccessfullyChecked - NotInEmby; } }
        internal int Updated { get { return Results.Count(r => r.SheetOutcome == "Updated" || r.SheetOutcome == "ConfirmedAfterWriteError"); } }
        internal int Unchanged { get { return Results.Count(r => r.SheetOutcome == "Unchanged"); } }
        internal int SheetFailures { get { return Results.Count(r => r.SheetOutcome == "Conflict" || r.SheetOutcome == "WriteFailed"); } }
        internal int ProposedChanges { get { return Results.Count(r => r.SheetOutcome == "WouldUpdate"); } }
        internal bool HasFailures { get { return Aborted || Failed > 0 || SheetFailures > 0; } }
        internal object Report(DateTime date)
        {
            return new {
                Mode = UpdateMode ? "LIVE UPDATE" : "PREVIEW", AsOfDate = date.ToString("yyyy-MM-dd"), Aborted, Error,
                TotalShows, SelectedShows, Examined = Results.Count, SuccessfullyChecked, WithMissing, ZeroMissing,
                TotalMissing, NotInEmby, Failed, UpdatedSheetRows = Updated, UnchangedSheetRows = Unchanged,
                ProposedChanges, SheetFailures, NotChecked = SelectedShows - Results.Count,
                UncheckedShows = SelectedRows.Where(s => !Results.Any(r => r.Show.Row == s.Row))
                    .Select(s => new { s.Row, s.Title, s.TvdbId }).ToArray(),
                Shows = Results.Select(r => new {
                    r.Show.Row, r.Show.Title, r.Show.Year, SheetTvdbId = r.Show.TvdbId,
                    SheetImdbId = r.Show.ImdbId, SheetTmdbId = r.Show.TmdbId, r.Show.OutputCell,
                    Status = r.Status.ToString(), r.Message, r.MatchAttempt, r.MatchCandidates, r.EmbySeries, r.TvdbId,
                    r.Warnings, r.ProposedValue, r.ProposedText, r.OwnedEpisodeCount, r.AiredEpisodeCount, r.IgnoredEmbyItems,
                    MissingEpisodes = r.MissingEpisodes.Select(p => p.ToString()).ToArray(),
                    OwnedOutsideAired = r.OwnedOutsideAired.Select(p => p.ToString()).ToArray(),
                    r.SheetOutcome, r.SheetMessage,
                    RequiresAttention = r.ProposedText == null || r.Warnings.Count > 0 || r.SheetOutcome == "Conflict" || r.SheetOutcome == "WriteFailed"
                }).ToArray()
            };
        }
    }

    internal static class MissingEpisodeWorkflow
    {
        internal const int BatchSize = 50;
        internal static bool SameValue(object a, object b)
        {
            string left = Convert.ToString(a, CultureInfo.InvariantCulture).Trim(), right = Convert.ToString(b, CultureInfo.InvariantCulture).Trim();
            int x, y;
            return int.TryParse(left, NumberStyles.Integer, CultureInfo.InvariantCulture, out x) && int.TryParse(right, NumberStyles.Integer, CultureInfo.InvariantCulture, out y) ? x == y : left == right;
        }

        internal static void Execute(MissingEpisodeRun run, IList<MissingEpisodeShow> shows, Func<MissingEpisodeShow, MissingEpisodeResult> check,
            IMissingEpisodeSheet sheet, CancellationToken cancellation, Action<MissingEpisodeResult, int> progress)
        {
            if (run.UpdateMode && sheet == null) throw new ArgumentException("Update mode needs a Sheet writer.");
            var pending = new List<MissingEpisodeResult>();
            run.SelectedShows = shows.Count;
            run.SelectedRows = shows;
            try
            {
                foreach (var show in shows)
                {
                    cancellation.ThrowIfCancellationRequested();
                    MissingEpisodeResult result;
                    try { result = check(show); }
                    catch (Exception) { result = new MissingEpisodeResult { Show = show, Status = MissingEpisodeStatus.InvalidMetadata, Message = "Unexpected show failure; existing value preserved." }; }
                    run.Results.Add(result);
                    if (result.ProposedText != null)
                    {
                        if (show.OutputIsFormula) Conflict(result, "Destination is a formula; preserved.");
                        else if (!run.UpdateMode) result.SheetOutcome = SameValue(show.ExistingValue, result.ProposedText) ? "Unchanged" : "WouldUpdate";
                        else { result.SheetOutcome = "Pending"; pending.Add(result); }
                    }
                    progress?.Invoke(result, run.Results.Count);
                    if (result.Abort) { run.Aborted = true; run.Error = result.Message; break; }
                    if (run.UpdateMode && pending.Count >= BatchSize) Flush(run, pending, sheet, cancellation);
                    if (run.Aborted) break;
                }
                // No further writes on abort, including previously verified but uncommitted results.
                if (!run.Aborted && run.UpdateMode) Flush(run, pending, sheet, cancellation);
            }
            catch (OperationCanceledException) { run.Aborted = true; run.Error = "Cancelled; uncommitted rows preserved."; }
            catch (Exception) { run.Aborted = true; run.Error = "Sheet preflight failed; no further writes attempted."; }
            finally
            {
                foreach (var result in pending.Where(r => r.SheetOutcome == "Pending"))
                { result.SheetOutcome = "NotAttempted"; result.SheetMessage = "Run stopped before this row was committed."; }
            }
        }

        private static bool SameDestination(MissingEpisodeShow original, MissingEpisodeShow current)
        {
            return current != null && original.Row == current.Row && original.OutputColumn == current.OutputColumn &&
                !string.IsNullOrWhiteSpace(original.TvdbId) && MissingEpisodePreview.Id(original.TvdbId) == MissingEpisodePreview.Id(current.TvdbId);
        }
        private static void Conflict(MissingEpisodeResult result, string reason)
        { result.SheetOutcome = "Conflict"; result.SheetMessage = reason; }

        private static void Flush(MissingEpisodeRun run, List<MissingEpisodeResult> pending, IMissingEpisodeSheet sheet, CancellationToken cancellation)
        {
            if (pending.Count == 0) return;
            cancellation.ThrowIfCancellationRequested();
            var current = sheet.Read().ToDictionary(s => s.Row);
            var changes = new List<MissingEpisodeCellChange>();
            foreach (var result in pending)
            {
                MissingEpisodeShow now;
                current.TryGetValue(result.Show.Row, out now);
                if (!SameDestination(result.Show, now)) { Conflict(result, "Row identity or output header moved; preserved. Rerun from a fresh snapshot."); continue; }
                if (now.OutputIsFormula) { Conflict(result, "Destination became a formula; preserved."); continue; }
                if (SameValue(now.ExistingValue, result.ProposedText)) { result.SheetOutcome = "Unchanged"; continue; }
                if (!SameValue(now.ExistingValue, result.Show.ExistingValue)) { Conflict(result, "Destination was edited during scan; preserved."); continue; }
                changes.Add(new MissingEpisodeCellChange { Result = result });
            }
            cancellation.ThrowIfCancellationRequested();
            if (changes.Count > 0)
            {
                try
                {
                    sheet.Write(changes);
                    foreach (var change in changes) change.Result.SheetOutcome = "Updated";
                }
                catch (Exception)
                {
                    // A timeout may happen after commit. Read back; never blindly retry a write.
                    run.Aborted = true; run.Error = "Sheet batch failed or acknowledgement was lost; inspected destinations where possible. Rerun safely.";
                    IDictionary<int, MissingEpisodeShow> observed = null;
                    try { observed = sheet.Read().ToDictionary(s => s.Row); } catch (Exception) { }
                    foreach (var change in changes)
                    {
                        MissingEpisodeShow now = null;
                        if (observed != null) observed.TryGetValue(change.Result.Show.Row, out now);
                        bool confirmed = SameDestination(change.Result.Show, now) && !now.OutputIsFormula && SameValue(now.ExistingValue, change.Value);
                        change.Result.SheetOutcome = confirmed ? "ConfirmedAfterWriteError" : "WriteFailed";
                        change.Result.SheetMessage = confirmed ? "Desired value confirmed by readback after write error." : "Write outcome not confirmed; inspect or rerun. No automatic write retry.";
                    }
                }
            }
            pending.Clear();
        }
    }
}
