using BachFlixNfo.Features;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace BachFlixNfo.Features
{
    internal sealed class TvMediaFileSystem : ITvMediaFileSystem
    {
        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        private static extern int WNetGetConnection(string localName, StringBuilder remoteName, ref int length);
        private readonly Dictionary<string, string> mappedDrives = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string CanonicalPath(string path)
        {
            if (!Path.IsPathRooted(path)) throw new IOException("Expected an absolute media path: " + path);
            string full = Path.GetFullPath(path);
            if (full.Length >= 3 && full[1] == ':')
            {
                string drive = full.Substring(0, 2), remote;
                if (!mappedDrives.TryGetValue(drive, out remote))
                {
                    var buffer = new StringBuilder(32768); int length = buffer.Capacity;
                    remote = WNetGetConnection(drive, buffer, ref length) == 0 ? buffer.ToString() : "";
                    mappedDrives[drive] = remote;
                }
                if (remote.Length > 0) full = Path.Combine(remote, full.Substring(3));
            }
            // Keep drive roots rooted (C:\), but normalize all other trailing separators.
            return full.Length == 3 && full[1] == ':' ? full : full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
        private static void VerifyDirectory(string path)
        {
            var attributes = File.GetAttributes(path); // Unlike Directory.Exists, access errors are not hidden.
            if ((attributes & FileAttributes.Directory) == 0) throw new IOException("Not a directory: " + path);
            if ((attributes & FileAttributes.ReparsePoint) != 0)
                throw new IOException("Directory link/junction requires review to avoid duplicate/incomplete inventory: " + path);
        }
        public IList<string> Directories(string path) { VerifyDirectory(path); return Directory.EnumerateDirectories(path).ToArray(); }
        public IList<string> Files(string path) { VerifyDirectory(path); return Directory.EnumerateFiles(path).ToArray(); }
        public TvEpisodeFile Metadata(string path)
        {
            var file = new FileInfo(path);
            if ((file.Attributes & FileAttributes.ReparsePoint) != 0) throw new IOException("Episode file link requires review: " + path);
            return new TvEpisodeFile { FullPath = CanonicalPath(path), SizeBytes = file.Length, LastWriteTimeUtc = file.LastWriteTimeUtc };
        }
    }

}

namespace SheetsQuickstart
{
    partial class Program
    {
        private static void UpdateTvShowResolutionAndFormat(bool force)
        {
            Environment.ExitCode = 0;
            var timer = Stopwatch.StartNew();
            string cachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BachFlixNfo", "TvEpisodeMediaCache.json");
            var logLines = new List<string>();
            var summary = new TvMediaScanReport();
            var pendingShows = new List<TvMediaShow>();
            TvMediaShow currentShow = null;
            bool interrupted = false;
            Action<string> log = message => {
                Type(message, 0, 0, 1, TvMediaScanReport.Color(message));
                logLines.Add(message);
                summary.Record(message, currentShow);
            };
            using (var cancellation = new CancellationTokenSource())
            {
                ConsoleCancelEventHandler cancel = (sender, e) => { e.Cancel = true; cancellation.Cancel(); };
                Console.CancelKeyPress += cancel;
                try
                {
                    log("61 - Update TV Show Resolution and Format (" + (force ? "Force Rescan" : "Smart Scan") + "). Ctrl+C stops safely.");
                    using (var sheets = CreateSheetsService())
                    using (var cache = new TvMediaCache(cachePath, log))
                    {
                        try
                        {
                            var headers = sheets.Spreadsheets.Values.Get(SPREADSHEET_ID, DB_TITLE_RANGE).Execute().Values;
                            var columns = TvMediaSheet.ValidateHeaders(headers);
                            int resolutionColumn = columns["Resolution"], formatColumn = columns["Format"];
                            string resolutionLetter = SheetHeaderMapping.ColumnLetter(resolutionColumn), formatLetter = SheetHeaderMapping.ColumnLetter(formatColumn);
                            log("DB columns mapped from headers: Resolution = " + resolutionLetter + ", Format = " + formatLetter);
                            var rows = sheets.Spreadsheets.Values.Get(SPREADSHEET_ID, "DB!A3:" + SheetHeaderMapping.ColumnLetter(headers[0].Count - 1)).Execute().Values;
                            var shows = TvMediaSheet.ReadShows(headers, rows);
                            var resolutionFormulas = ReadTvMediaFormulas(sheets, "DB!" + resolutionLetter + "3:" + resolutionLetter);
                            var formatFormulas = ReadTvMediaFormulas(sheets, "DB!" + formatLetter + "3:" + formatLetter);
                            var files = new TvMediaFileSystem();
                            var resolver = new TvMediaDirectoryResolver(files);
                            var probe = new NReco.VideoInfo.FFProbe { ExecutionTimeout = TimeSpan.FromSeconds(60) };
                            var scanner = new TvEpisodeMediaScanner(files, path => {
                                var info = probe.GetMediaInfo(path);
                                var stream = info.Streams.FirstOrDefault(s => s.CodecType == "video" && s.Width > 0 && s.Height > 0);
                                if (stream == null) throw new InvalidOperationException("No usable video stream.");
                                return TvEpisodeMediaScanner.ClassifyResolution(stream.Width, stream.Height);
                            }, cache, log);
                            var pending = new List<TvMediaCellChange>();
                            Action flush = () => {
                                if (pending.Count == 0) return;
                                cancellation.Token.ThrowIfCancellationRequested();
                                // Recheck headers and formulas immediately before every write.
                                // Queued cell addresses were built from the start-of-run mapping; refuse if a column moved since.
                                var currentColumns = TvMediaSheet.ValidateHeaders(sheets.Spreadsheets.Values.Get(SPREADSHEET_ID, DB_TITLE_RANGE).Execute().Values);
                                if (currentColumns["Resolution"] != resolutionColumn || currentColumns["Format"] != formatColumn)
                                    throw new InvalidOperationException("DB columns changed during the scan (Resolution was " + resolutionLetter + ", now " +
                                        SheetHeaderMapping.ColumnLetter(currentColumns["Resolution"]) + "; Format was " + formatLetter + ", now " +
                                        SheetHeaderMapping.ColumnLetter(currentColumns["Format"]) + "). Pending writes were not sent; rerun to remap");
                                var currentResolution = ReadTvMediaFormulas(sheets, "DB!" + resolutionLetter + "3:" + resolutionLetter);
                                var currentFormat = ReadTvMediaFormulas(sheets, "DB!" + formatLetter + "3:" + formatLetter);
                                var writable = pending.Where(c => {
                                    bool formula = IsTvMediaFormula(c.IsResolution ? currentResolution : currentFormat, c.Row);
                                    if (formula) log("Preserving formula in " + c.Range + ".");
                                    return !formula;
                                }).ToList();
                                var batch = new TvMediaSheetBatch(changes => BulkWriteToSheet(new BatchUpdateValuesRequest {
                                    ValueInputOption = "RAW", Data = changes.Select(c => new ValueRange {
                                        Range = c.Range, MajorDimension = "ROWS", Values = new List<IList<object>> { new List<object> { c.Value } }
                                    }).ToList()
                                }), resolutionColumn, formatColumn);
                                summary.Cells += batch.Write(writable);
                                summary.ResolutionChanges += writable.Count(c => c.IsResolution);
                                summary.FormatChanges += writable.Count(c => !c.IsResolution);
                                foreach (var show in pendingShows)
                                {
                                    if (writable.Any(c => c.Row == show.Row))
                                    { summary.Updated++; log("  Google Sheet: Updated - DB row " + show.Row + " " + show.Title); }
                                    else { summary.Skipped++; log("  SKIPPED: Formula cells preserved - DB row " + show.Row); }
                                }
                                pending.Clear(); pendingShows.Clear();
                            };
                            foreach (var show in shows)
                            {
                                cancellation.Token.ThrowIfCancellationRequested();
                                currentShow = show;
                                summary.Examined++;
                                log("[" + summary.Examined + "/" + shows.Count + "] " + show.Title + " (DB row " + show.Row + ")");
                                if (string.IsNullOrWhiteSpace(show.Locations))
                                {
                                    summary.Skipped++;
                                    log("  SKIPPED: Found Locations is empty. Existing values preserved.");
                                    continue;
                                }
                                show.ResolutionIsFormula = IsTvMediaFormula(resolutionFormulas, show.Row);
                                show.FormatIsFormula = IsTvMediaFormula(formatFormulas, show.Row);
                                TvMediaScanResult result;
                                try
                                {
                                    var directories = resolver.Resolve(show);
                                    if (directories.Count == 0) throw new IOException("Found Locations contains no usable show paths: " + show.Locations);
                                    result = scanner.Scan(show, directories, force, cancellation.Token);
                                }
                                catch (OperationCanceledException) { throw; }
                                catch (IOException ex) { summary.Errors++; log("  ERROR: " + ex.Message + " | Existing values preserved."); continue; }
                                catch (UnauthorizedAccessException ex) { summary.Errors++; log("  ERROR: " + ex.Message + " | Existing values preserved."); continue; }
                                catch (ArgumentException ex) { summary.Errors++; log("  ERROR: Invalid show location: " + ex.Message + " | Existing values preserved."); continue; }
                                summary.Add(result);
                                log("  Episodes found: " + result.Discovered + " | Cached: " + result.Reused + " | Changed: " + (result.New + result.Modified) + " | FFProbe scans: " + result.Probed);
                                if (!result.Success) { summary.Errors++; log("  ERROR: " + result.Error); continue; }
                                if (show.ResolutionIsFormula || show.FormatIsFormula) log("  Formula output cells preserved.");
                                var changes = TvMediaSheet.Changes(show, result);
                                if (changes.Count == 0) { summary.Unchanged++; log("  Google Sheet: Unchanged"); continue; }
                                foreach (var change in changes)
                                    log("  " + change.Label + " (" + change.Range + "): " + (change.Previous == "" ? "(blank)" : change.Previous) + " -> " + (change.Value == "" ? "(blank)" : change.Value));
                                pending.AddRange(changes); pendingShows.Add(show);
                                log("  Google Sheet: Queued for batch update");
                                if (pending.Count >= 100) flush();
                            }
                            currentShow = null;
                            flush();
                        }
                        finally { cache.Compact(); } // Metadata survives a Sheets failure or cancellation.
                    }
                }
                catch (OperationCanceledException) { interrupted = true; log("Scan interrupted. Completed episode checkpoints are available on the next run. Pending Sheet writes were not sent."); Environment.ExitCode = 1; }
                catch (Exception ex)
                {
                    summary.ScanErrors++;
                    log("Scan stopped (" + ex.GetType().Name + "): " + ex.Message + ". Collected episode checkpoints are retained.");
                    foreach (var show in pendingShows)
                    {
                        currentShow = show;
                        log("  ERROR: Sheet update did not complete: " + ex.Message + ". Check the sheet before retrying.");
                    }
                    Environment.ExitCode = 1;
                }
                finally { Console.CancelKeyPress -= cancel; }
            }
            // Includes examined shows whose queued writes failed or were interrupted.
            summary.CompleteUnfinished(interrupted);
            if (summary.Errors > 0 || summary.ScanErrors > 0) Environment.ExitCode = 1;
            log("TV Shows examined: " + summary.Examined);
            log("TV Shows unchanged: " + summary.Unchanged);
            log("TV Shows updated: " + summary.Updated);
            log("TV Shows skipped: " + summary.Skipped);
            log("TV Shows with errors: " + summary.Errors);
            log("Scan errors: " + summary.ScanErrors);
            log("Episode files discovered: " + summary.Discovered);
            log("Episode metadata reused from cache: " + summary.Reused);
            log("Episodes probed with FFProbe: " + summary.Probed);
            log("New episodes: " + summary.New);
            log("Modified episodes: " + summary.Modified);
            log("Deleted episodes: " + summary.Deleted);
            log("Resolution changes: " + summary.ResolutionChanges);
            log("Format changes: " + summary.FormatChanges);
            log("Google Sheets cells updated: " + summary.Cells);
            log("Cache location: " + cachePath);
            log("Total execution time: " + timer.Elapsed);
            // Print the recap without recording the same errors a second time.
            if (summary.ErrorDetails.Count > 0)
            {
                log("Errors to review:");
                foreach (string detail in summary.ErrorDetails)
                {
                    Type("  " + detail, 0, 0, 1, "Red");
                    logLines.Add("  " + detail);
                }
                string reportError;
                string reportPath = BachFlixLog.WriteBachFlixLog(summary.ErrorDetails, "TV Media Smart Scan", "TvMediaSmartScan_Errors", out reportError);
                if (reportPath != null) log("Errors-only log: " + reportPath);
                else if (reportError != null) DisplayMessage("error", "Could not save errors-only log: " + reportError);
            }
            string error;
            string logPath = BachFlixLog.WriteBachFlixLog(logLines, "TV Media Smart Scan", "TvMediaSmartScan", out error);
            if (logPath != null) Type("Scan log: " + logPath, 0, 0, 1, "Cyan");
            else if (error != null) DisplayMessage("error", "Could not save scan log: " + error);
        }

        private static IList<IList<object>> ReadTvMediaFormulas(SheetsService sheets, string range)
        {
            var request = sheets.Spreadsheets.Values.Get(SPREADSHEET_ID, range);
            request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMULA;
            return request.Execute().Values;
        }
        private static bool IsTvMediaFormula(IList<IList<object>> values, int row)
        {
            int index = row - 3;
            return values != null && index >= 0 && index < values.Count && values[index].Count > 0 && Convert.ToString(values[index][0]).StartsWith("=", StringComparison.Ordinal);
        }
    }
}
