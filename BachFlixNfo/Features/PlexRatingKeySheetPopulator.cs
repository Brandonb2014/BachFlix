using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Builds a mapping of Plex Movie ratingKeys by IMDB/TMDB IDs and writes Plex ratingKeys into the Movies sheet.
    /// Designed to match how Program.cs locates columns by header name (with a fallback to known column letters).
    /// </summary>
    public static class PlexRatingKeySheetPopulator
    {
        public class Options
        {
            public string PlexBaseUrl { get; set; } = "http://192.168.0.5:32400";
            public string PlexToken { get; set; } = "";
            public int MoviesSectionId { get; set; } = 1;

            // Fallback letters if headers aren't found
            public string TmdbIdColumnLetterFallback { get; set; } = "CF";
            public string PlexRatingKeyColumnLetterFallback { get; set; } = "CG";
            public string ImdbIdColumnLetterFallback { get; set; } = "CN";

            // Header names (case-insensitive match)
            public string TmdbHeaderName { get; set; } = "TMDB ID";
            public string ImdbHeaderName { get; set; } = "IMDB ID";
            public string PlexRatingKeyHeaderName { get; set; } = "Rating Key";
            public string StatusHeaderName { get; set; } = "Status";
            public string TitleHeaderName { get; set; } = "IMDB Title";
            public string QuickCreateHeaderName { get; set; } = "Quick Create";
            public string QuickCreateValue { get; set; } = "X";
            public string StatusReportFilterValue { get; set; } = "n";

            public bool OverwriteExistingRatingKeys { get; set; } = false;
            public bool MarkQuickCreateForWrittenRatingKeys { get; set; } = false;
        }

        public class Summary
        {
            public int RatingKeyCellsWritten { get; set; }
            public int QuickCreateCellsWritten { get; set; }
            public int MatchedIds { get; set; }
            public int SkippedExistingRatingKeys { get; set; }
            public List<int> RatingKeyRowNumbersWritten { get; } = new List<int>();
            public List<int> QuickCreateRowNumbersMarked { get; } = new List<int>();
        }

        public static Summary Run(
            SheetsService sheetsService,
            string spreadsheetId,
            string moviesTitleRange,
            string moviesDataRange,
            Action<string, string, int> log, // (type, message, indent)
            Options opt)
        {
            // Wrap async for older calling style.
            return RunAsync(sheetsService, spreadsheetId, moviesTitleRange, moviesDataRange, log, opt)
                .GetAwaiter().GetResult();
        }

        private static async Task<Summary> RunAsync(
            SheetsService sheetsService,
            string spreadsheetId,
            string moviesTitleRange,
            string moviesDataRange,
            Action<string, string, int> log,
            Options opt)
        {
            if (string.IsNullOrWhiteSpace(opt?.PlexBaseUrl))
                throw new ArgumentException("PlexBaseUrl is required.");
            if (string.IsNullOrWhiteSpace(opt?.PlexToken))
                throw new ArgumentException("PlexToken is required.");

            var summary = new Summary();

            log?.Invoke("info", "Step 1/4: Reading Movies headers + rows from Google Sheet...", 2);

            // Headers (row 2)
            IList<IList<object>> headerRows = GetValues(sheetsService, spreadsheetId, moviesTitleRange);
            var headers = headerRows != null && headerRows.Count > 0 ? headerRows[0].Select(h => (h ?? "").ToString()).ToList() : new List<string>();

            // Data (row 3+)
            IList<IList<object>> rows = GetValues(sheetsService, spreadsheetId, moviesDataRange);

            if (rows == null || rows.Count == 0)
            {
                log?.Invoke("warning", "Movies data range returned 0 rows. Nothing to do.", 2);
                return summary;
            }

            // Determine column indices + letters (prefer header match)
            int tmdbIdx = FindHeaderIndex(headers, opt.TmdbHeaderName);
            int imdbIdx = FindHeaderIndex(headers, opt.ImdbHeaderName);
            int plexKeyIdx = FindHeaderIndex(headers, opt.PlexRatingKeyHeaderName);
            int statusIdx = FindHeaderIndex(headers, opt.StatusHeaderName);
            int titleIdx = FindFirstHeaderIndex(headers, opt.TitleHeaderName, "Clean Title", "Title", "Auto Title");
            int quickCreateIdx = FindHeaderIndex(headers, opt.QuickCreateHeaderName);

            string tmdbCol = tmdbIdx >= 0 ? IndexToColumnLetter(tmdbIdx) : opt.TmdbIdColumnLetterFallback;
            string imdbCol = imdbIdx >= 0 ? IndexToColumnLetter(imdbIdx) : opt.ImdbIdColumnLetterFallback;
            string plexKeyCol = plexKeyIdx >= 0 ? IndexToColumnLetter(plexKeyIdx) : opt.PlexRatingKeyColumnLetterFallback;
            string quickCreateCol = quickCreateIdx >= 0 ? IndexToColumnLetter(quickCreateIdx) : "";

            if (tmdbIdx < 0) tmdbIdx = ColumnLetterToIndex(tmdbCol);
            if (imdbIdx < 0) imdbIdx = ColumnLetterToIndex(imdbCol);
            if (plexKeyIdx < 0) plexKeyIdx = ColumnLetterToIndex(plexKeyCol);

            log?.Invoke("data", $"Using columns: TMDB={tmdbCol}, IMDB={imdbCol}, RatingKey={plexKeyCol}", 3);

            if (opt.MarkQuickCreateForWrittenRatingKeys)
            {
                if (quickCreateIdx >= 0)
                    log?.Invoke("data", $"Will mark Quick Create={opt.QuickCreateValue ?? "X"} in column {quickCreateCol} for rows that receive new ratingKeys.", 3);
                else
                    log?.Invoke("warning", $"Quick Create marking is enabled, but '{opt.QuickCreateHeaderName}' was not found in Movies headers.", 3);
            }

            // Build row lookup by IDs from sheet
            // Sheet rows start at row 3.
            var imdbToRow = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var tmdbToRow = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var existingRatingKeysByRow = new Dictionary<int, string>();
            var reportCandidates = new List<SheetRatingKeyCandidate>();
            var matchedRows = new HashSet<int>();

            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];

                string imdb = GetCell(r, imdbIdx);
                string tmdb = GetCell(r, tmdbIdx);

                // Normalize
                imdb = NormalizeImdb(imdb);
                tmdb = NormalizeTmdb(tmdb);

                int sheetRowNumber = 3 + i;
                string existingRatingKey = GetCell(r, plexKeyIdx);
                existingRatingKeysByRow[sheetRowNumber] = existingRatingKey;

                if (!string.IsNullOrWhiteSpace(imdb) && !imdbToRow.ContainsKey(imdb))
                    imdbToRow[imdb] = sheetRowNumber;

                if (!string.IsNullOrWhiteSpace(tmdb) && !tmdbToRow.ContainsKey(tmdb))
                    tmdbToRow[tmdb] = sheetRowNumber;

                if (ShouldReportMissingRatingKey(r, statusIdx, existingRatingKey, imdb, tmdb, opt.StatusReportFilterValue))
                {
                    reportCandidates.Add(new SheetRatingKeyCandidate
                    {
                        RowNumber = sheetRowNumber,
                        Title = GetCell(r, titleIdx),
                        ImdbId = imdb,
                        TmdbId = tmdb
                    });
                }
            }

            log?.Invoke("info", "Step 2/4: Pulling Plex Movies and building ID -> ratingKey index...", 2);

            var plexIndex = await BuildPlexMovieIdIndexAsync(opt.PlexBaseUrl, opt.PlexToken, opt.MoviesSectionId);

            log?.Invoke("data", $"Plex index: imdb={plexIndex.ImdbToRatingKey.Count}, tmdb={plexIndex.TmdbToRatingKey.Count}", 3);

            log?.Invoke("info", "Step 3/4: Matching Plex IDs to sheet rows...", 2);

            // Build updates: range -> ratingKey
            var updates = new List<(string RangeA1, string Value)>();
            var rowsWithRatingKeyWrites = new HashSet<int>();
            var rowsToMarkQuickCreate = new HashSet<int>();

            int matched = 0, unmatched = 0, skippedExisting = 0;

            // Prefer IMDB mapping where possible, then TMDB
            foreach (var kv in plexIndex.ImdbToRatingKey)
            {
                if (imdbToRow.TryGetValue(kv.Key, out int rowNum))
                {
                    matchedRows.Add(rowNum);

                    string ratingKey = kv.Value.ToString(CultureInfo.InvariantCulture);
                    if (!ShouldWriteRatingKey(rowNum, ratingKey, existingRatingKeysByRow, opt.OverwriteExistingRatingKeys))
                    {
                        skippedExisting++;
                        continue;
                    }

                    string a1 = $"Movies!{plexKeyCol}{rowNum}";
                    updates.Add((a1, ratingKey));
                    rowsWithRatingKeyWrites.Add(rowNum);
                    if (opt.MarkQuickCreateForWrittenRatingKeys && quickCreateIdx >= 0)
                        rowsToMarkQuickCreate.Add(rowNum);
                    matched++;
                }
                else
                {
                    unmatched++;
                }
            }

            foreach (var kv in plexIndex.TmdbToRatingKey)
            {
                if (tmdbToRow.TryGetValue(kv.Key, out int rowNum))
                {
                    matchedRows.Add(rowNum);

                    string ratingKey = kv.Value.ToString(CultureInfo.InvariantCulture);
                    if (!ShouldWriteRatingKey(rowNum, ratingKey, existingRatingKeysByRow, opt.OverwriteExistingRatingKeys))
                    {
                        skippedExisting++;
                        continue;
                    }

                    // If IMDB already matched this row, don't stomp it (but it's the same ratingKey anyway).
                    string a1 = $"Movies!{plexKeyCol}{rowNum}";
                    if (!updates.Any(u => u.RangeA1.Equals(a1, StringComparison.OrdinalIgnoreCase)))
                    {
                        updates.Add((a1, ratingKey));
                        rowsWithRatingKeyWrites.Add(rowNum);
                        if (opt.MarkQuickCreateForWrittenRatingKeys && quickCreateIdx >= 0)
                            rowsToMarkQuickCreate.Add(rowNum);
                        matched++;
                    }
                }
            }

            log?.Invoke("data", $"Matched updates to write: {updates.Count}; skipped existing keys: {skippedExisting}", 3);

            log?.Invoke("info", "Step 4/4: Writing plex_ratingKey values back to Google Sheet (batch)...", 2);

            int ratingKeyUpdates = updates.Count;
            int quickCreateUpdates = 0;

            if (opt.MarkQuickCreateForWrittenRatingKeys && quickCreateIdx >= 0)
            {
                string quickCreateValue = string.IsNullOrWhiteSpace(opt.QuickCreateValue)
                    ? "X"
                    : opt.QuickCreateValue;

                foreach (int rowNum in rowsToMarkQuickCreate.OrderBy(r => r))
                {
                    updates.Add(($"Movies!{quickCreateCol}{rowNum}", quickCreateValue));
                    quickCreateUpdates++;
                }
            }

            int written = BatchUpdateCells(sheetsService, spreadsheetId, updates, log);

            summary.RatingKeyCellsWritten = ratingKeyUpdates;
            summary.QuickCreateCellsWritten = quickCreateUpdates;
            summary.MatchedIds = matched;
            summary.SkippedExistingRatingKeys = skippedExisting;
            summary.RatingKeyRowNumbersWritten.AddRange(rowsWithRatingKeyWrites.OrderBy(r => r));
            summary.QuickCreateRowNumbersMarked.AddRange(rowsToMarkQuickCreate.OrderBy(r => r));

            log?.Invoke("success", $"Done. Wrote {ratingKeyUpdates} ratingKey cells and {quickCreateUpdates} Quick Create cells. Total cells written={written}. (Matched {matched} ids)", 2);

            ReportMissingRatingKeys(reportCandidates, matchedRows, log, opt.StatusReportFilterValue);

            return summary;
        }

        private static bool ShouldReportMissingRatingKey(
            IList<object> row,
            int statusIdx,
            string existingRatingKey,
            string imdb,
            string tmdb,
            string statusFilterValue)
        {
            if (!string.IsNullOrWhiteSpace(existingRatingKey))
                return false;

            if (string.IsNullOrWhiteSpace(imdb) && string.IsNullOrWhiteSpace(tmdb))
                return false;

            if (statusIdx < 0)
                return true;

            string status = GetCell(row, statusIdx);
            if (string.IsNullOrWhiteSpace(statusFilterValue))
                return true;

            return string.Equals(status.Trim(), statusFilterValue.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static void ReportMissingRatingKeys(
            List<SheetRatingKeyCandidate> reportCandidates,
            HashSet<int> matchedRows,
            Action<string, string, int> log,
            string statusFilterValue)
        {
            if (reportCandidates == null || reportCandidates.Count == 0)
            {
                log?.Invoke("success", "No active Movies rows are missing a Plex ratingKey.", 2);
                return;
            }

            var missing = reportCandidates
                .Where(r => matchedRows == null || !matchedRows.Contains(r.RowNumber))
                .OrderBy(r => r.RowNumber)
                .ToList();

            if (missing.Count == 0)
            {
                log?.Invoke("success", "Every active Movies row missing a ratingKey was matched in Plex.", 2);
                return;
            }

            string statusText = string.IsNullOrWhiteSpace(statusFilterValue) ? "all statuses" : $"Status={statusFilterValue}";
            log?.Invoke("warning", $"Plex ratingKey missing after sync for {missing.Count} Movies row(s) ({statusText}). These may be mismatched in Plex:", 2);

            foreach (var row in missing)
            {
                string title = string.IsNullOrWhiteSpace(row.Title) ? "(no title)" : row.Title;
                string ids = BuildIdDisplay(row.ImdbId, row.TmdbId);
                log?.Invoke("warning", $"Movies row {row.RowNumber}: {title}{ids}", 3);
            }
        }

        private static string BuildIdDisplay(string imdbId, string tmdbId)
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(imdbId)) parts.Add("IMDB " + imdbId);
            if (!string.IsNullOrWhiteSpace(tmdbId)) parts.Add("TMDB " + tmdbId);

            return parts.Count == 0 ? "" : " (" + string.Join(", ", parts) + ")";
        }

        private static bool ShouldWriteRatingKey(
            int rowNum,
            string newRatingKey,
            Dictionary<int, string> existingRatingKeysByRow,
            bool overwriteExistingRatingKeys)
        {
            if (!existingRatingKeysByRow.TryGetValue(rowNum, out string existing))
                return true;

            if (string.IsNullOrWhiteSpace(existing))
                return true;

            if (!overwriteExistingRatingKeys)
                return false;

            return !RatingKeysEqual(existing, newRatingKey);
        }

        private static bool RatingKeysEqual(string existingRatingKey, string newRatingKey)
        {
            return string.Equals(
                NormalizeRatingKeyForCompare(existingRatingKey),
                NormalizeRatingKeyForCompare(newRatingKey),
                StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeRatingKeyForCompare(string value)
        {
            value = (value ?? "").Trim();
            if (value.EndsWith(".0", StringComparison.OrdinalIgnoreCase))
                value = value.Substring(0, value.Length - 2);

            return value;
        }

        private class PlexIndex
        {
            public Dictionary<string, int> ImdbToRatingKey { get; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, int> TmdbToRatingKey { get; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        }

        private class SheetRatingKeyCandidate
        {
            public int RowNumber { get; set; }
            public string Title { get; set; }
            public string ImdbId { get; set; }
            public string TmdbId { get; set; }
        }

        private static async Task<PlexIndex> BuildPlexMovieIdIndexAsync(string plexBaseUrl, string token, int moviesSectionId)
        {
            var idx = new PlexIndex();

            using (var http = new HttpClient())
            {
                // Ask Plex to include GUIDs in the section listing if supported.
                // If Plex doesn't include Guid nodes, we'll still have ratingKeys and can fall back later (not implemented here).
                string url =
                    $"{plexBaseUrl.TrimEnd('/')}/library/sections/{moviesSectionId}/all" +
                    $"?type=1&includeGuids=1&X-Plex-Token={Uri.EscapeDataString(token)}";

                string xml = await http.GetStringAsync(url).ConfigureAwait(false);

                var doc = XDocument.Parse(xml);

                var videos = doc.Descendants()
                                .Where(e => e.Name.LocalName == "Video");

                foreach (var v in videos)
                {
                    var rkAttr = v.Attribute("ratingKey");
                    if (rkAttr == null) continue;

                    if (!int.TryParse(rkAttr.Value, out int ratingKey)) continue;

                    // GUIDs can appear as <Guid id="imdb://tt..." /> under the Video
                    var guids = v.Elements().Where(e => e.Name.LocalName == "Guid")
                                 .Select(g => (g.Attribute("id")?.Value ?? "").Trim())
                                 .Where(s => !string.IsNullOrWhiteSpace(s))
                                 .ToList();

                    // Some Plex variants may put Guid elements at MediaContainer level; handle descendants under Video as well.
                    if (guids.Count == 0)
                    {
                        guids = v.Descendants().Where(e => e.Name.LocalName == "Guid")
                                   .Select(g => (g.Attribute("id")?.Value ?? "").Trim())
                                   .Where(s => !string.IsNullOrWhiteSpace(s))
                                   .ToList();
                    }

                    foreach (var guid in guids)
                    {
                        if (guid.StartsWith("imdb://", StringComparison.OrdinalIgnoreCase))
                        {
                            string imdb = NormalizeImdb(guid.Substring("imdb://".Length));
                            if (!string.IsNullOrWhiteSpace(imdb) && !idx.ImdbToRatingKey.ContainsKey(imdb))
                                idx.ImdbToRatingKey[imdb] = ratingKey;
                        }
                        else if (guid.StartsWith("tmdb://", StringComparison.OrdinalIgnoreCase))
                        {
                            string tmdb = NormalizeTmdb(guid.Substring("tmdb://".Length));
                            if (!string.IsNullOrWhiteSpace(tmdb) && !idx.TmdbToRatingKey.ContainsKey(tmdb))
                                idx.TmdbToRatingKey[tmdb] = ratingKey;
                        }
                    }
                }

                // Fallback: some Plex setups ignore includeGuids=1 on section listings.
                // If we got ratingKeys but no GUIDs, fetch per-item metadata to extract Guid IDs.
                if (idx.ImdbToRatingKey.Count == 0 && idx.TmdbToRatingKey.Count == 0)
                {
                    var ratingKeys = videos
                        .Select(v => v.Attribute("ratingKey")?.Value)
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .Select(s => int.TryParse(s, out var rk) ? rk : -1)
                        .Where(rk => rk > 0)
                        .Distinct()
                        .ToList();

                    // Limit concurrency so we don't hammer Plex
                    var gate = new SemaphoreSlim(8);
                    var tasks = new List<Task>();

                    foreach (var rk in ratingKeys)
                    {
                        await gate.WaitAsync().ConfigureAwait(false);

                        tasks.Add(Task.Run(async () =>
                        {
                            try
                            {
                                string metaUrl = $"{plexBaseUrl.TrimEnd('/')}/library/metadata/{rk}?X-Plex-Token={Uri.EscapeDataString(token)}";
                                string metaXml = await http.GetStringAsync(metaUrl).ConfigureAwait(false);
                                var metaDoc = XDocument.Parse(metaXml);

                                var guidNodes = metaDoc.Descendants().Where(e => e.Name.LocalName == "Guid")
                                    .Select(g => (g.Attribute("id")?.Value ?? "").Trim())
                                    .Where(s => !string.IsNullOrWhiteSpace(s));

                                foreach (var guid in guidNodes)
                                {
                                    if (guid.StartsWith("imdb://", StringComparison.OrdinalIgnoreCase))
                                    {
                                        string imdb = NormalizeImdb(guid.Substring("imdb://".Length));
                                        if (!string.IsNullOrWhiteSpace(imdb) && !idx.ImdbToRatingKey.ContainsKey(imdb))
                                            idx.ImdbToRatingKey[imdb] = rk;
                                    }
                                    else if (guid.StartsWith("tmdb://", StringComparison.OrdinalIgnoreCase))
                                    {
                                        string tmdb = NormalizeTmdb(guid.Substring("tmdb://".Length));
                                        if (!string.IsNullOrWhiteSpace(tmdb) && !idx.TmdbToRatingKey.ContainsKey(tmdb))
                                            idx.TmdbToRatingKey[tmdb] = rk;
                                    }
                                }
                            }
                            catch
                            {
                                // ignore individual failures; we'll still populate most items
                            }
                            finally
                            {
                                gate.Release();
                            }
                        }));
                    }

                    await Task.WhenAll(tasks).ConfigureAwait(false);
                }
            }

            return idx;
        }

        private static IList<IList<object>> GetValues(SheetsService sheetsService, string spreadsheetId, string range)
        {
            var req = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
            var resp = req.Execute();
            return resp.Values;
        }

        private static int BatchUpdateCells(SheetsService sheetsService, string spreadsheetId, List<(string RangeA1, string Value)> updates, Action<string, string, int> log)
        {
            if (updates == null || updates.Count == 0) return 0;

            int totalWritten = 0;

            const int BATCH_SIZE = 500;

            for (int i = 0; i < updates.Count; i += BATCH_SIZE)
            {
                var chunk = updates.Skip(i).Take(BATCH_SIZE).ToList();

                var data = new List<ValueRange>();
                foreach (var u in chunk)
                {
                    data.Add(new ValueRange
                    {
                        Range = u.RangeA1,
                        Values = new List<IList<object>> { new List<object> { u.Value } }
                    });
                }

                var body = new BatchUpdateValuesRequest
                {
                    ValueInputOption = "USER_ENTERED",
                    Data = data
                };

                var req = sheetsService.Spreadsheets.Values.BatchUpdate(body, spreadsheetId);
                var resp = req.Execute();

                // Each update is 1 cell
                totalWritten += chunk.Count;

                log?.Invoke("data", $"Wrote batch {i / BATCH_SIZE + 1}: {chunk.Count} cells", 3);
            }

            return totalWritten;
        }

        private static int FindHeaderIndex(List<string> headers, string headerName)
        {
            if (headers == null || headers.Count == 0 || string.IsNullOrWhiteSpace(headerName))
                return -1;

            for (int i = 0; i < headers.Count; i++)
            {
                string h = (headers[i] ?? "").Trim();
                if (h.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }

            // fallback loose match
            for (int i = 0; i < headers.Count; i++)
            {
                string h = (headers[i] ?? "").Trim();
                if (h.Replace(" ", "").Equals(headerName.Replace(" ", ""), StringComparison.OrdinalIgnoreCase))
                    return i;
            }

            return -1;
        }

        private static int FindFirstHeaderIndex(List<string> headers, params string[] headerNames)
        {
            if (headerNames == null)
                return -1;

            foreach (string headerName in headerNames)
            {
                int index = FindHeaderIndex(headers, headerName);
                if (index >= 0)
                    return index;
            }

            return -1;
        }

        private static string GetCell(IList<object> row, int idx)
        {
            if (row == null || idx < 0) return "";
            return idx < row.Count ? (row[idx]?.ToString() ?? "").Trim() : "";
        }

        private static string NormalizeImdb(string imdb)
        {
            if (string.IsNullOrWhiteSpace(imdb)) return "";
            imdb = imdb.Trim();

            // Accept imdb://tt123, tt123, or full URL fragments
            int ttPos = imdb.IndexOf("tt", StringComparison.OrdinalIgnoreCase);
            if (ttPos >= 0)
                imdb = imdb.Substring(ttPos);

            // Strip non-alnum
            imdb = new string(imdb.Where(char.IsLetterOrDigit).ToArray());

            if (!imdb.StartsWith("tt", StringComparison.OrdinalIgnoreCase))
                return "";

            return "tt" + imdb.Substring(2);
        }

        private static string NormalizeTmdb(string tmdb)
        {
            if (string.IsNullOrWhiteSpace(tmdb)) return "";
            tmdb = tmdb.Trim();

            // Strip non-digits
            tmdb = new string(tmdb.Where(char.IsDigit).ToArray());
            return tmdb;
        }

                // 0-based index -> A1 column letters
        private static string IndexToColumnLetter(int index)
        {
            if (index < 0) return "A";

            int dividend = index + 1;
            string col = "";

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                col = Convert.ToChar('A' + modulo) + col;
                dividend = (dividend - 1) / 26;
            }

            return col;
        }

        private static int ColumnLetterToIndex(string letter)
        {
            if (string.IsNullOrWhiteSpace(letter)) return -1;
            letter = letter.Trim().ToUpperInvariant();

            int sum = 0;
            for (int i = 0; i < letter.Length; i++)
            {
                char c = letter[i];
                if (c < 'A' || c > 'Z') return -1;
                sum *= 26;
                sum += (c - 'A' + 1);
            }
            return sum - 1;
        }
    }
}
