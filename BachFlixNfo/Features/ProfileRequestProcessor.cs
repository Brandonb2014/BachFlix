using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Processes rows from Profile Requests into the Movies Tags column, rebuilds Web Movies,
    /// archives completed queue items, and triggers configured media library refreshes.
    /// </summary>
    public static class ProfileRequestProcessor
    {
        private const string HeaderTags = "Tags";
        private const string HeaderCleanTitle = "Clean Title";
        private const string HeaderRowNum = "RowNum";
        private const string HeaderTmdbId = "TMDB ID";
        private const string HeaderImdbId = "IMDB ID";
        private const string HeaderRatingKey = "Rating Key";
        private const string HeaderTitle = "Title";
        private const string HeaderYear = "Year";
        private const string HeaderUser = "User";
        private const string HeaderStatus = "Status";
        private const string HeaderTrailer = "Trailer";
        private const string HeaderYouTubeTrailerId = "YouTube Trailer ID";
        private const string HeaderRating = "Rating";
        private const string HeaderContentRating = "Content Rating";
        private const string HeaderKids = "Kids";
        private const string HeaderTeens = "Teens";
        private const string HeaderQuickCreate = "Quick Create";
        private const string HeaderPlot = "Plot";
        private const string HeaderRuntime = "Runtime";

        private static readonly string[] RequiredWebMovieHeaders =
        {
            HeaderPlot,
            HeaderRuntime
        };

        private static readonly string[] RequestStatusDoneValues =
        {
            "done",
            "processed",
            "complete",
            "completed",
            "archived"
        };

        public class Options
        {
            public string MoviesSheetName { get; set; } = "Movies";
            public int MoviesHeaderRow { get; set; } = 2;
            public int MoviesDataStartRow { get; set; } = 3;

            public string RequestsSheetName { get; set; } = "Profile Requests";
            public int RequestsHeaderRow { get; set; } = 1;
            public int RequestsDataStartRow { get; set; } = 2;

            public string ArchiveSheetName { get; set; } = "Profile Requests Archive";
            public string WebMoviesSheetName { get; set; } = "Web Movies";
            public int WebMoviesHeaderRow { get; set; } = 1;
            public int WebMoviesDataStartRow { get; set; } = 2;
            public string WebMoviesStatusValue { get; set; } = "n";

            public bool RebuildWebMovies { get; set; } = true;
            public bool ArchiveProcessedRequests { get; set; } = true;
            public bool DeleteProcessedRequestsFromQueue { get; set; } = true;
            public bool RefreshLibraries { get; set; } = true;
            public bool MarkQuickCreateForChangedMovies { get; set; } = true;
            public string QuickCreateValue { get; set; } = "X";
            public Action<IList<int>> SyncQuickCreateMovieMetadata { get; set; }
            public Action RecreateSelectedMovieNfoFiles { get; set; }

            public int BatchSize { get; set; } = 500;
            public List<MediaServerOptions> MediaServers { get; } = new List<MediaServerOptions>();
        }

        public class MediaServerOptions
        {
            public string Name { get; set; }
            public string BaseUrl { get; set; }
            public string ApiKey { get; set; }
            public string LibraryItemId { get; set; }

            // Emby commonly uses /emby/Items/{id}/Refresh. Jellyfin commonly works with /Items/{id}/Refresh.
            public string ApiPathPrefix { get; set; }
        }

        public class Summary
        {
            public int RequestsRead { get; set; }
            public int RequestsProcessed { get; set; }
            public int RequestsFailed { get; set; }
            public int MoviesChanged { get; set; }
            public List<int> ChangedMovieRowNumbers { get; } = new List<int>();
            public int WebMoviesRowsWritten { get; set; }
            public int RequestsArchived { get; set; }
            public int RequestsDeleted { get; set; }
            public int LibrariesRefreshed { get; set; }
            public int QuickCreateMarked { get; set; }
            public List<int> QuickCreateMovieRowNumbers { get; } = new List<int>();
            public bool QuickCreateMovieMetadataSyncRun { get; set; }
            public bool QuickCreateMovieMetadataSyncFailed { get; set; }
            public bool SelectedMovieNfoFilesRun { get; set; }
        }

        private enum TagMutation
        {
            Add,
            Remove
        }

        public static int RebuildWebMoviesOnly(
            SheetsService sheetsService,
            string spreadsheetId,
            Action<string, string, int> log,
            Options options)
        {
            if (sheetsService == null)
                throw new ArgumentNullException(nameof(sheetsService));
            if (string.IsNullOrWhiteSpace(spreadsheetId))
                throw new ArgumentException("Spreadsheet ID is required.", nameof(spreadsheetId));
            if (options == null)
                options = new Options();

            log?.Invoke("info", "Rebuilding Web Movies from Movies only...", 2);

            EnsureSheetExists(sheetsService, spreadsheetId, options.WebMoviesSheetName);

            var moviesSheet = ReadSheet(sheetsService, spreadsheetId, options.MoviesSheetName, options.MoviesHeaderRow, options.MoviesDataStartRow);
            int written = RebuildWebMovies(sheetsService, spreadsheetId, options, moviesSheet);

            log?.Invoke("success", $"Web Movies rebuilt with {written} movie row(s).", 2);

            return written;
        }

        public static Summary Run(
            SheetsService sheetsService,
            string spreadsheetId,
            Action<string, string, int> log,
            Options options)
        {
            if (sheetsService == null)
                throw new ArgumentNullException(nameof(sheetsService));
            if (string.IsNullOrWhiteSpace(spreadsheetId))
                throw new ArgumentException("Spreadsheet ID is required.", nameof(spreadsheetId));
            if (options == null)
                options = new Options();

            var summary = new Summary();

            log?.Invoke("info", "Profile request processor: reading sheets...", 2);

            EnsureSheetExists(sheetsService, spreadsheetId, options.ArchiveSheetName);
            EnsureSheetExists(sheetsService, spreadsheetId, options.WebMoviesSheetName);

            var moviesSheet = ReadSheet(sheetsService, spreadsheetId, options.MoviesSheetName, options.MoviesHeaderRow, options.MoviesDataStartRow);
            var requestSheet = ReadSheet(sheetsService, spreadsheetId, options.RequestsSheetName, options.RequestsHeaderRow, options.RequestsDataStartRow);

            ValidateMoviesHeaders(moviesSheet, options);

            var pendingRequests = requestSheet.DataRows
                .Where(r => RowHasData(r.Values))
                .Select(r => ProfileRequest.FromRow(requestSheet, r))
                .Where(r => r.IsPending)
                .ToList();

            summary.RequestsRead = pendingRequests.Count;

            if (pendingRequests.Count == 0)
            {
                log?.Invoke("info", "No pending profile requests found.", 2);
                return summary;
            }

            log?.Invoke("data", $"Pending profile requests: {pendingRequests.Count}", 3);

            var movieIndex = MovieIndex.Build(moviesSheet);
            var results = new List<ProcessingResult>();

            foreach (var request in pendingRequests)
            {
                results.Add(ProcessRequest(request, movieIndex));
            }

            var succeeded = results.Where(r => r.Success).ToList();
            var failed = results.Where(r => !r.Success).ToList();
            var changed = succeeded.Where(r => r.TagsChanged).ToList();

            summary.RequestsProcessed = succeeded.Count;
            summary.RequestsFailed = failed.Count;
            summary.MoviesChanged = changed.Count;
            foreach (int rowNumber in changed
                .Select(r => r.MovieRowNumber)
                .Where(r => r > 0)
                .Distinct()
                .OrderBy(r => r))
            {
                summary.ChangedMovieRowNumbers.Add(rowNumber);
            }

            foreach (var failure in failed)
                log?.Invoke("warning", $"Row {failure.Request.SheetRowNumber}: {failure.Message}", 3);

            if (changed.Count > 0)
            {
                log?.Invoke("info", $"Writing {changed.Count} Movies Tags update(s)...", 2);
                summary.QuickCreateMarked = BatchUpdateChangedMovieCells(sheetsService, spreadsheetId, options, moviesSheet, changed);
                summary.QuickCreateMovieRowNumbers.AddRange(GetQuickCreateMovieRowNumbers(
                    moviesSheet,
                    options,
                    summary.ChangedMovieRowNumbers));
            }
            else
            {
                log?.Invoke("info", "No Movies Tags cells needed changes.", 2);
            }

            if (succeeded.Count > 0 && options.RebuildWebMovies)
            {
                log?.Invoke("info", "Rebuilding Web Movies from Movies...", 2);
                summary.WebMoviesRowsWritten = RebuildWebMovies(sheetsService, spreadsheetId, options, moviesSheet);
                log?.Invoke("success", $"Web Movies rebuilt with {summary.WebMoviesRowsWritten} movie row(s).", 2);
            }

            if (summary.QuickCreateMarked > 0 && options.SyncQuickCreateMovieMetadata != null)
            {
                var rowNumbers = summary.QuickCreateMovieRowNumbers
                    .Distinct()
                    .OrderBy(r => r)
                    .ToList();

                if (rowNumbers.Count > 0)
                {
                    try
                    {
                        log?.Invoke("info", $"Syncing Plex metadata for {rowNumbers.Count} Quick Create movie row(s) before selected NFO recreation...", 2);
                        options.SyncQuickCreateMovieMetadata(rowNumbers);
                        summary.QuickCreateMovieMetadataSyncRun = true;
                    }
                    catch (Exception ex)
                    {
                        summary.QuickCreateMovieMetadataSyncFailed = true;
                        log?.Invoke("error", "Quick Create Plex metadata sync failed: " + ex.Message, 2);
                    }
                }
            }

            if (summary.QuickCreateMarked > 0 && options.RecreateSelectedMovieNfoFiles != null)
            {
                log?.Invoke("info", "Recreating selected Movie NFO files from Quick Create marks...", 2);
                options.RecreateSelectedMovieNfoFiles();
                summary.SelectedMovieNfoFilesRun = true;
            }

            if (succeeded.Count > 0 && options.ArchiveProcessedRequests)
            {
                log?.Invoke("info", $"Archiving {succeeded.Count} processed request(s)...", 2);
                summary.RequestsArchived = ArchiveProcessedRequests(sheetsService, spreadsheetId, options, requestSheet, succeeded);
            }

            if (succeeded.Count > 0 && options.DeleteProcessedRequestsFromQueue)
            {
                log?.Invoke("info", "Removing processed requests from queue...", 2);
                summary.RequestsDeleted = DeleteProcessedRequestRows(sheetsService, spreadsheetId, options.RequestsSheetName, succeeded);
            }

            if (succeeded.Count > 0 && options.RefreshLibraries)
            {
                log?.Invoke("info", "Refreshing configured media libraries...", 2);
                summary.LibrariesRefreshed = RefreshMediaLibraries(options, log);
            }

            log?.Invoke(
                "success",
                $"Profile requests complete. Processed={summary.RequestsProcessed}, Failed={summary.RequestsFailed}, MoviesChanged={summary.MoviesChanged}, QuickCreateMarked={summary.QuickCreateMarked}, QuickCreatePlexSyncRun={summary.QuickCreateMovieMetadataSyncRun}, SelectedNfoRun={summary.SelectedMovieNfoFilesRun}.",
                2);

            return summary;
        }

        private static void ValidateMoviesHeaders(SheetData moviesSheet, Options options)
        {
            if (!moviesSheet.Headers.Contains(HeaderTags))
                throw new Exception("Movies sheet is missing the required 'Tags' column.");
            if (options.MarkQuickCreateForChangedMovies && !moviesSheet.Headers.Contains(HeaderQuickCreate))
                throw new Exception("Movies sheet is missing the required 'Quick Create' column.");
            if (!moviesSheet.Headers.Contains(HeaderCleanTitle) &&
                !moviesSheet.Headers.Contains(HeaderTmdbId) &&
                !moviesSheet.Headers.Contains(HeaderImdbId) &&
                !moviesSheet.Headers.Contains(HeaderRatingKey))
            {
                throw new Exception("Movies sheet needs at least one movie identity column: Clean Title, TMDB ID, IMDB ID, or Rating Key.");
            }
        }

        private static ProcessingResult ProcessRequest(ProfileRequest request, MovieIndex movieIndex)
        {
            var result = new ProcessingResult { Request = request };

            if (request.Tags.Count == 0)
                return result.Fail("No user/profile tag was found on the request.");

            if (!TryResolveMutation(request, out TagMutation mutation, out string mutationReason))
                return result.Fail(mutationReason);

            MovieRow movie = movieIndex.Find(request);
            if (movie == null)
                return result.Fail("Could not match this request to a Movies row.");

            result.MovieRowNumber = movie.SheetRowNumber;
            result.MovieCleanTitle = movie.CleanTitle;
            result.MovieTmdbId = movie.TmdbId;
            result.MovieImdbId = movie.ImdbId;
            result.MovieRatingKey = movie.RatingKey;
            result.TagsBefore = movie.TagsText;

            var tags = SplitTags(movie.TagsText);
            bool changed = false;

            foreach (string tag in request.Tags)
            {
                if (mutation == TagMutation.Add)
                {
                    if (!tags.Any(t => t.Equals(tag, StringComparison.OrdinalIgnoreCase)))
                    {
                        tags.Add(tag);
                        changed = true;
                    }
                }
                else
                {
                    int removed = tags.RemoveAll(t => t.Equals(tag, StringComparison.OrdinalIgnoreCase));
                    if (removed > 0)
                        changed = true;
                }
            }

            string newTagsText = FormatTags(tags);
            movie.TagsText = newTagsText;
            SetCell(movie.Values, movie.TagsColumnIndex, newTagsText);

            result.Success = true;
            result.TagsChanged = changed;
            result.TagsAfter = newTagsText;
            result.Message = changed
                ? $"{mutation} tag(s): {string.Join(", ", request.Tags)}"
                : $"No tag change needed for: {string.Join(", ", request.Tags)}";

            return result;
        }

        private static bool TryResolveMutation(ProfileRequest request, out TagMutation mutation, out string reason)
        {
            string action = NormalizeToken(request.Action);
            string tagMode = NormalizeToken(request.TagMode);

            if (IsAny(action, "requestaccess", "accessrequest"))
            {
                mutation = ResolveAccessMutation(true, tagMode);
                reason = "";
                return true;
            }

            if (IsAny(action, "hidefromprofile", "removefromprofile", "revokeaccess"))
            {
                mutation = ResolveAccessMutation(false, tagMode);
                reason = "";
                return true;
            }

            if (IsAny(action, "addtag", "tagadd", "add", "append"))
            {
                mutation = TagMutation.Add;
                reason = "";
                return true;
            }

            if (IsAny(action, "removetag", "tagremove", "remove", "delete", "drop"))
            {
                mutation = TagMutation.Remove;
                reason = "";
                return true;
            }

            if (IsAny(action, "allow", "grant", "include", "enable", "approve", "accessallowed"))
            {
                mutation = ResolveAccessMutation(true, tagMode);
                reason = "";
                return true;
            }

            if (IsAny(action, "block", "deny", "exclude", "disable", "hide", "accessblocked"))
            {
                mutation = ResolveAccessMutation(false, tagMode);
                reason = "";
                return true;
            }

            if (string.IsNullOrWhiteSpace(action) && IsAny(tagMode, "allow", "block"))
            {
                mutation = TagMutation.Add;
                reason = "";
                return true;
            }

            mutation = TagMutation.Add;
            reason = "No recognizable action was found. Expected add/remove, allow/block, grant/deny, or include/exclude.";
            return false;
        }

        private static TagMutation ResolveAccessMutation(bool grantAccess, string tagMode)
        {
            if (grantAccess)
                return tagMode == "block" ? TagMutation.Remove : TagMutation.Add;

            return tagMode == "allow" ? TagMutation.Remove : TagMutation.Add;
        }

        private static int BatchUpdateChangedMovieCells(
            SheetsService sheetsService,
            string spreadsheetId,
            Options options,
            SheetData moviesSheet,
            List<ProcessingResult> changed)
        {
            int tagsCol = moviesSheet.Headers.GetRequiredIndex(HeaderTags);
            string tagsColumnLetter = ColumnIndexToLetter(tagsCol);
            int quickCreateCol = options.MarkQuickCreateForChangedMovies
                ? moviesSheet.Headers.GetRequiredIndex(HeaderQuickCreate)
                : -1;
            string quickCreateColumnLetter = quickCreateCol >= 0
                ? ColumnIndexToLetter(quickCreateCol)
                : "";

            var updates = new List<ValueRange>();
            foreach (var result in changed)
            {
                updates.Add(new ValueRange
                {
                    Range = A1(options.MoviesSheetName, tagsColumnLetter + result.MovieRowNumber),
                    Values = new List<IList<object>>
                    {
                        new List<object> { result.TagsAfter ?? "" }
                    }
                });
            }

            int quickCreateMarked = 0;
            foreach (var result in changed
                .GroupBy(r => r.MovieRowNumber)
                .Select(g => g.First()))
            {
                if (quickCreateCol < 0)
                    continue;

                updates.Add(new ValueRange
                {
                    Range = A1(options.MoviesSheetName, quickCreateColumnLetter + result.MovieRowNumber),
                    Values = new List<IList<object>>
                    {
                        new List<object> { options.QuickCreateValue ?? "" }
                    }
                });
                quickCreateMarked++;
            }

            BatchUpdateValues(sheetsService, spreadsheetId, updates, options.BatchSize);
            return quickCreateMarked;
        }

        private static List<int> GetQuickCreateMovieRowNumbers(
            SheetData moviesSheet,
            Options options,
            IEnumerable<int> newlyMarkedRowNumbers)
        {
            var rowNumbers = new HashSet<int>();

            if (moviesSheet != null && options != null && options.MarkQuickCreateForChangedMovies)
            {
                int quickCreateCol = moviesSheet.Headers.FindIndex(HeaderQuickCreate);
                string quickCreateValue = string.IsNullOrWhiteSpace(options.QuickCreateValue)
                    ? "X"
                    : options.QuickCreateValue.Trim();

                if (quickCreateCol >= 0)
                {
                    foreach (var row in moviesSheet.DataRows)
                    {
                        if (row == null)
                            continue;

                        string value = GetCell(row.Values, quickCreateCol);
                        if (string.Equals(value, quickCreateValue, StringComparison.OrdinalIgnoreCase))
                            rowNumbers.Add(row.SheetRowNumber);
                    }
                }
            }

            if (newlyMarkedRowNumbers != null)
            {
                foreach (int rowNumber in newlyMarkedRowNumbers)
                {
                    if (rowNumber > 0)
                        rowNumbers.Add(rowNumber);
                }
            }

            return rowNumbers
                .Where(r => r > 0)
                .OrderBy(r => r)
                .ToList();
        }

        private static int RebuildWebMovies(
            SheetsService sheetsService,
            string spreadsheetId,
            Options options,
            SheetData moviesSheet)
        {
            var webMoviesSheet = ReadSheet(
                sheetsService,
                spreadsheetId,
                options.WebMoviesSheetName,
                options.WebMoviesHeaderRow,
                options.WebMoviesDataStartRow);

            if (!webMoviesSheet.Headers.OriginalHeaders.Any(h => !string.IsNullOrWhiteSpace(h)))
                throw new Exception($"'{options.WebMoviesSheetName}' needs a header row before it can be rebuilt.");

            webMoviesSheet.Headers = EnsureWebMovieHeaders(
                sheetsService,
                spreadsheetId,
                options,
                webMoviesSheet.Headers);

            int statusCol = moviesSheet.Headers.GetRequiredIndex(HeaderStatus);
            var columnMap = BuildWebMovieColumnMap(moviesSheet.Headers, webMoviesSheet.Headers);

            ClearRange(
                sheetsService,
                spreadsheetId,
                A1(options.WebMoviesSheetName, $"A{options.WebMoviesDataStartRow}:ZZ"));

            var exportRows = new List<IList<object>>();

            foreach (var row in moviesSheet.DataRows)
            {
                if (!RowHasData(row.Values))
                    continue;

                string status = GetCell(row.Values, statusCol);
                if (!string.Equals(status, options.WebMoviesStatusValue, StringComparison.OrdinalIgnoreCase))
                    continue;

                exportRows.Add(BuildWebMovieRow(row.Values, columnMap));
            }

            if (exportRows.Count == 0)
                return 0;

            UpdateValues(
                sheetsService,
                spreadsheetId,
                A1(options.WebMoviesSheetName, $"A{options.WebMoviesDataStartRow}"),
                exportRows);

            return exportRows.Count;
        }

        private static HeaderMap EnsureWebMovieHeaders(
            SheetsService sheetsService,
            string spreadsheetId,
            Options options,
            HeaderMap currentHeaders)
        {
            var desiredHeaders = new List<string>(currentHeaders.OriginalHeaders);

            foreach (string requiredHeader in RequiredWebMovieHeaders)
                AddHeaderIfMissing(desiredHeaders, requiredHeader);

            if (!HeadersEqual(currentHeaders.OriginalHeaders, desiredHeaders))
            {
                UpdateValues(
                    sheetsService,
                    spreadsheetId,
                    A1(options.WebMoviesSheetName, $"A{options.WebMoviesHeaderRow}"),
                    new List<IList<object>> { desiredHeaders.Cast<object>().ToList() });
            }

            return new HeaderMap(desiredHeaders);
        }

        private static List<WebMovieColumnMapping> BuildWebMovieColumnMap(HeaderMap moviesHeaders, HeaderMap webHeaders)
        {
            var columnMap = new List<WebMovieColumnMapping>();

            foreach (string webHeader in webHeaders.OriginalHeaders)
            {
                if (string.IsNullOrWhiteSpace(webHeader))
                {
                    columnMap.Add(new WebMovieColumnMapping());
                    continue;
                }

                if (NormalizeHeader(webHeader) == NormalizeHeader(HeaderTags))
                {
                    columnMap.Add(new WebMovieColumnMapping
                    {
                        WebHeader = webHeader,
                        MoviesIndex = moviesHeaders.GetRequiredIndex(HeaderTags),
                        KidsIndex = moviesHeaders.FindIndex(HeaderKids),
                        TeensIndex = moviesHeaders.FindIndex(HeaderTeens),
                        IsComputedTags = true
                    });
                    continue;
                }

                int moviesIndex = FindMovieColumnForWebHeader(moviesHeaders, webHeader);
                if (moviesIndex < 0)
                {
                    throw new Exception(
                        $"Web Movies header '{webHeader}' was not found in Movies. " +
                        $"Supported custom mappings are '{HeaderTrailer}' -> '{HeaderYouTubeTrailerId}' " +
                        $"and '{HeaderRating}' -> '{HeaderContentRating}'. " +
                        $"Web Movies '{HeaderTags}' is computed from Movies '{HeaderTags}', '{HeaderKids}', and '{HeaderTeens}'.");
                }

                columnMap.Add(new WebMovieColumnMapping
                {
                    WebHeader = webHeader,
                    MoviesIndex = moviesIndex
                });
            }

            return columnMap;
        }

        private static int FindMovieColumnForWebHeader(HeaderMap moviesHeaders, string webHeader)
        {
            if (NormalizeHeader(webHeader) == NormalizeHeader(HeaderTrailer))
                return moviesHeaders.FindIndex(HeaderYouTubeTrailerId);

            if (NormalizeHeader(webHeader) == NormalizeHeader(HeaderRating))
                return moviesHeaders.FindIndex(HeaderContentRating);

            return moviesHeaders.FindIndex(webHeader);
        }

        private static IList<object> BuildWebMovieRow(IList<object> movieRow, List<WebMovieColumnMapping> columnMap)
        {
            var exportRow = new List<object>();

            foreach (var mapping in columnMap)
            {
                if (mapping.IsComputedTags)
                {
                    exportRow.Add(BuildWebMovieTags(movieRow, mapping));
                    continue;
                }

                exportRow.Add(mapping.MoviesIndex >= 0 ? GetCell(movieRow, mapping.MoviesIndex) : "");
            }

            return exportRow;
        }

        private static string BuildWebMovieTags(IList<object> movieRow, WebMovieColumnMapping mapping)
        {
            var tags = SplitTags(GetCell(movieRow, mapping.MoviesIndex));

            AddWebVisibilityTagIfMarked(tags, GetCell(movieRow, mapping.KidsIndex), HeaderKids);
            AddWebVisibilityTagIfMarked(tags, GetCell(movieRow, mapping.TeensIndex), HeaderTeens);

            return FormatTags(tags);
        }

        private static void AddWebVisibilityTagIfMarked(List<string> tags, string marker, string tag)
        {
            if (!string.Equals((marker ?? "").Trim(), "X", StringComparison.OrdinalIgnoreCase))
                return;

            if (!tags.Any(t => t.Equals(tag, StringComparison.OrdinalIgnoreCase)))
                tags.Add(tag);
        }

        private static int ArchiveProcessedRequests(
            SheetsService sheetsService,
            string spreadsheetId,
            Options options,
            SheetData requestSheet,
            List<ProcessingResult> results)
        {
            var archiveValues = GetValues(sheetsService, spreadsheetId, A1(options.ArchiveSheetName, "A:ZZ"));
            var existingHeaders = archiveValues != null && archiveValues.Count > 0
                ? archiveValues[0].Select(v => (v ?? "").ToString()).Where(s => !string.IsNullOrWhiteSpace(s)).ToList()
                : new List<string>();

            var desiredHeaders = existingHeaders.Count > 0
                ? new List<string>(existingHeaders)
                : new List<string>();

            foreach (string header in requestSheet.Headers.OriginalHeaders.Where(h => !string.IsNullOrWhiteSpace(h)))
                AddHeaderIfMissing(desiredHeaders, header);

            AddHeaderIfMissing(desiredHeaders, "Processed At");
            AddHeaderIfMissing(desiredHeaders, "Processor Status");
            AddHeaderIfMissing(desiredHeaders, "Processor Message");
            AddHeaderIfMissing(desiredHeaders, "Movie Row");
            AddHeaderIfMissing(desiredHeaders, "Matched Clean Title");
            AddHeaderIfMissing(desiredHeaders, "Matched TMDB ID");
            AddHeaderIfMissing(desiredHeaders, "Matched IMDB ID");
            AddHeaderIfMissing(desiredHeaders, "Matched Rating Key");
            AddHeaderIfMissing(desiredHeaders, "Tags Before");
            AddHeaderIfMissing(desiredHeaders, "Tags After");

            if (!HeadersEqual(existingHeaders, desiredHeaders))
            {
                UpdateValues(
                    sheetsService,
                    spreadsheetId,
                    A1(options.ArchiveSheetName, "A1"),
                    new List<IList<object>> { desiredHeaders.Cast<object>().ToList() });
            }

            var rowsToAppend = new List<IList<object>>();
            foreach (var result in results)
            {
                var row = new List<object>();
                foreach (string header in desiredHeaders)
                {
                    string value;
                    if (TryGetArchiveExtraValue(result, header, out value))
                    {
                        row.Add(value);
                        continue;
                    }

                    int requestColumn = requestSheet.Headers.FindIndex(header);
                    row.Add(requestColumn >= 0 ? GetCell(result.Request.RawRow, requestColumn) : "");
                }

                rowsToAppend.Add(row);
            }

            AppendValues(sheetsService, spreadsheetId, A1(options.ArchiveSheetName, "A1"), rowsToAppend);
            return rowsToAppend.Count;
        }

        private static int DeleteProcessedRequestRows(
            SheetsService sheetsService,
            string spreadsheetId,
            string sheetName,
            List<ProcessingResult> results)
        {
            int? sheetId = GetSheetId(sheetsService, spreadsheetId, sheetName);
            if (!sheetId.HasValue)
                throw new Exception($"Could not find sheet ID for '{sheetName}'.");

            var deleteRequests = results
                .Select(r => r.Request.SheetRowNumber)
                .Distinct()
                .OrderByDescending(r => r)
                .Select(rowNumber => new Google.Apis.Sheets.v4.Data.Request
                {
                    DeleteDimension = new DeleteDimensionRequest
                    {
                        Range = new DimensionRange
                        {
                            SheetId = sheetId.Value,
                            Dimension = "ROWS",
                            StartIndex = rowNumber - 1,
                            EndIndex = rowNumber
                        }
                    }
                })
                .ToList();

            if (deleteRequests.Count == 0)
                return 0;

            var request = new BatchUpdateSpreadsheetRequest { Requests = deleteRequests };
            sheetsService.Spreadsheets.BatchUpdate(request, spreadsheetId).Execute();
            return deleteRequests.Count;
        }

        private static int RefreshMediaLibraries(Options options, Action<string, string, int> log)
        {
            int refreshed = 0;

            foreach (var server in options.MediaServers)
            {
                if (server == null ||
                    string.IsNullOrWhiteSpace(server.BaseUrl) ||
                    string.IsNullOrWhiteSpace(server.ApiKey) ||
                    string.IsNullOrWhiteSpace(server.LibraryItemId))
                {
                    log?.Invoke("info", "Media server refresh skipped because configuration is incomplete.", 3);
                    continue;
                }

                try
                {
                    using (var client = new HttpClient())
                    {
                        client.Timeout = TimeSpan.FromSeconds(30);
                        client.DefaultRequestHeaders.Add("X-Emby-Token", server.ApiKey);

                        string prefix = string.IsNullOrWhiteSpace(server.ApiPathPrefix)
                            ? ""
                            : "/" + server.ApiPathPrefix.Trim('/');

                        string url =
                            server.BaseUrl.TrimEnd('/') +
                            prefix +
                            "/Items/" +
                            Uri.EscapeDataString(server.LibraryItemId) +
                            "/Refresh?Recursive=true" +
                            "&ImageRefreshMode=FullRefresh" +
                            "&MetadataRefreshMode=Default" +
                            "&ReplaceAllMetadata=false" +
                            "&ReplaceAllImages=false" +
                            "&api_key=" +
                            Uri.EscapeDataString(server.ApiKey);

                        var response = client.PostAsync(url, null).Result;
                        if (response.IsSuccessStatusCode)
                        {
                            refreshed++;
                            log?.Invoke("success", $"{DisplayName(server)} library refresh triggered.", 3);
                        }
                        else
                        {
                            log?.Invoke("error", $"{DisplayName(server)} refresh failed: {(int)response.StatusCode} {response.ReasonPhrase}", 3);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log?.Invoke("error", $"{DisplayName(server)} refresh error: {ex.Message}", 3);
                }
            }

            return refreshed;
        }

        private static SheetData ReadSheet(
            SheetsService sheetsService,
            string spreadsheetId,
            string sheetName,
            int headerRow,
            int dataStartRow)
        {
            var rows = GetValues(sheetsService, spreadsheetId, A1(sheetName, "A:ZZ")) ?? new List<IList<object>>();

            if (rows.Count < headerRow)
                throw new Exception($"Sheet '{sheetName}' does not contain header row {headerRow}.");

            var headerValues = rows[headerRow - 1].Select(v => (v ?? "").ToString()).ToList();
            var dataRows = new List<SheetRow>();

            for (int rowIndex = dataStartRow - 1; rowIndex < rows.Count; rowIndex++)
            {
                dataRows.Add(new SheetRow
                {
                    SheetRowNumber = rowIndex + 1,
                    Values = rows[rowIndex]
                });
            }

            return new SheetData
            {
                SheetName = sheetName,
                HeaderRowNumber = headerRow,
                DataStartRowNumber = dataStartRow,
                Headers = new HeaderMap(headerValues),
                DataRows = dataRows
            };
        }

        private static IList<IList<object>> GetValues(SheetsService sheetsService, string spreadsheetId, string range)
        {
            var response = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range).Execute();
            return response.Values ?? new List<IList<object>>();
        }

        private static void UpdateValues(SheetsService sheetsService, string spreadsheetId, string range, IList<IList<object>> rows)
        {
            var body = new ValueRange { Values = rows };
            var request = sheetsService.Spreadsheets.Values.Update(body, spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            request.Execute();
        }

        private static void AppendValues(SheetsService sheetsService, string spreadsheetId, string range, IList<IList<object>> rows)
        {
            if (rows == null || rows.Count == 0)
                return;

            var body = new ValueRange { Values = rows };
            var request = sheetsService.Spreadsheets.Values.Append(body, spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.Execute();
        }

        private static void BatchUpdateValues(
            SheetsService sheetsService,
            string spreadsheetId,
            List<ValueRange> updates,
            int batchSize)
        {
            if (updates == null || updates.Count == 0)
                return;

            int size = batchSize > 0 ? batchSize : 500;
            for (int i = 0; i < updates.Count; i += size)
            {
                var body = new BatchUpdateValuesRequest
                {
                    ValueInputOption = "USER_ENTERED",
                    Data = updates.Skip(i).Take(size).ToList()
                };

                sheetsService.Spreadsheets.Values.BatchUpdate(body, spreadsheetId).Execute();
            }
        }

        private static void ClearRange(SheetsService sheetsService, string spreadsheetId, string range)
        {
            sheetsService.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadsheetId, range).Execute();
        }

        private static void EnsureSheetExists(SheetsService sheetsService, string spreadsheetId, string sheetName)
        {
            if (GetSheetId(sheetsService, spreadsheetId, sheetName).HasValue)
                return;

            var request = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Google.Apis.Sheets.v4.Data.Request>
                {
                    new Google.Apis.Sheets.v4.Data.Request
                    {
                        AddSheet = new AddSheetRequest
                        {
                            Properties = new SheetProperties
                            {
                                Title = sheetName
                            }
                        }
                    }
                }
            };

            sheetsService.Spreadsheets.BatchUpdate(request, spreadsheetId).Execute();
        }

        private static int? GetSheetId(SheetsService sheetsService, string spreadsheetId, string sheetName)
        {
            var spreadsheet = sheetsService.Spreadsheets.Get(spreadsheetId).Execute();
            var sheet = spreadsheet.Sheets == null
                ? null
                : spreadsheet.Sheets.FirstOrDefault(s =>
                    s.Properties != null &&
                    s.Properties.Title.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            return sheet?.Properties?.SheetId;
        }

        private static bool TryGetArchiveExtraValue(ProcessingResult result, string header, out string value)
        {
            switch (NormalizeHeader(header))
            {
                case "processedat":
                    value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    return true;
                case "processorstatus":
                    value = result.Success ? "Processed" : "Failed";
                    return true;
                case "processormessage":
                    value = result.Message ?? "";
                    return true;
                case "movierow":
                    value = result.MovieRowNumber > 0 ? result.MovieRowNumber.ToString() : "";
                    return true;
                case "matchedcleantitle":
                    value = result.MovieCleanTitle ?? "";
                    return true;
                case "matchedtmdbid":
                    value = result.MovieTmdbId ?? "";
                    return true;
                case "matchedimdbid":
                    value = result.MovieImdbId ?? "";
                    return true;
                case "matchedratingkey":
                    value = result.MovieRatingKey ?? "";
                    return true;
                case "tagsbefore":
                    value = result.TagsBefore ?? "";
                    return true;
                case "tagsafter":
                    value = result.TagsAfter ?? "";
                    return true;
                default:
                    value = "";
                    return false;
            }
        }

        private static bool LooksLikeMovieRow(IList<object> row, int cleanTitleCol, int tmdbCol, int imdbCol, int ratingKeyCol)
        {
            if (cleanTitleCol >= 0 && !string.IsNullOrWhiteSpace(GetCell(row, cleanTitleCol)))
                return true;
            if (tmdbCol >= 0 && !string.IsNullOrWhiteSpace(GetCell(row, tmdbCol)))
                return true;
            if (imdbCol >= 0 && !string.IsNullOrWhiteSpace(GetCell(row, imdbCol)))
                return true;
            if (ratingKeyCol >= 0 && !string.IsNullOrWhiteSpace(GetCell(row, ratingKeyCol)))
                return true;

            return false;
        }

        private static bool RowHasData(IList<object> row)
        {
            return row != null && row.Any(v => !string.IsNullOrWhiteSpace((v ?? "").ToString()));
        }

        private static List<object> NormalizeRowWidth(IList<object> row, int width)
        {
            var normalized = new List<object>();
            for (int i = 0; i < width; i++)
                normalized.Add(GetCell(row, i));

            return normalized;
        }

        private static List<string> SplitTags(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return new List<string>();

            return value
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string FormatTags(List<string> tags)
        {
            if (tags == null || tags.Count == 0)
                return "";

            return string.Join(", ", tags.Where(t => !string.IsNullOrWhiteSpace(t)).Select(t => t.Trim()));
        }

        private static void AddHeaderIfMissing(List<string> headers, string header)
        {
            if (string.IsNullOrWhiteSpace(header))
                return;

            if (!headers.Any(h => NormalizeHeader(h) == NormalizeHeader(header)))
                headers.Add(header);
        }

        private static bool HeadersEqual(List<string> left, List<string> right)
        {
            if (left.Count != right.Count)
                return false;

            for (int i = 0; i < left.Count; i++)
            {
                if (!string.Equals(left[i], right[i], StringComparison.Ordinal))
                    return false;
            }

            return true;
        }

        private static bool IsAny(string value, params string[] candidates)
        {
            return candidates.Any(c => string.Equals(value, NormalizeToken(c), StringComparison.OrdinalIgnoreCase));
        }

        private static string GetCell(IList<object> row, int index)
        {
            if (row == null || index < 0 || index >= row.Count)
                return "";

            return (row[index] ?? "").ToString().Trim();
        }

        private static void SetCell(IList<object> row, int index, string value)
        {
            if (row == null || index < 0)
                return;

            while (row.Count <= index)
                row.Add("");

            row[index] = value ?? "";
        }

        private static string NormalizeTitle(string value)
        {
            return (value ?? "").Trim().ToLowerInvariant();
        }

        private static string BuildCleanTitle(string title, string year)
        {
            title = (title ?? "").Trim();
            year = NormalizeNumberText(year);

            if (string.IsNullOrWhiteSpace(title))
                return "";

            if (string.IsNullOrWhiteSpace(year))
                return title;

            string yearSuffix = "(" + year + ")";
            if (title.EndsWith(yearSuffix, StringComparison.OrdinalIgnoreCase))
                return title;

            return title + " " + yearSuffix;
        }

        private static string NormalizeToken(string value)
        {
            return NormalizeHeader(value);
        }

        private static string NormalizeHeader(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            var chars = value
                .Trim()
                .Where(char.IsLetterOrDigit)
                .Select(char.ToLowerInvariant)
                .ToArray();

            return new string(chars);
        }

        private static string NormalizeImdb(string value)
        {
            value = (value ?? "").Trim();
            if (value.StartsWith("imdb://", StringComparison.OrdinalIgnoreCase))
                value = value.Substring("imdb://".Length);
            return value.ToLowerInvariant();
        }

        private static string NormalizeNumberText(string value)
        {
            value = (value ?? "").Trim();
            if (value.EndsWith(".0", StringComparison.OrdinalIgnoreCase))
                value = value.Substring(0, value.Length - 2);
            return value;
        }

        private static string ColumnIndexToLetter(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = "";

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private static string A1(string sheetName, string range)
        {
            return QuoteSheetName(sheetName) + "!" + range;
        }

        private static string QuoteSheetName(string sheetName)
        {
            return "'" + (sheetName ?? "").Replace("'", "''") + "'";
        }

        private static string DisplayName(MediaServerOptions server)
        {
            return string.IsNullOrWhiteSpace(server.Name) ? "Media server" : server.Name.Trim();
        }

        private class SheetData
        {
            public string SheetName { get; set; }
            public int HeaderRowNumber { get; set; }
            public int DataStartRowNumber { get; set; }
            public HeaderMap Headers { get; set; }
            public List<SheetRow> DataRows { get; set; }
        }

        private class SheetRow
        {
            public int SheetRowNumber { get; set; }
            public IList<object> Values { get; set; }
        }

        private class WebMovieColumnMapping
        {
            public string WebHeader { get; set; }
            public int MoviesIndex { get; set; } = -1;
            public int KidsIndex { get; set; } = -1;
            public int TeensIndex { get; set; } = -1;
            public bool IsComputedTags { get; set; }
        }

        private class HeaderMap
        {
            private readonly Dictionary<string, int> _indices = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            public HeaderMap(List<string> headers)
            {
                OriginalHeaders = headers ?? new List<string>();

                for (int i = 0; i < OriginalHeaders.Count; i++)
                {
                    string normalized = NormalizeHeader(OriginalHeaders[i]);
                    if (!string.IsNullOrWhiteSpace(normalized) && !_indices.ContainsKey(normalized))
                        _indices[normalized] = i;
                }
            }

            public List<string> OriginalHeaders { get; }
            public int Count => OriginalHeaders.Count;

            public bool Contains(params string[] aliases)
            {
                return FindIndex(aliases) >= 0;
            }

            public int FindIndex(params string[] aliases)
            {
                foreach (string alias in aliases)
                {
                    string normalized = NormalizeHeader(alias);
                    if (_indices.TryGetValue(normalized, out int index))
                        return index;
                }

                return -1;
            }

            public int GetRequiredIndex(params string[] aliases)
            {
                int index = FindIndex(aliases);
                if (index >= 0)
                    return index;

                throw new Exception("Missing required column: " + string.Join(" / ", aliases));
            }
        }

        private class ProfileRequest
        {
            public int SheetRowNumber { get; set; }
            public IList<object> RawRow { get; set; }
            public bool IsPending { get; set; }
            public string TagMode { get; set; }
            public string Action { get; set; }
            public List<string> Tags { get; set; }
            public string User { get; set; }
            public string Title { get; set; }
            public string Year { get; set; }
            public string CleanTitle { get; set; }
            public string TmdbId { get; set; }
            public string ImdbId { get; set; }
            public string RatingKey { get; set; }
            public int MovieRowNumber { get; set; }

            public static ProfileRequest FromRow(SheetData sheet, SheetRow row)
            {
                var headers = sheet.Headers;
                string status = GetFirst(row.Values, headers, "Processor Status", "Status", "Request Status");
                string user = GetFirst(row.Values, headers, HeaderUser, "Username", "Profile", "Profile Name");
                string title = GetFirst(row.Values, headers, HeaderTitle, "Movie Title", "Clean Title", "IMDB Title");
                string year = NormalizeNumberText(GetFirst(row.Values, headers, HeaderYear, "Movie Year", "Release Year"));

                string tagsText = GetFirst(
                    row.Values,
                    headers,
                    "Tags",
                    "Tag",
                    "Profile Tags",
                    "Profile Tag",
                    "Access Tags",
                    "Access Tag",
                    "User Tag",
                    "User Tags");

                if (string.IsNullOrWhiteSpace(tagsText))
                    tagsText = user;

                string cleanTitle = GetFirst(row.Values, headers, "Clean Title", "Movie Title", "IMDB Title");
                if (string.IsNullOrWhiteSpace(cleanTitle))
                    cleanTitle = BuildCleanTitle(title, year);

                int movieRowNumber = 0;
                int.TryParse(NormalizeNumberText(GetFirst(
                    row.Values,
                    headers,
                    "Movie Row",
                    "Movie RowNum",
                    "Movies Row",
                    "Movies RowNum",
                    "Movie Sheet Row",
                    "Movie Sheet RowNum")), out movieRowNumber);

                return new ProfileRequest
                {
                    SheetRowNumber = row.SheetRowNumber,
                    RawRow = row.Values,
                    IsPending = !RequestStatusDoneValues.Contains(NormalizeToken(status)),
                    TagMode = GetFirst(row.Values, headers, "tagMode", "Tag Mode", "Mode", "Access Mode"),
                    Action = GetFirst(row.Values, headers, "Action", "Operation", "Request", "Request Type", "Requested Access", "Access", "Change"),
                    Tags = SplitTags(tagsText),
                    User = user,
                    Title = title,
                    Year = year,
                    CleanTitle = cleanTitle,
                    TmdbId = NormalizeNumberText(GetFirst(row.Values, headers, "TMDB ID", "tmdbId", "TMDB")),
                    ImdbId = NormalizeImdb(GetFirst(row.Values, headers, "IMDB ID", "imdbId", "IMDB")),
                    RatingKey = NormalizeNumberText(GetFirst(row.Values, headers, "Rating Key", "ratingKey", "Plex Rating Key")),
                    MovieRowNumber = movieRowNumber
                };
            }

            private static string GetFirst(IList<object> row, HeaderMap headers, params string[] aliases)
            {
                int index = headers.FindIndex(aliases);
                return index >= 0 ? GetCell(row, index) : "";
            }
        }

        private class MovieRow
        {
            public int SheetRowNumber { get; set; }
            public IList<object> Values { get; set; }
            public string CleanTitle { get; set; }
            public string TmdbId { get; set; }
            public string ImdbId { get; set; }
            public string RatingKey { get; set; }
            public string TagsText { get; set; }
            public int TagsColumnIndex { get; set; }
        }

        private class MovieIndex
        {
            private readonly Dictionary<int, MovieRow> _byRow = new Dictionary<int, MovieRow>();
            private readonly Dictionary<string, MovieRow> _byCleanTitle = new Dictionary<string, MovieRow>(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _ambiguousCleanTitles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, MovieRow> _byTmdbId = new Dictionary<string, MovieRow>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, MovieRow> _byImdbId = new Dictionary<string, MovieRow>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, MovieRow> _byRatingKey = new Dictionary<string, MovieRow>(StringComparer.OrdinalIgnoreCase);

            public static MovieIndex Build(SheetData moviesSheet)
            {
                var index = new MovieIndex();

                int cleanTitleCol = moviesSheet.Headers.FindIndex(HeaderCleanTitle, "Title", "Movie Title");
                int tagsCol = moviesSheet.Headers.GetRequiredIndex(HeaderTags);
                int tmdbCol = moviesSheet.Headers.FindIndex(HeaderTmdbId, "tmdbId");
                int imdbCol = moviesSheet.Headers.FindIndex(HeaderImdbId, "imdbId");
                int ratingKeyCol = moviesSheet.Headers.FindIndex(HeaderRatingKey, "ratingKey", "Plex Rating Key");

                foreach (var row in moviesSheet.DataRows)
                {
                    if (!RowHasData(row.Values))
                        continue;

                    var movie = new MovieRow
                    {
                        SheetRowNumber = row.SheetRowNumber,
                        Values = row.Values,
                        CleanTitle = GetCell(row.Values, cleanTitleCol),
                        TmdbId = NormalizeNumberText(GetCell(row.Values, tmdbCol)),
                        ImdbId = NormalizeImdb(GetCell(row.Values, imdbCol)),
                        RatingKey = NormalizeNumberText(GetCell(row.Values, ratingKeyCol)),
                        TagsText = GetCell(row.Values, tagsCol),
                        TagsColumnIndex = tagsCol
                    };

                    index._byRow[movie.SheetRowNumber] = movie;
                    index.AddTitle(movie.CleanTitle, movie);
                    AddIfMissing(index._byTmdbId, movie.TmdbId, movie);
                    AddIfMissing(index._byImdbId, movie.ImdbId, movie);
                    AddIfMissing(index._byRatingKey, movie.RatingKey, movie);
                }

                return index;
            }

            public MovieRow Find(ProfileRequest request)
            {
                MovieRow movie;

                if (request.MovieRowNumber > 0 && _byRow.TryGetValue(request.MovieRowNumber, out movie))
                    return movie;

                if (!string.IsNullOrWhiteSpace(request.TmdbId) && _byTmdbId.TryGetValue(request.TmdbId, out movie))
                    return movie;

                if (!string.IsNullOrWhiteSpace(request.ImdbId) && _byImdbId.TryGetValue(request.ImdbId, out movie))
                    return movie;

                if (!string.IsNullOrWhiteSpace(request.RatingKey) && _byRatingKey.TryGetValue(request.RatingKey, out movie))
                    return movie;

                string titleKey = NormalizeTitle(request.CleanTitle);
                if (!string.IsNullOrWhiteSpace(titleKey) && _ambiguousCleanTitles.Contains(titleKey))
                    return null;

                if (!string.IsNullOrWhiteSpace(titleKey) && _byCleanTitle.TryGetValue(titleKey, out movie))
                    return movie;

                return null;
            }

            private void AddTitle(string title, MovieRow movie)
            {
                string key = NormalizeTitle(title);
                if (string.IsNullOrWhiteSpace(key))
                    return;

                MovieRow existing;
                if (_byCleanTitle.TryGetValue(key, out existing))
                {
                    if (existing.SheetRowNumber != movie.SheetRowNumber)
                        _ambiguousCleanTitles.Add(key);

                    return;
                }

                _byCleanTitle[key] = movie;
            }

            private static void AddIfMissing(Dictionary<string, MovieRow> index, string key, MovieRow movie)
            {
                if (string.IsNullOrWhiteSpace(key))
                    return;

                if (!index.ContainsKey(key))
                    index[key] = movie;
            }
        }

        private class ProcessingResult
        {
            public ProfileRequest Request { get; set; }
            public bool Success { get; set; }
            public bool TagsChanged { get; set; }
            public int MovieRowNumber { get; set; }
            public string MovieCleanTitle { get; set; }
            public string MovieTmdbId { get; set; }
            public string MovieImdbId { get; set; }
            public string MovieRatingKey { get; set; }
            public string TagsBefore { get; set; }
            public string TagsAfter { get; set; }
            public string Message { get; set; }

            public ProcessingResult Fail(string message)
            {
                Success = false;
                Message = message;
                return this;
            }
        }
    }
}
