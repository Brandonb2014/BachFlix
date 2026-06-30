using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Pushes selected Movies sheet metadata into Plex movie items.
    /// </summary>
    public static class PlexMetadataSheetUpdater
    {
        public class Options
        {
            public string PlexBaseUrl { get; set; } = "http://192.168.0.5:32400";
            public string PlexToken { get; set; } = "";
            public int MoviesSectionId { get; set; } = 1;

            public string MoviesSheetName { get; set; } = "Movies";
            public int MoviesHeaderRow { get; set; } = 2;
            public int MoviesDataStartRow { get; set; } = 3;
            public string StatusFilterValue { get; set; } = "n";
            public int PlexRequestTimeoutSeconds { get; set; } = 45;

            public bool LockEditedFields { get; set; } = true;
            public bool ReplaceLabelsFromTags { get; set; } = true;
            public bool IncludeNfoBodyTagsAsLabels { get; set; } = false;
            public bool ClearEmptySortTitle { get; set; } = false;
            public bool ClearEmptyPlot { get; set; } = false;
            public bool ClearEmptyContentRating { get; set; } = false;

            public IList<int> MovieRowNumbersToUpdate { get; set; } = new List<int>();
        }

        public class Summary
        {
            public int RowsRead { get; set; }
            public int RowsEligible { get; set; }
            public int RowsSkippedStatus { get; set; }
            public int RowsSkippedMissingRatingKey { get; set; }
            public int RowsUnchanged { get; set; }
            public int RowsUpdated { get; set; }
            public int RowsFailed { get; set; }
            public int LabelsChanged { get; set; }
            public int SortTitlesChanged { get; set; }
            public int PlotsChanged { get; set; }
            public int ContentRatingsChanged { get; set; }
            public int RowsWithNfoBodyTags { get; set; }
        }

        public static Summary Run(
            SheetsService sheetsService,
            string spreadsheetId,
            Action<string, string, int> log,
            Options options)
        {
            return RunAsync(sheetsService, spreadsheetId, log, options)
                .GetAwaiter()
                .GetResult();
        }

        private static async Task<Summary> RunAsync(
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
            if (string.IsNullOrWhiteSpace(options.PlexBaseUrl))
                throw new ArgumentException("PlexBaseUrl is required.");
            if (string.IsNullOrWhiteSpace(options.PlexToken))
                throw new ArgumentException("PlexToken is required.");

            var summary = new Summary();

            log?.Invoke("info", "Plex metadata sync: reading Movies sheet...", 2);

            var sheet = ReadMoviesSheet(sheetsService, spreadsheetId, options);
            var headers = new HeaderMap(sheet.Headers);

            int ratingKeyCol = headers.GetRequiredIndex("Rating Key", "ratingKey", "Plex Rating Key");
            int statusCol = headers.GetRequiredIndex("Status");
            int tagsCol = headers.FindIndex("Tags", "Tag", "Labels", "Plex Labels");
            int nfoBodyCol = options.IncludeNfoBodyTagsAsLabels
                ? headers.FindIndex("NFO Body", "Nfo Body", "NFO")
                : -1;
            int sortTitleCol = headers.GetRequiredIndex("Sort Title", "SortTitle", "Plex Sort Title");
            int altPlotCol = headers.GetRequiredIndex("Alt Plot", "Alternate Plot", "Plex Plot", "Plex Summary");
            int contentRatingCol = headers.GetRequiredIndex("Content Rating", "Plex Content Rating");
            int titleCol = headers.FindIndex("IMDB Title", "Clean Title", "Title", "Auto Title");

            if (tagsCol < 0 && nfoBodyCol < 0)
                throw new Exception("Missing required Movies column for Plex labels: Tags / Tag / Labels / Plex Labels, or NFO Body with <tag> values.");

            if (nfoBodyCol >= 0)
                log?.Invoke("info", "Plex metadata sync: using sheet labels plus <tag> values found in NFO Body. This writes labels through the Plex API; it does not prove Plex read the NFO file.", 2);

            HashSet<int> rowFilter = BuildRowFilter(options);
            var rowsToProcess = new List<MovieSheetRow>();

            foreach (var row in sheet.Rows)
            {
                if (!RowHasData(row.Values))
                    continue;

                summary.RowsRead++;

                if (rowFilter != null && !rowFilter.Contains(row.SheetRowNumber))
                    continue;

                string status = GetCell(row.Values, statusCol);
                if (!ShouldIncludeStatus(status, options.StatusFilterValue))
                {
                    summary.RowsSkippedStatus++;
                    continue;
                }

                summary.RowsEligible++;

                string ratingKey = NormalizeRatingKey(GetCell(row.Values, ratingKeyCol));
                if (string.IsNullOrWhiteSpace(ratingKey))
                {
                    summary.RowsSkippedMissingRatingKey++;
                    log?.Invoke("warning", $"Movies row {row.SheetRowNumber} skipped: missing Rating Key.", 3);
                    continue;
                }

                rowsToProcess.Add(new MovieSheetRow
                {
                    SheetRowNumber = row.SheetRowNumber,
                    RatingKey = ratingKey,
                    Title = GetCell(row.Values, titleCol),
                    Tags = BuildDesiredLabels(row.Values, tagsCol, nfoBodyCol, summary),
                    SortTitle = GetCell(row.Values, sortTitleCol),
                    Plot = GetCell(row.Values, altPlotCol),
                    ContentRating = GetCell(row.Values, contentRatingCol)
                });
            }

            if (rowsToProcess.Count == 0)
            {
                log?.Invoke("info", "Plex metadata sync: no Movies rows to update.", 2);
                return summary;
            }

            TimeSpan requestTimeout = GetRequestTimeout(options);
            log?.Invoke("info", $"Plex metadata sync: updating {rowsToProcess.Count} movie row(s)... Request timeout={requestTimeout.TotalSeconds:0}s.", 2);

            using (var http = new HttpClient())
            {
                http.Timeout = requestTimeout;

                foreach (var row in rowsToProcess)
                {
                    try
                    {
                        var current = await FetchPlexMovieAsync(http, options, row.RatingKey).ConfigureAwait(false);
                        var changes = DetermineChanges(row, current, options);

                        if (!changes.HasAnyChange)
                        {
                            summary.RowsUnchanged++;
                            continue;
                        }

                        await ApplyPlexChangesAsync(http, options, row, current, changes).ConfigureAwait(false);

                        summary.RowsUpdated++;
                        if (changes.LabelsChanged) summary.LabelsChanged++;
                        if (changes.SortTitleChanged) summary.SortTitlesChanged++;
                        if (changes.PlotChanged) summary.PlotsChanged++;
                        if (changes.ContentRatingChanged) summary.ContentRatingsChanged++;

                        string title = string.IsNullOrWhiteSpace(row.Title) ? current.Title : row.Title;
                        log?.Invoke("success", $"Plex updated: {title} (Movies row {row.SheetRowNumber}, ratingKey {row.RatingKey}; changed {DescribeChanges(row, changes)})", 3);
                    }
                    catch (Exception ex)
                    {
                        summary.RowsFailed++;
                        string title = string.IsNullOrWhiteSpace(row.Title) ? row.RatingKey : row.Title;
                        log?.Invoke("error", $"Plex update failed for {title} (Movies row {row.SheetRowNumber}, ratingKey {row.RatingKey}): {DescribeException(ex)}", 3);
                    }
                }
            }

            log?.Invoke(
                "success",
                $"Plex metadata sync complete. ActiveStatus={options.StatusFilterValue}, Eligible={summary.RowsEligible}, Updated={summary.RowsUpdated}, Unchanged={summary.RowsUnchanged}, Failed={summary.RowsFailed}, MissingRatingKey={summary.RowsSkippedMissingRatingKey}, SkippedOtherStatus={summary.RowsSkippedStatus}. Changed: Labels={summary.LabelsChanged}, SortTitle={summary.SortTitlesChanged}, AltPlot={summary.PlotsChanged}, ContentRating={summary.ContentRatingsChanged}. RowsWithNfoBodyTags={summary.RowsWithNfoBodyTags}.",
                2);

            if (summary.RowsSkippedMissingRatingKey > 0)
                log?.Invoke("warning", "Rows skipped for missing Rating Key cannot be edited in Plex by menu 60. Run menu 58/58o to populate ratingKeys, then run 60 again.", 2);

            return summary;
        }

        private static bool ShouldIncludeStatus(string status, string requiredStatus)
        {
            requiredStatus = (requiredStatus ?? "").Trim();
            if (string.IsNullOrWhiteSpace(requiredStatus))
                return true;

            return string.Equals((status ?? "").Trim(), requiredStatus, StringComparison.OrdinalIgnoreCase);
        }

        private static PlexChanges DetermineChanges(MovieSheetRow desired, PlexMovieMetadata current, Options options)
        {
            var changes = new PlexChanges();

            if (options.ReplaceLabelsFromTags)
                changes.LabelsChanged = !TagSetsEqual(desired.Tags, current.Labels);

            bool hasSortTitleValue = !string.IsNullOrWhiteSpace(desired.SortTitle);
            if (hasSortTitleValue || options.ClearEmptySortTitle)
            {
                changes.SortTitleChanged = !TextsEqual(
                    hasSortTitleValue ? desired.SortTitle : "",
                    current.TitleSort);
            }

            bool hasPlotValue = !string.IsNullOrWhiteSpace(desired.Plot);
            if (hasPlotValue || options.ClearEmptyPlot)
            {
                changes.PlotChanged = !TextsEqual(
                    hasPlotValue ? desired.Plot : "",
                    current.Summary);
            }

            bool hasContentRatingValue = !string.IsNullOrWhiteSpace(desired.ContentRating);
            if (hasContentRatingValue || options.ClearEmptyContentRating)
            {
                changes.ContentRatingChanged = !TextsEqual(
                    hasContentRatingValue ? desired.ContentRating : "",
                    current.ContentRating);
            }

            return changes;
        }

        private static async Task ApplyPlexChangesAsync(
            HttpClient http,
            Options options,
            MovieSheetRow desired,
            PlexMovieMetadata current,
            PlexChanges changes)
        {
            var updateParams = BuildEditBaseParams(desired.RatingKey);
            bool lockValue = options.LockEditedFields;

            if (changes.SortTitleChanged)
            {
                updateParams["titleSort.value"] = string.IsNullOrWhiteSpace(desired.SortTitle) ? "" : desired.SortTitle;
                updateParams["titleSort.locked"] = lockValue ? "1" : "0";
            }

            if (changes.PlotChanged)
            {
                updateParams["summary.value"] = string.IsNullOrWhiteSpace(desired.Plot) ? "" : desired.Plot;
                updateParams["summary.locked"] = lockValue ? "1" : "0";
            }

            if (changes.ContentRatingChanged)
            {
                updateParams["contentRating.value"] = string.IsNullOrWhiteSpace(desired.ContentRating) ? "" : desired.ContentRating;
                updateParams["contentRating.locked"] = lockValue ? "1" : "0";
            }

            if (changes.LabelsChanged && desired.Tags.Count > 0)
            {
                updateParams["label.locked"] = lockValue ? "1" : "0";
                for (int i = 0; i < desired.Tags.Count; i++)
                    updateParams[$"label[{i}].tag.tag"] = desired.Tags[i];
            }

            if (updateParams.Count > 2)
                await PutPlexEditAsync(http, options, updateParams).ConfigureAwait(false);

            if (changes.LabelsChanged && current.Labels.Count > 0)
            {
                var desiredSet = new HashSet<string>(desired.Tags, StringComparer.OrdinalIgnoreCase);
                var labelsToRemove = current.Labels
                    .Where(label => !desiredSet.Contains(label))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (labelsToRemove.Count > 0)
                {
                    var removeParams = BuildEditBaseParams(desired.RatingKey);
                    removeParams["label.locked"] = lockValue ? "1" : "0";
                    removeParams["label[].tag.tag-"] = string.Join(",", labelsToRemove);

                    await PutPlexEditAsync(http, options, removeParams).ConfigureAwait(false);
                }
            }
        }

        private static async Task<PlexMovieMetadata> FetchPlexMovieAsync(HttpClient http, Options options, string ratingKey)
        {
            string endpoint = "/library/metadata/" + ratingKey + "?includeGuids=1";
            string url =
                options.PlexBaseUrl.TrimEnd('/') +
                "/library/metadata/" +
                Uri.EscapeDataString(ratingKey) +
                "?includeGuids=1&X-Plex-Token=" +
                Uri.EscapeDataString(options.PlexToken);

            HttpResponseMessage response;
            string body;
            try
            {
                response = await http.GetAsync(url).ConfigureAwait(false);
                body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
            catch (TaskCanceledException ex)
            {
                throw BuildPlexRequestCanceledException("GET", endpoint, ratingKey, options, ex);
            }
            catch (HttpRequestException ex)
            {
                throw BuildPlexHttpException("GET", endpoint, ratingKey, ex);
            }

            if (!response.IsSuccessStatusCode)
                throw new Exception($"Plex metadata GET failed: {(int)response.StatusCode} {response.ReasonPhrase}");

            var doc = XDocument.Parse(body);
            XElement video = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "Video");
            if (video == null && doc.Root != null && doc.Root.Name.LocalName == "Video")
                video = doc.Root;

            if (video == null)
                throw new Exception("Plex metadata response did not contain a movie Video node.");

            return new PlexMovieMetadata
            {
                RatingKey = ratingKey,
                Title = (string)video.Attribute("title") ?? "",
                TitleSort = (string)video.Attribute("titleSort") ?? "",
                Summary = (string)video.Attribute("summary") ?? "",
                ContentRating = (string)video.Attribute("contentRating") ?? "",
                Labels = video
                    .Descendants()
                    .Where(e => e.Name.LocalName == "Label")
                    .Select(e => ((string)e.Attribute("tag") ?? "").Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList()
            };
        }

        private static async Task PutPlexEditAsync(HttpClient http, Options options, Dictionary<string, string> parameters)
        {
            string ratingKey = parameters != null && parameters.ContainsKey("id") ? parameters["id"] : "";
            string endpoint = "/library/sections/" + options.MoviesSectionId + "/all";
            string editedFields = DescribeEditedFields(parameters);
            string url =
                options.PlexBaseUrl.TrimEnd('/') +
                "/library/sections/" +
                options.MoviesSectionId.ToString() +
                "/all" +
                BuildQueryString(parameters) +
                "&X-Plex-Token=" +
                Uri.EscapeDataString(options.PlexToken);

            HttpResponseMessage response;
            try
            {
                response = await http.PutAsync(url, new StringContent("", Encoding.UTF8)).ConfigureAwait(false);
            }
            catch (TaskCanceledException ex)
            {
                throw BuildPlexRequestCanceledException("PUT", endpoint + " (" + editedFields + ")", ratingKey, options, ex);
            }
            catch (HttpRequestException ex)
            {
                throw BuildPlexHttpException("PUT", endpoint + " (" + editedFields + ")", ratingKey, ex);
            }

            if (!response.IsSuccessStatusCode)
            {
                string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!string.IsNullOrWhiteSpace(body) && body.Length > 300)
                    body = body.Substring(0, 300);

                throw new Exception($"Plex metadata PUT failed: {(int)response.StatusCode} {response.ReasonPhrase} {body}");
            }
        }

        private static TimeSpan GetRequestTimeout(Options options)
        {
            int seconds = options == null ? 45 : options.PlexRequestTimeoutSeconds;
            if (seconds < 5)
                seconds = 5;

            return TimeSpan.FromSeconds(seconds);
        }

        private static Exception BuildPlexRequestCanceledException(
            string method,
            string endpoint,
            string ratingKey,
            Options options,
            TaskCanceledException ex)
        {
            TimeSpan timeout = GetRequestTimeout(options);
            string message =
                $"Plex metadata {method} timed out or was canceled after {timeout.TotalSeconds:0}s " +
                $"(ratingKey {ratingKey}, endpoint {endpoint}). " +
                "Because this sync does not pass its own cancellation token, this usually means Plex or the network did not respond before the request timeout.";

            return new Exception(message, ex);
        }

        private static Exception BuildPlexHttpException(
            string method,
            string endpoint,
            string ratingKey,
            HttpRequestException ex)
        {
            return new Exception(
                $"Plex metadata {method} network failure (ratingKey {ratingKey}, endpoint {endpoint}): {ex.Message}",
                ex);
        }

        private static string DescribeException(Exception ex)
        {
            if (ex == null)
                return "";

            string message = ex.Message;
            if (ex.InnerException != null)
            {
                message += $" | Inner {ex.InnerException.GetType().Name}: {ex.InnerException.Message}";
            }

            return message;
        }

        private static string DescribeEditedFields(Dictionary<string, string> parameters)
        {
            if (parameters == null || parameters.Count == 0)
                return "fields unknown";

            var fields = new List<string>();
            if (parameters.Keys.Any(k => k.StartsWith("label", StringComparison.OrdinalIgnoreCase)))
                fields.Add("Labels");
            if (parameters.ContainsKey("titleSort.value"))
                fields.Add("SortTitle");
            if (parameters.ContainsKey("summary.value"))
                fields.Add("AltPlot");
            if (parameters.ContainsKey("contentRating.value"))
                fields.Add("ContentRating");

            return fields.Count == 0 ? "fields unknown" : string.Join(", ", fields);
        }

        private static Dictionary<string, string> BuildEditBaseParams(string ratingKey)
        {
            return new Dictionary<string, string>
            {
                { "type", "1" },
                { "id", ratingKey }
            };
        }

        private static string BuildQueryString(Dictionary<string, string> parameters)
        {
            if (parameters == null || parameters.Count == 0)
                return "";

            var parts = parameters.Select(kv =>
                kv.Key + "=" + Uri.EscapeDataString(kv.Value ?? ""));

            return "?" + string.Join("&", parts);
        }

        private static SheetData ReadMoviesSheet(SheetsService sheetsService, string spreadsheetId, Options options)
        {
            var request = sheetsService.Spreadsheets.Values.Get(
                spreadsheetId,
                A1(options.MoviesSheetName, "A:ZZ"));
            var response = request.Execute();
            var rows = response.Values ?? new List<IList<object>>();

            if (rows.Count < options.MoviesHeaderRow)
                throw new Exception($"Sheet '{options.MoviesSheetName}' does not contain header row {options.MoviesHeaderRow}.");

            var headers = rows[options.MoviesHeaderRow - 1]
                .Select(v => (v ?? "").ToString())
                .ToList();

            var dataRows = new List<SheetRow>();
            for (int i = options.MoviesDataStartRow - 1; i < rows.Count; i++)
            {
                dataRows.Add(new SheetRow
                {
                    SheetRowNumber = i + 1,
                    Values = rows[i]
                });
            }

            return new SheetData
            {
                Headers = headers,
                Rows = dataRows
            };
        }

        private static HashSet<int> BuildRowFilter(Options options)
        {
            if (options.MovieRowNumbersToUpdate == null || options.MovieRowNumbersToUpdate.Count == 0)
                return null;

            return new HashSet<int>(options.MovieRowNumbersToUpdate.Where(r => r > 0));
        }

        private static string A1(string sheetName, string range)
        {
            return "'" + (sheetName ?? "").Replace("'", "''") + "'!" + range;
        }

        private static bool RowHasData(IList<object> row)
        {
            return row != null && row.Any(v => !string.IsNullOrWhiteSpace((v ?? "").ToString()));
        }

        private static string GetCell(IList<object> row, int index)
        {
            if (row == null || index < 0 || index >= row.Count)
                return "";

            return (row[index] ?? "").ToString().Trim();
        }

        private static List<string> SplitTags(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return new List<string>();

            return DistinctTags(value
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => !string.IsNullOrWhiteSpace(t)));
        }

        private static List<string> BuildDesiredLabels(IList<object> row, int tagsCol, int nfoBodyCol, Summary summary)
        {
            var labels = new List<string>();

            if (tagsCol >= 0)
                labels.AddRange(SplitTags(GetCell(row, tagsCol)));

            if (nfoBodyCol >= 0)
            {
                var nfoTags = ExtractNfoBodyTags(GetCell(row, nfoBodyCol));
                if (nfoTags.Count > 0)
                {
                    summary.RowsWithNfoBodyTags++;
                    labels.AddRange(nfoTags);
                }
            }

            return DistinctTags(labels);
        }

        private static List<string> ExtractNfoBodyTags(string nfoBody)
        {
            nfoBody = TrimSurroundingQuotes(nfoBody);
            if (string.IsNullOrWhiteSpace(nfoBody))
                return new List<string>();

            try
            {
                var doc = XDocument.Parse(nfoBody);
                return DistinctTags(doc
                    .Descendants()
                    .Where(e => e.Name.LocalName == "tag")
                    .Select(e => (e.Value ?? "").Trim())
                    .Where(t => !string.IsNullOrWhiteSpace(t)));
            }
            catch
            {
                var matches = Regex.Matches(
                    nfoBody,
                    @"<\s*tag(?:\s+[^>]*)?\s*>(?<tag>.*?)<\s*/\s*tag\s*>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                return DistinctTags(matches
                    .Cast<Match>()
                    .Select(m => System.Net.WebUtility.HtmlDecode(m.Groups["tag"].Value).Trim())
                    .Where(t => !string.IsNullOrWhiteSpace(t)));
            }
        }

        private static string TrimSurroundingQuotes(string value)
        {
            value = (value ?? "").Trim();
            if (value.Length >= 2)
            {
                char first = value[0];
                char last = value[value.Length - 1];
                if ((first == '"' && last == '"') || (first == '\'' && last == '\''))
                    value = value.Substring(1, value.Length - 2).Trim();
            }

            return value;
        }

        private static List<string> DistinctTags(IEnumerable<string> tags)
        {
            return (tags ?? Enumerable.Empty<string>())
                .Select(t => (t ?? "").Trim())
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool TagSetsEqual(List<string> left, List<string> right)
        {
            var leftSet = new HashSet<string>(left ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
            var rightSet = new HashSet<string>(right ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
            return leftSet.SetEquals(rightSet);
        }

        private static bool TextsEqual(string left, string right)
        {
            return string.Equals(NormalizeText(left), NormalizeText(right), StringComparison.Ordinal);
        }

        private static string DescribeChanges(MovieSheetRow desired, PlexChanges changes)
        {
            var parts = new List<string>();

            if (changes.LabelsChanged)
                parts.Add("Labels=" + (desired.Tags.Count == 0 ? "(none)" : string.Join(", ", desired.Tags)));
            if (changes.SortTitleChanged)
                parts.Add("SortTitle");
            if (changes.PlotChanged)
                parts.Add("AltPlot");
            if (changes.ContentRatingChanged)
                parts.Add("ContentRating");

            return string.Join(", ", parts);
        }

        private static string NormalizeText(string value)
        {
            return (value ?? "")
                .Replace("\r\n", "\n")
                .Replace("\r", "\n")
                .Trim();
        }

        private static string NormalizeRatingKey(string value)
        {
            value = (value ?? "").Trim();
            if (value.EndsWith(".0", StringComparison.OrdinalIgnoreCase))
                value = value.Substring(0, value.Length - 2);

            return value;
        }

        private static string NormalizeHeader(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            return new string(
                value.Trim()
                    .Where(char.IsLetterOrDigit)
                    .Select(char.ToLowerInvariant)
                    .ToArray());
        }

        private class HeaderMap
        {
            private readonly Dictionary<string, int> _indices = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            public HeaderMap(List<string> headers)
            {
                Headers = headers ?? new List<string>();
                for (int i = 0; i < Headers.Count; i++)
                {
                    string normalized = NormalizeHeader(Headers[i]);
                    if (!string.IsNullOrWhiteSpace(normalized) && !_indices.ContainsKey(normalized))
                        _indices[normalized] = i;
                }
            }

            public List<string> Headers { get; }

            public int FindIndex(params string[] aliases)
            {
                foreach (string alias in aliases)
                {
                    int index;
                    if (_indices.TryGetValue(NormalizeHeader(alias), out index))
                        return index;
                }

                return -1;
            }

            public int GetRequiredIndex(params string[] aliases)
            {
                int index = FindIndex(aliases);
                if (index >= 0)
                    return index;

                throw new Exception("Missing required Movies column: " + string.Join(" / ", aliases));
            }
        }

        private class SheetData
        {
            public List<string> Headers { get; set; }
            public List<SheetRow> Rows { get; set; }
        }

        private class SheetRow
        {
            public int SheetRowNumber { get; set; }
            public IList<object> Values { get; set; }
        }

        private class MovieSheetRow
        {
            public int SheetRowNumber { get; set; }
            public string RatingKey { get; set; }
            public string Title { get; set; }
            public List<string> Tags { get; set; }
            public string SortTitle { get; set; }
            public string Plot { get; set; }
            public string ContentRating { get; set; }
        }

        private class PlexMovieMetadata
        {
            public string RatingKey { get; set; }
            public string Title { get; set; }
            public string TitleSort { get; set; }
            public string Summary { get; set; }
            public string ContentRating { get; set; }
            public List<string> Labels { get; set; }
        }

        private class PlexChanges
        {
            public bool LabelsChanged { get; set; }
            public bool SortTitleChanged { get; set; }
            public bool PlotChanged { get; set; }
            public bool ContentRatingChanged { get; set; }

            public bool HasAnyChange
            {
                get { return LabelsChanged || SortTitleChanged || PlotChanged || ContentRatingChanged; }
            }
        }
    }
}
