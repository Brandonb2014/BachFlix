using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net.Http;

namespace BachFlixNfo
{
    internal sealed class EmbyException : InvalidOperationException
    {
        internal bool Systemic { get; private set; }
        internal EmbyException(string message, bool systemic = true) : base(message) { Systemic = systemic; }
    }

    // Shared transport for movie refresh and read-only TV inventory. Never put credentials in URLs.
    internal sealed class EmbyClient : IDisposable
    {
        private readonly HttpClient client;
        private readonly Action<TimeSpan> delay;
        internal EmbyClient(string server, string apiKey, HttpMessageHandler handler = null, Action<TimeSpan> delay = null)
        {
            this.delay = delay;
            Uri uri;
            if (!Uri.TryCreate(server, UriKind.Absolute, out uri) ||
                (uri.Scheme != "http" && uri.Scheme != "https") || !string.IsNullOrEmpty(uri.UserInfo) ||
                !string.IsNullOrEmpty(uri.Query) || !string.IsNullOrEmpty(uri.Fragment))
                throw new EmbyException("Set EMBY_SERVER_URL (or EMBY_SERVER) to a valid Emby server URL.");
            if (string.IsNullOrWhiteSpace(apiKey)) throw new EmbyException("Set EMBY_API_KEY in BachFlixNfo.local.env or the environment.");
            string root = server.TrimEnd('/');
            if (!root.EndsWith("/emby", StringComparison.OrdinalIgnoreCase)) root += "/emby";
            client = new HttpClient(handler ?? new HttpClientHandler { AllowAutoRedirect = false })
            { BaseAddress = new Uri(root + "/"), Timeout = TimeSpan.FromSeconds(30) };
            client.DefaultRequestHeaders.Add("X-Emby-Token", apiKey);
            client.DefaultRequestHeaders.Add("Accept", "application/json");
        }

        private HttpResponseMessage Send(HttpMethod method, string resource)
        {
            for (int attempt = 0; ; attempt++)
            {
                HttpResponseMessage response;
                try
                {
                    using (var request = new HttpRequestMessage(method, resource))
                        response = client.SendAsync(request).GetAwaiter().GetResult();
                }
                catch (Exception)
                {
                    if (method == HttpMethod.Get && attempt < 2) { ApiRetry.Pause(attempt, null, delay); continue; }
                    throw new EmbyException("Emby network failure after bounded retries; inventory unavailable.");
                }
                if (response.IsSuccessStatusCode) return response;
                int status = (int)response.StatusCode;
                var retry = response.Headers.RetryAfter;
                TimeSpan? wait = retry == null ? null : retry.Delta ?? (retry.Date.HasValue ? retry.Date.Value - DateTimeOffset.UtcNow : (TimeSpan?)null);
                response.Dispose();
                if (method == HttpMethod.Get && ApiRetry.Transient(status) && attempt < 2)
                {
                    try { ApiRetry.Pause(attempt, wait, delay); }
                    catch (InvalidOperationException) { throw new EmbyException("Emby requested a long retry delay; retry later."); }
                    continue;
                }
                throw new EmbyException("Emby request failed (HTTP " + status + ").", status != 404);
            }
        }

        internal void RefreshMovies(string libraryId)
        {
            using (Send(HttpMethod.Post, "Items/" + Uri.EscapeDataString(libraryId) +
                "/Refresh?Recursive=true&ImageRefreshMode=FullRefresh&MetadataRefreshMode=Default&ReplaceAllMetadata=false&ReplaceAllImages=false")) { }
        }

        internal IList<JObject> GetSeries() { return GetItems("Series", null); }
        internal IList<JObject> GetEpisodes(string seriesId) { return GetItems("Episode", seriesId); }

        private IList<JObject> GetItems(string type, string parent)
        {
            var result = new List<JObject>();
            var ids = new HashSet<string>(StringComparer.Ordinal);
            int? expectedTotal = null;
            for (int page = 0; page < 10000; page++)
            {
                // Request media evidence explicitly; current servers may omit LocationType.
                // Do not filter virtual/placeholder records out of the response: validate locally.
                string fields = type == "Episode" ? "ProviderIds,Path,MediaSources" : "ProviderIds";
                string resource = "Items?IncludeItemTypes=" + type + "&Recursive=true&Fields=" + fields + "&SortBy=SortName&SortOrder=Ascending&EnableImages=false&EnableTotalRecordCount=true&Limit=500&StartIndex=" + result.Count;
                if (parent != null) resource += "&ParentId=" + Uri.EscapeDataString(parent);
                JObject json;
                using (var response = Send(HttpMethod.Get, resource))
                {
                    try { json = JObject.Parse(response.Content.ReadAsStringAsync().GetAwaiter().GetResult()); }
                    catch (Exception) { throw new EmbyException("Emby returned unreadable inventory data.", false); }
                }
                var items = json["Items"] as JArray;
                int total;
                if (items == null || !int.TryParse((string)json["TotalRecordCount"], out total) || total < 0)
                    throw new EmbyException("Emby returned incomplete inventory metadata.", false);
                if (expectedTotal.HasValue && expectedTotal != total)
                    throw new EmbyException("Emby inventory changed during pagination; run again.", false);
                expectedTotal = total;
                foreach (var item in items)
                {
                    var obj = item as JObject;
                    string id = obj == null ? null : (string)obj["Id"];
                    if (string.IsNullOrWhiteSpace(id) || !ids.Add(id))
                        throw new EmbyException("Emby returned invalid or repeated inventory items.", false);
                    if (parent != null && ((string)obj["SeriesId"] != parent || (string)obj["Type"] != "Episode"))
                        throw new EmbyException("Emby episode belongs to a different or unknown series.", false);
                    result.Add(obj);
                }
                if (result.Count == total) return result;
                if (items.Count == 0 || result.Count > total)
                    throw new EmbyException("Emby returned a partial inventory; no count can be trusted.", false);
            }
            throw new EmbyException("Emby inventory pagination limit exceeded.");
        }

        public void Dispose() { client.Dispose(); }
    }
}
