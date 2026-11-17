using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Linq;
using System.Net;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using BachFlixNfo;

namespace TvdbApiCall
{
    // The synchronous public facade retains the existing callers' signatures and JSON fields.
    class TvdbApi
    {
        private static readonly TvdbClient Client = new TvdbClient(
            () => LocalEnvironment.GetRequired("TVDB_API_KEY"));

        public static bool LogIntoTvdbAsync(ref string token)
        {
            token = Client.Login();
            return true;
        }

        public static string GetSeriesIdAsync(ref string token, string TvdbSlug)
        {
            try { return Client.GetSeriesId(TvdbSlug); }
            finally { token = Client.Token; }
        }

        public static dynamic GetSeriesDetailsAsync(ref string token, string TvdbId)
        {
            try { return Client.GetSeriesDetails(TvdbId); }
            finally { token = Client.Token; }
        }

        public static dynamic GetTvEpisodeDetails(ref string token, string TvdbID, string seasonNum, string episodeNum)
        {
            try { return Client.GetEpisodeDetails(TvdbID, seasonNum, episodeNum); }
            finally { token = Client.Token; }
        }
    }

    internal enum TvdbFailure { Authentication, SeriesNotFound, EpisodeNotFound, Api, Inventory }

    internal sealed class TvdbException : InvalidOperationException
    {
        public TvdbFailure Failure { get; private set; }
        public TvdbException(TvdbFailure failure, string message) : base(message) { Failure = failure; }
    }

    internal sealed class TvdbClient
    {
        internal const string BaseUrl = "https://api4.thetvdb.com/v4";
        private readonly Func<string> apiKey;
        private readonly Func<RestRequest, IRestResponse> execute;
        private readonly Action<TimeSpan> delay;
        internal string Token { get; private set; }

        // Inject the transport for offline tests; credentials and tokens never enter diagnostics.
        internal TvdbClient(Func<string> apiKey, Func<RestRequest, IRestResponse> execute = null, Action<TimeSpan> delay = null)
        {
            this.apiKey = apiKey;
            this.delay = delay;
            var client = new RestClient(BaseUrl) { Timeout = 30000, FollowRedirects = false };
            this.execute = execute ?? (request => client.Execute(request));
        }

        internal string Login()
        {
            Token = null;
            var request = new RestRequest("login", Method.POST);
            request.AddHeader("Accept", "application/json");
            request.AddJsonBody(new { apikey = apiKey() });
            var json = ReadResponse(Send(request), "login", TvdbFailure.Authentication);
            string token = (string)(json["data"] as JObject)?["token"];
            if (string.IsNullOrWhiteSpace(token))
                throw new TvdbException(TvdbFailure.Authentication, "TVDB authentication failed: login returned no data.token.");
            Token = token;
            return Token;
        }

        private IRestResponse Send(RestRequest request)
        {
            for (int attempt = 0; ; attempt++)
            {
                IRestResponse response = null;
                try { response = execute(request); } catch (Exception) { }
                bool transient = response == null || response.ResponseStatus != ResponseStatus.Completed || (int)response.StatusCode == 0 || ApiRetry.Transient((int)response.StatusCode);
                if (!transient || attempt == 2)
                {
                    if (response != null) return response;
                    throw new TvdbException(TvdbFailure.Api, "TVDB network failure during " + request.Method + " " + request.Resource + ".");
                }
                TimeSpan? wait = null;
                var header = response?.Headers?.FirstOrDefault(h => string.Equals(h.Name, "Retry-After", StringComparison.OrdinalIgnoreCase));
                double seconds; DateTimeOffset until;
                if (header != null && double.TryParse(Convert.ToString(header.Value), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out seconds) && !double.IsNaN(seconds) && !double.IsInfinity(seconds))
                    wait = TimeSpan.FromSeconds(Math.Min(Math.Max(seconds, 0), 31));
                else if (header != null && DateTimeOffset.TryParse(Convert.ToString(header.Value), out until)) wait = until - DateTimeOffset.UtcNow;
                try { ApiRetry.Pause(attempt, wait, delay); }
                catch (InvalidOperationException) { throw new TvdbException(TvdbFailure.Api, "TVDB requested a long retry delay; retry later."); }
            }
        }

        private JObject Get(string resource, TvdbFailure notFound)
        {
            if (string.IsNullOrWhiteSpace(Token)) Login();
            for (int attempt = 0; attempt < 2; attempt++)
            {
                var request = new RestRequest(resource, Method.GET);
                request.AddHeader("Accept", "application/json");
                request.AddHeader("Authorization", "Bearer " + Token);
                var response = Send(request);
                if (response.StatusCode == HttpStatusCode.Unauthorized)
                {
                    Token = null;
                    if (attempt == 0) { Login(); continue; }
                }
                return ReadResponse(response, resource, notFound);
            }
            throw new TvdbException(TvdbFailure.Authentication, "TVDB authentication retry exhausted.");
        }

        private static JObject ReadResponse(IRestResponse response, string resource, TvdbFailure notFound)
        {
            int status = (int)response.StatusCode;
            if (response.ResponseStatus != ResponseStatus.Completed || status == 0)
                throw new TvdbException(TvdbFailure.Api, "TVDB network/API failure for " + resource + " (HTTP " + status + ", " + response.ResponseStatus + ").");
            if (status < 200 || status >= 300)
            {
                var failure = status == 401 || status == 403 ? TvdbFailure.Authentication
                    : status == 404 ? notFound : TvdbFailure.Api;
                throw new TvdbException(failure, "TVDB " + failure + " for " + resource + " (HTTP " + status + ").");
            }
            try
            {
                var json = JObject.Parse(response.Content ?? "");
                if (!string.Equals((string)json["status"], "success", StringComparison.OrdinalIgnoreCase))
                    throw new TvdbException(TvdbFailure.Api, "TVDB API reported failure for " + resource + " (HTTP " + status + ").");
                return json;
            }
            catch (JsonException)
            {
                throw new TvdbException(TvdbFailure.Api, "TVDB returned invalid JSON for " + resource + " (HTTP " + status + ").");
            }
        }

        internal string GetSeriesId(string slug)
        {
            if (string.IsNullOrWhiteSpace(slug)) throw new ArgumentException("A TVDB slug or ID is required.");
            slug = slug.Trim();
            int id;
            if (int.TryParse(slug, out id) && id > 0)
            {
                GetSeriesDetails(slug);
                return slug;
            }
            // Exact slug lookup first. TVDB's title search misses some shows by their full name
            // (e.g. "the secret lives of mormon wives orange county" returns nothing).
            try
            {
                var bySlug = Get("series/slug/" + Uri.EscapeDataString(slug), TvdbFailure.SeriesNotFound);
                string slugId = (string)(bySlug["data"] as JObject)?["id"];
                if (!string.IsNullOrWhiteSpace(slugId)) return Number(slugId, "series ID", false);
            }
            catch (TvdbException ex) when (ex.Failure == TvdbFailure.SeriesNotFound) { }

            // Fallback: search indexes titles, not slug filters. Require an exact slug match among results.
            var json = Get("search?type=series&query=" + Uri.EscapeDataString(slug.Replace('-', ' ')), TvdbFailure.SeriesNotFound);
            var results = json["data"] as JArray;
            if (results == null) throw InvalidShape("series search");
            var matches = results.OfType<JObject>().Where(r =>
                string.Equals((string)r["slug"], slug, StringComparison.OrdinalIgnoreCase)).ToList();
            if (matches.Count == 0)
                throw new TvdbException(TvdbFailure.SeriesNotFound, "TVDB series not found for slug '" + slug + "'; provide a TVDB ID if the slug has changed.");
            if (matches.Count != 1)
                throw new TvdbException(TvdbFailure.Api, "TVDB series search is ambiguous for slug '" + slug + "'; provide a TVDB ID.");
            return Number((string)matches[0]["tvdb_id"], "series ID", false);
        }

        internal JObject GetSeriesDetails(string id)
        {
            var json = Get("series/" + Number(id, "series ID", false) + "/extended", TvdbFailure.SeriesNotFound);
            var data = json["data"] as JObject;
            if (data == null) throw InvalidShape("series details");
            var seasons = data["seasons"] as JArray;
            if (seasons == null) throw InvalidShape("series seasons");
            // Official order matches the legacy airedSeason/episode semantics; exclude specials.
            int count = seasons.OfType<JObject>()
                .Where(s => string.Equals((string)(s["type"] as JObject)?["type"], "official", StringComparison.OrdinalIgnoreCase))
                .Select(s => (int?)s["number"] ?? 0).Where(n => n > 0).Distinct().Count();
            var ratings = data["contentRatings"] as JArray;
            var rating = ratings == null ? null : ratings.OfType<JObject>().FirstOrDefault(r =>
                string.Equals((string)r["country"], "usa", StringComparison.OrdinalIgnoreCase));
            return new JObject { ["data"] = new JObject {
                ["id"] = data["id"], ["seriesName"] = (string)data["name"] ?? "",
                ["firstAired"] = (string)data["firstAired"] ?? "", ["season"] = count,
                ["rating"] = (string)rating?["name"] ?? "",
                ["status"] = (string)(data["status"] as JObject)?["name"] ?? ""
            }};
        }

        internal JObject GetEpisodeDetails(string id, string season, string episode)
        {
            id = Number(id, "series ID", false);
            season = Number(season, "season", true);
            episode = Number(episode, "episode", false);
            string resource = "series/" + id + "/episodes/official?page=0&season=" + season + "&episodeNumber=" + episode;
            var json = Get(resource, TvdbFailure.EpisodeNotFound);
            var episodes = (json["data"] as JObject)?["episodes"] as JArray;
            if (episodes == null) throw InvalidShape("episode lookup");
            var matches = episodes.OfType<JObject>().Where(e => (string)e["seasonNumber"] == season && (string)e["number"] == episode).ToList();
            if (matches.Count == 0)
                throw new TvdbException(TvdbFailure.EpisodeNotFound, "TVDB episode not found: series " + id + " S" + season + "E" + episode + " (official order).");
            if (matches.Count != 1) throw InvalidShape("ambiguous episode lookup");
            var result = matches[0];
            return new JObject { ["data"] = new JArray(new JObject {
                ["id"] = result["id"], ["episodeName"] = (string)result["name"] ?? "",
                ["overview"] = (string)result["overview"] ?? ""
            }) };
        }

        internal IList<JObject> GetOfficialEpisodes(string id)
        {
            return GetOfficialEpisodes(id, null);
        }

        // language: a TVDB 3-letter code such as "eng". Without it TVDB returns names in the show's
        // original language (e.g. Japanese for anime), which can't match English recordings.
        internal IList<JObject> GetOfficialEpisodes(string id, string language)
        {
            id = Number(id, "series ID", false);
            string languagePath = "";
            if (!string.IsNullOrWhiteSpace(language))
            {
                if (!Regex.IsMatch(language, "^[a-z]{3}$")) throw new ArgumentException("TVDB language must be a 3-letter code like 'eng'.");
                languagePath = "/" + language;
            }
            var result = new List<JObject>();
            var pageBodies = new HashSet<string>();
            var episodeIds = new HashSet<int>();
            int? expectedTotal = null;
            for (int page = 0; page < 10000; page++)
            {
                var json = Get("series/" + id + "/episodes/official" + languagePath + "?page=" + page, TvdbFailure.SeriesNotFound);
                var episodes = (json["data"] as JObject)?["episodes"] as JArray;
                var links = json["links"] as JObject;
                if (episodes == null || links == null || links.Property("next") == null ||
                    episodes.Any(e => !(e is JObject))) throw InvalidInventory("complete episode inventory");
                // The language endpoint doesn't include the series block; every episode's seriesId is still checked below.
                if (languagePath.Length == 0 && (string)(json["data"] as JObject)?["series"]?["id"] != id) throw InvalidInventory("series identity");
                if (links["total_items"] != null && links["total_items"].Type != JTokenType.Null)
                {
                    int total;
                    if (!int.TryParse((string)links["total_items"], out total) || total < 0 || (expectedTotal.HasValue && expectedTotal != total))
                        throw InvalidInventory("changing episode total");
                    expectedTotal = total;
                }
                foreach (var episode in episodes.Cast<JObject>())
                {
                    int episodeId;
                    if (!int.TryParse((string)episode["id"], out episodeId) || episodeId <= 0 || !episodeIds.Add(episodeId) || (string)episode["seriesId"] != id)
                        throw InvalidInventory("invalid, repeated or mismatched episode identity");
                }
                if (episodes.Count > 0 && !pageBodies.Add(episodes.ToString(Formatting.None)))
                    throw InvalidInventory("repeated episode page");
                result.AddRange(episodes.Cast<JObject>());
                string next = (string)links["next"];
                if (string.IsNullOrWhiteSpace(next))
                {
                    if (result.Count == 0 || (expectedTotal.HasValue && expectedTotal != result.Count))
                        throw InvalidInventory("empty or incomplete episode inventory");
                    return result;
                }
                // Use only the page number; never forward the bearer token to a returned URL.
                var match = Regex.Match(next, @"[?&]page=(\d+)(?:&|$)");
                int nextPage;
                if (episodes.Count == 0 || !match.Success || !int.TryParse(match.Groups[1].Value, out nextPage) || nextPage != page + 1)
                    throw InvalidInventory("episode pagination");
            }
            throw InvalidInventory("episode pagination limit");
        }

        private static TvdbException InvalidInventory(string operation)
        { return new TvdbException(TvdbFailure.Inventory, "TVDB inventory cannot be verified: " + operation + "."); }

        private static string Number(string value, string name, bool allowZero)
        {
            int number;
            if (!int.TryParse(value, out number) || number < (allowZero ? 0 : 1))
                throw new ArgumentException("TVDB " + name + " must be " + (allowZero ? "nonnegative." : "positive."));
            return number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static TvdbException InvalidShape(string operation)
        {
            return new TvdbException(TvdbFailure.Api, "TVDB returned an unexpected response shape for " + operation + ".");
        }
    }
}
