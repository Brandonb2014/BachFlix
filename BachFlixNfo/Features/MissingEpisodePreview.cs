using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using TvdbApiCall;

namespace BachFlixNfo.Features
{
    internal sealed class MissingEpisodeSheetException : InvalidOperationException
    {
        internal MissingEpisodeSheetException(string message) : base(message) { }
    }

    internal sealed class EpisodePair : IEquatable<EpisodePair>, IComparable<EpisodePair>
    {
        internal int Season { get; private set; }
        internal int Episode { get; private set; }
        internal EpisodePair(int season, int episode) { Season = season; Episode = episode; }
        public bool Equals(EpisodePair other) { return other != null && Season == other.Season && Episode == other.Episode; }
        public override bool Equals(object other) { return Equals(other as EpisodePair); }
        public override int GetHashCode() { return unchecked(Season * 397 ^ Episode); }
        public int CompareTo(EpisodePair other) { int s = Season.CompareTo(other.Season); return s == 0 ? Episode.CompareTo(other.Episode) : s; }
        public override string ToString() { return "S" + Season.ToString("D2") + "E" + Episode.ToString("D2"); }
    }

    internal sealed class MissingEpisodeShow
    {
        internal int Row;
        internal int OutputColumn;
        internal string Title, TvdbId, ImdbId, TmdbId, Year, Continuing;
        internal object ExistingValue;
        internal bool OutputIsFormula;
        internal string OutputCell
        {
            get
            {
                return "DB!" + SheetHeaderMapping.ColumnLetter(OutputColumn) + Row;
            }
        }
    }

    internal enum MissingEpisodeStatus { Success, NoEmbyMatch, AmbiguousEmbyMatch, EmbyEpisodesUnavailable, TvdbUnavailable, TvdbAuthentication, MissingTvdbId, IdentityConflict, NumberingConflict, InvalidMetadata }

    internal sealed class MissingEpisodeResult
    {
        internal MissingEpisodeShow Show;
        internal MissingEpisodeStatus Status;
        internal string Message, EmbySeries, TvdbId;
        internal bool Abort;
        internal int AiredEpisodeCount, OwnedEpisodeCount, IgnoredEmbyItems;
        internal List<EpisodePair> MissingEpisodes = new List<EpisodePair>();
        internal List<EpisodePair> OwnedOutsideAired = new List<EpisodePair>();
        internal string MatchAttempt;
        internal List<string> MatchCandidates = new List<string>();
        internal List<string> Warnings = new List<string>();
        internal string SheetOutcome = "NotAttempted", SheetMessage;
        internal int MissingCount { get { return MissingEpisodes.Count; } }
        // Failure never has a proposed value, including a misleading zero.
        internal int? ProposedValue { get { return Status == MissingEpisodeStatus.Success ? (int?)MissingCount : null; } }
        internal bool NotInEmby { get { return Status == MissingEpisodeStatus.NoEmbyMatch; } }
        // The Sheet cell text: a compact episode list, "None" when complete, "Not in Emby" when the
        // series is absent from Emby, null on any other failure.
        internal string ProposedText
        {
            get
            {
                return Status == MissingEpisodeStatus.Success ? MissingEpisodeFormat.Describe(MissingEpisodes)
                    : NotInEmby ? MissingEpisodeFormat.NotInEmby : null;
            }
        }
    }

    internal static class MissingEpisodeFormat
    {
        internal const string NoneMissing = "None";
        internal const string NotInEmby = "Not in Emby";

        // e.g. "S01E04-08, 10 & 13, S02E03 & 04". Runs of three or more collapse to a range.
        internal static string Describe(IEnumerable<EpisodePair> episodes)
        {
            var seasons = (episodes ?? Enumerable.Empty<EpisodePair>()).Distinct().OrderBy(e => e)
                .GroupBy(e => e.Season).ToList();
            if (seasons.Count == 0) return NoneMissing;
            return string.Join(", ", seasons.Select(season =>
                "S" + season.Key.ToString("D2", CultureInfo.InvariantCulture) + "E" + JoinParts(Runs(season.Select(e => e.Episode).ToList()))));
        }

        private static List<string> Runs(List<int> numbers)
        {
            var parts = new List<string>();
            for (int i = 0; i < numbers.Count;)
            {
                int j = i;
                while (j + 1 < numbers.Count && numbers[j + 1] == numbers[j] + 1) j++;
                if (j - i >= 2) parts.Add(Pad(numbers[i]) + "-" + Pad(numbers[j]));
                else for (int k = i; k <= j; k++) parts.Add(Pad(numbers[k]));
                i = j + 1;
            }
            return parts;
        }

        private static string JoinParts(List<string> parts)
        {
            return parts.Count == 1 ? parts[0] : string.Join(", ", parts.Take(parts.Count - 1)) + " & " + parts[parts.Count - 1];
        }

        private static string Pad(int episode) { return episode.ToString("D2", CultureInfo.InvariantCulture); }
    }

    internal sealed class MissingEpisodePreview
    {
        private readonly IList<JObject> series;
        private readonly Func<string, IList<JObject>> embyEpisodes;
        private readonly Func<string, IList<JObject>> tvdbEpisodes;
        // Cache compact normalized data, never the large Emby MediaSources payloads.
        private readonly Dictionary<string, HashSet<EpisodePair>> ownedCache = new Dictionary<string, HashSet<EpisodePair>>();
        private readonly Dictionary<string, TvdbInventory> airedCache = new Dictionary<string, TvdbInventory>();
        internal MissingEpisodePreview(IList<JObject> series, Func<string, IList<JObject>> embyEpisodes, Func<string, IList<JObject>> tvdbEpisodes)
        { this.series = series; this.embyEpisodes = embyEpisodes; this.tvdbEpisodes = tvdbEpisodes; }

        internal static List<MissingEpisodeShow> ReadShows(IList<IList<object>> headers, IList<IList<object>> rows)
        {
            if (headers == null || headers.Count != 1 || headers[0].Count == 0)
                throw new MissingEpisodeSheetException("Sheet/header mapping failure: DB row 2 is unavailable.");
            var map = new Dictionary<string, int>();
            foreach (string name in new[] { "Missing Episodes", "TVDB ID", "IMDb ID", "IMDB ID", "TMDB ID", "Series Name", "Clean Name", "Clean Name with Year", "Year", "Continuing?" })
            {
                if (headers[0].Count(h => Convert.ToString(h) == name) > 1)
                    throw new MissingEpisodeSheetException("Sheet/header mapping failure: duplicate " + name + " header.");
                map[name] = -1;
            }
            SheetHeaderMapping.Update(map, headers);
            if (map["Missing Episodes"] < 0)
                throw new MissingEpisodeSheetException("Add a \"Missing Episodes\" header to an unused column in DB row 2.");
            if (map["TVDB ID"] < 0 || (map["Series Name"] < 0 && map["Clean Name"] < 0))
                throw new MissingEpisodeSheetException("Sheet/header mapping failure: DB needs TVDB ID and Series Name or Clean Name headers.");
            var result = new List<MissingEpisodeShow>();
            int rowNumber = 2;
            foreach (var row in rows ?? new List<IList<object>>())
            {
                rowNumber++;
                Func<string, string> cell = key => map[key] >= 0 && map[key] < row.Count ? Convert.ToString(row[map[key]]).Trim() : "";
                string title = cell("Series Name");
                if (title.Length == 0) title = cell("Clean Name");
                string id = cell("TVDB ID");
                if (title.Length == 0 && id.Length == 0) continue;
                string imdb = cell("IMDb ID");
                if (imdb.Length == 0) imdb = cell("IMDB ID");
                result.Add(new MissingEpisodeShow { Row = rowNumber, OutputColumn = map["Missing Episodes"], Title = title,
                    TvdbId = id, ImdbId = imdb, TmdbId = cell("TMDB ID"), Year = cell("Year"), Continuing = cell("Continuing?"),
                    ExistingValue = map["Missing Episodes"] < row.Count ? row[map["Missing Episodes"]] : null,
                    OutputIsFormula = cell("Missing Episodes").StartsWith("=", StringComparison.Ordinal) });
            }
            if (result.Where(s => !string.IsNullOrWhiteSpace(s.TvdbId)).GroupBy(s => Id(s.TvdbId, "Tvdb")).Any(g => g.Count() > 1))
                throw new MissingEpisodeSheetException("Duplicate TVDB IDs in DB; resolve row identities before running.");
            return result;
        }

        internal static int? Number(JToken value)
        {
            int n;
            return value != null && value.Type != JTokenType.Object && value.Type != JTokenType.Array &&
                int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? (int?)n : null;
        }

        internal static string Id(string value, string provider = "Tvdb")
        {
            value = (value ?? "").Trim();
            int n;
            return provider != "Imdb" && int.TryParse(value, out n) && n > 0 ? n.ToString(CultureInfo.InvariantCulture) : value;
        }

        internal static string Provider(JObject item, string name)
        {
            var providers = item["ProviderIds"] as JObject;
            return Id((string)providers?.GetValue(name, StringComparison.OrdinalIgnoreCase), name);
        }

        private static string TitleKey(string title)
        {
            title = Regex.Replace(title ?? "", @"\s*\(\d{4}\)\s*$", "").Normalize(NormalizationForm.FormD);
            return new string(title.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
        }

        internal MissingEpisodeResult Check(MissingEpisodeShow show, DateTime today)
        {
            var result = new MissingEpisodeResult { Show = show, Status = MissingEpisodeStatus.NoEmbyMatch };
            string provider = !string.IsNullOrWhiteSpace(show.TvdbId) ? "Tvdb" : !string.IsNullOrWhiteSpace(show.ImdbId) ? "Imdb" : !string.IsNullOrWhiteSpace(show.TmdbId) ? "Tmdb" : null;
            string id = provider == "Tvdb" ? show.TvdbId : provider == "Imdb" ? show.ImdbId : show.TmdbId;
            result.MatchAttempt = provider == null ? "Title/year: " + show.Title + " / " + show.Year : provider + " ID: " + Id(id, provider);
            var matches = provider != null
                ? series.Where(s => string.Equals(Provider(s, provider), Id(id, provider), StringComparison.OrdinalIgnoreCase)).ToList()
                : series.Where(s => string.Equals(((string)s["Name"] ?? "").Trim(), show.Title, StringComparison.OrdinalIgnoreCase) &&
                    (string.IsNullOrWhiteSpace(show.Year) || (string)s["ProductionYear"] == show.Year)).ToList();
            // The same provider ID in several Emby series means copies of one show (e.g. on two
            // drives); their episodes are combined. Title-only matches stay ambiguous.
            bool sameShowCopies = provider != null && matches.Count > 1;
            if (matches.Count != 1 && !sameShowCopies)
            {
                result.Status = matches.Count == 0 ? MissingEpisodeStatus.NoEmbyMatch : MissingEpisodeStatus.AmbiguousEmbyMatch;
                result.Message = matches.Count == 0 ? "No Emby series matches this Google Sheet row; marked Not in Emby." : "Ambiguous Emby series match.";
                // Alternatives are evidence for manual review, never automatic replacements for an ID match.
                var candidates = matches.Count > 0 ? matches : series.Where(s =>
                    string.Equals(((string)s["Name"] ?? "").Trim(), show.Title, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrWhiteSpace(show.ImdbId) && string.Equals(Provider(s, "Imdb"), Id(show.ImdbId, "Imdb"), StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrWhiteSpace(show.TmdbId) && string.Equals(Provider(s, "Tmdb"), Id(show.TmdbId, "Tmdb"), StringComparison.OrdinalIgnoreCase))).ToList();
                result.MatchCandidates = candidates.Select(s => "Emby " + (string)s["Id"] + ": " + (string)s["Name"] +
                    " (" + (string)s["ProductionYear"] + ") | TVDB=" + Provider(s, "Tvdb") +
                    " | IMDb=" + Provider(s, "Imdb") + " | TMDB=" + Provider(s, "Tmdb")).ToList();
                return result;
            }
            var match = matches[0];
            result.EmbySeries = (string)match["Name"];
            string tvdbFromEmby = Provider(match, "Tvdb");
            if (matches.Any(m => Provider(m, "Tvdb") != tvdbFromEmby))
            {
                result.Status = MissingEpisodeStatus.AmbiguousEmbyMatch;
                result.Message = "Emby copies of this show disagree on TVDB ID; review identity.";
                result.MatchCandidates = matches.Select(s => "Emby " + (string)s["Id"] + ": " + (string)s["Name"] + " | TVDB=" + Provider(s, "Tvdb")).ToList();
                return result;
            }
            if (matches.Count > 1) result.Warnings.Add("Show is in " + matches.Count + " Emby series (Emby " +
                string.Join(", ", matches.Select(m => (string)m["Id"])) + "); their episodes were combined.");
            result.TvdbId = Id(string.IsNullOrWhiteSpace(show.TvdbId) ? tvdbFromEmby : show.TvdbId, "Tvdb");
            int tvdb;
            if (!int.TryParse(result.TvdbId, out tvdb) || tvdb <= 0)
            {
                result.Status = MissingEpisodeStatus.MissingTvdbId;
                result.Message = "No usable TVDB ID in the Sheet row or matched Emby series.";
                return result;
            }
            foreach (var copy in matches)
            foreach (var alternate in new[] { new { Name = "Imdb", Value = show.ImdbId }, new { Name = "Tmdb", Value = show.TmdbId } })
                if (!string.IsNullOrWhiteSpace(alternate.Value) && !string.IsNullOrWhiteSpace(Provider(copy, alternate.Name)) &&
                    !string.Equals(Id(alternate.Value, alternate.Name), Provider(copy, alternate.Name), StringComparison.OrdinalIgnoreCase))
                {
                    result.Status = MissingEpisodeStatus.IdentityConflict;
                    result.Message = "TVDB match has a conflicting " + alternate.Name + " provider ID (Sheet " + Id(alternate.Value, alternate.Name) + ", Emby " + Provider(copy, alternate.Name) + "); review identity.";
                    return result;
                }
            bool titleDiffers = !string.IsNullOrWhiteSpace(show.Title) && TitleKey(show.Title) != TitleKey(result.EmbySeries);
            int sheetYear, embyYear;
            bool yearDiffers = int.TryParse(show.Year, out sheetYear) && int.TryParse((string)match["ProductionYear"], out embyYear) && Math.Abs((long)sheetYear - embyYear) > 1;
            if (titleDiffers) result.Warnings.Add("Title differs from Emby; matching provider ID retained (possible alias).");
            if (yearDiffers) result.Warnings.Add("Production year differs from Emby by more than one year.");
            if (titleDiffers && yearDiffers)
            {
                result.Status = MissingEpisodeStatus.IdentityConflict;
                result.Message = "Both title and year conflict despite the provider-ID match; review identity.";
                return result;
            }
            try
            {
                var owned = new HashSet<EpisodePair>();
                foreach (var copy in matches)
                {
                    string seriesId = (string)copy["Id"];
                    if (string.IsNullOrWhiteSpace(seriesId)) throw new EmbyException("Matched Emby series has no item ID.", false);
                    HashSet<EpisodePair> copyOwned;
                    if (!ownedCache.TryGetValue(seriesId, out copyOwned))
                    {
                        copyOwned = Owned(embyEpisodes(seriesId), out result.IgnoredEmbyItems);
                        if (result.IgnoredEmbyItems > 0)
                            throw new EmbyException("Emby inventory contains " + result.IgnoredEmbyItems + " items with uncertain ownership or invalid numbering; existing value preserved.", false);
                        ownedCache[seriesId] = copyOwned;
                    }
                    owned.UnionWith(copyOwned);
                }
                result.OwnedEpisodeCount = owned.Count;
                TvdbInventory tvdbInventory;
                if (!airedCache.TryGetValue(result.TvdbId, out tvdbInventory))
                    airedCache[result.TvdbId] = tvdbInventory = ValidateTvdbInventory(tvdbEpisodes(result.TvdbId), today);
                var inventory = tvdbInventory.Dates;
                if (tvdbInventory.UndatedAired > 0) result.Warnings.Add(tvdbInventory.UndatedAired + " TVDB episodes have no air date but come before aired episodes; counted as aired.");
                if (tvdbInventory.ConflictingDates > 0) result.Warnings.Add(tvdbInventory.ConflictingDates + " TVDB episode numbers repeat with different air dates; earliest date used.");
                var aired = new HashSet<EpisodePair>(inventory.Where(e => e.Value.Date <= today.Date).Select(e => e.Key));
                result.AiredEpisodeCount = aired.Count;
                result.MissingEpisodes = aired.Except(owned).OrderBy(e => e).ToList();
                result.OwnedOutsideAired = owned.Except(aired).OrderBy(e => e).ToList();
                if (owned.Except(inventory.Keys).Any())
                {
                    result.Status = MissingEpisodeStatus.NumberingConflict;
                    result.Message = "Owned episode numbers do not exist in the official TVDB inventory; review numbering/order.";
                    return result;
                }
                if (result.OwnedOutsideAired.Count > 0) result.Warnings.Add("Some owned episodes have future official air dates; excluded from missing tally.");
                result.Status = MissingEpisodeStatus.Success;
            }
            catch (EmbyException ex) { result.Status = MissingEpisodeStatus.EmbyEpisodesUnavailable; result.Message = ex.Message; result.Abort = ex.Systemic; }
            catch (TvdbException ex)
            {
                result.Status = ex.Failure == TvdbFailure.Authentication ? MissingEpisodeStatus.TvdbAuthentication : MissingEpisodeStatus.TvdbUnavailable;
                result.Message = ex.Message;
                result.Abort = ex.Failure == TvdbFailure.Authentication || ex.Failure == TvdbFailure.Api;
            }
            catch (Exception)
            {
                result.Status = MissingEpisodeStatus.InvalidMetadata;
                result.Message = "Unexpected or malformed episode metadata; existing value preserved.";
            }
            return result;
        }

        // Known excluded records are safe to ignore. Missing evidence is an error, not absence.
        private static bool IsOwned(JObject item, out bool uncertain)
        {
            uncertain = false;
            if (item == null) { uncertain = true; return false; }
            if (string.Equals((string)item["LocationType"], "Virtual", StringComparison.OrdinalIgnoreCase) ||
                string.Equals((string)item["LocationType"], "Remote", StringComparison.OrdinalIgnoreCase) ||
                (item["IsPlaceHolder"]?.Type == JTokenType.Boolean && (bool)item["IsPlaceHolder"]) ||
                (item["IsVirtualItem"]?.Type == JTokenType.Boolean && (bool)item["IsVirtualItem"])) return false;
            var sources = item["MediaSources"] as JArray;
            if (sources == null) { uncertain = true; return false; }
            foreach (var source in sources)
            {
                var s = source as JObject;
                if (s == null) { uncertain = true; continue; }
                string protocol = (string)s["Protocol"], type = (string)s["Type"], path = (string)s["Path"];
                if (string.Equals(type, "Placeholder", StringComparison.OrdinalIgnoreCase) ||
                    (path ?? "").EndsWith(".strm", StringComparison.OrdinalIgnoreCase) ||
                    (s["IsRemote"]?.Type == JTokenType.Boolean && (bool)s["IsRemote"]) ||
                    (!string.IsNullOrEmpty(protocol) && !string.Equals(protocol, "File", StringComparison.OrdinalIgnoreCase))) continue;
                long size;
                bool playable = new[] { "SupportsDirectPlay", "SupportsDirectStream", "SupportsTranscoding" }.Any(k => s[k]?.Type == JTokenType.Boolean && (bool)s[k]);
                if (string.Equals(protocol, "File", StringComparison.OrdinalIgnoreCase) && string.Equals(type, "Default", StringComparison.OrdinalIgnoreCase) &&
                    s["IsRemote"]?.Type == JTokenType.Boolean && !(bool)s["IsRemote"] && !string.IsNullOrWhiteSpace(path) &&
                    System.IO.Path.IsPathRooted(path) &&
                    !path.EndsWith(".strm", StringComparison.OrdinalIgnoreCase) && long.TryParse((string)s["Size"], out size) && size > 0 && playable)
                    return true;
                uncertain = true;
            }
            return false;
        }

        internal static HashSet<EpisodePair> Owned(IList<JObject> items, out int ignored)
        {
            var pairs = new HashSet<EpisodePair>(); ignored = 0;
            foreach (var item in items)
            {
                if (Number(item?["ParentIndexNumber"]) == 0) continue;
                bool uncertain;
                if (!IsOwned(item, out uncertain)) { if (uncertain) ignored++; continue; }
                int? season = Number(item["ParentIndexNumber"]), first = Number(item["IndexNumber"]);
                var endToken = item["IndexNumberEnd"];
                int? last = endToken == null || endToken.Type == JTokenType.Null ? first : Number(endToken);
                if (!season.HasValue || season < 0 || !first.HasValue || first <= 0 || !last.HasValue || last < first || last.Value - (long)first.Value > 10000)
                { ignored++; continue; }
                for (long episode = first.Value; episode <= last.Value; episode++) pairs.Add(new EpisodePair(season.Value, (int)episode));
            }
            return pairs;
        }

        internal sealed class TvdbInventory
        {
            internal readonly Dictionary<EpisodePair, DateTime> Dates = new Dictionary<EpisodePair, DateTime>();
            internal int UndatedAired, UndatedUpcoming, ConflictingDates;
        }

        // TVDB leaves some episodes undated: upcoming ones, and some old ones it never dated.
        // An undated episode that comes before an episode that has already aired is treated as
        // aired; one after the last aired episode is treated as not aired yet. When an episode
        // number repeats with different dates, the earliest date wins.
        internal static TvdbInventory ValidateTvdbInventory(IList<JObject> episodes, DateTime today)
        {
            if (episodes == null || episodes.Count == 0)
                throw new TvdbException(TvdbFailure.Inventory, "TVDB returned an empty inventory; cannot verify a count.");
            var inventory = new TvdbInventory();
            var undated = new HashSet<EpisodePair>();
            foreach (var episode in episodes)
            {
                int? season = Number(episode?["seasonNumber"]), number = Number(episode?["number"]);
                if (season == 0) continue;
                if (!(season > 0 && number > 0))
                    throw new TvdbException(TvdbFailure.Inventory, "TVDB has invalid regular-episode numbers; review inventory.");
                var pair = new EpisodePair(season.Value, number.Value);
                string aired = (string)episode?["aired"];
                DateTime date;
                if (string.IsNullOrWhiteSpace(aired)) { undated.Add(pair); continue; }
                if (!DateTime.TryParseExact(aired, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                    throw new TvdbException(TvdbFailure.Inventory, "TVDB has an unreadable air date; review inventory.");
                DateTime previous;
                if (inventory.Dates.TryGetValue(pair, out previous))
                {
                    if (previous != date) inventory.ConflictingDates++;
                    if (previous <= date) continue;
                }
                inventory.Dates[pair] = date;
            }
            if (inventory.Dates.Count == 0)
                throw new TvdbException(TvdbFailure.Inventory, "TVDB has no air dates for any regular episode; cannot verify a count.");
            var airedPairs = inventory.Dates.Where(e => e.Value.Date <= today.Date).Select(e => e.Key).ToList();
            EpisodePair lastAired = airedPairs.Count == 0 ? null : airedPairs.Max();
            foreach (var pair in undated.Where(p => !inventory.Dates.ContainsKey(p)))
            {
                bool before = lastAired != null && pair.CompareTo(lastAired) < 0;
                inventory.Dates[pair] = before ? DateTime.MinValue : DateTime.MaxValue;
                if (before) inventory.UndatedAired++; else inventory.UndatedUpcoming++;
            }
            return inventory;
        }

        internal static HashSet<EpisodePair> Aired(IList<JObject> episodes, DateTime today)
        {
            var pairs = new HashSet<EpisodePair>();
            foreach (var episode in episodes)
            {
                int? season = Number(episode["seasonNumber"]), number = Number(episode["number"]);
                DateTime aired;
                if (season > 0 && number > 0 && DateTime.TryParseExact((string)episode["aired"], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out aired) && aired.Date <= today.Date)
                    pairs.Add(new EpisodePair(season.Value, number.Value));
            }
            return pairs;
        }
    }
}
