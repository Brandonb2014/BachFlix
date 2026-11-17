using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    // One episode as TVDB lists it (official order).
    internal sealed class TvdbEpisodeInfo
    {
        public int Season;
        public int Number;
        public string Title = "";
        public string Code { get { return "S" + Season.ToString("00") + "E" + Number.ToString("00"); } }
    }

    // What we could read out of a recorded file name.
    internal sealed class RecordedEpisodeName
    {
        public string FileName = "";
        public int? Season;
        public int? Episode;
        public int? LastEpisode;   // set for double episodes like S01E01-E02
        public string Title = "";
        public List<string> TitleParts = new List<string>();   // "Title A _ Title B" -> both titles (2+ only)
        public string Code
        {
            get
            {
                if (!Season.HasValue || !Episode.HasValue) return "(none)";
                string code = "S" + Season.Value.ToString("00") + "E" + Episode.Value.ToString("00");
                return LastEpisode.HasValue ? code + "-E" + LastEpisode.Value.ToString("00") : code;
            }
        }
    }

    internal enum EpisodeConfidence { Green, Yellow, Red }

    internal sealed class EpisodeMatchResult
    {
        public EpisodeConfidence Confidence;
        public List<TvdbEpisodeInfo> Episodes = new List<TvdbEpisodeInfo>();      // chosen episode(s)
        public List<TvdbEpisodeInfo> Alternatives = new List<TvdbEpisodeInfo>();  // other possibilities for review
        public string Reason = "";
        public bool NumberOnly;   // the title gave no support; the episode came from the SxxEyy number alone
        public string Code
        {
            get
            {
                if (Episodes.Count == 0) return "";
                var first = Episodes[0];
                return Episodes.Count == 1 ? first.Code : first.Code + "-E" + Episodes[Episodes.Count - 1].Number.ToString("00");
            }
        }
    }

    // Matches recorded episodes to TVDB. The episode TITLE is the source of truth;
    // the SxxEyy number in the file name only breaks ties or fills in when the title is missing or generic.
    internal static class TvEpisodeMatcher
    {
        private static readonly Regex SeasonEpisode = new Regex(
            @"(?<![A-Za-z0-9])[Ss](?<s>\d{1,2})[ ._-]?[Ee](?<e>\d{1,3})(?:-[Ee]?(?<e2>\d{1,3})|[ ._-]{0,3}[Ee](?<e2>\d{1,3}))?(?![0-9])",
            RegexOptions.Compiled);
        private static readonly Regex CrossFormat = new Regex(
            @"(?<![A-Za-z0-9])(?<s>\d{1,2})[xX](?<e>\d{2,3})(?![0-9])", RegexOptions.Compiled);
        private static readonly Regex QualityTail = new Regex(
            @"[\s.\[(_-]+(?:2160p|1080p|720p|480p|WEB[-. ]?DL|WEBRip|BluRay|[xXhH]\.?26[45]|HEVC)\b.*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex GenericTitle = new Regex(
            @"^(?:pilot|finale|season finale|series finale|premiere|season premiere|series premiere|untitled|tba|tbd|" +
            @"(?:episode|ep|part|chapter|show|program|programme) \d+|\d+)$",
            RegexOptions.Compiled);
        private static readonly string[] NumberWords = { "zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
            "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen", "twenty" };
        // StreamFab writes "Title A / Title B" as "Title A _ Title B" (the slash becomes an underscore).
        private static readonly Regex TitleSeparator = new Regex(@"\s+_\s+|\s+/\s+", RegexOptions.Compiled);
        private static readonly Regex DubPrefix = new Regex(@"^\((?:dub|dubbed|sub|subbed)\)\s*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex SeasonEpisodePrefix = new Regex(@"^season\s*\d+\s*,?\s*episode\s*\d+\b[\s,:-]*", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        internal const double CloseTitle = 0.90;     // near-identical spelling
        internal const double PossibleTitle = 0.75;  // worth suggesting, never trusted alone

        // ---------------------------------------------------------------
        // Reading file names
        // ---------------------------------------------------------------
        public static RecordedEpisodeName Parse(string fileNameWithoutExtension, string showName)
        {
            var result = new RecordedEpisodeName { FileName = fileNameWithoutExtension ?? "" };
            string name = (fileNameWithoutExtension ?? "").Trim();
            if (name.Length == 0) return result;

            var match = SeasonEpisode.Match(name);
            if (!match.Success) match = CrossFormat.Match(name);

            string titlePart = name;
            if (match.Success)
            {
                result.Season = int.Parse(match.Groups["s"].Value, CultureInfo.InvariantCulture);
                result.Episode = int.Parse(match.Groups["e"].Value, CultureInfo.InvariantCulture);
                if (match.Groups["e2"].Success)
                {
                    int last = int.Parse(match.Groups["e2"].Value, CultureInfo.InvariantCulture);
                    if (last > result.Episode.Value && last - result.Episode.Value <= 5) result.LastEpisode = last;
                }
                titlePart = name.Substring(match.Index + match.Length);
            }
            else if (!string.IsNullOrWhiteSpace(showName))
            {
                // No number at all: drop a leading show name so only the episode title is left.
                titlePart = StripLeadingShowName(name, showName);
            }

            titlePart = QualityTail.Replace(titlePart, "");
            // "Show.S01E02.The.Title" style: dots stand in for spaces.
            if (titlePart.IndexOf(' ') < 0 && titlePart.Count(c => c == '.') >= 2) titlePart = titlePart.Replace('.', ' ');

            var parts = TitleSeparator.Split(titlePart).Select(CleanTitle).Where(t => t.Length > 0).ToList();
            if (parts.Count >= 2) result.TitleParts = parts;

            result.Title = CleanTitle(titlePart);
            if (!string.IsNullOrWhiteSpace(showName) && NormalizeTitle(result.Title) == NormalizeTitle(showName))
                result.Title = "";
            return result;
        }

        private static string CleanTitle(string part)
        {
            string t = Regex.Replace((part ?? "").Replace('_', ' ').Trim(' ', '-', '.', '–', '—'), @"\s+", " ").Trim();
            // Junk StreamFab puts in front of the real title: "(Dub) Title", "Season 1, Episode 4: Title".
            t = DubPrefix.Replace(t, "");
            t = SeasonEpisodePrefix.Replace(t, "");
            return t.Trim(' ', '-', '.', '–', '—');
        }

        private static string StripLeadingShowName(string name, string showName)
        {
            string show = Regex.Replace(showName, @"\s*\(\d{4}\)\s*$", "").Trim();
            if (show.Length == 0) return name;
            var words = Regex.Split(show, @"[^\p{L}\p{Nd}]+").Where(w => w.Length > 0).Select(Regex.Escape);
            var prefix = new Regex(@"^\W*" + string.Join(@"[^\p{L}\p{Nd}]+", words) + @"(?:[^\p{L}\p{Nd}]*\(?\d{4}\)?)?[^\p{L}\p{Nd}]+",
                RegexOptions.IgnoreCase);
            var m = prefix.Match(name);
            return m.Success ? name.Substring(m.Length) : name;
        }

        // ---------------------------------------------------------------
        // Comparing titles
        // ---------------------------------------------------------------
        public static string NormalizeTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title)) return "";
            string t = RemoveDiacritics(title).ToLowerInvariant();
            t = t.Replace("&", " and ").Replace("’", "'");
            // Part markers: "(1)", "Pt. 1", "Part One" all become "part 1".
            t = Regex.Replace(t, @"\(\s*(\d+)\s*\)\s*$", " part $1");
            t = Regex.Replace(t, @"\bpt\.?\s*(\d+)\b", "part $1");
            // "Part One" / "Episode One" / "Chapter One" -> "... 1"
            for (int i = 0; i < NumberWords.Length; i++)
                t = Regex.Replace(t, @"\b(part|episode|chapter)\s+" + NumberWords[i] + @"\b", "$1 " + i);
            t = t.Replace("'", "");
            t = Regex.Replace(t, @"[^\p{L}\p{Nd}]+", " ");
            t = Regex.Replace(t, @"\s+", " ").Trim();
            t = Regex.Replace(t, @"^the ", "");
            return t;
        }

        public static bool IsGeneric(string title)
        {
            return GenericTitle.IsMatch(NormalizeTitle(title));
        }

        public static double Similarity(string a, string b)
        {
            a = NormalizeTitle(a); b = NormalizeTitle(b);
            if (a.Length == 0 || b.Length == 0) return 0;
            if (a == b) return 1;
            int[] prev = new int[b.Length + 1], cur = new int[b.Length + 1];
            for (int j = 0; j <= b.Length; j++) prev[j] = j;
            for (int i = 1; i <= a.Length; i++)
            {
                cur[0] = i;
                for (int j = 1; j <= b.Length; j++)
                    cur[j] = Math.Min(Math.Min(cur[j - 1] + 1, prev[j] + 1), prev[j - 1] + (a[i - 1] == b[j - 1] ? 0 : 1));
                var swap = prev; prev = cur; cur = swap;
            }
            return 1.0 - (double)prev[b.Length] / Math.Max(a.Length, b.Length);
        }

        private static string RemoveDiacritics(string s)
        {
            var sb = new StringBuilder();
            foreach (char c in s.Normalize(NormalizationForm.FormD))
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark) sb.Append(c);
            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        // ---------------------------------------------------------------
        // Matching
        // ---------------------------------------------------------------
        public static EpisodeMatchResult Match(RecordedEpisodeName file, IList<TvdbEpisodeInfo> tvdb)
        {
            tvdb = (tvdb ?? new List<TvdbEpisodeInfo>()).Where(e => e != null).ToList();
            if (tvdb.Count == 0) return Red("TVDB returned no episodes for this show");

            if (file.LastEpisode.HasValue) return MatchMulti(file, tvdb);
            if (file.TitleParts.Count >= 2)
            {
                // One number but two titles ("Gloria Green _ Daddy Ditty"): a double episode.
                var byTitles = MatchTitleParts(file, file.TitleParts, tvdb, null);
                if (byTitles != null) return byTitles;
            }

            TvdbEpisodeInfo byNumber = AtNumber(tvdb, file.Season, file.Episode);
            string key = NormalizeTitle(file.Title);
            bool generic = key.Length > 0 && IsGeneric(file.Title);

            if (key.Length > 0)
            {
                var exact = tvdb.Where(e => NormalizeTitle(e.Title) == key).ToList();

                if (exact.Count == 1)
                {
                    var ep = exact[0];
                    if (ep == byNumber) return Green(ep, generic ? "generic title, but number agrees" : "title and number agree");
                    if (generic)
                        return Yellow(ep, "generic title '" + file.Title + "' points to " + ep.Code + " but file says " + file.Code, byNumber);
                    if (MovesIntoSpecials(file, new[] { ep }))
                        return Yellow(ep, "title matches special " + ep.Code + " but file says " + file.Code + "; check before moving into Specials", byNumber);
                    if (!file.Episode.HasValue) return Green(ep, "title match (no number in file name)");
                    return Green(ep, "title match; number corrected from " + file.Code);
                }

                if (exact.Count > 1)
                {
                    if (byNumber != null && exact.Contains(byNumber))
                        return Green(byNumber, "title shared by " + exact.Count + " episodes; file number picked " + byNumber.Code);
                    var sameSeason = exact.Where(e => file.Season.HasValue && e.Season == file.Season.Value).ToList();
                    if (sameSeason.Count == 1 && !generic)
                        return Yellow(sameSeason[0], "title shared by " + exact.Count + " episodes; only one is in season " + file.Season, exact.Where(e => e != sameSeason[0]));
                    return Yellow(null, "title matches " + exact.Count + " episodes; number doesn't pick one", exact);
                }

                if (!generic)
                {
                    var ranked = tvdb.Select(e => new { Ep = e, Score = Similarity(file.Title, e.Title) })
                        .OrderByDescending(x => x.Score).ToList();
                    var best = ranked[0];
                    double second = ranked.Count > 1 ? ranked[1].Score : 0;
                    if (best.Score >= CloseTitle && second < PossibleTitle && best.Ep == byNumber)
                        return Green(best.Ep, "near-exact title (" + Percent(best.Score) + ") and number agree");
                    if (best.Score >= PossibleTitle)
                        return Yellow(best.Ep, "closest title is " + Percent(best.Score) + " similar: '" + best.Ep.Title + "'",
                            ranked.Skip(1).Where(x => x.Score >= PossibleTitle).Select(x => x.Ep).Concat(new[] { byNumber }));
                }

                if (byNumber != null)
                    return NumberGuess(Yellow(byNumber, "title '" + file.Title + "' matched nothing; number says " + byNumber.Code + " '" + byNumber.Title + "'"));
                return Red("title '" + file.Title + "' matched nothing and " + file.Code + " isn't on TVDB");
            }

            if (byNumber != null)
                return NumberGuess(Yellow(byNumber, "no title in file name; number only (" + byNumber.Code + " '" + byNumber.Title + "')"));
            return Red("no title and no usable SxxEyy in file name");
        }

        private static EpisodeMatchResult MatchMulti(RecordedEpisodeName file, IList<TvdbEpisodeInfo> tvdb)
        {
            int count = file.LastEpisode.Value - file.Episode.Value + 1;
            var byNumber = Enumerable.Range(file.Episode.Value, count).Select(n => AtNumber(tvdb, file.Season, n)).ToList();

            // Titles of double episodes are usually joined with " / ", " _ ", " + " or "; ".
            var parts = file.TitleParts.Count >= 2 ? file.TitleParts
                : Regex.Split(file.Title ?? "", @"\s*(?:\s/\s|/|\s\+\s|;)\s*").Where(p => p.Trim().Length > 0).ToList();
            if (parts.Count == count)
            {
                var byTitles = MatchTitleParts(file, parts, tvdb, byNumber);
                if (byTitles != null) return byTitles;
            }

            if (byNumber.All(e => e != null))
            {
                var result = new EpisodeMatchResult { Confidence = EpisodeConfidence.Yellow, Episodes = byNumber, NumberOnly = true };
                result.Reason = "double episode: titles couldn't be confirmed; number says " + byNumber[0].Code + "-E" + byNumber[count - 1].Number.ToString("00");
                return result;
            }
            return Red("double episode " + file.Code + " couldn't be matched by title or number");
        }

        // Every title part must match exactly one TVDB episode. Returns null when they don't,
        // so the caller can fall back to other evidence.
        private static EpisodeMatchResult MatchTitleParts(RecordedEpisodeName file, IList<string> parts,
            IList<TvdbEpisodeInfo> tvdb, IList<TvdbEpisodeInfo> byNumber)
        {
            if (parts.Count < 2 || parts.Count > 6) return null;
            var found = parts.Select(p => tvdb.Where(e => NormalizeTitle(e.Title) == NormalizeTitle(p)).ToList()).ToList();
            if (!found.All(f => f.Count == 1)) return null;

            var eps = found.Select(f => f[0]).ToList();
            bool consecutive = eps.All(e => e.Season == eps[0].Season) &&
                               eps.Select((e, i) => e.Number - i).Distinct().Count() == 1;
            if (byNumber == null)
                byNumber = Enumerable.Range(0, eps.Count)
                    .Select(i => file.Episode.HasValue ? AtNumber(tvdb, file.Season, file.Episode.Value + i) : null).ToList();

            var result = new EpisodeMatchResult { Episodes = eps };
            if (!consecutive)
            {
                result.Confidence = EpisodeConfidence.Yellow;
                result.Reason = "double episode titles match episodes that aren't back to back";
            }
            else if (MovesIntoSpecials(file, eps))
            {
                result.Confidence = EpisodeConfidence.Yellow;
                result.Reason = "double episode titles match specials but file says " + file.Code + "; check before moving into Specials";
            }
            else
            {
                result.Confidence = EpisodeConfidence.Green;
                result.Reason = eps.SequenceEqual(byNumber) ? "double episode: titles and numbers agree"
                    : "double episode: titles match; numbers corrected from " + file.Code;
            }
            return result;
        }

        // A regular-season file being moved into Specials (S00) is a big jump: always ask.
        private static bool MovesIntoSpecials(RecordedEpisodeName file, IEnumerable<TvdbEpisodeInfo> eps)
        {
            return file.Season.HasValue && file.Season.Value != 0 && eps.Any(e => e.Season == 0);
        }

        // StreamFab sometimes only knows "Episode 1", "Episode 2"... for a whole season. Trust the numbers when
        // every file in that season is named that way, each "Episode N" matches its own SxxEyy number, and the
        // files cover exactly the episodes TVDB lists for the season (10 files for TVDB's 10 episodes, no gaps).
        // Seasons with any real title, a missing or extra episode, or a double episode are left for review.
        private static readonly Regex EpisodeNumberTitle = new Regex(@"^episode (?<n>\d+)$", RegexOptions.Compiled);

        public static void AcceptNumberedSeasons(IList<RecordedEpisodeName> files, IList<EpisodeMatchResult> results, IList<TvdbEpisodeInfo> tvdb)
        {
            if (files == null || results == null || files.Count != results.Count || tvdb == null) return;
            var seasons = Enumerable.Range(0, files.Count).Where(i => files[i].Season.HasValue)
                .GroupBy(i => files[i].Season.Value);
            foreach (var season in seasons)
            {
                var idx = season.ToList();
                bool allNumbered = idx.All(i =>
                {
                    var m = EpisodeNumberTitle.Match(NormalizeTitle(files[i].Title));
                    return m.Success && !files[i].LastEpisode.HasValue && files[i].Episode.HasValue
                        && int.Parse(m.Groups["n"].Value, CultureInfo.InvariantCulture) == files[i].Episode.Value
                        && results[i].NumberOnly && results[i].Episodes.Count == 1
                        && results[i].Episodes[0].Season == season.Key && results[i].Episodes[0].Number == files[i].Episode.Value;
                });
                if (!allNumbered) continue;

                var tvdbNumbers = tvdb.Where(e => e.Season == season.Key).Select(e => e.Number).OrderBy(n => n).ToList();
                var fileNumbers = idx.Select(i => files[i].Episode.Value).OrderBy(n => n).ToList();
                if (tvdbNumbers.Count == 0 || !tvdbNumbers.SequenceEqual(fileNumbers)) continue;

                foreach (int i in idx)
                {
                    results[i].Confidence = EpisodeConfidence.Green;
                    results[i].Reason = "StreamFab only gave 'Episode N'; all " + fileNumbers.Count + " episodes of season " + season.Key +
                        " line up with TVDB, so number says " + results[i].Code + " '" + results[i].Episodes[0].Title + "'";
                }
            }
        }

        // Two files must never land on the same episode: downgrade every file involved.
        // Exception: when exactly one of them is backed by its title and the others only by their
        // (often wrong) SxxEyy number, the title-backed file keeps its result. The number guesses
        // are already yellow and get flagged.
        public static void FlagCollisions(IList<EpisodeMatchResult> results)
        {
            var groups = results.Where(r => r.Confidence != EpisodeConfidence.Red)
                .SelectMany(r => r.Episodes.Select(e => new { r, e }))
                .GroupBy(x => x.e.Code).Where(g => g.Select(x => x.r).Distinct().Count() > 1).ToList();
            // Decide everything first, then change results, so one group's downgrade can't affect another's decision.
            var downgrade = new List<KeyValuePair<string, EpisodeMatchResult>>();
            foreach (var g in groups)
            {
                var involved = g.Select(x => x.r).Distinct().ToList();
                var titleBacked = involved.Where(r => !r.NumberOnly).ToList();
                foreach (var r in involved)
                    if (!(titleBacked.Count == 1 && r == titleBacked[0]))
                        downgrade.Add(new KeyValuePair<string, EpisodeMatchResult>(g.Key, r));
            }
            foreach (var d in downgrade)
            {
                var r = d.Value;
                r.Confidence = EpisodeConfidence.Yellow;
                if (r.Reason.IndexOf("another file", StringComparison.Ordinal) < 0)
                    r.Reason = "another file also matched " + d.Key + "; " + r.Reason;
            }
        }

        // ---------------------------------------------------------------
        // Building the FileBot-style name: "Show (2022) - S02E06 - Title"
        // ---------------------------------------------------------------
        public static string BuildFileName(string showName, EpisodeMatchResult match)
        {
            if (match == null || match.Episodes.Count == 0) return "";
            string titles = string.Join(" & ", match.Episodes.Select(e => e.Title).Where(t => !string.IsNullOrWhiteSpace(t)));
            string name = showName + " - " + match.Code + (titles.Length > 0 ? " - " + titles : "");
            return SafeFileName(name);
        }

        public static string SeasonFolder(EpisodeMatchResult match)
        {
            return match == null || match.Episodes.Count == 0 ? "" : "Season " + match.Episodes[0].Season.ToString("00");
        }

        public static string SafeFileName(string name)
        {
            name = Regex.Replace(name, @"\s*:\s*", " - ");
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c.ToString(), "");
            return Regex.Replace(name, @"\s+", " ").Trim().TrimEnd('.');
        }

        // ---------------------------------------------------------------
        private static TvdbEpisodeInfo AtNumber(IList<TvdbEpisodeInfo> tvdb, int? season, int? episode)
        {
            if (!season.HasValue || !episode.HasValue) return null;
            var hits = tvdb.Where(e => e.Season == season.Value && e.Number == episode.Value).ToList();
            return hits.Count == 1 ? hits[0] : null;
        }

        private static string Percent(double score) { return ((int)Math.Round(score * 100)).ToString(CultureInfo.InvariantCulture) + "%"; }

        private static EpisodeMatchResult Green(TvdbEpisodeInfo ep, string reason)
        {
            return new EpisodeMatchResult { Confidence = EpisodeConfidence.Green, Episodes = { ep }, Reason = reason };
        }

        private static EpisodeMatchResult Yellow(TvdbEpisodeInfo ep, string reason, IEnumerable<TvdbEpisodeInfo> alternatives)
        {
            var r = new EpisodeMatchResult { Confidence = EpisodeConfidence.Yellow, Reason = reason };
            if (ep != null) r.Episodes.Add(ep);
            if (alternatives != null) r.Alternatives.AddRange(alternatives.Where(a => a != null && a != ep).Distinct());
            return r;
        }

        private static EpisodeMatchResult Yellow(TvdbEpisodeInfo ep, string reason, TvdbEpisodeInfo alternative = null)
        {
            return Yellow(ep, reason, new[] { alternative });
        }

        private static EpisodeMatchResult NumberGuess(EpisodeMatchResult r)
        {
            r.NumberOnly = true;
            return r;
        }

        private static EpisodeMatchResult Red(string reason)
        {
            return new EpisodeMatchResult { Confidence = EpisodeConfidence.Red, Reason = reason };
        }
    }
}
