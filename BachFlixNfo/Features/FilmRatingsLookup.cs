using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Looks up the MPA rating reason (e.g. "Rated R for sci-fi violence and brief language.")
    /// from FilmRatings.com, the MPA's official rating search.
    /// </summary>
    internal static class FilmRatingsLookup
    {
        internal sealed class RatedFilm
        {
            public string Title { get; set; }
            public string Reason { get; set; }
            public int YearRated { get; set; }
        }

        private static readonly Regex ItemRegex = new Regex(
            "item-title\">(?<title>[^<]*)</div>(?<body>.*?)<!--end-item-->",
            RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex ReasonRegex = new Regex(
            "Reason:\\s*</span>\\s*<span class=\"text\">(?<v>[^<]*)</span>",
            RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex YearRegex = new Regex(
            "Year Rated:\\s*</span>\\s*<span class=\"text\">(?<v>\\d{4})</span>",
            RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex TitleYearSuffix = new Regex("\\s*\\((\\d{4})\\)\\s*$", RegexOptions.Compiled);

        private static DateTime lastRequest = DateTime.MinValue;

        // Results for this run, including misses, so retry loops don't search the same title again.
        private static readonly Dictionary<string, string> Cache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Gets the rating reason for a sheet title like "The Matrix (1999)". Returns null when no single match is found.
        /// </summary>
        internal static string GetReasonForSheetTitle(string sheetTitle)
        {
            if (string.IsNullOrWhiteSpace(sheetTitle)) return null;
            Match m = TitleYearSuffix.Match(sheetTitle);
            if (!m.Success) return null;

            string key = sheetTitle.Trim();
            if (Cache.TryGetValue(key, out string cached)) return cached;

            string reason = GetReason(sheetTitle.Substring(0, m.Index), int.Parse(m.Groups[1].Value));
            Cache[key] = reason;
            return reason;
        }

        /// <summary>
        /// Gets the rating reason for a title and release year. Films are often rated the year before release,
        /// so it checks the release year, then the year before, then the year after.
        /// </summary>
        internal static string GetReason(string title, int releaseYear)
        {
            if (string.IsNullOrWhiteSpace(title) || releaseYear < 1900) return null;

            foreach (int yearRated in new[] { releaseYear, releaseYear - 1, releaseYear + 1 })
            {
                string html = Search(title, yearRated);
                if (html == null) return null; // Site unreachable: don't keep hammering it.

                string reason = PickReason(ParseResults(html), title);
                if (reason != null) return reason;
            }

            return null;
        }

        internal static List<RatedFilm> ParseResults(string html)
        {
            var films = new List<RatedFilm>();
            foreach (Match item in ItemRegex.Matches(html))
            {
                string body = item.Groups["body"].Value;
                Match reason = ReasonRegex.Match(body);
                Match year = YearRegex.Match(body);
                films.Add(new RatedFilm
                {
                    Title = WebUtility.HtmlDecode(item.Groups["title"].Value).Trim(),
                    Reason = reason.Success ? WebUtility.HtmlDecode(reason.Groups["v"].Value).Trim() : null,
                    YearRated = year.Success ? int.Parse(year.Groups["v"].Value) : 0
                });
            }
            return films;
        }

        /// <summary>
        /// Returns the reason only when exactly one distinct reason belongs to a film with the same title.
        /// </summary>
        internal static string PickReason(IEnumerable<RatedFilm> films, string title)
        {
            string wanted = NormalizeTitle(title);
            var reasons = films
                .Where(f => NormalizeTitle(f.Title) == wanted && !string.IsNullOrWhiteSpace(f.Reason) && f.Reason.StartsWith("Rated ", StringComparison.OrdinalIgnoreCase))
                .Select(f => f.Reason)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            return reasons.Count == 1 ? reasons[0] : null;
        }

        /// <summary>
        /// Lowercases, moves a trailing ", The"/", A"/", An" to the front, and drops punctuation,
        /// so "Matrix, The" and "The Matrix" match.
        /// </summary>
        internal static string NormalizeTitle(string title)
        {
            string t = (title ?? "").Trim().ToLowerInvariant().Replace("&", " and ");
            Match article = Regex.Match(t, ",\\s*(the|a|an)$");
            if (article.Success) t = article.Groups[1].Value + " " + t.Substring(0, article.Index);
            t = Regex.Replace(t, "[^a-z0-9 ]", " ");
            return Regex.Replace(t, "\\s+", " ").Trim();
        }

        private static string Search(string title, int yearRated)
        {
            // Be polite: at most one request every 700 ms.
            TimeSpan sinceLast = DateTime.UtcNow - lastRequest;
            if (sinceLast < TimeSpan.FromMilliseconds(700)) Thread.Sleep(TimeSpan.FromMilliseconds(700) - sinceLast);
            lastRequest = DateTime.UtcNow;

            try
            {
                var client = new RestClient("https://www.filmratings.com/search-results/") { UserAgent = "BachFlixNfo/1.0" };
                var request = new RestRequest(Method.GET);
                request.AddQueryParameter("search", title.Trim());
                request.AddQueryParameter("my", yearRated.ToString());
                IRestResponse response = client.Execute(request);
                return response.IsSuccessful ? response.Content : null;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
