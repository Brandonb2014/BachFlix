using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    internal static class MovieTitleMatcher
    {
        private static readonly Regex TitleYear = new Regex(
            @"^(?<title>.+?)\s*\((?<year>\d{4})\)\s*$", RegexOptions.Compiled);

        public static List<T> FindCandidates<T>(string fileTitle, IEnumerable<T> movies,
            Func<T, string> cleanTitle, out bool ignoredYear)
        {
            ignoredYear = false;
            if (string.IsNullOrWhiteSpace(fileTitle)) return new List<T>();

            var available = movies.Where(movie => !string.IsNullOrWhiteSpace(cleanTitle(movie)))
                .GroupBy(cleanTitle, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First()).ToList();

            // Prefer the original spelling, then try aliases with the year intact.
            // Only fall back to ignoring the year when neither full-title pass matches.
            foreach (bool titleOnly in new[] { false, true })
            {
                string candidate = titleOnly ? WithoutYear(fileTitle) : fileTitle;
                foreach (bool aliases in new[] { false, true })
                {
                    string key = Normalize(candidate, aliases);
                    if (key.Length == 0) continue;
                    var matches = available.Where(movie => Normalize(
                        titleOnly ? WithoutYear(cleanTitle(movie)) : cleanTitle(movie), aliases) == key)
                        .ToList();
                    if (matches.Count == 0) continue;
                    ignoredYear = titleOnly;
                    // Return all equally good matches so the caller can ask the user.
                    return matches;
                }
            }
            return new List<T>();
        }

        private static string WithoutYear(string title)
        {
            var match = TitleYear.Match(title.Trim());
            return match.Success ? match.Groups["title"].Value.Trim() : title.Trim();
        }

        private static string Normalize(string title, bool aliases)
        {
            title = title.Replace('_', ' ').ToLowerInvariant();
            if (aliases)
            {
                // Match the standalone word only: "Candy" must not become "C&y".
                title = Regex.Replace(title, @"\band\b|&", " and ");
            }
            // Keep & in the exact pass so "A & B" cannot silently match "A B".
            title = Regex.Replace(title, @"[^\w\s&]", "");
            title = Regex.Replace(title, @"\s+", " ").Trim();
            if (aliases) title = Regex.Replace(title, @"^the\s+", "");
            return title;
        }
    }
}
