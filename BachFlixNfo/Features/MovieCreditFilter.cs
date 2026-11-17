using System;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Decides whether a TMDB movie credit is really a short or a TV entry, so the actor
    /// credits import doesn't add it to the Movies sheet.
    /// </summary>
    internal static class MovieCreditFilter
    {
        /// <summary>The Academy's definition of a short: 40 minutes or less.</summary>
        internal const int MaxShortRuntimeMinutes = 40;

        /// <summary>
        /// Returns a plain-English reason to skip the title, or null to keep it.
        /// </summary>
        /// <param name="tmdbRuntime">TMDB runtime in minutes; 0 or less means unknown.</param>
        /// <param name="imdbType">IMDb type from OMDb ("movie", "series", "episode", ...); null if unknown.</param>
        /// <param name="imdbGenre">IMDb genre list from OMDb, e.g. "Animation, Short, Comedy"; null if unknown.</param>
        /// <param name="imdbRuntime">IMDb runtime in minutes; null if unknown.</param>
        internal static string GetSkipReason(int tmdbRuntime, string imdbType, string imdbGenre, int? imdbRuntime)
        {
            if (!string.IsNullOrWhiteSpace(imdbType))
            {
                string type = imdbType.Trim().ToLowerInvariant();
                if (type == "episode")
                    return "IMDb lists it as a TV episode";
                if (type == "series")
                    return "IMDb lists it as a TV series";
                if (type != "movie")
                    return "IMDb lists it as a " + type + ", not a movie";
            }

            if (HasGenre(imdbGenre, "Short"))
                return "IMDb lists it as a short";

            int runtime = tmdbRuntime > 0 ? tmdbRuntime : (imdbRuntime ?? 0);
            if (runtime > 0 && runtime <= MaxShortRuntimeMinutes)
                return "the runtime is only " + runtime + " minutes";

            return null;
        }

        /// <summary>Parses OMDb runtimes like "22 min"; returns null for "N/A" or anything unreadable.</summary>
        internal static int? ParseRuntimeMinutes(string runtime)
        {
            if (string.IsNullOrWhiteSpace(runtime)) return null;
            string digits = runtime.Trim().Split(' ')[0];
            int minutes;
            return int.TryParse(digits, out minutes) && minutes > 0 ? minutes : (int?)null;
        }

        private static bool HasGenre(string genres, string genre)
        {
            if (string.IsNullOrWhiteSpace(genres)) return false;
            foreach (string g in genres.Split(','))
            {
                if (string.Equals(g.Trim(), genre, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
    }
}
