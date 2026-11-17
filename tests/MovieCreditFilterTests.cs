using System;
using BachFlixNfo.Features;

internal static class MovieCreditFilterTests
{
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    public static int Main()
    {
        try
        {
            // Kept: normal features, and titles with nothing known.
            Check(MovieCreditFilter.GetSkipReason(112, "movie", "Comedy, Drama", 112) == null, "Feature should be kept");
            Check(MovieCreditFilter.GetSkipReason(0, null, null, null) == null, "Unknown title should be kept");
            Check(MovieCreditFilter.GetSkipReason(41, "movie", "Comedy", null) == null, "41 minutes is not a short");

            // Skipped: TV entries.
            Check(MovieCreditFilter.GetSkipReason(44, "episode", "Drama", 44) != null, "Episode should be skipped");
            Check(MovieCreditFilter.GetSkipReason(0, "series", "Drama", null) != null, "Series should be skipped");
            Check(MovieCreditFilter.GetSkipReason(90, "game", null, null) != null, "Non-movie type should be skipped");

            // Skipped: shorts by genre or runtime.
            Check(MovieCreditFilter.GetSkipReason(0, "movie", "Animation, Short, Comedy", null) != null, "Short genre should be skipped");
            Check(MovieCreditFilter.GetSkipReason(40, "movie", "Comedy", null) != null, "40 minutes is a short");
            Check(MovieCreditFilter.GetSkipReason(0, "movie", "Comedy", 12) != null, "IMDb runtime used when TMDB has none");
            Check(MovieCreditFilter.GetSkipReason(22, null, null, null) != null, "Short runtime alone should be skipped");
            Check(MovieCreditFilter.GetSkipReason(0, "movie", "Shorts Drama", null) == null, "Only the exact Short genre counts");

            // Runtime parsing.
            Check(MovieCreditFilter.ParseRuntimeMinutes("22 min") == 22, "Parse 22 min");
            Check(MovieCreditFilter.ParseRuntimeMinutes("N/A") == null, "N/A is unknown");
            Check(MovieCreditFilter.ParseRuntimeMinutes(null) == null, "null is unknown");

            Console.WriteLine("Movie credit filter tests passed.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
    }
}
