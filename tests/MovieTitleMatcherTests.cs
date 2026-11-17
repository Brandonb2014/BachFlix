using System;
using System.Linq;
using BachFlixNfo.Features;

internal static class MovieTitleMatcherTests
{
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    private static void Matches(string input, string expected, params string[] sheet)
    {
        bool ignoredYear;
        var result = MovieTitleMatcher.FindCandidates(input, sheet, title => title, out ignoredYear);
        Check(result.Count == 1 && result[0] == expected, "Unexpected match for: " + input);
    }

    private static void NoMatch(string input, params string[] sheet)
    {
        bool ignoredYear;
        Check(MovieTitleMatcher.FindCandidates(input, sheet, title => title, out ignoredYear).Count == 0,
            "Unexpected automatic match for: " + input);
    }

    public static int Main()
    {
        try
        {
            Matches("Angels & Demons (2009)", "Angels and Demons (2009)", "Angels and Demons (2009)");
            Matches("Angels and Demons (2009)", "Angels & Demons (2009)", "Angels & Demons (2009)");
            Matches("The Avengers (2012)", "Avengers (2012)", "Avengers (2012)");
            Matches("Avengers (2012)", "The Avengers (2012)", "The Avengers (2012)");
            Matches("The Angels & Demons (2009)", "Angels and Demons (2009)", "Angels and Demons (2009)");
            Matches("Angels and Demons (2009)", "The Angels & Demons (2009)", "The Angels & Demons (2009)");
            Matches("  THE_Angels_&_Demons   (2009) ", "Angels and Demons (2009)", "Angels and Demons (2009)");
            Matches("The_Angels&Demons (2009)", "Angels and Demons (2009)", "Angels and Demons (2009)");
            Matches("Spider-Man (2002)", "SpiderMan (2002)", "SpiderMan (2002)");
            Matches("Léon: The Professional (1994)", "Léon The Professional (1994)", "Léon The Professional (1994)");

            // Exact spelling wins when both spellings are actual sheet entries.
            Matches("The Thing (1982)", "The Thing (1982)", "Thing (1982)", "The Thing (1982)");
            Matches("Thing (1982)", "Thing (1982)", "The Thing (1982)", "Thing (1982)");
            Matches("Angels & Demons (2009)", "Angels & Demons (2009)", "Angels and Demons (2009)", "Angels & Demons (2009)");
            Matches("Angels and Demons (2009)", "Angels and Demons (2009)", "Angels & Demons (2009)", "Angels and Demons (2009)");

            // A correct-year alias beats an exact title with a different year.
            bool ignoredYear;
            var correctYear = MovieTitleMatcher.FindCandidates("The Angels & Demons (2009)",
                new[] { "The Angels & Demons (2020)", "Angels and Demons (2009)" }, title => title, out ignoredYear);
            Check(correctYear.Single() == "Angels and Demons (2009)" && !ignoredYear, "Lost year preference");

            var sameYearAmbiguous = MovieTitleMatcher.FindCandidates("The Angels and Demons (2009)",
                new[] { "Angels & Demons (2009)", "Angels and Demons (2009)" }, title => title, out ignoredYear);
            Check(sameYearAmbiguous.Count == 2 && !ignoredYear, "Equally good aliases must require a choice");
            var remakes = MovieTitleMatcher.FindCandidates("The Angels & Demons",
                new[] { "Angels and Demons (2009)", "Angels and Demons (2020)" }, title => title, out ignoredYear);
            Check(remakes.Count == 2 && ignoredYear, "Remakes must require a choice");
            var uniqueTitle = MovieTitleMatcher.FindCandidates("The Angels & Demons (2010)",
                new[] { "Angels and Demons (2009)" }, title => title, out ignoredYear);
            Check(uniqueTitle.Single() == "Angels and Demons (2009)" && ignoredYear, "Existing title-only fallback lost");
            Matches("Angels & Demons (2009)", "Angels and Demons (2009)", "Angels and Demons (2009)", "Angels and Demons (2009)");

            // Only standalone "and" and leading "The " qualify; keep other words significant.
            NoMatch("Candy (2006)", "C&y (2006)");
            NoMatch("Theatre (2020)", "atre (2020)");
            NoMatch("Breathe (2017)", "Brea (2017)");
            NoMatch("Into the Woods (2014)", "Into Woods (2014)");
            NoMatch("Angels & Demons (2009)", "Angels Demons (2009)");
            NoMatch("Angels and Demons (2009)", "Angels Demons (2009)");
            NoMatch("A Movie (2020)", "Movie (2020)");
            NoMatch("The Unknown (2020)", "Another Movie (2020)");
            NoMatch("   ", "Angels and Demons (2009)");
            NoMatch(null, "Angels and Demons (2009)");
            NoMatch("The (2020)", "");
            NoMatch("Anything (2020)");

            Console.WriteLine("Movie title matching tests passed: aliases, leading The, exact/year preference, ambiguity, and word boundaries.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
    }
}
