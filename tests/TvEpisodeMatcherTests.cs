using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using BachFlixNfo.Features;
using Newtonsoft.Json.Linq;

internal static class TvEpisodeMatcherTests
{
    private static int failures;

    private static void Check(bool condition, string message)
    {
        if (!condition) { failures++; Console.WriteLine("FAIL: " + message); }
    }

    private static TvdbEpisodeInfo Ep(int s, int e, string title) { return new TvdbEpisodeInfo { Season = s, Number = e, Title = title }; }

    private static readonly List<TvdbEpisodeInfo> Show = new List<TvdbEpisodeInfo>
    {
        Ep(1, 1, "Pilot"),
        Ep(1, 2, "The Long Night"),
        Ep(1, 3, "Angels & Demons"),
        Ep(1, 4, "Halloween"),
        Ep(1, 5, "Into the Woods (1)"),
        Ep(1, 6, "Into the Woods (2)"),
        Ep(2, 1, "Halloween"),
        Ep(2, 2, "Smallfolk"),
        Ep(2, 3, "Léon's Return"),
        Ep(2, 4, "Fire: The Sequel"),
        Ep(0, 1, "Behind the Scenes"),
        Ep(3, 1, "Episode 1"),
    };

    private static EpisodeMatchResult M(string fileName, string show = "Show (2020)")
    {
        return TvEpisodeMatcher.Match(TvEpisodeMatcher.Parse(fileName, show), Show);
    }

    public static int Main()
    {
        try
        {
            // --- Whole seasons named "Episode N" are accepted when they line up with TVDB
            {
                var lastShip = new List<TvdbEpisodeInfo> { Ep(1, 1, "Phase Six"), Ep(1, 2, "Welcome to Gitmo"), Ep(1, 3, "Dead Reckoning"),
                    Ep(2, 1, "Unreal City"), Ep(2, 2, "Fight the Ship"), Ep(2, 3, "It's Not a Rumor"), Ep(3, 1, "Sea Legs"), Ep(3, 2, "Flags"),
                    Ep(4, 1, "Alone"), Ep(4, 2, "Together") };
                string[] names = {
                    "The Last Ship(2014) - S01E01 - Episode 1", "The Last Ship(2014) - S01E02 - Episode 2", "The Last Ship(2014) - S01E03 - Episode 3",  // full season
                    "The Last Ship(2015) - S02E01 - Episode 1", "The Last Ship(2015) - S02E02 - Episode 2",                                             // S02E03 missing
                    "The Last Ship(2016) - S03E01 - Episode 1", "The Last Ship(2016) - S03E02 - Flags",                                                 // one real title
                    "The Last Ship(2017) - S04E01 - Episode 1", "The Last Ship(2017) - S04E02 - Episode 3" };                                           // title != number
                var files = names.Select(n => TvEpisodeMatcher.Parse(n, "The Last Ship (2014)")).ToList();
                var results = files.Select(f => TvEpisodeMatcher.Match(f, lastShip)).ToList();
                Check(results.All(x => x.Confidence == EpisodeConfidence.Yellow || x.Confidence == EpisodeConfidence.Green), "numbered season: setup");
                TvEpisodeMatcher.AcceptNumberedSeasons(files, results, lastShip);
                Check(results.Take(3).All(x => x.Confidence == EpisodeConfidence.Green), "full 'Episode N' season lines up with TVDB: green");
                Check(results[0].Code == "S01E01" && TvEpisodeMatcher.BuildFileName("The Last Ship (2014)", results[0]) == "The Last Ship (2014) - S01E01 - Phase Six", "numbered season gets TVDB titles");
                Check(results[3].Confidence == EpisodeConfidence.Yellow && results[4].Confidence == EpisodeConfidence.Yellow, "season missing an episode stays yellow");
                Check(results[5].Confidence == EpisodeConfidence.Yellow, "season with a real title: 'Episode 1' file stays yellow");
                Check(results[7].Confidence == EpisodeConfidence.Yellow && results[8].Confidence == EpisodeConfidence.Yellow, "'Episode 3' as S04E02 doesn't line up: stays yellow");
            }

            // --- Parsing StreamFab-style names
            var p =TvEpisodeMatcher.Parse("Show S01E02 The Long Night", "Show (2020)");
            Check(p.Season == 1 && p.Episode == 2 && p.Title == "The Long Night", "space style parse");
            p = TvEpisodeMatcher.Parse("Show.S01E02.The.Long.Night.1080p.WEB-DL", "Show (2020)");
            Check(p.Episode == 2 && p.Title == "The Long Night", "dotted style parse: " + p.Title);
            p = TvEpisodeMatcher.Parse("Show (2020) - S02E06 - Smallfolk", "Show (2020)");
            Check(p.Season == 2 && p.Episode == 6 && p.Title == "Smallfolk", "FileBot style parse");
            p = TvEpisodeMatcher.Parse("Show - S01E05-E06 - Into the Woods (1) / Into the Woods (2)", "Show (2020)");
            Check(p.Episode == 5 && p.LastEpisode == 6, "double episode parse");
            p = TvEpisodeMatcher.Parse("Show S01E01E02 Title", "Show (2020)");
            Check(p.Episode == 1 && p.LastEpisode == 2, "E01E02 parse");
            p = TvEpisodeMatcher.Parse("Show - S01E01 - 3 Days Later", "Show (2020)");
            Check(p.Episode == 1 && !p.LastEpisode.HasValue && p.Title == "3 Days Later", "number in title is not a double episode: " + p.Title);
            p = TvEpisodeMatcher.Parse("Show 1x03 Angels and Demons", "Show (2020)");
            Check(p.Season == 1 && p.Episode == 3, "1x03 parse");
            p = TvEpisodeMatcher.Parse("Show (2020) - Smallfolk", "Show (2020)");
            Check(!p.Episode.HasValue && p.Title == "Smallfolk", "no number parse strips show: " + p.Title);
            p = TvEpisodeMatcher.Parse("Show S01E02", "Show (2020)");
            Check(p.Title == "", "no title parse");

            // --- The title is the source of truth
            var r = M("Show S01E02 The Long Night");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S01E02", "title + number agree");

            r = M("Show S01E01 The Long Night");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S01E02" && r.Reason.Contains("corrected"), "title corrects number (StreamFab off by one)");

            r = M("Show S01E09 Smallfolk");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S02E02", "title corrects season too");

            r = M("Show S01E01 Angels and Demons");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S01E03", "& vs and");

            r = M("Show S02E03 Leon's Return");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S02E03", "accents and apostrophes");

            r = M("Show S02E04 Fire - The Sequel");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S02E04", "colon vs dash");

            r = M("Show S01E05 Into the Woods Part 1");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S01E05", "(1) vs Part 1");
            r = M("Show S01E05 Into the Woods, Part Two");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S01E06", "Part Two -> (2), corrected number");

            r = M("Show (2020) - Smallfolk");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S02E02", "title only, no number");

            // --- Same title in two seasons: number decides, otherwise ask
            r = M("Show S02E01 Halloween");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S02E01", "shared title, number picks");
            r = M("Show S03E07 Halloween");
            Check(r.Confidence == EpisodeConfidence.Yellow && r.Alternatives.Count == 2, "shared title, number doesn't pick");

            // --- Generic titles never override the number on their own
            r = M("Show S01E01 Pilot");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S01E01", "generic title + number agree");
            r = M("Show S01E03 Pilot");
            Check(r.Confidence == EpisodeConfidence.Yellow, "generic title disagrees with number -> ask");
            Check(TvEpisodeMatcher.IsGeneric("Episode 5") && TvEpisodeMatcher.IsGeneric("Part 2") && !TvEpisodeMatcher.IsGeneric("Smallfolk"), "generic detection");

            // --- Near misses are suggestions, not facts
            r = M("Show S01E02 The Longest Night");
            Check(r.Confidence == EpisodeConfidence.Yellow && r.Code == "S01E02", "close title -> ask: " + r.Confidence + " " + r.Reason);
            r = M("Show S01E04 Something Totally Different");
            Check(r.Confidence == EpisodeConfidence.Yellow && r.Code == "S01E04", "unmatched title, number exists -> ask");
            r = M("Show S09E09 Something Totally Different");
            Check(r.Confidence == EpisodeConfidence.Red, "nothing matches -> hold");
            r = M("Show S01E02");
            Check(r.Confidence == EpisodeConfidence.Yellow, "no title -> number only -> ask");
            r = TvEpisodeMatcher.Match(TvEpisodeMatcher.Parse("Show S01E02 Title", "Show"), new List<TvdbEpisodeInfo>());
            Check(r.Confidence == EpisodeConfidence.Red, "no TVDB data -> hold");

            // --- Double episodes
            r = M("Show - S01E05-E06 - Into the Woods (1) / Into the Woods (2)");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S01E05-E06", "double episode by titles");
            r = M("Show - S01E01-E02 - Into the Woods (1) / Into the Woods (2)");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S01E05-E06" && r.Reason.Contains("corrected"), "double episode corrected by titles");
            r = M("Show - S01E01-E02 - Pilot and More");
            Check(r.Confidence == EpisodeConfidence.Yellow && r.Code == "S01E01-E02", "double episode without split titles -> ask");

            // --- StreamFab double episodes: one number, two titles joined by " _ " (was " / ")
            p = TvEpisodeMatcher.Parse("Show (2020) - S01E05 - Into the Woods (1) _ Into the Woods (2)", "Show (2020)");
            Check(p.TitleParts.Count == 2 && p.TitleParts[1] == "Into the Woods (2)", "' _ ' splits titles");
            p = TvEpisodeMatcher.Parse("Batman_ The Animated Series () - S01E02 - Fire_ The Sequel", "Show (2020)");
            Check(p.TitleParts.Count == 0 && p.Title == "Fire The Sequel", "colon-underscore is not a separator: " + p.Title);
            r = M("Show (2020) - S01E05 - Into the Woods (1) _ Into the Woods (2)");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S01E05-E06", "' _ ' double episode: " + r.Confidence + " " + r.Code + " " + r.Reason);
            r = M("Show (2020) - S01E01 - The Long Night _ Angels & Demons");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S01E02-E03" && r.Reason.Contains("corrected"), "' _ ' double episode corrected: " + r.Code);
            r = M("Show (2020) - S01E01 - The Long Night _ Smallfolk");
            Check(r.Confidence == EpisodeConfidence.Yellow, "' _ ' titles not back to back -> ask");

            // --- Junk in front of the title
            r = M("Show (2020) - S02E14 - (Dub) Smallfolk");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S02E02", "(Dub) prefix stripped: " + r.Reason);
            r = M("Show (2020) - S02E02 - Season 2, Episode 2_ Smallfolk");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S02E02", "'Season N, Episode N:' prefix stripped: " + r.Reason);

            // --- Number words
            r = M("Show (2020) - S03E01 - Episode One");
            Check(r.Confidence == EpisodeConfidence.Green && r.Code == "S03E01", "Episode One = Episode 1: " + r.Reason);

            // --- Moving a regular episode into Specials always asks
            r = M("Show (2020) - S03E10 - Behind the Scenes");
            Check(r.Confidence == EpisodeConfidence.Yellow && r.Code == "S00E01" && r.Reason.Contains("Specials"), "move into specials -> ask: " + r.Confidence);
            r = M("Show (2020) - S00E01 - Behind the Scenes");
            Check(r.Confidence == EpisodeConfidence.Green, "special already numbered S00 stays green");

            // --- Two files landing on the same episode are both held back
            var a = M("Show S01E01 The Long Night");
            var b = M("Show S01E02 The Long Night");
            TvEpisodeMatcher.FlagCollisions(new List<EpisodeMatchResult> { a, b });
            Check(a.Confidence == EpisodeConfidence.Yellow && b.Confidence == EpisodeConfidence.Yellow, "collision downgrades both");

            // A title match beats a file that only has a (wrong) number pointing at the same episode
            var titled = M("Show S01E01 The Long Night");                       // title -> S01E02, green
            var guessed = M("Show S01E02 Something Totally Different");          // number only -> S01E02, yellow
            TvEpisodeMatcher.FlagCollisions(new List<EpisodeMatchResult> { titled, guessed });
            Check(titled.Confidence == EpisodeConfidence.Green, "title match kept against number-only guess");
            Check(guessed.Confidence == EpisodeConfidence.Yellow && guessed.Reason.Contains("another file"), "number-only guess flagged");
            // ...but two title matches plus a guess on one episode still downgrade everyone
            var t1 = M("Show S01E01 The Long Night"); var t2 = M("Show S01E02 The Long Night"); var t3 = M("Show S01E02 Something Totally Different");
            TvEpisodeMatcher.FlagCollisions(new List<EpisodeMatchResult> { t1, t2, t3 });
            Check(t1.Confidence == EpisodeConfidence.Yellow && t2.Confidence == EpisodeConfidence.Yellow, "two title matches still collide");

            // --- New names match the FileBot library format
            r = M("Show S01E01 The Long Night");
            Check(TvEpisodeMatcher.BuildFileName("Show (2020)", r) == "Show (2020) - S01E02 - The Long Night", "file name");
            Check(TvEpisodeMatcher.SeasonFolder(r) == "Season 01", "season folder");
            r = M("Show S02E04 Fire - The Sequel");
            Check(TvEpisodeMatcher.BuildFileName("Show (2020)", r) == "Show (2020) - S02E04 - Fire - The Sequel", "colon made safe");
            r = M("Show - S01E05-E06 - Into the Woods (1) / Into the Woods (2)");
            Check(TvEpisodeMatcher.BuildFileName("Show (2020)", r) == "Show (2020) - S01E05-E06 - Into the Woods (1) & Into the Woods (2)", "double episode name");

            ShadowRunNeverTouchesFiles();
        }
        catch (Exception ex)
        {
            failures++;
            Console.WriteLine("FAIL (exception): " + ex);
        }

        Console.WriteLine(failures == 0 ? "All TV episode matcher tests passed." : failures + " failure(s).");
        return failures == 0 ? 0 : 1;
    }

    private static void ShadowRunNeverTouchesFiles()
    {
        string root = Path.Combine(Path.GetTempPath(), "TvShadowTest_" + Guid.NewGuid().ToString("N"));
        try
        {
            string showDir = Path.Combine(root, "Show (2020)");
            string season = Path.Combine(showDir, "Season 01");
            Directory.CreateDirectory(season);
            string f1 = Path.Combine(season, "Show S01E01 The Long Night.mkv");
            string f2 = Path.Combine(season, "Show S01E03 Angels & Demons.mkv");
            File.WriteAllText(f1, "x");
            File.WriteAllText(f2, "y");
            string unknownDir = Path.Combine(root, "Mystery Show");
            Directory.CreateDirectory(Path.Combine(unknownDir, "Season 01"));
            File.WriteAllText(Path.Combine(unknownDir, "Season 01", "Mystery S01E01.mkv"), "z");

            var raw = Show.Select(e => new JObject(new JProperty("seasonNumber", e.Season), new JProperty("number", e.Number), new JProperty("name", e.Title))).ToList();
            var logs = new List<string>();
            string reportDir = Path.Combine(root, "reports");

            var summary = TvEpisodeShadowRun.Run(new[] { showDir, unknownDir },
                folder => folder == "Show (2020)" ? new KnownTvShow { TvdbId = "123", Name = "Show (2020)" } : null,
                id => raw,
                (kind, msg) => logs.Add(kind + ": " + msg),
                reportDir);

            Check(File.Exists(f1) && File.Exists(f2), "shadow run must not rename files");
            Check(summary.Files == 3 && summary.Green == 2 && summary.Red == 1 && summary.Corrected == 1, "shadow counts: G" + summary.Green + " Y" + summary.Yellow + " R" + summary.Red + " C" + summary.Corrected);
            Check(summary.UnknownShows == 1, "unknown show counted");
            Check(summary.ReportPath != null && File.ReadAllText(summary.ReportPath).Contains("Show (2020) - S01E02 - The Long Night.mkv"), "report has proposed name");
        }
        finally
        {
            try { Directory.Delete(root, true); } catch { }
        }
    }
}
