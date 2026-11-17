using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using BachFlixNfo.Features;

// Pure-logic tests for the TV Episode Review sheet flow (no Google, no real files).
internal static class TvEpisodeReviewTests
{
    private static int failures;
    private static void Check(bool condition, string message) { if (!condition) { failures++; Console.WriteLine("FAIL: " + message); } }

    private const string Wa = @"\\QUADPLEX\S\Working Area\Hulu\Show\S01\";
    private const string Lib = @"\\QUADPLEX\S\TV Shows\Show (2020)";

    private static ReviewPlan Plan(string newTitle, string file, IEnumerable<string> folderFiles, IEnumerable<string> existing = null)
    {
        var exists = new HashSet<string>((existing ?? Enumerable.Empty<string>()).Concat(folderFiles), StringComparer.OrdinalIgnoreCase);
        var row = new ReviewRow { SheetRow = 5, NewTitle = newTitle, FullPath = file, OriginalTitle = Path.GetFileName(file) };
        return TvEpisodeReview.Plan(row, name => name == "Show (2020)" ? Lib : @"\\QUADPLEX\S\TV Shows\" + name, exists.Contains, f => folderFiles);
    }

    public static int Main()
    {
        try
        {
            // --- Rows written by 55m
            var green = new ShadowEpisodeRow
            {
                ShowName = "Show (2020)", SourceFile = @"\\QUADPLEX\\S\\Working Area\Hulu\Show\S01\Show () - S01E01 - The Long Night.mkv",
                ProposedPath = @"Season 01\Show (2020) - S01E02 - The Long Night.mkv",
                Match = new EpisodeMatchResult { Confidence = EpisodeConfidence.Green, Reason = "title match; number corrected from S01E01" }
            };
            var yellow = new ShadowEpisodeRow
            {
                ShowName = "Show (2020)", SourceFile = Wa + "Show () - S01E04 - Mystery.mkv", ProposedPath = @"Season 01\Show (2020) - S01E04 - Halloween.mkv",
                Match = new EpisodeMatchResult { Confidence = EpisodeConfidence.Yellow, Reason = "title matched nothing", Alternatives = { new TvdbEpisodeInfo { Season = 2, Number = 1, Title = "Halloween" } } }
            };
            var red = new ShadowEpisodeRow
            {
                ShowName = "Show (2020)", SourceFile = Wa + "Show () - S09E09 - Nope.mkv", ProposedPath = "",
                Match = new EpisodeMatchResult { Confidence = EpisodeConfidence.Red, Reason = "nothing" }
            };
            var rows = TvEpisodeReview.NewRows(new[] { green, yellow, red }, new List<string>());
            Check(rows.Count == 3, "three new rows");
            var g = rows.First(r => (string)r[1] == "Green");
            Check((string)g[0] == "", "green is approved by default");
            Check((string)g[2] == "Show () - S01E01 - The Long Night.mkv", "original title is the file name");
            Check((string)g[3] == "Show (2020) - S01E02 - The Long Night", "new title has no folder or extension: " + g[3]);
            Check((string)g[6] == @"\\QUADPLEX\S\Working Area\Hulu\Show\S01\Show () - S01E01 - The Long Night.mkv", "full path normalized: " + g[6]);
            Check(rows.Where(r => (string)r[1] != "Green").All(r => (string)r[0] == "X"), "yellow and red pre-marked X");
            Check(((string)rows.First(r => (string)r[1] == "Yellow")[4]).Contains("also possible: S02E01"), "alternatives in Why");
            var again = TvEpisodeReview.NewRows(new[] { green, yellow }, new List<string> { (string)g[6] });
            Check(again.Count == 1, "files already on the sheet are not added twice");

            // --- Planning a rename
            string video = Wa + "Show () - S01E01 - The Long Night.mkv";
            var files = new[] { video, Wa + "Show () - S01E01 - The Long Night.en.srt", Wa + "Show () - S01E01 - The Long Night.en.cc.srt", Wa + "Show () - S01E03 - Other.mkv" };
            var p = Plan("Show (2020) - S01E02 - The Long Night", video, files);
            Check(p.Ok, "plan ok: " + p.Problem);
            Check(p.Moves[0].Target == Lib + @"\Season 01\Show (2020) - S01E02 - The Long Night.mkv", "video target: " + p.Moves[0].Target);
            Check(p.Moves.Count == 3, "video + 2 subtitles, not the other episode: " + p.Moves.Count);
            Check(p.Moves.Any(m => m.Target.EndsWith("S01E02 - The Long Night.eng.srt")), "subtitle gets .eng: " + string.Join(" | ", p.Moves.Select(m => Path.GetFileName(m.Target))));
            Check(p.Moves.Any(m => m.Target.EndsWith("S01E02 - The Long Night.eng.cc.srt")), "CC subtitle gets .eng.cc");

            // en-gb + en-us SDH both become .eng.sdh.srt: keep en-us, leave en-gb behind.
            string pv = Wa + "Poldark (2015) - S01E01 - Episode One.mkv";
            var pFiles = new[] { pv, Wa + "Poldark (2015) - S01E01 - Episode One.en-gb.Dialog.Sdh.srt", Wa + "Poldark (2015) - S01E01 - Episode One.en-us.Dialog.Sdh.srt" };
            p = Plan("Poldark (2015) - S01E01 - Episode 1", pv, pFiles);
            Check(p.Ok, "regional duplicates don't block the row: " + p.Problem);
            Check(p.Moves.Count == 2 && p.Moves[1].Source.Contains(".en-us.") && p.Moves[1].Target.EndsWith("Episode 1.eng.sdh.srt"), "en-us subtitle kept: " + string.Join(" | ", p.Moves.Select(m => Path.GetFileName(m.Source))));
            Check(p.LeftBehind.Count == 1 && p.LeftBehind[0].Contains(".en-gb."), "en-gb left in place");
            Check(TvEpisodeReview.EnglishSubtitleRank("x.en.cc.srt") == 1 && TvEpisodeReview.EnglishSubtitleRank("x.eng-us.srt") == 0, "cc is not read as a region");

            p = Plan("Show (2020) - S01E02 - The Long Night.mkv", video, files);
            Check(p.Ok && p.Moves[0].Target.EndsWith("The Long Night.mkv") && !p.Moves[0].Target.EndsWith(".mkv.mkv"), "extension typed in New Title is not doubled");
            p = Plan("Show (2020) - S02E04 - Fire: The Sequel", video, files);
            Check(p.Ok && p.Moves[0].Target == Lib + @"\Season 02\Show (2020) - S02E04 - Fire - The Sequel.mkv", "edited title picks season + makes colon safe: " + p.Moves[0].Target);
            p = Plan("Show (2020) - S01E01-E02 - A & B", video, files);
            Check(p.Ok && p.Moves[0].Target.Contains(@"\Season 01\Show (2020) - S01E01-E02 - A & B.mkv"), "double episode title");

            // --- Things that must stop a row
            Check(Plan("", video, files).Problem == "no New Title", "blank new title");
            Check(Plan("Just a title", video, files).Problem.StartsWith("New Title must look like"), "new title without SxxEyy");
            Check(Plan("Show (2020) - S01E02 - The Long Night", Wa + "gone.mkv", files).SourceMissing, "missing source flagged");
            var taken = Plan("Show (2020) - S01E02 - The Long Night", video, files, new[] { Lib + @"\Season 01\Show (2020) - S01E02 - The Long Night.mkv" });
            Check(!taken.Ok && taken.Problem.StartsWith("already in library"), "never overwrite a library file");
            var takenSub = Plan("Show (2020) - S01E02 - The Long Night", video, files, new[] { Lib + @"\Season 01\Show (2020) - S01E02 - The Long Night.eng.srt" });
            Check(!takenSub.Ok, "never overwrite a library subtitle either");

            // --- Library show folder matching
            var lib = new[] { @"\\Q\S\TV Shows\Batman The Animated Series (1992)", @"\\Q\S\TV Shows\The Secret Lives of Mormon Wives (2024)" };
            Check(TvEpisodeReview.FindShowFolder("Batman: The Animated Series (1992)", lib) == lib[0], "punctuation-insensitive show folder match");
            Check(TvEpisodeReview.FindShowFolder("The Secret Lives of Mormon Wives Orange County (2026)", lib) == null, "different show is not matched");
            var noYear = new[] { @"\\Q\S\TV Shows\PJ Masks", @"\\Q\S\TV Shows\Scrubs (2001)", @"\\Q\S\TV Shows\The Beverly Hillbillies" };
            Check(TvEpisodeReview.FindShowFolder("PJ Masks (2015)", noYear) == noYear[0], "library folder without a year still matches");
            Check(TvEpisodeReview.FindShowFolder("Beverly Hillbillies (1962)", noYear) == noYear[2], "no-year folder with leading The");
            Check(TvEpisodeReview.FindShowFolder("Scrubs (2026)", noYear) == null, "different years never match");

            // --- Upgrades: an old library episode with the same SxxEyy gets replaced
            var library = new[]
            {
                Lib + @"\Season 01\Show (2020) - S01E02 - The Long Night.mp4",
                Lib + @"\Season 01\Show (2020) - S01E02 - The Long Night.eng.srt",
                Lib + @"\Season 01\Show (2020) - S01E02 - The Long Night-320-10.bif",
                Lib + @"\Season 01\Show (2020) - S01E02 - The Long Night-thumb.jpg",
                Lib + @"\Season 01\Show (2020) - S01E03 - Angels & Demons.mp4",
                Lib + @"\Season 01\Show (2020) - S01E05-E06 - Into the Woods.mp4",
            };
            var row = new ReviewRow { NewTitle = "Show (2020) - S01E02 - The Long Night", FullPath = video };
            var allExisting = new HashSet<string>(files.Concat(library), StringComparer.OrdinalIgnoreCase);
            var up = TvEpisodeReview.Plan(row, n => Lib, allExisting.Contains, f => files, f => library);
            Check(up.Ok && up.Replaces.Count == 1 && up.Replaces[0].Video.EndsWith("S01E02 - The Long Night.mp4"), "old MP4 found for replacement: " + up.Problem);
            Check(up.Replaces[0].Sidecars.Count == 3, "old srt, bif and thumb go with it: " + up.Replaces[0].Sidecars.Count);
            // Same name + extension already in the library is fine when it's the file being replaced
            var sameName = new[] { Lib + @"\Season 01\Show (2020) - S01E02 - The Long Night.mkv" };
            var up2 = TvEpisodeReview.Plan(row, n => Lib, new HashSet<string>(files.Concat(sameName), StringComparer.OrdinalIgnoreCase).Contains, f => files, f => sameName);
            Check(up2.Ok && up2.Replaces.Count == 1, "re-record with the identical name replaces it: " + up2.Problem);
            // A library double episode is only deleted when both halves come in
            var half = TvEpisodeReview.Plan(new ReviewRow { NewTitle = "Show (2020) - S01E05 - Into the Woods (1)", FullPath = video }, n => Lib, allExisting.Contains, f => files, f => library);
            var halves = new List<ReviewPlan> { half };
            TvEpisodeReview.CheckReplacementCoverage(halves);
            Check(!half.Ok && half.Problem.Contains("holds more episodes"), "half of a library double episode is held");
            var h1 = TvEpisodeReview.Plan(new ReviewRow { NewTitle = "Show (2020) - S01E05 - Into the Woods (1)", FullPath = video }, n => Lib, allExisting.Contains, f => files, f => library);
            string video2 = Wa + "Show () - S01E03 - Other.mkv";
            var h2 = TvEpisodeReview.Plan(new ReviewRow { NewTitle = "Show (2020) - S01E06 - Into the Woods (2)", FullPath = video2 }, n => Lib, allExisting.Contains, f => files, f => library);
            var both = new List<ReviewPlan> { h1, h2 };
            TvEpisodeReview.CheckReplacementCoverage(both);
            Check(h1.Ok && h2.Ok, "both halves approved -> the old double episode can go: " + h1.Problem + " / " + h2.Problem);

            DbRules();
        }
        catch (Exception ex)
        {
            failures++;
            Console.WriteLine("FAIL (exception): " + ex);
        }
        Console.WriteLine(failures == 0 ? "All TV episode review tests passed." : failures + " failure(s).");
        return failures == 0 ? 0 : 1;
    }

    private static Dictionary<string, string> DbRow(string res, string fmt, string streamFab, string note, string locations)
    {
        return new Dictionary<string, string>
        {
            { "Record Source", "Netflix" }, { "Found Locations", locations }, { "Size", "21.08" }, { "Resolution", res }, { "Format", fmt },
            { "StreamFab", streamFab }, { "PlayonPal", "Y - S5" }, { "Removed Splashes", "Y" }, { "Subtitles", "Y" },
            { "Verify Subtitles line up correctly", "Y" }, { "Full Credits", "Y" }, { "Note", note }
        };
    }

    private static string New(List<DbCellChange> changes, string column)
    {
        var c = changes.FirstOrDefault(x => x.Column == column);
        return c == null ? null : c.New;
    }

    private static void DbRules()
    {
        var today = new DateTime(2026, 11, 16);
        string lib = @"\\QUADPLEX\#-B\TV Shows\Abbott Elementary (2021)";

        // Upgrade 720p MP4 -> 1080p MKV, recorded from Hulu
        var upgrade = new ShowRunFacts { ShowName = "Abbott Elementary (2021)", LibraryFolder = @"\\QUADPLEX\#-b\TV Shows\Abbott Elementary (2021)",
            Replaced = 10, RecordSources = { "Hulu", "Hulu" }, Resolution = "1080p", Format = "MKV", ResolutionKnown = true, Today = today };
        var ch = TvShowDbUpdate.Changes(DbRow("720p", "MP4", "R - 7.0.5.8", "8/27/25 - Added in 720p.", lib), upgrade);
        Check(New(ch, "Record Source") == "Hulu", "record source from the Working Area folder");
        Check(New(ch, "StreamFab") == "Y - 7.0.5.8", "StreamFab R -> Y: " + New(ch, "StreamFab"));
        Check(New(ch, "Size") == "", "size cleared for option 24");
        Check(New(ch, "Resolution") == "1080p" && New(ch, "Format") == "MKV", "resolution + format");
        Check(New(ch, "Found Locations") == null, "same location (different case) is left alone");
        Check(TvShowDbUpdate.PlayOnColumns.All(c => New(ch, c) == ""), "PlayOn-era columns cleared");
        Check(New(ch, "Note") == "8/27/25 - Added in 720p. 11/16/26 - Upgraded to 1080p. Re-added in MKV format.", "note appended: " + New(ch, "Note"));

        // Found Locations with two drives splits into both folders; commas inside a folder name stay
        var split = TvShowDbUpdate.SplitLocations(@"\\QUADPLEX\E-I\TV Shows\Gabby's Dollhouse (2021),\\QUADPLEX\Other\The Bus\TV Shows\Gabby's Dollhouse (2021)");
        Check(split.Count == 2 && split[1] == @"\\QUADPLEX\Other\The Bus\TV Shows\Gabby's Dollhouse (2021)", "two locations split on the comma");
        Check(TvShowDbUpdate.SplitLocations(@"\\QUADPLEX\S\TV Shows\Crazy, Stupid Show (2010)").Count == 1, "comma inside a folder name is kept");
        Check(TvShowDbUpdate.SplitLocations("  ").Count == 0, "blank cell: no locations");

        // A Record Source list that already names the source is kept as-is
        var listed = DbRow("720p", "MP4", "", "", lib); listed["Record Source"] = "Prime, Hulu";
        Check(New(TvShowDbUpdate.Changes(listed, upgrade), "Record Source") == null, "'Prime, Hulu' already lists Hulu: left alone");
        listed["Record Source"] = "Prime, Netflix";
        Check(New(TvShowDbUpdate.Changes(listed, upgrade), "Record Source") == "Hulu", "list without Hulu: replaced");


        // Any StreamFab version, not just one
        foreach (var v in new[] { "R - 7.0.5.8", "R - 6.2.1.0", "R - 7.1.0.12", "r - 7.0.0.3" })
        {
            var c2 = TvShowDbUpdate.Changes(DbRow("720p", "MP4", v, "", lib), upgrade);
            Check(New(c2, "StreamFab") == "Y" + v.Substring(1), "StreamFab " + v + " -> " + New(c2, "StreamFab"));
        }
        foreach (var v in new[] { "Y - 7.0.5.8", "", "Rec", "Y - S1-4" })
            Check(New(TvShowDbUpdate.Changes(DbRow("720p", "MP4", v, "", lib), upgrade), "StreamFab") == null, "StreamFab '" + v + "' left alone");

        // Only the format changed
        ch = TvShowDbUpdate.Changes(DbRow("1080p", "MP4", "Y - 7.0.0.3", "", lib), upgrade);
        Check(New(ch, "Note") == "11/16/26 - Re-added in MKV format.", "format-only upgrade note: " + New(ch, "Note"));
        Check(New(ch, "StreamFab") == null, "StreamFab already Y is left alone");

        // Own note
        ch = TvShowDbUpdate.Changes(DbRow("720p", "MP4", "", "Some history", lib), upgrade, "Hulu doesn't have S03E02.");
        Check(New(ch, "Note") == "Some history. 11/16/26 - Hulu doesn't have S03E02.", "own note appended: " + New(ch, "Note"));
        ch = TvShowDbUpdate.Changes(DbRow("720p", "MP4", "", "Some history", lib), upgrade, "");
        Check(New(ch, "Note") == null, "blank own note = no note");

        // New show
        var brandNew = new ShowRunFacts { ShowName = "X (2026)", LibraryFolder = @"\\QUADPLEX\T-Z\TV Shows\X (2026)", NewShow = true, Added = 8,
            RecordSources = { "Hulu" }, Resolution = "1080p", Format = "MKV", ResolutionKnown = true, Today = today };
        ch = TvShowDbUpdate.Changes(DbRow("", "", "R - 7.0.5.8", "", ""), brandNew);
        Check(New(ch, "Note") == "11/16/26 - Added in 1080p MKV format.", "new show note: " + New(ch, "Note"));
        Check(New(ch, "Found Locations") == null, "blank Found Locations stays blank");

        // Mixed resolution
        brandNew.Resolution = "Mixed";
        ch = TvShowDbUpdate.Changes(DbRow("", "", "", "", ""), brandNew);
        Check(New(ch, "Resolution") == "Mixed" && New(ch, "Note") == "11/16/26 - Added in MKV format.", "mixed resolution: " + New(ch, "Note"));

        // Just new episodes for a show already done: no note, StreamFab / PlayOn untouched, location moved -> cleared
        var episodes = new ShowRunFacts { ShowName = "Abbott Elementary (2021)", LibraryFolder = lib, Added = 4, RecordSources = { "Hulu" },
            Resolution = "1080p", Format = "MKV", ResolutionKnown = true, Today = today };
        ch = TvShowDbUpdate.Changes(DbRow("1080p", "MKV", "R - 7.0.5.8", "old", @"\\QUADPLEX\#-B\TV Shows\Abbott Elementary"), episodes);
        Check(New(ch, "Note") == null && New(ch, "StreamFab") == null && New(ch, "PlayonPal") == null && New(ch, "Record Source") == null, "new episodes only: no note/StreamFab/PlayOn/source");
        Check(New(ch, "Size") == "" && New(ch, "Found Locations") == "", "new episodes: size cleared, moved location cleared");

        // Multiple Found Locations, and the Working Area service folder
        Check(TvShowDbUpdate.LocationsInclude(@"\\QUADPLEX\S\TV Shows\A (2001), \\QUADPLEX\#-B\TV Shows\Abbott Elementary (2021)", lib), "location found in a list");
        Check(TvShowDbUpdate.RecordSourceFromPath(@"\\QUADPLEX\S\Working Area\Disney+\Scrubs\S01\x.mkv") == "Disney+", "record source path");
    }
}
