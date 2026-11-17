using System;
using System.IO;
using System.Linq;
using BachFlixNfo.Features;

internal static class OutputMediaFoldersTests
{
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    private static string Folder(string root, string relative)
    {
        string path = Path.Combine(root, relative);
        Directory.CreateDirectory(path);
        return path;
    }

    public static int Main()
    {
        string testBase = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
        string root = Folder(testBase, "OutputMediaFoldersTests-" + Guid.NewGuid().ToString("N"));
        try
        {
            string output = Folder(root, "Output");
            string amazon = Folder(output, "Amazon");
            string disney = Folder(output, "Disney+");
            string custom = Folder(output, "Custom Source");
            string legacy = Folder(output, @"Ready\Prime");
            string skip = Folder(output, @"sKiP\Amazon\Unready Show\Season 01");
            string sync = Folder(output, @".sync\Amazon");
            File.WriteAllText(Path.Combine(skip, "episode.mkv"), "unready");
            File.WriteAllText(Path.Combine(sync, "state"), "sync state");

            string[] providers = OutputMediaFolders.GetProviderFolders(output).ToArray();
            Check(providers.Length == 4, "Provider discovery included Skip/.sync or treated Ready as a provider");
            Check(new[] { amazon, disney, custom, legacy }.All(providers.Contains), "Output or legacy provider missing");
            Check(OutputMediaFolders.GetRecordSource(amazon) == "Prime", "Amazon must map to Prime");
            Check(OutputMediaFolders.GetRecordSource(Path.Combine(output, "aMaZoN") + "\\") == "Prime", "Alias must ignore case and trailing separators");
            Check(OutputMediaFolders.GetRecordSource(disney) == "Disney+", "Existing provider changed");
            Check(OutputMediaFolders.GetRecordSource(custom) == "Custom Source", "Unknown provider must pass through");
            Check(OutputMediaFolders.GetRecordSource(legacy) == "Prime", "Prime alias must stay Prime");

            string movie = Folder(amazon, "Movie (2026)");
            File.WriteAllText(Path.Combine(movie, "Movie (2026).mkv"), "movie video");
            File.WriteAllText(Path.Combine(amazon, "Loose Movie.mkv"), "loose movie");
            string show = Folder(amazon, "Show");
            string season = Folder(show, "Season 01");
            File.WriteAllText(Path.Combine(season, "Show S01E01.mkv"), "episode one");
            string legacySeason = Folder(legacy, @"Show\Season 01");
            File.WriteAllText(Path.Combine(legacySeason, "Show S01E02.mkv"), "episode two");
            string providerSkip = Folder(amazon, @"Skip\Held Movie");
            File.WriteAllText(Path.Combine(providerSkip, "Held Movie.mkv"), "held movie");
            string heldShow = Folder(amazon, "Held Show");
            Folder(heldShow, @"Season 01\Skip");
            File.WriteAllText(Path.Combine(heldShow, "poster.jpg"), "held poster");
            Folder(amazon, @".sync\internal");
            Folder(amazon, @"Empty Show\Empty Season");
            string[] titles = OutputMediaFolders.GetMovableFolders(amazon).ToArray();
            Check(titles.Length == 2 && titles.Contains(movie) && titles.Contains(show), "Protected or empty folders entered the move plan");
            Check(File.ReadAllText(Path.Combine(amazon, "Loose Movie.mkv")) == "loose movie", "Loose movie changed during discovery");

            // Direct and legacy providers merge into the same canonical destination.
            string working = Folder(root, "Working");
            foreach (string provider in providers)
            {
                foreach (string title in OutputMediaFolders.GetMovableFolders(provider))
                {
                    string target = Path.Combine(working, OutputMediaFolders.GetRecordSource(provider), Path.GetFileName(title));
                    MediaMoveProgress.Move(title, target, 1, 1);
                }
            }
            Check(File.ReadAllText(Path.Combine(working, @"Prime\Show\Season 01\Show S01E01.mkv")) == "episode one", "Output show not moved to Prime");
            Check(File.ReadAllText(Path.Combine(working, @"Prime\Show\Season 01\Show S01E02.mkv")) == "episode two", "Legacy show did not merge into Prime");
            Check(File.Exists(Path.Combine(working, @"Prime\Movie (2026)\Movie (2026).mkv")), "Movie folder not moved");
            Check(!Directory.Exists(Path.Combine(working, "Amazon")), "StreamFab name leaked into destination");
            Check(Directory.GetDirectories(Path.Combine(working, "Prime")).Length == 2, "Merge created duplicate title folders");

            // Cleanup removes empty descendants, preserving roots and all excluded trees.
            Folder(output, @"Empty Provider\Empty Show\Empty Season");
            Folder(output, @"Skip\Empty Child");
            Folder(output, @".sync\Empty Child");
            Folder(disney, "Skip");
            string nonMedia = Folder(output, "Non Media");
            File.WriteAllText(Path.Combine(nonMedia, "notes.txt"), "keep");
            int errors = 0;
            OutputMediaFolders.DeleteEmptyFolders(output, (path, error) => errors++);
            Check(errors == 0, "Cleanup failed");
            Check(Directory.Exists(output), "Configured output root removed");
            Check(!Directory.Exists(Path.Combine(output, "Empty Provider")), "Empty hierarchy remains");
            Check(!Directory.Exists(Path.Combine(output, "Ready")), "Emptied legacy Ready hierarchy remains");
            Check(!Directory.Exists(custom), "Empty provider remains");
            Check(!Directory.Exists(Path.Combine(amazon, "Empty Show")), "Empty title hierarchy remains");
            Check(Directory.Exists(Path.Combine(output, @"Skip\Empty Child")), "Cleanup entered Skip");
            Check(Directory.Exists(Path.Combine(output, @".sync\Empty Child")), "Cleanup entered .sync");
            Check(Directory.Exists(Path.Combine(disney, "Skip")), "Empty Skip removed");
            Check(File.ReadAllText(Path.Combine(skip, "episode.mkv")) == "unready", "Root Skip contents changed");
            Check(File.ReadAllText(Path.Combine(providerSkip, "Held Movie.mkv")) == "held movie", "Provider Skip contents changed");
            Check(Directory.Exists(Path.Combine(heldShow, @"Season 01\Skip")), "Nested Skip removed");
            Check(File.Exists(Path.Combine(nonMedia, "notes.txt")), "Nonempty folder removed");
            Check(File.Exists(Path.Combine(amazon, "Loose Movie.mkv")), "Unmoved file removed");
            OutputMediaFolders.DeleteEmptyFolders(Path.Combine(root, "Missing"), (path, error) => errors++);
            Check(errors == 0, "Missing/offline root should be ignored during cleanup");
            Console.WriteLine("Output folder tests passed.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
        finally
        {
            // Only remove this test's uniquely named workspace inside the test bin.
            string fullRoot = Path.GetFullPath(root);
            if (fullRoot.StartsWith(testBase.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar,
                StringComparison.OrdinalIgnoreCase)
                && Path.GetFileName(fullRoot).StartsWith("OutputMediaFoldersTests-", StringComparison.Ordinal)
                && Directory.Exists(fullRoot))
                Directory.Delete(fullRoot, true);
        }
    }
}
