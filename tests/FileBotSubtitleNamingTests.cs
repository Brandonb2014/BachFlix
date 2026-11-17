using System;
using System.IO;
using System.Linq;
using BachFlixNfo.Features;

internal static class FileBotSubtitleNamingTests
{
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    public static int Main()
    {
        string root = Path.Combine(Path.GetTempPath(), "BachFlixSrtTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            string title = "Show (2022) - S02E06 - Smallfolk";
            foreach (var pair in new[]
            {
                new[] { ".en.srt", ".eng.srt" },
                new[] { ".en.sdh.srt", ".eng.sdh.srt" },
                new[] { ".eng-SDH.srt", ".eng.sdh.srt" },
                new[] { ".ger.SDH.srt", ".eng.sdh.srt" },
                new[] { ".en.forced.srt", ".eng.forced.srt" },
                new[] { ".en.closedcaptions.srt", ".eng.cc.srt" },
                new[] { ".en.closed_captions.srt", ".eng.cc.srt" },
                new[] { ".en.closed.captions.srt", ".eng.cc.srt" },
                new[] { ".en.cc.srt", ".eng.cc.srt" },
                new[] { ".en.hoh.srt", ".eng.sdh.srt" },
                new[] { ".en.sdh.closedcaptions.forced.srt", ".eng.forced.srt" },
                new[] { ".en.sdh.closedcaptions.srt", ".eng.cc.srt" },
                new[] { ".so.en.srt", ".eng.srt" },
                new[] { ".en-US.srt", ".eng.srt" },
                new[] { ".srt", ".eng.srt" }
            })
            {
                string source = Path.Combine(root, title + pair[0]);
                string expected = Path.Combine(root, title + pair[1]);
                Check(FileBotSubtitleNaming.RenameTarget(source) == expected, "Unexpected target for " + pair[0]);
                Check(FileBotSubtitleNaming.RenameTarget(expected) == null, "Naming must be idempotent for " + pair[1]);
            }
            Check(FileBotSubtitleNaming.RenameTarget(Path.Combine(root, title + ".mkv")) == null, "Non-SRT affected");
            string dotted = Path.Combine(root, "Mr. Robot - S03E03 - eps3.2_legacy.so.en.srt");
            Check(FileBotSubtitleNaming.RenameTarget(dotted) == Path.Combine(root, "Mr. Robot - S03E03 - eps3.2_legacy.eng.srt"), "Episode-title dots damaged");
            string video = Path.Combine(root, title + ".mkv");
            string embyName = Path.Combine(root, "Wrong show - S02E06 - Different Title.en.sdh.srt");
            Check(FileBotSubtitleNaming.RenameTarget(embyName, new[] { video }) == Path.Combine(root, title + ".eng.sdh.srt"), "Subtitle title must match the unique matching episode video");

            string subtitle = Path.Combine(root, title + ".en.sdh.srt");
            string target = FileBotSubtitleNaming.RenameTarget(subtitle);
            File.WriteAllText(subtitle, "subtitle contents");
            File.WriteAllText(target, "keep existing subtitle");
            Check(!FileBotSubtitleNaming.TryRename(subtitle, target), "Collision must not overwrite");
            Check(File.ReadAllText(subtitle) == "subtitle contents" && File.ReadAllText(target) == "keep existing subtitle", "Collision changed contents");
            File.Delete(target);
            Check(FileBotSubtitleNaming.TryRename(subtitle, target) && !File.Exists(subtitle), "Rename failed");
            Check(File.ReadAllText(target) == "subtitle contents", "Rename modified subtitle contents");
            string uppercase = Path.Combine(root, "Case.ENG.SDH.SRT");
            File.WriteAllText(uppercase, "case test");
            string lowercase = FileBotSubtitleNaming.RenameTarget(uppercase);
            Check(FileBotSubtitleNaming.TryRename(uppercase, lowercase), "Case normalization failed");
            Check(Directory.GetFiles(root).Contains(lowercase, StringComparer.Ordinal), "On-disk case not normalized");
            Console.WriteLine("SRT naming tests passed: custom FileBot suffixes, forced/CC/SDH priority, legacy languages, video title matching, idempotence, collisions, case-only renames.");
            return 0;
        }
        finally { Directory.Delete(root, true); }
    }
}
