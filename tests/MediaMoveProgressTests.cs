using System;
using System.IO;
using System.Linq;
using BachFlixNfo.Features;

internal static class MediaMoveProgressTests
{
    private sealed class TransferLog : StringWriter
    {
        public Action<string> OnLine;

        public override void WriteLine(string value)
        {
            base.WriteLine(value);
            OnLine?.Invoke(value);
        }
    }

    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    public static int Main()
    {
        string root = Path.Combine(Path.GetTempPath(), "BachFlixTvMoveTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var original = Console.Out;
        var log = new TransferLog();
        try
        {
            Console.SetOut(log);
            string source = Path.Combine(root, "Show");
            string destination = Path.Combine(root, "Target");
            Directory.CreateDirectory(Path.Combine(source, "S01", "Empty"));
            Directory.CreateDirectory(Path.Combine(destination, "S01"));
            byte[] episode = Enumerable.Range(0, 3 * 1024 * 1024 + 7).Select(i => (byte)(i % 251)).ToArray();
            File.WriteAllBytes(Path.Combine(source, "S01", "Episode.mkv"), episode);
            File.WriteAllText(Path.Combine(source, "S01", "Episode.srt"), "subtitle");
            File.WriteAllBytes(Path.Combine(source, "empty.nfo"), new byte[0]);
            File.WriteAllText(Path.Combine(destination, "S01", "Episode.mkv"), "existing episode");
            MediaMoveProgress.Move(source, destination, 1, 4);
            Check(!Directory.Exists(source), "Completed source must be removed");
            Check(File.ReadAllBytes(Path.Combine(destination, "S01", "Episode.mkv")).SequenceEqual(episode), "Partial episode was not replaced correctly");
            Check(!File.Exists(Path.Combine(destination, "S01", "Episode (2).mkv")), "Duplicate episode created");
            Check(File.ReadAllText(Path.Combine(destination, "S01", "Episode.srt")) == "subtitle", "Subtitle missing");
            Check(Directory.Exists(Path.Combine(destination, "S01", "Empty")), "Empty folder lost");
            Check(new FileInfo(Path.Combine(destination, "empty.nfo")).Length == 0, "Empty file lost");
            string output = log.ToString();
            Check(output.Contains("[TV] Show 1 of 4: Show"), "Show position missing");
            Check(output.Contains("Current file:") && output.Contains("Speed:") && output.Contains("Remaining:"), "Live progress missing");
            Check(output.Contains("[TV] Completed: S01") && output.Contains("3 / 3 files") && output.Contains("100%"), "Completion progress missing");
            Check(output.IndexOf("Removing source folder...") < output.IndexOf("[TV] Finished: Show"), "Summary before source cleanup");

            log.GetStringBuilder().Clear();
            source = Path.Combine(root, "ResumeShow");
            destination = Path.Combine(root, "ResumeTarget");
            Directory.CreateDirectory(source);
            Directory.CreateDirectory(destination);
            File.WriteAllBytes(Path.Combine(source, "complete.mkv"), episode);
            File.WriteAllBytes(Path.Combine(destination, "complete.mkv"), episode);
            DateTime savedTime = new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc);
            File.SetLastWriteTimeUtc(Path.Combine(destination, "complete.mkv"), savedTime);
            File.WriteAllBytes(Path.Combine(source, "partial.mkv"), episode);
            File.WriteAllText(Path.Combine(destination, "partial.mkv"), "interrupted final file");
            File.WriteAllText(Path.Combine(destination, "partial.mkv.bachflix-copying"), "interrupted temporary file");
            File.WriteAllText(Path.Combine(source, "missing.mkv"), "missing episode");
            File.WriteAllText(Path.Combine(destination, "unrelated.mkv"), "keep existing extra episode");
            MediaMoveProgress.Move(source, destination, 2, 4);
            Check(!Directory.Exists(source), "Resume did not remove verified source");
            Check(File.GetLastWriteTimeUtc(Path.Combine(destination, "complete.mkv")) == savedTime, "Matching episode recopied");
            Check(File.ReadAllBytes(Path.Combine(destination, "partial.mkv")).SequenceEqual(episode), "Interrupted episode not repaired");
            Check(File.ReadAllText(Path.Combine(destination, "missing.mkv")) == "missing episode", "Missing episode not copied");
            Check(File.ReadAllText(Path.Combine(destination, "unrelated.mkv")) == "keep existing extra episode", "Unrelated destination episode changed");
            Check(!Directory.GetFiles(destination).Any(path => path.Contains("(2)") || path.EndsWith(".bachflix-copying")), "Resume left duplicates or temporary files");
            Check(log.ToString().Contains("Skipped: complete.mkv") && log.ToString().Contains("Replacing: partial.mkv") && log.ToString().Contains("1 skipped"), "Resume reporting incorrect");
            string resumeOutput = log.ToString();
            int skippedSummary = resumeOutput.IndexOf("Skipped files (1; already at destination with matching sizes):");
            Check(skippedSummary > resumeOutput.IndexOf("Finished: ResumeShow"), "Skipped filenames summary must follow show completion");
            Check(resumeOutput.Substring(skippedSummary).Contains("  complete.mkv"), "Skipped filename missing from final summary");

            // Reproduce output folders receiving new downloads after the snapshot.
            log.GetStringBuilder().Clear();
            source = Path.Combine(root, "GrowingShow");
            destination = Path.Combine(root, "GrowingTarget");
            Directory.CreateDirectory(Path.Combine(source, "S01"));
            Directory.CreateDirectory(destination);
            string originalEpisode = Path.Combine(source, "S01", "episode-one.mkv");
            File.WriteAllText(originalEpisode, "finished episode");
            log.OnLine = line =>
            {
                if (!line.Contains("Completed: S01")) return;
                log.OnLine = null;
                File.WriteAllText(Path.Combine(source, "S01", "episode-two.mkv"), "new download");
                Directory.CreateDirectory(Path.Combine(source, "Skip"));
                File.WriteAllText(Path.Combine(source, "Skip", "hold.mkv"), "hold");
            };
            MediaMoveProgress.Move(source, destination, 1, 1);
            Check(!File.Exists(originalEpisode), "New download prevented completed episode cleanup");
            Check(File.ReadAllText(Path.Combine(destination, "S01", "episode-one.mkv")) == "finished episode", "Completed episode missing at destination");
            Check(File.ReadAllText(Path.Combine(source, "S01", "episode-two.mkv")) == "new download", "New episode was deleted");
            Check(File.ReadAllText(Path.Combine(source, "Skip", "hold.mkv")) == "hold", "New protected subtree was deleted");
            Check(!File.Exists(Path.Combine(destination, "S01", "episode-two.mkv")), "New files must wait for the next run");
            Check(log.ToString().Contains("New source files/folders remain for the next run"), "New source files were incorrectly reported as an error");

            // Same-size edits must also survive cleanup, while other files finish.
            log.GetStringBuilder().Clear();
            source = Path.Combine(root, "ChangingShow");
            destination = Path.Combine(root, "ChangingTarget");
            Directory.CreateDirectory(source);
            Directory.CreateDirectory(destination);
            string changingFile = Path.Combine(source, "changed.mkv");
            File.WriteAllText(changingFile, "before");
            File.WriteAllText(Path.Combine(source, "stable.mkv"), "stable episode");
            log.OnLine = line =>
            {
                if (!line.Contains("Completed: changed.mkv")) return;
                log.OnLine = null;
                DateTime oldTime = File.GetLastWriteTimeUtc(changingFile);
                File.WriteAllText(changingFile, "after!");
                File.SetLastWriteTimeUtc(changingFile, oldTime.AddMinutes(1));
            };
            bool changed = false;
            try { MediaMoveProgress.Move(source, destination, 1, 1); }
            catch (IOException) { changed = true; }
            Check(changed && File.ReadAllText(changingFile) == "after!", "Changed source was deleted");
            Check(File.ReadAllText(Path.Combine(destination, "changed.mkv")) == "before", "Original snapshot not copied correctly");
            Check(!File.Exists(Path.Combine(source, "stable.mkv"))
                && File.ReadAllText(Path.Combine(destination, "stable.mkv")) == "stable episode", "Changed file prevented independent successful cleanup");
            Check(log.ToString().Contains("Successfully moved source files were removed"), "Partial cleanup result not explained");

            source = Path.Combine(root, "PublishFailureShow");
            destination = Path.Combine(root, "PublishFailureTarget");
            Directory.CreateDirectory(source);
            Directory.CreateDirectory(destination);
            string retainedSource = Path.Combine(source, "episode.mkv");
            string retainedTarget = Path.Combine(destination, "episode.mkv");
            File.WriteAllBytes(retainedSource, episode);
            File.WriteAllText(retainedTarget, "old incomplete episode");
            bool publishFailed = false;
            using (var held = new FileStream(retainedTarget, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                try { MediaMoveProgress.Move(source, destination, 1, 1); }
                catch (IOException) { publishFailed = true; }
            }
            Check(publishFailed && File.ReadAllBytes(retainedSource).SequenceEqual(episode), "Failed replacement lost source");
            Check(File.ReadAllText(retainedTarget) == "old incomplete episode", "Failed replacement damaged old destination");
            Check(!File.Exists(retainedTarget + ".bachflix-copying"), "Failed replacement left temporary file");
            MediaMoveProgress.Move(source, destination, 1, 1);
            Check(!Directory.Exists(source) && File.ReadAllBytes(retainedTarget).SequenceEqual(episode), "Retry after failed replacement did not recover");

            source = Path.Combine(root, "LockedShow");
            destination = Path.Combine(root, "LockedTarget");
            Directory.CreateDirectory(source);
            Directory.CreateDirectory(destination);
            string locked = Path.Combine(source, "locked.mkv");
            File.WriteAllText(locked, "keep source");
            File.WriteAllText(Path.Combine(source, "unlocked.mkv"), "move this episode");
            bool failed = false;
            using (var held = new FileStream(locked, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                try { MediaMoveProgress.Move(source, destination, 2, 4); }
                catch (IOException) { failed = true; }
            }
            Check(failed && File.ReadAllText(locked) == "keep source", "Failed copy must preserve source");
            Check(!File.Exists(Path.Combine(destination, "locked.mkv")), "Failed copy left incomplete destination");
            Check(!File.Exists(Path.Combine(source, "unlocked.mkv"))
                && File.ReadAllText(Path.Combine(destination, "unlocked.mkv")) == "move this episode", "Locked file prevented successful sibling cleanup");

            source = Path.Combine(root, "FastShow");
            destination = Path.Combine(root, "FastTarget");
            Directory.CreateDirectory(source);
            File.WriteAllText(Path.Combine(source, "episode.mkv"), "fast move");
            MediaMoveProgress.Move(source, destination, 3, 4);
            Check(!Directory.Exists(source) && File.ReadAllText(Path.Combine(destination, "episode.mkv")) == "fast move", "Direct move failed");
            Check(log.ToString().Contains("Moved directly"), "Direct move summary missing");

            log.GetStringBuilder().Clear();
            source = Path.Combine(root, "MovieFolder");
            destination = Path.Combine(root, "MovieTarget");
            Directory.CreateDirectory(source);
            Directory.CreateDirectory(destination);
            File.WriteAllBytes(Path.Combine(source, "movie.mkv"), episode);
            MediaMoveProgress.Move(source, destination, 1, 2, true);
            Check(File.ReadAllBytes(Path.Combine(destination, "movie.mkv")).SequenceEqual(episode), "Movie folder copy differs");
            Check(!Directory.Exists(source), "Movie folder source not removed");
            Check(log.ToString().Contains("[MOVIE] Movie 1 of 2: MovieFolder") && log.ToString().Contains("Movie progress:"), "Movie folder progress missing");
            Check(log.ToString().Contains("[MOVIE] Finished: MovieFolder") && !log.ToString().Contains("[TV]"), "Movie folder labels incorrect");

            log.GetStringBuilder().Clear();
            source = Path.Combine(root, "LooseMovies");
            destination = Path.Combine(root, "RenamedMovie");
            Directory.CreateDirectory(source);
            Directory.CreateDirectory(destination);
            string video = Path.Combine(source, "original.mkv");
            string subtitle = Path.Combine(source, "original.en.srt");
            string trailer = Path.Combine(source, "original - trailer.mp4");
            File.WriteAllBytes(video, episode);
            File.WriteAllText(subtitle, "movie subtitle");
            File.WriteAllText(trailer, "trailer");
            File.WriteAllText(Path.Combine(source, "other.mkv"), "unrelated movie");
            var movie = MediaMoveProgress.ForMovieFiles(source, new[] { video, subtitle, trailer }, "Renamed Movie", 2, 2);
            movie.MoveMovieFile(video, Path.Combine(destination, "Renamed Movie.mkv"));
            movie.MoveMovieFile(subtitle, Path.Combine(destination, "Renamed Movie.en.srt"));
            movie.MoveMovieFile(trailer, Path.Combine(destination, "Renamed Movie-trailer.mp4"));
            movie.FinishMovieFiles();
            Check(!File.Exists(video) && !File.Exists(subtitle) && !File.Exists(trailer), "Loose movie sources not removed");
            Check(File.ReadAllBytes(Path.Combine(destination, "Renamed Movie.mkv")).SequenceEqual(episode), "Loose video differs");
            Check(File.ReadAllText(Path.Combine(destination, "Renamed Movie.en.srt")) == "movie subtitle", "Renamed subtitle missing");
            Check(File.ReadAllText(Path.Combine(destination, "Renamed Movie-trailer.mp4")) == "trailer", "Normalized trailer missing");
            Check(File.ReadAllText(Path.Combine(source, "other.mkv")) == "unrelated movie", "Unrelated movie removed");
            Check(log.ToString().Contains("[MOVIE] Movie 2 of 2: Renamed Movie") && log.ToString().Contains("3 / 3 files ready"), "Loose movie totals incorrect");

            log.GetStringBuilder().Clear();
            File.WriteAllBytes(video, episode);
            File.WriteAllText(subtitle, "new subtitle with different size");
            savedTime = new DateTime(2021, 2, 3, 4, 5, 6, DateTimeKind.Utc);
            File.SetLastWriteTimeUtc(Path.Combine(destination, "Renamed Movie.mkv"), savedTime);
            movie = MediaMoveProgress.ForMovieFiles(source, new[] { video, subtitle }, "Renamed Movie", 1, 1);
            movie.MoveMovieFile(video, Path.Combine(destination, "Renamed Movie.mkv"));
            movie.MoveMovieFile(subtitle, Path.Combine(destination, "Renamed Movie.en.srt"));
            movie.FinishMovieFiles();
            Check(!File.Exists(video) && !File.Exists(subtitle), "Resumed loose movie sources not removed");
            Check(File.GetLastWriteTimeUtc(Path.Combine(destination, "Renamed Movie.mkv")) == savedTime, "Resumed loose video was recopied");
            Check(File.ReadAllText(Path.Combine(destination, "Renamed Movie.en.srt")) == "new subtitle with different size", "Mismatched movie sidecar not replaced");
            Check(log.ToString().Contains("1 skipped") && log.ToString().Contains("2 / 2 files ready"), "Loose movie resume totals incorrect");
            Check(log.ToString().Contains("Skipped files (1;") && log.ToString().Contains("  original.mkv"), "Movie skipped filenames summary missing");

            log.GetStringBuilder().Clear();
            File.WriteAllText(video, "retain failed movie");
            movie = MediaMoveProgress.ForMovieFiles(source, new[] { video }, "Failed Movie", 1, 1);
            failed = false;
            using (var held = new FileStream(video, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                try { movie.MoveMovieFile(video, Path.Combine(destination, "failed.mkv")); }
                catch (IOException) { failed = true; }
            }
            movie.FinishMovieFiles();
            Check(failed && File.ReadAllText(video) == "retain failed movie", "Failed loose movie lost source");
            Check(log.ToString().Contains("[MOVIE] Incomplete: Failed Movie") && !log.ToString().Contains("[MOVIE] Finished:"), "Failed movie incorrectly marked finished");
            Console.SetOut(original);
            Console.WriteLine("Media move tests passed: growing sources, same-size edits, partial cleanup, interrupted transfer recovery, matching-size skips, replacement preservation, TV/movie progress, sidecars.");
            return 0;
        }
        catch (Exception error)
        {
            Console.SetOut(original);
            Console.Error.WriteLine(error);
            return 1;
        }
        finally
        {
            Console.SetOut(original);
            string fullRoot = Path.GetFullPath(root);
            string tempRoot = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (fullRoot.StartsWith(tempRoot, StringComparison.OrdinalIgnoreCase)
                && Path.GetFileName(fullRoot).StartsWith("BachFlixTvMoveTests-", StringComparison.Ordinal))
                Directory.Delete(fullRoot, true);
        }
    }
}
