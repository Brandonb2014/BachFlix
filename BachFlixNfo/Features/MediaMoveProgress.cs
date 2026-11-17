using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace BachFlixNfo.Features
{
    // Remove each source file only after its destination is ready and the source is unchanged.
    internal sealed class MediaMoveProgress
    {
        private readonly Stopwatch elapsed = Stopwatch.StartNew();
        private readonly string source;
        private readonly string name;
        private readonly int index;
        private readonly int totalItems;
        private readonly string tag;
        private readonly string noun;
        private int failedFiles;
        private long totalBytes;
        private long copiedBytes;
        private long skippedBytes;
        private int skippedFiles;
        private readonly List<string> skippedNames = new List<string>();
        private int totalFiles;
        private int completedFiles;
        private double lastUpdate = -1;
        private int visibleLines;
        private bool interactive = !Console.IsOutputRedirected;

        private MediaMoveProgress(string source, int index, int totalItems, bool movie = false, string title = null)
        {
            this.source = Path.GetFullPath(source).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            name = title ?? Path.GetFileName(this.source);
            tag = movie ? "[MOVIE]" : "[TV]";
            noun = movie ? "Movie" : "Show";
            this.index = index;
            this.totalItems = totalItems;
        }

        public static void Move(string source, string destination, int index, int totalItems, bool movie = false)
        {
            var progress = new MediaMoveProgress(source, index, totalItems, movie);
            try
            {
                progress.Transfer(destination);
            }
            catch
            {
                progress.Clear();
                Log(progress.tag + " Transfer did not finish: " + progress.name + ". Check the error below before retrying.", ConsoleColor.Red);
                throw;
            }
        }

        public static MediaMoveProgress ForMovieFiles(string source, string[] files, string title, int index, int movies)
        {
            var progress = new MediaMoveProgress(source, index, movies, true, title);
            Log("[MOVIE] Calculating total size...", ConsoleColor.Blue);
            progress.totalFiles = files.Length;
            progress.totalBytes = files.Sum(file => new FileInfo(file).Length);
            return progress;
        }

        public void MoveMovieFile(string sourceFile, string destination)
        {
            try
            {
                CopyAndRemove(new FileInfo(sourceFile), destination, new byte[1024 * 1024]);
            }
            catch
            {
                failedFiles++;
                Clear();
                throw;
            }
        }

        public void FinishMovieFiles()
        {
            Clear();
            if (failedFiles > 0)
                Log(tag + " Incomplete: " + name + " | " + failedFiles + " failed files", ConsoleColor.Red);
            else
                Log(tag + " Finished: " + name, ConsoleColor.Green);
            Log(tag + " " + completedFiles + " / " + totalFiles + " files ready (" + skippedFiles + " skipped) | " + Size(copiedBytes) + " copied | Total time: " + Duration(elapsed.Elapsed), ConsoleColor.Cyan);
            PrintSkippedSummary();
        }

        private void Transfer(string destination)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(destination));
            // Same-volume moves can finish instantly without a copy.
            try
            {
                Directory.Move(source, destination);
                Log(tag + " Finished: " + name + " | Moved directly | Total time: " + Duration(elapsed.Elapsed), ConsoleColor.Green);
                Log(tag + " Source folder removed.", ConsoleColor.Green);
                return;
            }
            catch (IOException)
            {
                // Existing destination folders and cross-volume moves need a merge/copy.
            }

            Log(tag + " Calculating total size...", ConsoleColor.Blue);
            var files = Directory.GetFiles(source, "*", SearchOption.AllDirectories)
                .Select(path => new FileInfo(path)).ToArray();
            totalFiles = files.Length;
            totalBytes = files.Sum(file => file.Length);
            Directory.CreateDirectory(destination);
            string[] directories = Directory.GetDirectories(source, "*", SearchOption.AllDirectories);
            foreach (string directory in directories)
                Directory.CreateDirectory(Path.Combine(destination, Relative(directory)));

            var buffer = new byte[1024 * 1024];
            foreach (var file in files)
            {
                string relative = Relative(file.FullName);
                string target = Path.Combine(destination, relative);
                try
                {
                    CopyAndRemove(file, target, buffer);
                }
                catch (IOException ex)
                {
                    failedFiles++;
                    Clear();
                    Log(tag + " Source retained: " + relative + " | " + ex.Message, ConsoleColor.Red);
                }
                catch (UnauthorizedAccessException ex)
                {
                    failedFiles++;
                    Clear();
                    Log(tag + " Source retained: " + relative + " | " + ex.Message, ConsoleColor.Red);
                }
            }

            Clear();
            if (failedFiles == 0)
                Log(tag + " Move complete: " + completedFiles + " / " + totalFiles + " files | " + Size(copiedBytes + skippedBytes) + " | 100% (" + skippedFiles + " skipped)", ConsoleColor.Green);
            Log(tag + " Removing source folder...", ConsoleColor.Blue);
            // Prune only empty directories from the original snapshot. Never recursively
            // delete the tree: StreamFab may have added files or folders during the copy.
            foreach (string directory in directories.OrderByDescending(path => path.Length).Concat(new[] { source }))
            {
                if (Directory.Exists(directory) && Directory.GetFileSystemEntries(directory).Length == 0)
                {
                    try { Directory.Delete(directory, false); }
                    catch (IOException)
                    {
                        // A new file may have arrived after the emptiness check.
                        if (!Directory.Exists(directory) || Directory.GetFileSystemEntries(directory).Length == 0) throw;
                    }
                }
            }
            if (failedFiles > 0)
            {
                Log(tag + " Successfully moved source files were removed; " + failedFiles + " file(s) retained or failed.", ConsoleColor.Yellow);
                PrintSkippedSummary();
                throw new IOException(failedFiles + " file(s) could not be moved or changed during transfer. "
                    + "Successfully moved source files were removed; remaining files retained: " + source);
            }
            Log(tag + " Finished: " + name, ConsoleColor.Green);
            Log(tag + " " + completedFiles + " files | " + Size(copiedBytes) + " copied | " + skippedFiles + " skipped | Total time: " + Duration(elapsed.Elapsed), ConsoleColor.Cyan);
            if (Directory.Exists(source))
                Log(tag + " New source files/folders remain for the next run: " + source, ConsoleColor.Yellow);
            else
                Log(tag + " Source folder removed.", ConsoleColor.Green);
            PrintSkippedSummary();
        }

        private void CopyAndRemove(FileInfo file, string target, byte[] buffer)
        {
            long expectedLength = file.Length;
            DateTime expectedWriteTime = file.LastWriteTimeUtc;
            CopyFile(file, target, buffer);
            var current = new FileInfo(file.FullName);
            if (current.Length != expectedLength || current.LastWriteTimeUtc != expectedWriteTime
                || new FileInfo(target).Length != expectedLength)
                throw new IOException("Source changed or destination verification failed; source retained: " + file.FullName);
            File.Delete(file.FullName);
            Clear();
            Log(tag + " Source file removed: " + Relative(file.FullName), ConsoleColor.Green);
        }

        private void CopyFile(FileInfo file, string target, byte[] buffer)
        {
            string relative = Relative(file.FullName);
            long expectedLength = file.Length;
            if (File.Exists(target) && new FileInfo(target).Length == expectedLength)
            {
                string staleTemporary = target + ".bachflix-copying";
                if (File.Exists(staleTemporary)) File.Delete(staleTemporary);
                completedFiles++;
                skippedFiles++;
                skippedNames.Add(relative);
                skippedBytes += expectedLength;
                Clear();
                Log(tag + " Skipped: " + relative + " | Already at destination; size matches (" + Size(expectedLength) + ")", ConsoleColor.Yellow);
                return;
            }
            if (File.Exists(target))
            {
                Clear();
                Log(tag + " Replacing: " + relative + " | Destination size differs from source", ConsoleColor.Yellow);
            }
            // A stale temporary file from a stopped transfer is always recopied.
            string temporary = target + ".bachflix-copying";
            var fileTime = Stopwatch.StartNew();
            long fileBytes = 0;
            bool created = false;
            try
            {
                using (var input = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read, buffer.Length, FileOptions.SequentialScan))
                using (var output = new FileStream(temporary, FileMode.Create, FileAccess.Write, FileShare.None, buffer.Length, FileOptions.SequentialScan))
                {
                    created = true;
                    if (input.Length != expectedLength || new FileInfo(file.FullName).LastWriteTimeUtc != file.LastWriteTimeUtc)
                        throw new IOException("Source file changed before copying: " + file.FullName);
                    Render(relative, fileBytes, input.Length, true);
                    int count;
                    while ((count = input.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        output.Write(buffer, 0, count);
                        fileBytes += count;
                        copiedBytes += count;
                        Render(relative, fileBytes, input.Length, false);
                    }
                    output.Flush(true);
                    if (fileBytes != expectedLength || output.Length != expectedLength)
                        throw new IOException("Copied file size does not match source: " + file.FullName);
                }
                File.SetLastWriteTimeUtc(temporary, file.LastWriteTimeUtc);
                File.SetAttributes(temporary, FileAttributes.Normal);
                if (File.Exists(target)) File.Replace(temporary, target, null);
                else File.Move(temporary, target);
            }
            catch
            {
                copiedBytes -= fileBytes;
                // Never leave an incomplete file masquerading as a finished media file.
                if (created)
                {
                    try { File.Delete(temporary); }
                    catch (Exception cleanupError)
                    {
                        Clear();
                        Log(tag + " Could not remove incomplete file: " + temporary + " | " + cleanupError.Message, ConsoleColor.Red);
                    }
                }
                throw;
            }
            completedFiles++;
            Clear();
            Log(tag + " Completed: " + relative + " | " + Size(fileBytes) + " | " + Duration(fileTime.Elapsed), ConsoleColor.Green);
        }

        private string Relative(string path)
        {
            return path.Substring(source.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        private void Render(string file, long bytes, long length, bool force)
        {
            double seconds = elapsed.Elapsed.TotalSeconds;
            if (!force && seconds - lastUpdate < (interactive ? 1 : 10)) return;
            lastUpdate = seconds;
            double speed = seconds > 0 ? copiedBytes / seconds : 0;
            long readyBytes = copiedBytes + skippedBytes;
            double percent = totalBytes == 0 ? 100 : 100.0 * readyBytes / totalBytes;
            string remaining = speed > 0 ? "~" + Duration(TimeSpan.FromSeconds(Math.Max(0, totalBytes - readyBytes) / speed)) : "calculating...";
            Clear();
            WriteLine(tag + " " + noun + " " + index + " of " + totalItems + ": " + name, ConsoleColor.Blue);
            WriteLine("Current file: " + file + " - " + (length == 0 ? 100 : 100.0 * bytes / length).ToString("0") + "%", ConsoleColor.Cyan);
            WriteLine(noun + " progress: " + completedFiles + " / " + totalFiles + " files (" + skippedFiles + " skipped) | " + Size(readyBytes) + " / " + Size(totalBytes) + " - " + percent.ToString("0") + "%", ConsoleColor.Green);
            WriteLine("Speed: " + Size(speed) + "/s | Elapsed: " + Duration(elapsed.Elapsed) + " | Remaining: " + remaining, ConsoleColor.Cyan);
        }

        private void WriteLine(string text, ConsoleColor color)
        {
            if (interactive)
            {
                try
                {
                    int width = Math.Max(1, Console.BufferWidth - 1);
                    if (text.Length > width) text = text.Substring(0, Math.Max(0, width - 3)) + (width >= 3 ? "..." : "");
                }
                catch (IOException) { interactive = false; }
            }
            Log(text, color);
            if (interactive) visibleLines++;
        }

        private void PrintSkippedSummary()
        {
            if (skippedNames.Count == 0) return;
            Log(tag + " Skipped files (" + skippedNames.Count + "; already at destination with matching sizes):", ConsoleColor.Yellow);
            foreach (string file in skippedNames)
                Log("  " + file, ConsoleColor.Yellow);
        }

        private static void Log(string text, ConsoleColor color = ConsoleColor.Blue)
        {
            if (Console.IsOutputRedirected)
            {
                Console.WriteLine(text);
                return;
            }
            ConsoleColor previous = ConsoleColor.Gray;
            bool restore = false;
            try
            {
                try
                {
                    previous = Console.ForegroundColor;
                    Console.ForegroundColor = color;
                    restore = true;
                }
                catch (IOException) { }
                Console.WriteLine(text);
            }
            finally
            {
                if (restore)
                {
                    try { Console.ForegroundColor = previous; }
                    catch (IOException) { }
                }
            }
        }

        private void Clear()
        {
            if (interactive && visibleLines > 0)
            {
                try
                {
                    int top = Math.Max(0, Console.CursorTop - visibleLines);
                    int width = Math.Max(1, Console.BufferWidth - 1);
                    for (int i = 0; i < visibleLines; i++)
                    {
                        Console.SetCursorPosition(0, top + i);
                        Console.Write(new string(' ', width));
                    }
                    Console.SetCursorPosition(0, top);
                }
                catch (IOException) { interactive = false; }
                catch (ArgumentOutOfRangeException) { interactive = false; }
            }
            visibleLines = 0;
        }

        private static string Size(double bytes)
        {
            string[] units = { "B", "KB", "MB", "GB", "TB" };
            int unit = 0;
            while (bytes >= 1024 && unit < units.Length - 1) { bytes /= 1024; unit++; }
            return bytes.ToString("0.##") + " " + units[unit];
        }

        private static string Duration(TimeSpan time)
        {
            if (time.TotalHours >= 1) return (int)time.TotalHours + "h " + time.Minutes + "m " + time.Seconds + "s";
            if (time.TotalMinutes >= 1) return (int)time.TotalMinutes + "m " + time.Seconds + "s";
            return (int)time.TotalSeconds + "s";
        }
    }
}
