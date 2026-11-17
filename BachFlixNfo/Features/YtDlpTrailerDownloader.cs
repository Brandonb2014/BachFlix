using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace BachFlixNfo.Features
{
    internal sealed class TrailerCommandResult
    {
        internal int ExitCode;

        internal string Output = "";
        internal string Error = "";
    }

    internal static class YtDlpTrailerDownloader
    {
        private static string executable;
        private static bool cancelled;
        private static readonly object consoleLock = new object();
        private static bool progressVisible;

        internal static TrailerCommandResult Download(string url, string directory, string title)
        {
            string path = ResolveExecutable();
            if (path == null)
                return new TrailerCommandResult { ExitCode = -1, Error = "yt-dlp.exe was not selected; trailer download cancelled." };
            string arguments = "-f \"bv*[ext=mp4][height<=1080]+ba[ext=m4a]/b[ext=mp4]\" " +
                "--merge-output-format mp4 -o \"" + Path.Combine(directory, title + "-trailer.%(ext)s") +
                "\" --newline --no-part --no-keep-fragments --restrict-filenames --force-overwrites " + url;
            return DownloadWithRecovery(path, arguments, Run, Log);
        }

        internal static TrailerCommandResult DownloadWithRecovery(string path, string arguments,
            Func<string, string, bool, TrailerCommandResult> run, Action<string, ConsoleColor> log)
        {
            var result = run(path, arguments, true);
            if (result.ExitCode == 0) return result;

            log("[TRAILER] Download failed. Checking yt-dlp for an update...", ConsoleColor.Yellow);
            var before = run(path, "--ignore-config --version", false);
            // -U checks the current release channel and installs an update only if available.
            var update = run(path, "--ignore-config -U", true);
            if (update.ExitCode != 0)
            {
                log("[TRAILER] Update check/install failed. Keeping the failed-download result.", ConsoleColor.Red);
                return result;
            }
            var after = run(path, "--ignore-config --version", false);
            if (before.ExitCode != 0 || after.ExitCode != 0 || string.IsNullOrWhiteSpace(before.Output)
                || string.IsNullOrWhiteSpace(after.Output))
            {
                log("[TRAILER] Could not verify an updated yt-dlp version. Keeping the failed-download result.", ConsoleColor.Yellow);
                return result;
            }
            if (string.Equals(before.Output.Trim(), after.Output.Trim(), StringComparison.Ordinal))
            {
                log("[TRAILER] yt-dlp is already up to date. No retry; keeping the failed-download result.", ConsoleColor.Yellow);
                return result;
            }
            log("[TRAILER] yt-dlp updated: " + before.Output.Trim() + " -> " + after.Output.Trim() + ". Retrying download once...", ConsoleColor.Green);
            return run(path, arguments, true);
        }

        private static string ResolveExecutable()
        {
            if (executable != null && File.Exists(executable)) return executable;
            if (cancelled) return null;
            executable = FindInFolder(AppDomain.CurrentDomain.BaseDirectory);
            if (executable != null) return executable;
            executable = FindInFolder(@"C:\Dev\yt-dlp");
            if (executable != null) return executable;
            foreach (string folder in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator))
            {
                executable = FindInFolder(folder);
                if (executable != null) return executable;
            }
            executable = PromptForExecutable(Console.ReadLine, text => Log(text, ConsoleColor.Yellow));
            cancelled = executable == null;
            return executable;
        }

        internal static string FindInFolder(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;
            try
            {
                string folder = Environment.ExpandEnvironmentVariables(input.Trim().Trim('"', '\''));
                string path = Path.GetFullPath(Path.Combine(folder, "yt-dlp.exe"));
                return File.Exists(path) ? path : null;
            }
            catch (ArgumentException) { return null; }
            catch (NotSupportedException) { return null; }
            catch (PathTooLongException) { return null; }
        }

        internal static string PromptForExecutable(Func<string> read, Action<string> write)
        {
            write("[TRAILER] Could not find yt-dlp.exe. Enter the folder containing yt-dlp.exe (for example C:\\Tools\\yt-dlp). Enter q to skip trailer downloads this session:");
            while (true)
            {
                string input = read();
                if (input == null || input.Trim().Equals("q", StringComparison.OrdinalIgnoreCase)) return null;
                string found = FindInFolder(input);
                if (found != null)
                {
                    write("[TRAILER] Using " + found + " (remembered for this session).");
                    return found;
                }
                write("[TRAILER] yt-dlp.exe was not found in that folder. Enter another folder path, or q to skip:");
            }
        }

        private static TrailerCommandResult Run(string path, string arguments, bool display)
        {
            var output = new StringBuilder();
            var error = new StringBuilder();
            try
            {
                using (var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = path,
                        Arguments = arguments,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        StandardOutputEncoding = Encoding.UTF8,
                        StandardErrorEncoding = Encoding.UTF8
                    }
                })
                {
                    process.OutputDataReceived += (sender, args) =>
                    {
                        if (args.Data == null) return;
                        output.AppendLine(args.Data);
                        if (display) Log(args.Data, ConsoleColor.Cyan);
                    };
                    process.ErrorDataReceived += (sender, args) =>
                    {
                        if (args.Data == null) return;
                        error.AppendLine(args.Data);
                        if (display) Log(args.Data, ConsoleColor.Red);
                    };
                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    process.WaitForExit();
                    lock (consoleLock)
                    {
                        if (progressVisible) { Console.WriteLine(); progressVisible = false; }
                    }
                    return new TrailerCommandResult { ExitCode = process.ExitCode, Output = output.ToString(), Error = error.ToString() };
                }
            }
            catch (Win32Exception ex)
            {
                return new TrailerCommandResult { ExitCode = -1, Error = "Could not start yt-dlp.exe at " + path + ": " + ex.Message };
            }
        }

        private static void Log(string text, ConsoleColor color)
        {
            lock (consoleLock)
            {
                ConsoleColor original = Console.ForegroundColor;
                try
                {
                    if (!Console.IsOutputRedirected) Console.ForegroundColor = color;
                    if (!Console.IsOutputRedirected && text.Contains("[download]") && text.Contains("%"))
                    {
                        int width = Math.Max(1, Console.BufferWidth - 1);
                        if (text.Length > width) text = text.Substring(0, width);
                        Console.Write("\r" + text.PadRight(width));
                        progressVisible = true;
                    }
                    else
                    {
                        if (progressVisible) { Console.WriteLine(); progressVisible = false; }
                        Console.WriteLine(text);
                    }
                }
                finally { if (!Console.IsOutputRedirected) Console.ForegroundColor = original; }
            }
        }
    }
}
