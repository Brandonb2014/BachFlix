using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Text;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Provides utilities for scanning video files with FFmpeg and classifying
    /// file health as OK, WARNING, or BAD based on detected error output.
    /// </summary>
    public static class VideoHealthCheck
    {
        /// <summary>
        /// Represents the final health status assigned to a scanned video file.
        /// </summary>
        public enum VideoHealthStatus
        {
            OK,
            WARNING,
            BAD
        }

        /// <summary>
        /// Stores the results of a completed video health scan.
        /// </summary>
        public sealed class VideoHealthResult
        {
            /// <summary>
            /// Gets or sets the full path of the scanned video file.
            /// </summary>
            public string VideoPath { get; set; }

            /// <summary>
            /// Gets or sets the FFmpeg process exit code.
            /// </summary>
            public int ExitCode { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether FFmpeg produced any error output.
            /// </summary>
            public bool HadErrors { get; set; }

            /// <summary>
            /// Gets or sets the raw FFmpeg error text captured during the scan.
            /// </summary>
            public string Errors { get; set; }

            /// <summary>
            /// Gets or sets the classified health status of the file.
            /// </summary>
            public VideoHealthStatus Status { get; set; }

            /// <summary>
            /// Gets the warning patterns found in the FFmpeg output.
            /// </summary>
            public List<string> MatchedWarningPatterns { get; private set; }

            /// <summary>
            /// Gets the bad/error patterns found in the FFmpeg output.
            /// </summary>
            public List<string> MatchedBadPatterns { get; private set; }

            /// <summary>
            /// Initializes a new instance of the <see cref="VideoHealthResult"/> class.
            /// </summary>
            public VideoHealthResult()
            {
                VideoPath = string.Empty;
                Errors = string.Empty;
                MatchedWarningPatterns = new List<string>();
                MatchedBadPatterns = new List<string>();
            }
        }

        private static readonly string[] WarningPatterns =
        {
            "non monotonically increasing dts",
            "pts has no value",
            "timestamp discontinuity",
            "application provided invalid, non monotonically increasing dts",
            "invalid dts",
            "invalid pts"
        };

        private static readonly string[] BadPatterns =
        {
            "error while decoding",
            "invalid data found",
            "corrupt",
            "corrupted",
            "missing reference picture",
            "missing picture in access unit",
            "packet too short",
            "moov atom not found",
            "error splitting the input into nalu",
            "cabac decode of qscale diff failed",
            "concealing",
            "truncated",
            "decode_slice_header error",
            "no frame",
            "invalid nal",
            "sps_id",
            "pps_id",
            "crc mismatch",
            "invalid argument",
            "could not find codec parameters",
            "conversion failed"
        };

        /// <summary>
        /// Cleans a user-supplied path by trimming whitespace and removing
        /// surrounding single or double quotes.
        /// </summary>
        /// <param name="rawInput">The raw path entered or drag-dropped by the user.</param>
        /// <returns>A normalized path string, or an empty string if input is blank.</returns>
        private static string NormalizeUserPath(string rawInput)
        {
            if (string.IsNullOrWhiteSpace(rawInput))
                return string.Empty;

            string s = rawInput.Trim();
            bool changed = true;

            while (changed && s.Length > 1)
            {
                changed = false;

                if (s.StartsWith("\"") && s.EndsWith("\""))
                {
                    s = s.Substring(1, s.Length - 2).Trim();
                    changed = true;
                }
                else if (s.StartsWith("'") && s.EndsWith("'"))
                {
                    s = s.Substring(1, s.Length - 2).Trim();
                    changed = true;
                }
            }

            return s;
        }

        /// <summary>
        /// Interactive entry point that prompts the user to scan either a single
        /// video file or all supported video files within a folder.
        /// </summary>
        public static void FullMovieHealthCheck()
        {
            Console.WriteLine();
            Console.WriteLine("=== FULL MOVIE HEALTH CHECK ===");
            Console.WriteLine("Drag & drop a VIDEO FILE or a FOLDER, then press Enter.");
            Console.WriteLine("Type 'q' to cancel.");
            Console.Write("> ");

            string rawInput = Console.ReadLine();
            if (rawInput == null)
            {
                Console.WriteLine("[ERR] No input received.");
                return;
            }

            string input = NormalizeUserPath(rawInput);

            if (input.Equals("q", StringComparison.OrdinalIgnoreCase))
                return;

            Console.WriteLine("[DEBUG] Raw input : " + rawInput);
            Console.WriteLine("[DEBUG] Clean path: " + input);

            if (File.Exists(input))
            {
                Console.WriteLine();
                Console.WriteLine("[MODE] Single-file scan");
                ScanSingleFileFull(input);
            }
            else if (Directory.Exists(input))
            {
                Console.WriteLine();
                Console.WriteLine("[MODE] Folder scan");
                ScanFolderFull(input);
            }
            else
            {
                Console.WriteLine("[ERR] '" + input + "' is not a valid file or folder.");
            }
        }

        /// <summary>
        /// Scans all supported video files located in the specified folder.
        /// </summary>
        /// <param name="folderPath">The folder containing video files to scan.</param>
        public static void ScanFolderFull(string folderPath)
        {
            Console.WriteLine("[SCAN] Scanning folder: " + folderPath);

            string[] videoExtensions = { ".mkv", ".mp4", ".mov", ".avi", ".wmv", ".mpg", ".mpeg", ".m4v" };

            var files = Directory.GetFiles(folderPath)
                .Where(f => videoExtensions.Contains(Path.GetExtension(f).ToLowerInvariant()))
                .ToList();

            int total = files.Count;
            int index = 1;

            foreach (string file in files)
            {
                Console.WriteLine();
                Console.WriteLine("[FILE {0}/{1}] {2}", index, total, Path.GetFileName(file));
                Console.WriteLine("---------------------------------------");

                ScanSingleFileFull(file);
                index++;
            }
        }

        /// <summary>
        /// Scans a single video file and writes a formatted summary to the console.
        /// </summary>
        /// <param name="videoPath">The full path to the video file.</param>
        public static void ScanSingleFileFull(string videoPath)
        {
            try
            {
                VideoHealthResult result = AnalyzeVideoFile(videoPath);

                Console.WriteLine();
                WriteStatusLabel(result.Status);
                Console.Write(" " + Path.GetFileName(videoPath));

                if (result.Status == VideoHealthStatus.OK)
                {
                    Console.WriteLine(" appears healthy.");
                }
                else if (result.Status == VideoHealthStatus.WARNING)
                {
                    Console.WriteLine(" has warning-level issues.");
                }
                else
                {
                    Console.WriteLine(" appears CORRUPT or problematic.");
                }

                Console.WriteLine("      exitCode = {0}, hadErrors = {1}", result.ExitCode, result.HadErrors);

                if (result.MatchedWarningPatterns.Count > 0)
                    Console.WriteLine("      warning patterns: " + string.Join(", ", result.MatchedWarningPatterns.Distinct(StringComparer.OrdinalIgnoreCase)));

                if (result.MatchedBadPatterns.Count > 0)
                    Console.WriteLine("      bad patterns: " + string.Join(", ", result.MatchedBadPatterns.Distinct(StringComparer.OrdinalIgnoreCase)));
            }
            catch (Exception ex)
            {
                Console.WriteLine("[ERR] Failed to scan " + videoPath + ": " + ex.Message);
            }
        }

        /// <summary>
        /// Runs a full scan against a video file and returns a structured result object.
        /// </summary>
        /// <param name="videoPath">The full path to the video file.</param>
        /// <returns>A populated <see cref="VideoHealthResult"/> instance.</returns>
        public static VideoHealthResult AnalyzeVideoFile(string videoPath)
        {
            bool hadErrors;
            string errors;
            int exitCode = RunFfmpegFullScan(videoPath, out hadErrors, out errors);

            VideoHealthResult result = new VideoHealthResult
            {
                VideoPath = videoPath,
                ExitCode = exitCode,
                HadErrors = hadErrors,
                Errors = errors ?? string.Empty,
                Status = ClassifyResult(exitCode, hadErrors, errors ?? string.Empty)
            };

            FillMatchedPatterns(result);

            return result;
        }

        /// <summary>
        /// Runs FFmpeg against a file and returns the process exit code.
        /// </summary>
        /// <param name="videoPath">The full path to the video file.</param>
        /// <param name="hadErrors">True if FFmpeg produced any error output.</param>
        /// <returns>The FFmpeg exit code.</returns>
        public static int RunFfmpegFullScan(string videoPath, out bool hadErrors)
        {
            string errors;
            return RunFfmpegFullScan(videoPath, out hadErrors, out errors);
        }

        /// <summary>
        /// Runs FFmpeg against a file and captures both the exit code and raw error text.
        /// </summary>
        /// <param name="videoPath">The full path to the video file.</param>
        /// <param name="hadErrors">True if FFmpeg produced any error output.</param>
        /// <param name="errors">The raw FFmpeg error output.</param>
        /// <returns>The FFmpeg exit code.</returns>
        public static int RunFfmpegFullScan(string videoPath, out bool hadErrors, out string errors)
        {
            Console.WriteLine("[FFMPEG] FULL SCAN: " + Path.GetFileName(videoPath));
            double? duration = FfmpegHealthProgress.ReadDuration(videoPath);

            var psi = new ProcessStartInfo
            {
                FileName = "ffmpeg",
                Arguments = "-v error -nostdin -nostats -stats_period 1 -progress pipe:1 -i \"" + videoPath + "\" -f null -",
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
                StandardErrorEncoding = Encoding.UTF8,
                StandardOutputEncoding = Encoding.UTF8
            };

            StringBuilder errorBuilder = new StringBuilder();

            using (var process = new Process { StartInfo = psi })
            {
                process.ErrorDataReceived += delegate (object sender, DataReceivedEventArgs e)
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        errorBuilder.AppendLine(e.Data);
                };

                process.Start();
                process.BeginErrorReadLine();

                var progress = new FfmpegHealthProgress(duration);
                string line;
                while ((line = process.StandardOutput.ReadLine()) != null)
                    progress.ReadLine(line);

                process.WaitForExit();
                progress.Finish(process.ExitCode == 0);

                errors = errorBuilder.ToString().Trim();
                hadErrors = !string.IsNullOrWhiteSpace(errors);

                if (hadErrors)
                {
                    Console.WriteLine("[FFMPEG ERRORS]");
                    Console.WriteLine(errors);
                }

                return process.ExitCode;
            }
        }

        /// <summary>
        /// Determines the final health classification based on FFmpeg results.
        /// </summary>
        /// <param name="exitCode">The FFmpeg process exit code.</param>
        /// <param name="hadErrors">Whether FFmpeg produced error output.</param>
        /// <param name="errors">The captured FFmpeg error text.</param>
        /// <returns>The classified <see cref="VideoHealthStatus"/> value.</returns>
        private static VideoHealthStatus ClassifyResult(int exitCode, bool hadErrors, string errors)
        {
            if (exitCode != 0)
                return VideoHealthStatus.BAD;

            if (!hadErrors || string.IsNullOrWhiteSpace(errors))
                return VideoHealthStatus.OK;

            string lower = errors.ToLowerInvariant();

            bool matchedBad = ContainsAny(lower, BadPatterns);
            bool matchedWarning = ContainsAny(lower, WarningPatterns);

            if (matchedBad)
                return VideoHealthStatus.BAD;

            if (matchedWarning)
                return VideoHealthStatus.WARNING;

            return VideoHealthStatus.BAD;
        }

        /// <summary>
        /// Determines whether the supplied text contains any value from the pattern list.
        /// </summary>
        /// <param name="text">The text to inspect.</param>
        /// <param name="patterns">Patterns to search for.</param>
        /// <returns>True if any pattern is found; otherwise false.</returns>
        private static bool ContainsAny(string text, IEnumerable<string> patterns)
        {
            foreach (string pattern in patterns)
            {
                if (text.Contains(pattern))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Populates matched warning and bad pattern lists on the result object.
        /// </summary>
        /// <param name="result">The result object to update.</param>
        private static void FillMatchedPatterns(VideoHealthResult result)
        {
            string lower = (result.Errors ?? string.Empty).ToLowerInvariant();

            foreach (string pattern in WarningPatterns)
            {
                if (lower.Contains(pattern))
                    result.MatchedWarningPatterns.Add(pattern);
            }

            foreach (string pattern in BadPatterns)
            {
                if (lower.Contains(pattern))
                    result.MatchedBadPatterns.Add(pattern);
            }
        }

        /// <summary>
        /// Writes a color-coded status label to the console.
        /// </summary>
        /// <param name="status">The status value to display.</param>
        private static void WriteStatusLabel(VideoHealthStatus status)
        {
            ConsoleColor original = Console.ForegroundColor;

            if (status == VideoHealthStatus.OK)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("[OK]");
            }
            else if (status == VideoHealthStatus.WARNING)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write("[WARNING]");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("[BAD]");
            }

            Console.ForegroundColor = original;
        }
    }

    // Progress goes to stdout; only stderr participates in health classification.
    internal sealed class FfmpegHealthProgress
    {
        private readonly double? duration;
        private readonly Stopwatch elapsed = Stopwatch.StartNew();
        private double position;
        private double speed;
        private double lastRender = -10;
        private bool interactive = !Console.IsOutputRedirected;

        internal FfmpegHealthProgress(double? duration)
        {
            this.duration = duration;
            Render(false);
        }

        internal void ReadLine(string line)
        {
            int separator = line.IndexOf('=');
            if (separator < 0) return;
            string key = line.Substring(0, separator);
            string value = line.Substring(separator + 1);
            double parsed;
            if ((key == "out_time_us" || key == "out_time_ms") &&
                double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed) && parsed >= 0)
                position = Math.Max(position, parsed / 1000000.0);
            else if (key == "speed" && double.TryParse(value.TrimEnd('x'), NumberStyles.Float, CultureInfo.InvariantCulture, out parsed) && parsed > 0)
                speed = parsed;
            else if (key == "progress") Render(false);
        }

        internal void Finish(bool succeeded)
        {
            string text = succeeded ? "[FFMPEG] Scan finished | Elapsed: " + Clock(elapsed.Elapsed.TotalSeconds)
                : "[FFMPEG] Scan stopped with errors | Elapsed: " + Clock(elapsed.Elapsed.TotalSeconds);
            if (succeeded && duration.HasValue) text = "[FFMPEG] Scan finished: 100% | Elapsed: " + Clock(elapsed.Elapsed.TotalSeconds);
            Write(text, succeeded ? ConsoleColor.Green : ConsoleColor.Red);
            if (interactive) Console.WriteLine();
        }

        internal string CurrentText { get { return Format(position, duration, speed, elapsed.Elapsed.TotalSeconds); } }

        private void Render(bool force)
        {
            double seconds = elapsed.Elapsed.TotalSeconds;
            if (!force && seconds - lastRender < (interactive ? 1 : 10)) return;
            lastRender = seconds;
            Write(CurrentText, ConsoleColor.Cyan);
        }

        internal static string Format(double position, double? duration, double speed, double elapsed)
        {
            string percent = duration.HasValue && duration.Value > 0
                ? Math.Min(99.9, 100 * position / duration.Value).ToString("0.0", CultureInfo.InvariantCulture) + "%"
                : "duration unknown";
            string remaining = duration.HasValue && speed > 0
                ? "~" + Clock(Math.Max(0, duration.Value - position) / speed) : "calculating...";
            return "[FFMPEG] Scan: " + percent + " | Checked: " + Clock(position) +
                (duration.HasValue ? " / " + Clock(duration.Value) : "") +
                " | Speed: " + (speed > 0 ? speed.ToString("0.0", CultureInfo.InvariantCulture) + "x" : "calculating...") +
                " | Elapsed: " + Clock(elapsed) + " | Remaining: " + remaining;
        }

        private void Write(string text, ConsoleColor color)
        {
            if (!interactive) { Console.WriteLine(text); return; }
            ConsoleColor original = Console.ForegroundColor;
            try
            {
                int width = Math.Max(1, Console.BufferWidth - 1);
                if (text.Length > width) text = text.Substring(0, width);
                Console.ForegroundColor = color;
                Console.Write("\r" + text.PadRight(width));
            }
            catch (IOException)
            {
                interactive = false;
                Console.WriteLine(text);
            }
            finally { Console.ForegroundColor = original; }
        }

        private static string Clock(double seconds)
        {
            var time = TimeSpan.FromSeconds(Math.Max(0, seconds));
            return ((int)time.TotalHours).ToString("00") + ":" + time.Minutes.ToString("00") + ":" + time.Seconds.ToString("00");
        }

        internal static double? ReadDuration(string videoPath)
        {
            try
            {
                using (var probe = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "ffprobe",
                        Arguments = "-v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 \"" + videoPath + "\"",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    }
                })
                {
                    probe.Start();
                    var output = probe.StandardOutput.ReadToEndAsync();
                    var error = probe.StandardError.ReadToEndAsync();
                    if (!probe.WaitForExit(10000)) { probe.Kill(); probe.WaitForExit(); return null; }
                    double seconds;
                    if (probe.ExitCode == 0 && double.TryParse(output.Result.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out seconds)
                        && seconds > 0 && !double.IsInfinity(seconds) && !double.IsNaN(seconds)) return seconds;
                }
            }
            catch (System.ComponentModel.Win32Exception) { }
            catch (IOException) { }
            catch (InvalidOperationException) { }
            return null;
        }
    }
}
