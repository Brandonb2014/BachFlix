using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BachFlixNfo.Features
{
    public static class AudioCensorRunner
    {
        public static void RunInteractive(Action<string, string, int> log)
        {
            Write(log, "info", "=== AUDIO CENSOR PREP ===", 1);
            Write(log, "warning", "This phase scans subtitles for profanity and prepares the review list. No MKV will be changed or created yet.", 2);

            string mkvPath = PromptForExistingMkv(log);
            if (string.IsNullOrWhiteSpace(mkvPath))
                return;

            var scanner = new SubtitleScanner();
            string subtitlePath = PromptForSubtitlePath(log, scanner, mkvPath);
            if (string.IsNullOrWhiteSpace(subtitlePath))
                return;

            string dictionaryPath = PromptForDictionaryPath(log, mkvPath);
            if (string.IsNullOrWhiteSpace(dictionaryPath))
                return;

            int audioTrackIndex = PromptForAudioTrackIndex(log);

            var commandLog = new List<string>();
            AudioCensorScanSummary summary;

            try
            {
                summary = RunScan(new AudioCensorScanOptions
                {
                    MkvPath = mkvPath,
                    SubtitlePath = subtitlePath,
                    ProfanityDictionaryPath = dictionaryPath,
                    AudioTrackIndex = audioTrackIndex,
                    DetectExternalTools = true,
                    Log = log,
                    CommandLog = commandLog
                });
            }
            catch (Exception ex)
            {
                Write(log, "error", "Audio censor scan failed.", 1);
                Write(log, "harderror", ex.Message, 2);
                return;
            }

            summary.ApprovedOccurrences = ReviewOccurrences(summary.ScanResult.Occurrences, log);
            WriteSummary(summary, log);
            WriteLogFile(summary, commandLog, log);
        }

        public static AudioCensorScanSummary RunScan(AudioCensorScanOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            if (string.IsNullOrWhiteSpace(options.MkvPath))
                throw new ArgumentException("MKV path is required.", nameof(options));

            if (!File.Exists(options.MkvPath))
                throw new FileNotFoundException("Input MKV was not found.", options.MkvPath);

            if (!string.Equals(Path.GetExtension(options.MkvPath), ".mkv", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Input file must be an MKV.", nameof(options));

            var commandLog = options.CommandLog ?? new List<string>();
            var muxService = new FfmpegMuxService();
            IWhisperAlignmentService alignmentService = new WhisperAlignmentService();

            ExternalToolResolution ffmpeg = new ExternalToolResolution { ToolName = "ffmpeg", Found = false };
            ExternalToolResolution ffprobe = new ExternalToolResolution { ToolName = "ffprobe", Found = false };
            ExternalToolResolution whisperx = new ExternalToolResolution { ToolName = "whisperx", Found = false };

            if (options.DetectExternalTools)
            {
                Write(options.Log, "info", "Detecting FFmpeg, FFprobe, and WhisperX executables...", 1);
                ffmpeg = muxService.ResolveFfmpeg(options.FfmpegPath, options.Log, commandLog);
                ffprobe = muxService.ResolveFfprobe(options.FfprobePath, options.Log, commandLog);
                whisperx = alignmentService.ResolveWhisperX(options.WhisperXPath, options.Log, commandLog);

                if (!ffmpeg.Found)
                    Write(options.Log, "warning", ffmpeg.Message, 1);

                if (!ffprobe.Found)
                    Write(options.Log, "warning", ffprobe.Message, 1);

                if (!whisperx.Found)
                    Write(options.Log, "warning", whisperx.Message, 1);
            }

            string subtitlePath = options.SubtitlePath;
            if (string.IsNullOrWhiteSpace(subtitlePath))
            {
                var scannerForLocate = new SubtitleScanner();
                subtitlePath = scannerForLocate.FindMatchingSrtForVideo(options.MkvPath);
            }

            if (string.IsNullOrWhiteSpace(subtitlePath) || !File.Exists(subtitlePath))
                throw new FileNotFoundException("Matching SRT subtitle file was not found.", subtitlePath ?? "");

            Write(options.Log, "info", "Loading profanity dictionary...", 1);
            ProfanityDictionary dictionary = ProfanityDictionary.LoadFromFile(options.ProfanityDictionaryPath);
            Write(options.Log, "success", "Loaded " + dictionary.Count + " profanity terms.", 1);

            Write(options.Log, "info", "Scanning subtitles: " + subtitlePath, 1);
            var scanner = new SubtitleScanner();
            SubtitleScanResult scanResult = scanner.ScanFile(subtitlePath, dictionary);
            Write(options.Log, "success", "Found " + scanResult.Occurrences.Count + " profanity occurrence(s).", 2);

            return new AudioCensorScanSummary
            {
                MkvPath = options.MkvPath,
                SubtitlePath = subtitlePath,
                ProfanityDictionaryPath = options.ProfanityDictionaryPath,
                DictionaryTermCount = dictionary.Count,
                AudioTrackIndex = options.AudioTrackIndex < 0 ? 0 : options.AudioTrackIndex,
                OutputMkvPath = muxService.BuildCleanOutputPath(options.MkvPath),
                ScanResult = scanResult,
                ApprovedOccurrences = new List<ProfanityOccurrence>(scanResult.Occurrences),
                Ffmpeg = ffmpeg,
                Ffprobe = ffprobe,
                WhisperX = whisperx
            };
        }

        private static string PromptForExistingMkv(Action<string, string, int> log)
        {
            while (true)
            {
                Write(log, "question", "Enter MKV file path, or 0 to cancel:", 1);
                string input = CleanConsolePath(Console.ReadLine());

                if (input == "0")
                    return null;

                if (File.Exists(input) && string.Equals(Path.GetExtension(input), ".mkv", StringComparison.OrdinalIgnoreCase))
                    return input;

                Write(log, "error", "Please enter an existing .mkv file path.", 1);
            }
        }

        private static string PromptForSubtitlePath(Action<string, string, int> log, SubtitleScanner scanner, string mkvPath)
        {
            string suggested = scanner.FindMatchingSrtForVideo(mkvPath);

            if (!string.IsNullOrWhiteSpace(suggested))
            {
                Write(log, "info", "Found matching SRT:", 1);
                Write(log, "data", suggested, 1);
                Write(log, "question", "Press Enter to use it, enter another SRT path, or 0 to cancel:", 1);
                string response = CleanConsolePath(Console.ReadLine());

                if (response == "0")
                    return null;

                if (string.IsNullOrWhiteSpace(response))
                    return suggested;

                if (File.Exists(response) && string.Equals(Path.GetExtension(response), ".srt", StringComparison.OrdinalIgnoreCase))
                    return response;

                Write(log, "error", "That SRT path was not found.", 1);
            }

            while (true)
            {
                Write(log, "question", "Enter SRT subtitle file path, or 0 to cancel:", 1);
                string input = CleanConsolePath(Console.ReadLine());

                if (input == "0")
                    return null;

                if (File.Exists(input) && string.Equals(Path.GetExtension(input), ".srt", StringComparison.OrdinalIgnoreCase))
                    return input;

                Write(log, "error", "Please enter an existing .srt file path.", 1);
            }
        }

        private static string PromptForDictionaryPath(Action<string, string, int> log, string mkvPath)
        {
            List<string> suggestions = GetDictionarySuggestions(mkvPath)
                .Where(File.Exists)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (suggestions.Count > 0)
            {
                string suggested = suggestions[0];
                Write(log, "info", "Found profanity dictionary:", 1);
                Write(log, "data", suggested, 1);
                Write(log, "question", "Press Enter to use it, enter another dictionary path, or 0 to cancel:", 1);
                string response = CleanConsolePath(Console.ReadLine());

                if (response == "0")
                    return null;

                if (string.IsNullOrWhiteSpace(response))
                    return suggested;

                if (File.Exists(response))
                    return response;

                Write(log, "error", "That dictionary path was not found.", 1);
            }

            while (true)
            {
                Write(log, "question", "Enter profanity dictionary text file path, or 0 to cancel:", 1);
                string input = CleanConsolePath(Console.ReadLine());

                if (input == "0")
                    return null;

                if (File.Exists(input))
                    return input;

                Write(log, "error", "Please enter an existing profanity dictionary file path.", 1);
            }
        }

        private static int PromptForAudioTrackIndex(Action<string, string, int> log)
        {
            Write(log, "question", "Audio track index to censor later, 0-based audio track number. Press Enter for 0:", 1);
            string input = (Console.ReadLine() ?? "").Trim();

            int audioTrackIndex;
            if (!int.TryParse(input, out audioTrackIndex) || audioTrackIndex < 0)
                audioTrackIndex = 0;

            return audioTrackIndex;
        }

        private static List<ProfanityOccurrence> ReviewOccurrences(
            IReadOnlyList<ProfanityOccurrence> occurrences,
            Action<string, string, int> log)
        {
            if (occurrences == null || occurrences.Count == 0)
            {
                Write(log, "success", "Review list is empty. No profanity was detected in the subtitle scan.", 2);
                return new List<ProfanityOccurrence>();
            }

            Write(log, "info", "Review detected profanity:", 1);

            foreach (ProfanityOccurrence occurrence in occurrences)
            {
                string line = string.Format(
                    "{0}. {1} --> {2} | SRT #{3} | \"{4}\" (dictionary: {5})",
                    occurrence.ReviewNumber,
                    FormatTime(occurrence.Cue.Start),
                    FormatTime(occurrence.Cue.End),
                    occurrence.SubtitleSequenceNumber,
                    occurrence.Word,
                    occurrence.DictionaryTerm);

                Write(log, "data", line, 1);
                Write(log, "default", "   " + Truncate(occurrence.Cue.Text, 180), 1);
            }

            Write(log, "question", "Enter comma-separated review numbers to exclude, or press Enter to keep all:", 1);
            string response = (Console.ReadLine() ?? "").Trim();

            if (string.IsNullOrWhiteSpace(response))
                return occurrences.ToList();

            HashSet<int> excludedNumbers = ParseNumberList(response);
            List<ProfanityOccurrence> approved = occurrences
                .Where(o => !excludedNumbers.Contains(o.ReviewNumber))
                .ToList();

            Write(log, "success", "Approved " + approved.Count + " of " + occurrences.Count + " detected occurrence(s) for future alignment.", 2);
            return approved;
        }

        private static void WriteSummary(AudioCensorScanSummary summary, Action<string, string, int> log)
        {
            Write(log, "info", "Audio censor scan complete.", 1);
            Write(log, "data", "Input MKV: " + summary.MkvPath, 1);
            Write(log, "data", "SRT: " + summary.SubtitlePath, 1);
            Write(log, "data", "Dictionary terms: " + summary.DictionaryTermCount, 1);
            Write(log, "data", "Detected profanity occurrences: " + summary.ScanResult.Occurrences.Count, 1);
            Write(log, "data", "Approved for future alignment: " + summary.ApprovedOccurrences.Count, 1);
            Write(log, "data", "Selected audio track index: " + summary.AudioTrackIndex, 1);
            Write(log, "data", "Future clean output path: " + summary.OutputMkvPath, 1);
            Write(log, "warning", "WhisperX alignment and FFmpeg mux execution are placeholders in this commit, so no media file was created.", 2);
        }

        private static void WriteLogFile(
            AudioCensorScanSummary summary,
            IList<string> commandLog,
            Action<string, string, int> log)
        {
            var lines = new List<string>
            {
                "BachFlixNfo Audio Censor Scan",
                "Created: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                "",
                "Input MKV: " + summary.MkvPath,
                "Subtitle: " + summary.SubtitlePath,
                "Dictionary: " + summary.ProfanityDictionaryPath,
                "Selected audio track index: " + summary.AudioTrackIndex,
                "Future output MKV: " + summary.OutputMkvPath,
                "",
                "FFmpeg: " + DescribeTool(summary.Ffmpeg),
                "FFprobe: " + DescribeTool(summary.Ffprobe),
                "WhisperX: " + DescribeTool(summary.WhisperX),
                "",
                "Commands executed:"
            };

            if (commandLog == null || commandLog.Count == 0)
            {
                lines.Add("  (none)");
            }
            else
            {
                foreach (string command in commandLog)
                    lines.Add("  " + command);
            }

            lines.Add("");
            lines.Add("Detected profanity occurrences:");

            if (summary.ScanResult.Occurrences.Count == 0)
            {
                lines.Add("  (none)");
            }
            else
            {
                foreach (ProfanityOccurrence occurrence in summary.ScanResult.Occurrences)
                {
                    bool approved = summary.ApprovedOccurrences.Any(o => o.ReviewNumber == occurrence.ReviewNumber);
                    lines.Add(string.Format(
                        "  {0}. {1} --> {2} | SRT #{3} | {4} | approved={5} | {6}",
                        occurrence.ReviewNumber,
                        FormatTime(occurrence.Cue.Start),
                        FormatTime(occurrence.Cue.End),
                        occurrence.SubtitleSequenceNumber,
                        occurrence.Word,
                        approved ? "yes" : "no",
                        occurrence.Cue.Text));
                }
            }

            string error;
            string logPath = global::BachFlixLog.WriteBachFlixLog(lines, "Audio Censor", "AudioCensor", out error);

            if (!string.IsNullOrWhiteSpace(logPath))
                Write(log, "success", "Audio censor scan log written: " + logPath, 2);
            else if (!string.IsNullOrWhiteSpace(error))
                Write(log, "warning", "Could not write audio censor scan log: " + error, 2);
        }

        private static List<string> GetDictionarySuggestions(string mkvPath)
        {
            var suggestions = new List<string>();
            string directory = Path.GetDirectoryName(mkvPath);
            string baseName = Path.GetFileNameWithoutExtension(mkvPath);

            if (!string.IsNullOrWhiteSpace(directory))
            {
                if (!string.IsNullOrWhiteSpace(baseName))
                    suggestions.Add(Path.Combine(directory, baseName + ".profanity.txt"));

                suggestions.Add(Path.Combine(directory, "profanity.txt"));
            }

            suggestions.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "profanity.txt"));

            return suggestions;
        }

        private static HashSet<int> ParseNumberList(string response)
        {
            var numbers = new HashSet<int>();

            if (string.IsNullOrWhiteSpace(response))
                return numbers;

            string[] parts = response.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                int value;
                if (int.TryParse(part.Trim(), out value))
                    numbers.Add(value);
            }

            return numbers;
        }

        private static string DescribeTool(ExternalToolResolution resolution)
        {
            if (resolution == null || !resolution.Found)
                return "not found";

            if (!string.IsNullOrWhiteSpace(resolution.VersionLine))
                return resolution.Path + " | " + resolution.VersionLine;

            return resolution.Path;
        }

        private static string CleanConsolePath(string value)
        {
            if (value == null)
                return "";

            return value.Trim().Trim('"');
        }

        private static string FormatTime(TimeSpan value)
        {
            return string.Format(
                "{0:D2}:{1:D2}:{2:D2},{3:D3}",
                (int)value.TotalHours,
                value.Minutes,
                value.Seconds,
                value.Milliseconds);
        }

        private static string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            if (value.Length <= maxLength)
                return value;

            return value.Substring(0, Math.Max(0, maxLength - 3)) + "...";
        }

        private static void Write(Action<string, string, int> log, string type, string message, int lines)
        {
            try
            {
                if (log != null)
                    log(type, message, lines);
                else
                    Console.WriteLine(message);
            }
            catch
            {
                Console.WriteLine(message);
            }
        }
    }

    public sealed class AudioCensorScanOptions
    {
        public string MkvPath { get; set; }
        public string SubtitlePath { get; set; }
        public string ProfanityDictionaryPath { get; set; }
        public int AudioTrackIndex { get; set; }
        public bool DetectExternalTools { get; set; }
        public string FfmpegPath { get; set; }
        public string FfprobePath { get; set; }
        public string WhisperXPath { get; set; }
        public Action<string, string, int> Log { get; set; }
        public IList<string> CommandLog { get; set; }

        public AudioCensorScanOptions()
        {
            MkvPath = "";
            SubtitlePath = "";
            ProfanityDictionaryPath = "";
            DetectExternalTools = true;
            FfmpegPath = "";
            FfprobePath = "";
            WhisperXPath = "";
            CommandLog = new List<string>();
        }
    }

    public sealed class AudioCensorScanSummary
    {
        public string MkvPath { get; set; }
        public string SubtitlePath { get; set; }
        public string ProfanityDictionaryPath { get; set; }
        public int DictionaryTermCount { get; set; }
        public int AudioTrackIndex { get; set; }
        public string OutputMkvPath { get; set; }
        public SubtitleScanResult ScanResult { get; set; }
        public List<ProfanityOccurrence> ApprovedOccurrences { get; set; }
        public ExternalToolResolution Ffmpeg { get; set; }
        public ExternalToolResolution Ffprobe { get; set; }
        public ExternalToolResolution WhisperX { get; set; }

        public AudioCensorScanSummary()
        {
            MkvPath = "";
            SubtitlePath = "";
            ProfanityDictionaryPath = "";
            OutputMkvPath = "";
            ScanResult = new SubtitleScanResult();
            ApprovedOccurrences = new List<ProfanityOccurrence>();
            Ffmpeg = new ExternalToolResolution { ToolName = "ffmpeg" };
            Ffprobe = new ExternalToolResolution { ToolName = "ffprobe" };
            WhisperX = new ExternalToolResolution { ToolName = "whisperx" };
        }
    }
}
