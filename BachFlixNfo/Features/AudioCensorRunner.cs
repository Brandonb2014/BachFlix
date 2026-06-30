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
            Write(log, "info", "=== AUDIO CENSOR DRY RUN ===", 1);
            Write(log, "warning", "This phase scans subtitles, prepares mute windows, and prints the FFmpeg plan. No MKV will be changed or created yet.", 2);

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

            var commandLog = new List<string>();
            AudioCensorScanSummary summary;

            try
            {
                summary = RunScan(new AudioCensorScanOptions
                {
                    MkvPath = mkvPath,
                    SubtitlePath = subtitlePath,
                    ProfanityDictionaryPath = dictionaryPath,
                    AudioTrackIndex = 0,
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

            summary.AudioTrackIndex = PromptForAudioTrackIndex(log, summary.AudioTracks);
            summary.ApprovedOccurrences = ReviewOccurrences(summary.ScanResult.Occurrences, log);
            summary.CensorPlan = BuildDryRunPlan(summary, commandLog, log);

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
            var audioTracks = new List<AudioTrackInfo>();

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

                if (ffprobe.Found)
                {
                    try
                    {
                        Write(options.Log, "info", "Reading audio tracks with FFprobe...", 1);
                        audioTracks = muxService.ProbeAudioTracks(ffprobe.Path, options.MkvPath, options.Log, commandLog);
                        Write(options.Log, "success", "Found " + audioTracks.Count + " audio track(s).", 1);
                    }
                    catch (Exception ex)
                    {
                        Write(options.Log, "warning", "Could not read audio tracks: " + ex.Message, 1);
                    }
                }
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
                AudioTracks = audioTracks,
                OutputMkvPath = muxService.BuildCleanOutputPath(options.MkvPath),
                ScanResult = scanResult,
                ApprovedOccurrences = new List<ProfanityOccurrence>(scanResult.Occurrences),
                Ffmpeg = ffmpeg,
                Ffprobe = ffprobe,
                WhisperX = whisperx
            };
        }
        private static AudioCensorPlan BuildDryRunPlan(
            AudioCensorScanSummary summary,
            IList<string> commandLog,
            Action<string, string, int> log)
        {
            var plan = new AudioCensorPlan
            {
                InputMkvPath = summary.MkvPath,
                OutputMkvPath = summary.OutputMkvPath,
                AudioTrackIndex = summary.AudioTrackIndex
            };

            if (summary.ApprovedOccurrences == null || summary.ApprovedOccurrences.Count == 0)
            {
                plan.Message = "No approved profanity occurrences; no mute filter or FFmpeg command is needed.";
                Write(log, "success", plan.Message, 2);
                return plan;
            }

            var muteSegments = new List<AudioMuteSegment>();

            if (summary.Ffmpeg != null && summary.Ffmpeg.Found && summary.WhisperX != null && summary.WhisperX.Found)
            {
                plan.AlignmentAttempted = true;
                try
                {
                    Write(log, "info", "Running WhisperX alignment for approved profanity occurrences...", 1);
                    var alignmentService = new WhisperAlignmentService();
                    WhisperAlignmentResult alignmentResult = alignmentService.AlignWords(new WhisperAlignmentRequest
                    {
                        MkvPath = summary.MkvPath,
                        SubtitlePath = summary.SubtitlePath,
                        AudioTrackIndex = summary.AudioTrackIndex,
                        OccurrencesToAlign = summary.ApprovedOccurrences,
                        FfmpegPath = summary.Ffmpeg.Path,
                        WhisperXPath = summary.WhisperX.Path,
                        Model = GetEnvOrDefault("AUDIO_CENSOR_WHISPERX_MODEL", "small"),
                        Device = GetEnvOrDefault("AUDIO_CENSOR_WHISPERX_DEVICE", "cpu"),
                        ComputeType = GetEnvOrDefault("AUDIO_CENSOR_WHISPERX_COMPUTE_TYPE", "int8"),
                        Language = GetEnvOrDefault("AUDIO_CENSOR_WHISPERX_LANGUAGE", "en"),
                        KeepWorkFiles = GetBooleanEnv("AUDIO_CENSOR_KEEP_WORK_FILES"),
                        Log = log,
                        CommandLog = commandLog
                    });

                    plan.AlignmentResult = alignmentResult;
                    plan.AlignmentSucceeded = alignmentResult.AlignedWords.Count > 0;
                    plan.AlignmentMessage = "Aligned " + alignmentResult.AlignedWords.Count + " of " + summary.ApprovedOccurrences.Count + " approved occurrence(s).";
                    Write(log, plan.AlignmentSucceeded ? "success" : "warning", plan.AlignmentMessage, 1);

                    var builder = new AudioMuteFilterBuilder();
                    muteSegments.AddRange(builder.CreateSegments(alignmentResult));

                    if (alignmentResult.UnalignedOccurrences.Count > 0)
                    {
                        plan.FallbackSegmentsUsed = true;
                        Write(log, "warning", "Using subtitle cue fallback for " + alignmentResult.UnalignedOccurrences.Count + " unaligned occurrence(s).", 1);
                        muteSegments.AddRange(CreateSubtitleFallbackSegments(alignmentResult.UnalignedOccurrences));
                    }
                }
                catch (Exception ex)
                {
                    plan.AlignmentMessage = "WhisperX alignment failed: " + ex.Message;
                    plan.FallbackSegmentsUsed = true;
                    Write(log, "warning", plan.AlignmentMessage, 1);
                    Write(log, "warning", "Using subtitle cue timing as the dry-run fallback.", 1);
                    muteSegments.AddRange(CreateSubtitleFallbackSegments(summary.ApprovedOccurrences));
                }
            }
            else
            {
                plan.FallbackSegmentsUsed = true;
                plan.AlignmentMessage = "WhisperX alignment was skipped because FFmpeg or WhisperX was not available.";
                Write(log, "warning", plan.AlignmentMessage, 1);
                Write(log, "warning", "Using subtitle cue timing as the dry-run fallback.", 1);
                muteSegments.AddRange(CreateSubtitleFallbackSegments(summary.ApprovedOccurrences));
            }

            plan.MuteSegments.AddRange(MergeSegments(muteSegments));
            if (plan.MuteSegments.Count == 0)
            {
                plan.Message = "No valid mute segments were produced; no FFmpeg command was generated.";
                Write(log, "warning", plan.Message, 2);
                return plan;
            }

            var muteFilterBuilder = new AudioMuteFilterBuilder();
            plan.AudioFilter = muteFilterBuilder.BuildMuteFilter(plan.MuteSegments);
            if (string.IsNullOrWhiteSpace(plan.AudioFilter))
            {
                plan.Message = "Mute filter was empty; no FFmpeg command was generated.";
                Write(log, "warning", plan.Message, 2);
                return plan;
            }

            if (summary.Ffmpeg == null || !summary.Ffmpeg.Found)
            {
                plan.Message = "FFmpeg was not found, so the mute filter was built but no command line was generated.";
                Write(log, "warning", plan.Message, 2);
                return plan;
            }

            var muxService = new FfmpegMuxService();
            plan.MuxPlan = muxService.BuildMuxPlan(new FfmpegMuxRequest
            {
                FfmpegPath = summary.Ffmpeg.Path,
                InputMkvPath = summary.MkvPath,
                OutputMkvPath = summary.OutputMkvPath,
                AudioTrackIndex = summary.AudioTrackIndex,
                AudioFilter = plan.AudioFilter,
                AudioCodec = "aac"
            });

            plan.FfmpegCommandLine = plan.MuxPlan.CommandLine;
            plan.MuxPlanBuilt = true;
            plan.Message = "Dry-run FFmpeg command generated. No media file was created.";

            Write(log, "success", "Built " + plan.MuteSegments.Count + " mute segment(s).", 1);
            Write(log, "info", "FFmpeg dry-run command:", 1);
            Write(log, "data", plan.FfmpegCommandLine, 2);

            return plan;
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

        private static int PromptForAudioTrackIndex(Action<string, string, int> log, IReadOnlyList<AudioTrackInfo> audioTracks)
        {
            if (audioTracks != null && audioTracks.Count > 0)
            {
                Write(log, "info", "Audio tracks:", 1);
                foreach (AudioTrackInfo track in audioTracks)
                    Write(log, "data", track.Describe(), 1);

                while (true)
                {
                    Write(log, "question", "Audio track index to censor. Press Enter for 0:", 1);
                    string input = (Console.ReadLine() ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(input))
                        return 0;

                    int selectedIndex;
                    if (int.TryParse(input, out selectedIndex) && audioTracks.Any(t => t.AudioTrackIndex == selectedIndex))
                        return selectedIndex;

                    Write(log, "error", "Please choose one of the listed audio track indexes.", 1);
                }
            }

            Write(log, "question", "Audio track index to censor later, 0-based audio track number. Press Enter for 0:", 1);
            string fallbackInput = (Console.ReadLine() ?? "").Trim();

            int audioTrackIndex;
            if (!int.TryParse(fallbackInput, out audioTrackIndex) || audioTrackIndex < 0)
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
            Write(log, "info", "Audio censor dry run complete.", 1);
            Write(log, "data", "Input MKV: " + summary.MkvPath, 1);
            Write(log, "data", "SRT: " + summary.SubtitlePath, 1);
            Write(log, "data", "Dictionary terms: " + summary.DictionaryTermCount, 1);
            Write(log, "data", "Detected profanity occurrences: " + summary.ScanResult.Occurrences.Count, 1);
            Write(log, "data", "Approved for censor plan: " + summary.ApprovedOccurrences.Count, 1);
            Write(log, "data", "Selected audio track index: " + summary.AudioTrackIndex, 1);
            Write(log, "data", "Future clean output path: " + summary.OutputMkvPath, 1);

            if (summary.CensorPlan != null)
            {
                Write(log, "data", "Mute segments: " + summary.CensorPlan.MuteSegments.Count, 1);
                Write(log, summary.CensorPlan.FallbackSegmentsUsed ? "warning" : "success", "Subtitle cue fallback used: " + (summary.CensorPlan.FallbackSegmentsUsed ? "yes" : "no"), 1);
                if (!string.IsNullOrWhiteSpace(summary.CensorPlan.Message))
                    Write(log, summary.CensorPlan.MuxPlanBuilt ? "success" : "warning", summary.CensorPlan.Message, 1);
            }

            Write(log, "warning", "Dry-run only: no media file was created.", 2);
        }
        private static void WriteLogFile(
            AudioCensorScanSummary summary,
            IList<string> commandLog,
            Action<string, string, int> log)
        {
            var lines = new List<string>
            {
                "BachFlixNfo Audio Censor Dry Run",
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
                "Audio tracks:"
            };

            if (summary.AudioTracks.Count == 0)
            {
                lines.Add("  (none read)");
            }
            else
            {
                foreach (AudioTrackInfo track in summary.AudioTracks)
                    lines.Add("  " + track.Describe());
            }

            lines.Add("");
            lines.Add("Commands executed:");

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

            WritePlanLogLines(summary.CensorPlan, lines);

            string error;
            string logPath = global::BachFlixLog.WriteBachFlixLog(lines, "Audio Censor", "AudioCensor", out error);

            if (!string.IsNullOrWhiteSpace(logPath))
                Write(log, "success", "Audio censor dry-run log written: " + logPath, 2);
            else if (!string.IsNullOrWhiteSpace(error))
                Write(log, "warning", "Could not write audio censor dry-run log: " + error, 2);
        }

        private static void WritePlanLogLines(AudioCensorPlan plan, List<string> lines)
        {
            lines.Add("");
            lines.Add("Dry-run censor plan:");

            if (plan == null)
            {
                lines.Add("  (not built)");
                return;
            }

            lines.Add("  Alignment attempted: " + (plan.AlignmentAttempted ? "yes" : "no"));
            lines.Add("  Alignment succeeded: " + (plan.AlignmentSucceeded ? "yes" : "no"));
            lines.Add("  Subtitle fallback used: " + (plan.FallbackSegmentsUsed ? "yes" : "no"));
            lines.Add("  Mux plan built: " + (plan.MuxPlanBuilt ? "yes" : "no"));
            lines.Add("  Message: " + plan.Message);
            if (!string.IsNullOrWhiteSpace(plan.AlignmentMessage))
                lines.Add("  Alignment message: " + plan.AlignmentMessage);

            if (plan.AlignmentResult != null)
            {
                lines.Add("  Transcript words: " + plan.AlignmentResult.TranscriptWordCount);
                lines.Add("  Aligned words: " + plan.AlignmentResult.AlignedWords.Count);
                lines.Add("  Unaligned occurrences: " + plan.AlignmentResult.UnalignedOccurrences.Count);
                if (plan.AlignmentResult.UnalignedOccurrences.Count > 0)
                {
                    lines.Add("  Unaligned review numbers: " + string.Join(", ", plan.AlignmentResult.UnalignedOccurrences.Select(o => o.ReviewNumber)));
                }
                lines.Add("  Work directory: " + plan.AlignmentResult.WorkDirectory);
                lines.Add("  Transcript JSON: " + plan.AlignmentResult.TranscriptJsonPath);
                lines.Add("  Temporary files cleaned: " + (plan.AlignmentResult.TemporaryFilesCleaned ? "yes" : "no"));
            }

            lines.Add("");
            lines.Add("Mute segments:");
            if (plan.MuteSegments.Count == 0)
            {
                lines.Add("  (none)");
            }
            else
            {
                foreach (AudioMuteSegment segment in plan.MuteSegments)
                {
                    string source = segment.SourceOccurrence == null ? "fallback" : "review #" + segment.SourceOccurrence.ReviewNumber;
                    if (segment.SourceOccurrence != null)
                        source += segment.IsFallback ? " (subtitle fallback)" : " (aligned)";
                    lines.Add("  " + FormatTime(segment.Start) + " --> " + FormatTime(segment.End) + " | " + source);
                }
            }

            lines.Add("");
            lines.Add("Audio filter:");
            lines.Add(string.IsNullOrWhiteSpace(plan.AudioFilter) ? "  (none)" : "  " + plan.AudioFilter);

            lines.Add("");
            lines.Add("FFmpeg dry-run command:");
            lines.Add(string.IsNullOrWhiteSpace(plan.FfmpegCommandLine) ? "  (none)" : "  " + plan.FfmpegCommandLine);
        }
        private static List<AudioMuteSegment> CreateSubtitleFallbackSegments(IEnumerable<ProfanityOccurrence> occurrences)
        {
            var segments = new List<AudioMuteSegment>();
            if (occurrences == null)
                return segments;

            foreach (ProfanityOccurrence occurrence in occurrences)
            {
                if (occurrence == null || occurrence.Cue == null || occurrence.Cue.End <= occurrence.Cue.Start)
                    continue;

                segments.Add(new AudioMuteSegment
                {
                    Start = occurrence.Cue.Start,
                    End = occurrence.Cue.End,
                    SourceOccurrence = occurrence,
                    IsFallback = true
                });
            }

            return segments;
        }

        private static List<AudioMuteSegment> MergeSegments(IEnumerable<AudioMuteSegment> segments)
        {
            var ordered = (segments ?? Enumerable.Empty<AudioMuteSegment>())
                .Where(s => s != null && s.End > s.Start)
                .OrderBy(s => s.Start)
                .ThenBy(s => s.End)
                .ToList();

            var merged = new List<AudioMuteSegment>();
            foreach (AudioMuteSegment segment in ordered)
            {
                AudioMuteSegment last = merged.LastOrDefault();
                if (last != null && segment.Start <= last.End)
                {
                    if (segment.End > last.End)
                        last.End = segment.End;
                    last.IsFallback = last.IsFallback || segment.IsFallback;
                    continue;
                }

                merged.Add(new AudioMuteSegment
                {
                    Start = segment.Start,
                    End = segment.End,
                    SourceOccurrence = segment.SourceOccurrence,
                    IsFallback = segment.IsFallback
                });
            }

            return merged;
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

            return value
                .Trim()
                .Trim('"')
                .Trim('\uFEFF', '\u200B', '\u200E', '\u200F')
                .Trim()
                .Trim('"');
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

        private static string GetEnvOrDefault(string name, string fallback)
        {
            string value = Environment.GetEnvironmentVariable(name);
            return string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        }

        private static bool GetBooleanEnv(string name)
        {
            string value = Environment.GetEnvironmentVariable(name);
            if (string.IsNullOrWhiteSpace(value))
                return false;

            value = value.Trim();
            return value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("y", StringComparison.OrdinalIgnoreCase);
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
        public List<AudioTrackInfo> AudioTracks { get; set; }
        public string OutputMkvPath { get; set; }
        public SubtitleScanResult ScanResult { get; set; }
        public List<ProfanityOccurrence> ApprovedOccurrences { get; set; }
        public AudioCensorPlan CensorPlan { get; set; }
        public ExternalToolResolution Ffmpeg { get; set; }
        public ExternalToolResolution Ffprobe { get; set; }
        public ExternalToolResolution WhisperX { get; set; }

        public AudioCensorScanSummary()
        {
            MkvPath = "";
            SubtitlePath = "";
            ProfanityDictionaryPath = "";
            OutputMkvPath = "";
            AudioTracks = new List<AudioTrackInfo>();
            ScanResult = new SubtitleScanResult();
            ApprovedOccurrences = new List<ProfanityOccurrence>();
            CensorPlan = new AudioCensorPlan();
            Ffmpeg = new ExternalToolResolution { ToolName = "ffmpeg" };
            Ffprobe = new ExternalToolResolution { ToolName = "ffprobe" };
            WhisperX = new ExternalToolResolution { ToolName = "whisperx" };
        }
    }

    public sealed class AudioCensorPlan
    {
        public string InputMkvPath { get; set; }
        public string OutputMkvPath { get; set; }
        public int AudioTrackIndex { get; set; }
        public bool AlignmentAttempted { get; set; }
        public bool AlignmentSucceeded { get; set; }
        public bool FallbackSegmentsUsed { get; set; }
        public bool MuxPlanBuilt { get; set; }
        public string AlignmentMessage { get; set; }
        public string Message { get; set; }
        public WhisperAlignmentResult AlignmentResult { get; set; }
        public List<AudioMuteSegment> MuteSegments { get; private set; }
        public string AudioFilter { get; set; }
        public FfmpegMuxPlan MuxPlan { get; set; }
        public string FfmpegCommandLine { get; set; }

        public AudioCensorPlan()
        {
            InputMkvPath = "";
            OutputMkvPath = "";
            AlignmentMessage = "";
            Message = "";
            MuteSegments = new List<AudioMuteSegment>();
            AudioFilter = "";
            MuxPlan = new FfmpegMuxPlan();
            FfmpegCommandLine = "";
        }
    }
}