using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace BachFlixNfo.Features
{
    public static class AudioCensorRunner
    {
        public static void RunInteractive(Action<string, string, int> log)
        {
            Write(log, "info", "=== AUDIO CENSOR: WHISPERX WORD TIMESTAMPS ===", 1);
            Write(log, "warning", "WhisperX word timestamps are the source of truth. Existing subtitles are optional comparison and review context only.", 2);

            string mkvPath = PromptForExistingMkv(log);
            if (string.IsNullOrWhiteSpace(mkvPath))
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
                    ProfanityDictionaryPath = dictionaryPath,
                    DetectExternalTools = true,
                    PromptForMissingEnglishAudio = true,
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

            summary.ApprovedTranscriptHits = ReviewTranscriptHits(summary.TranscriptProfanityHits, summary.SelectedAudioTracks, log);
            summary.CensorPlan = BuildCensorPlan(summary, commandLog, log);

            WriteSummary(summary, log);
            WriteLogFile(summary, commandLog, log);

            if (summary.CensorPlan != null && summary.CensorPlan.MuxPlanBuilt)
            {
                Write(log, "question", "Create the _Clean MKV now? Type y to run FFmpeg, or press Enter to leave the plan only:", 1);
                string response = (Console.ReadLine() ?? "").Trim();
                if (response.Equals("y", StringComparison.OrdinalIgnoreCase) || response.Equals("yes", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        new FfmpegMuxService().MuxCensoredMedia(summary.CensorPlan.MuxPlan, log, commandLog);
                        Write(log, "success", "Created clean MKV: " + summary.CensorPlan.OutputMkvPath, 2);
                    }
                    catch (Exception ex)
                    {
                        Write(log, "error", "FFmpeg mux failed.", 1);
                        Write(log, "harderror", ex.Message, 2);
                    }
                }
            }
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
            IWordTranscriptionService transcriptionService = new WhisperXTranscriptionService();

            ExternalToolResolution ffmpeg = new ExternalToolResolution { ToolName = "ffmpeg" };
            ExternalToolResolution ffprobe = new ExternalToolResolution { ToolName = "ffprobe" };
            ExternalToolResolution transcriptionTool = new ExternalToolResolution { ToolName = "whisperx" };
            var media = new MediaProbeResult { InputPath = options.MkvPath };

            if (options.DetectExternalTools)
            {
                Write(options.Log, "info", "Detecting FFmpeg, FFprobe, and WhisperX executables...", 1);
                ffmpeg = muxService.ResolveFfmpeg(options.FfmpegPath, options.Log, commandLog);
                ffprobe = muxService.ResolveFfprobe(options.FfprobePath, options.Log, commandLog);
                transcriptionTool = transcriptionService.ResolveTranscriptionTool(options.WhisperXPath, options.Log, commandLog);

                if (!ffmpeg.Found)
                    throw new InvalidOperationException(ffmpeg.Message);
                if (!ffprobe.Found)
                    throw new InvalidOperationException(ffprobe.Message);
                if (!transcriptionTool.Found)
                    throw new InvalidOperationException(transcriptionTool.Message);

                Write(options.Log, "info", "Detecting media streams with FFprobe...", 1);
                media = muxService.ProbeMedia(ffprobe.Path, options.MkvPath, options.Log, commandLog);
                Write(options.Log, "success", "Found " + media.AudioTracks.Count + " audio, " + media.SubtitleTracks.Count + " subtitle, " + media.AttachmentStreams.Count + " attachment stream(s), and " + media.ChapterCount + " chapter(s).", 1);
            }

            if (media.AudioTracks.Count == 0)
                throw new InvalidOperationException("No audio streams were detected in the MKV.");

            List<AudioTrackInfo> selectedAudioTracks = SelectEnglishAudioTracks(media.AudioTracks, options.SelectedAudioTrackIndexes);
            if (selectedAudioTracks.Count == 0 && options.PromptForMissingEnglishAudio)
                selectedAudioTracks = PromptForAudioTrackIndexes(options.Log, media.AudioTracks);

            if (selectedAudioTracks.Count == 0)
                throw new InvalidOperationException("No English audio track could be confidently identified. Choose tracks manually before running censorship.");

            AudioTrackInfo primaryTrack = SelectPrimaryTranscriptionTrack(selectedAudioTracks);
            Write(options.Log, "success", "Selected " + selectedAudioTracks.Count + " English audio track(s). WhisperX source track: " + primaryTrack.AudioTrackIndex + ".", 1);

            Write(options.Log, "info", "Loading profanity dictionary...", 1);
            ProfanityDictionary dictionary = ProfanityDictionary.LoadFromFile(options.ProfanityDictionaryPath);
            Write(options.Log, "success", "Loaded " + dictionary.Count + " profanity terms.", 1);

            string outputPath = muxService.BuildCleanOutputPath(options.MkvPath);
            var summary = new AudioCensorScanSummary
            {
                MkvPath = options.MkvPath,
                ProfanityDictionaryPath = options.ProfanityDictionaryPath,
                DictionaryTermCount = dictionary.Count,
                MediaProbe = media,
                AudioTracks = media.AudioTracks,
                SelectedAudioTracks = selectedAudioTracks,
                PrimaryTranscriptionTrack = primaryTrack,
                OutputMkvPath = outputPath,
                Ffmpeg = ffmpeg,
                Ffprobe = ffprobe,
                TranscriptionTool = transcriptionTool
            };

            var transcriptScanner = new TranscriptProfanityScanner();
            var reusableTracks = selectedAudioTracks
                .Where(track => track.AudioTrackIndex == primaryTrack.AudioTrackIndex || !NeedsSeparateTranscript(primaryTrack, track, media.Duration))
                .OrderBy(track => track.AudioTrackIndex)
                .ToList();
            var separateTracks = selectedAudioTracks
                .Where(track => reusableTracks.All(r => r.AudioTrackIndex != track.AudioTrackIndex))
                .OrderBy(track => track.AudioTrackIndex)
                .ToList();

            WordTranscriptionResult primaryTranscript = RunTranscription(options, transcriptionService, ffmpeg, transcriptionTool, primaryTrack.AudioTrackIndex, true, commandLog);
            summary.TranscriptionResults.Add(primaryTrack.AudioTrackIndex, primaryTranscript);
            TranscriptProfanityScanResult primaryScan = transcriptScanner.Scan(primaryTranscript, dictionary, primaryTrack.AudioTrackIndex, "audio " + primaryTrack.AudioTrackIndex);
            foreach (TranscriptProfanityHit hit in primaryScan.Hits)
            {
                hit.AppliesToAudioTrackIndexes.Clear();
                foreach (AudioTrackInfo track in reusableTracks)
                    hit.AppliesToAudioTrackIndexes.Add(track.AudioTrackIndex);
                summary.TranscriptProfanityHits.Add(hit);
            }

            foreach (AudioTrackInfo track in separateTracks)
            {
                Write(options.Log, "warning", "Audio track " + track.AudioTrackIndex + " looks different enough to transcribe separately: " + track.Describe(), 1);
                WordTranscriptionResult transcript = RunTranscription(options, transcriptionService, ffmpeg, transcriptionTool, track.AudioTrackIndex, false, commandLog);
                summary.TranscriptionResults.Add(track.AudioTrackIndex, transcript);
                TranscriptProfanityScanResult scan = transcriptScanner.Scan(transcript, dictionary, track.AudioTrackIndex, "audio " + track.AudioTrackIndex);
                summary.TranscriptProfanityHits.AddRange(scan.Hits);
            }

            RenumberHits(summary.TranscriptProfanityHits);
            Write(options.Log, "success", "WhisperX transcript scan found " + summary.TranscriptProfanityHits.Count + " profanity hit(s).", 1);

            ProcessSubtitles(summary, dictionary, ffmpeg, options, commandLog);
            summary.ApprovedTranscriptHits = new List<TranscriptProfanityHit>(summary.TranscriptProfanityHits.Where(h => h.Approved));
            return summary;
        }

        private static WordTranscriptionResult RunTranscription(
            AudioCensorScanOptions options,
            IWordTranscriptionService transcriptionService,
            ExternalToolResolution ffmpeg,
            ExternalToolResolution transcriptionTool,
            int audioTrackIndex,
            bool canonicalOutputNames,
            IList<string> commandLog)
        {
            Write(options.Log, "info", "Running WhisperX on audio track " + audioTrackIndex + "...", 1);
            return transcriptionService.Transcribe(new WordTranscriptionRequest
            {
                MediaPath = options.MkvPath,
                AudioTrackIndex = audioTrackIndex,
                FfmpegPath = ffmpeg.Path,
                TranscriptionToolPath = transcriptionTool.Path,
                Model = GetEnvOrDefault("AUDIO_CENSOR_WHISPERX_MODEL", "small"),
                Device = GetEnvOrDefault("AUDIO_CENSOR_WHISPERX_DEVICE", "cpu"),
                ComputeType = GetEnvOrDefault("AUDIO_CENSOR_WHISPERX_COMPUTE_TYPE", "int8"),
                Language = GetEnvOrDefault("AUDIO_CENSOR_WHISPERX_LANGUAGE", "en"),
                KeepWorkFiles = GetBooleanEnv("AUDIO_CENSOR_KEEP_WORK_FILES"),
                UseCanonicalOutputNames = canonicalOutputNames,
                Log = options.Log,
                CommandLog = commandLog
            });
        }

        private static void ProcessSubtitles(
            AudioCensorScanSummary summary,
            ProfanityDictionary dictionary,
            ExternalToolResolution ffmpeg,
            AudioCensorScanOptions options,
            IList<string> commandLog)
        {
            var detector = new SubtitleSourceDetector();
            summary.SubtitleSources = detector.DetectAllSubtitleSources(summary.MediaProbe, summary.MkvPath);
            summary.EnglishSubtitleSources = summary.SubtitleSources.Where(s => s.IsEnglish).ToList();

            if (summary.EnglishSubtitleSources.Count == 0)
            {
                Write(options.Log, "warning", "No English subtitle sources were detected. Continuing with WhisperX-only censorship.", 1);
                return;
            }

            Write(options.Log, "info", "Detected " + summary.EnglishSubtitleSources.Count + " English subtitle source(s) for comparison/review context.", 1);
            string subtitleWorkDirectory = Path.Combine(Path.GetTempPath(), "BachFlixNfo-AudioCensor-Subtitles-" + Guid.NewGuid().ToString("N"));
            summary.SubtitleWorkDirectory = subtitleWorkDirectory;

            new EmbeddedSubtitleExtractor().ExtractEmbeddedSubtitles(
                summary.EnglishSubtitleSources,
                summary.MkvPath,
                ffmpeg.Path,
                subtitleWorkDirectory,
                options.Log,
                commandLog);

            var subtitleScanner = new SubtitleProfanityService();
            var comparisonService = new SubtitleComparisonService();
            foreach (SubtitleSourceInfo source in summary.EnglishSubtitleSources)
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(source.ScanError))
                        throw new InvalidOperationException(source.ScanError);

                    source.ScanResult = subtitleScanner.Scan(source, dictionary);
                    source.Comparison = comparisonService.Compare(source, summary.TranscriptProfanityHits, source.ScanResult);
                    Write(options.Log, "data", BuildSubtitleCoverageLine(source), 1);
                }
                catch (Exception ex)
                {
                    source.ScanError = ex.Message;
                    Write(options.Log, "warning", "Subtitle comparison skipped for " + source.Describe() + ": " + ex.Message, 1);
                }
            }
        }

        private static AudioCensorPlan BuildCensorPlan(
            AudioCensorScanSummary summary,
            IList<string> commandLog,
            Action<string, string, int> log)
        {
            var plan = new AudioCensorPlan
            {
                InputMkvPath = summary.MkvPath,
                OutputMkvPath = summary.OutputMkvPath
            };

            List<TranscriptProfanityHit> approvedHits = (summary.ApprovedTranscriptHits ?? new List<TranscriptProfanityHit>())
                .Where(h => h != null && h.Approved)
                .ToList();

            if (approvedHits.Count == 0)
            {
                plan.Message = "No approved WhisperX profanity hits; no clean tracks were generated.";
                Write(log, "success", plan.Message, 2);
                return plan;
            }

            var muteBuilder = new AudioMuteFilterBuilder();
            foreach (AudioTrackInfo track in summary.SelectedAudioTracks.OrderBy(t => t.AudioTrackIndex))
            {
                List<TranscriptProfanityHit> hitsForTrack = approvedHits
                    .Where(h => h.AppliesToAudioTrackIndexes.Contains(track.AudioTrackIndex))
                    .ToList();
                List<AudioMuteSegment> segments = MergeSegments(muteBuilder.CreateSegments(hitsForTrack));
                string filter = muteBuilder.BuildMuteFilter(segments);
                if (string.IsNullOrWhiteSpace(filter))
                    continue;

                var trackPlan = new AudioTrackCensorPlan
                {
                    Track = track,
                    AudioFilter = filter
                };
                trackPlan.MuteSegments.AddRange(segments);
                plan.AudioTrackPlans.Add(trackPlan);
            }

            string replacement = GetEnvOrDefault("AUDIO_CENSOR_SUBTITLE_REPLACEMENT", "asterisks");
            var subtitleWriter = new CensoredSubtitleWriter();
            foreach (SubtitleSourceInfo subtitle in summary.EnglishSubtitleSources)
            {
                if (!string.IsNullOrWhiteSpace(subtitle.ScanError))
                    continue;

                try
                {
                    CensoredSubtitleOutput output = subtitleWriter.WriteCensoredSubtitle(subtitle, approvedHits, summary.MkvPath, replacement);
                    if (output != null)
                        plan.CensoredSubtitleOutputs.Add(output);
                }
                catch (Exception ex)
                {
                    Write(log, "warning", "Could not create censored subtitle for " + subtitle.Describe() + ": " + ex.Message, 1);
                }
            }

            if (plan.AudioTrackPlans.Count == 0 && plan.CensoredSubtitleOutputs.Count == 0)
            {
                plan.Message = "Approved hits existed, but no valid clean audio or subtitle outputs could be generated.";
                Write(log, "warning", plan.Message, 2);
                return plan;
            }

            var muxService = new FfmpegMuxService();
            var audioOutputs = plan.AudioTrackPlans.Select(trackPlan => new AudioTrackCensorOutput
            {
                SourceAudioTrackIndex = trackPlan.Track.AudioTrackIndex,
                SourceTrack = trackPlan.Track,
                AudioFilter = trackPlan.AudioFilter,
                AudioCodec = GetEnvOrDefault("AUDIO_CENSOR_CENSORED_AUDIO_CODEC", "flac"),
                Title = BuildCleanAudioTitle(trackPlan.Track),
                IsDefault = true
            }).ToList();

            plan.MuxPlan = muxService.BuildMuxPlan(new FfmpegMuxRequest
            {
                FfmpegPath = summary.Ffmpeg.Path,
                InputMkvPath = summary.MkvPath,
                OutputMkvPath = summary.OutputMkvPath,
                MediaProbe = summary.MediaProbe,
                CensoredAudioTracks = audioOutputs,
                CensoredSubtitleTracks = plan.CensoredSubtitleOutputs
            });
            plan.FfmpegCommandLine = plan.MuxPlan.CommandLine;
            plan.MuxPlanBuilt = true;
            plan.Message = "FFmpeg plan generated. Original MKV and external subtitle files remain untouched.";

            Write(log, "success", "Built clean outputs for " + plan.AudioTrackPlans.Count + " audio track(s) and " + plan.CensoredSubtitleOutputs.Count + " subtitle track(s).", 1);
            Write(log, "info", "FFmpeg command:", 1);
            Write(log, "data", plan.FfmpegCommandLine, 2);

            return plan;
        }

        private static List<AudioTrackInfo> SelectEnglishAudioTracks(IReadOnlyList<AudioTrackInfo> tracks, IReadOnlyList<int> explicitTrackIndexes)
        {
            if (tracks == null)
                return new List<AudioTrackInfo>();

            if (explicitTrackIndexes != null && explicitTrackIndexes.Count > 0)
            {
                HashSet<int> selected = new HashSet<int>(explicitTrackIndexes);
                return tracks.Where(t => selected.Contains(t.AudioTrackIndex)).OrderBy(t => t.AudioTrackIndex).ToList();
            }

            List<AudioTrackInfo> english = tracks
                .Where(t => SubtitleSourceDetector.IsEnglishLanguage(t.Language))
                .OrderBy(t => t.AudioTrackIndex)
                .ToList();

            if (english.Count > 0)
                return english;

            if (tracks.Count == 1 && string.IsNullOrWhiteSpace(tracks[0].Language))
                return new List<AudioTrackInfo> { tracks[0] };

            return new List<AudioTrackInfo>();
        }

        private static AudioTrackInfo SelectPrimaryTranscriptionTrack(IReadOnlyList<AudioTrackInfo> selectedTracks)
        {
            return selectedTracks
                .OrderByDescending(t => t.IsDefault)
                .ThenByDescending(t => t.Channels)
                .ThenByDescending(t => CodecQualityScore(t.CodecName))
                .ThenBy(t => t.AudioTrackIndex)
                .First();
        }

        private static int CodecQualityScore(string codecName)
        {
            string value = (codecName ?? "").Trim().ToLowerInvariant();
            if (value.Contains("truehd") || value.Contains("flac") || value.Contains("pcm"))
                return 50;
            if (value.Contains("dts") || value.Contains("eac3"))
                return 40;
            if (value.Contains("ac3"))
                return 30;
            if (value.Contains("aac"))
                return 20;
            return 10;
        }

        private static bool NeedsSeparateTranscript(AudioTrackInfo primaryTrack, AudioTrackInfo candidate, TimeSpan mediaDuration)
        {
            if (candidate.AudioTrackIndex == primaryTrack.AudioTrackIndex)
                return false;

            string title = (candidate.Title ?? "").ToLowerInvariant();
            if (title.Contains("commentary") || title.Contains("descriptive") || title.Contains("description") || title.Contains("narration"))
                return true;

            TimeSpan primaryDuration = primaryTrack.Duration > TimeSpan.Zero ? primaryTrack.Duration : mediaDuration;
            TimeSpan candidateDuration = candidate.Duration > TimeSpan.Zero ? candidate.Duration : mediaDuration;
            if (primaryDuration > TimeSpan.Zero && candidateDuration > TimeSpan.Zero)
                return Math.Abs((primaryDuration - candidateDuration).TotalSeconds) > 2;

            return false;
        }

        private static List<AudioTrackInfo> PromptForAudioTrackIndexes(Action<string, string, int> log, IReadOnlyList<AudioTrackInfo> audioTracks)
        {
            Write(log, "warning", "No English audio track could be confidently identified.", 1);
            Write(log, "info", "Audio tracks:", 1);
            foreach (AudioTrackInfo track in audioTracks)
                Write(log, "data", track.Describe(), 1);

            while (true)
            {
                Write(log, "question", "Enter comma-separated audio track indexes to censor, or 0 to cancel:", 1);
                string input = (Console.ReadLine() ?? "").Trim();
                if (input == "0")
                    return new List<AudioTrackInfo>();

                HashSet<int> selected = ParseNumberList(input);
                List<AudioTrackInfo> tracks = audioTracks.Where(t => selected.Contains(t.AudioTrackIndex)).OrderBy(t => t.AudioTrackIndex).ToList();
                if (tracks.Count > 0)
                    return tracks;

                Write(log, "error", "Please choose at least one listed audio track index.", 1);
            }
        }

        private static List<TranscriptProfanityHit> ReviewTranscriptHits(
            IReadOnlyList<TranscriptProfanityHit> hits,
            IReadOnlyList<AudioTrackInfo> selectedAudioTracks,
            Action<string, string, int> log)
        {
            var editableHits = (hits ?? new List<TranscriptProfanityHit>()).ToList();
            if (editableHits.Count == 0)
            {
                Write(log, "success", "Review list is empty. No profanity was detected in the WhisperX transcript.", 2);
                return new List<TranscriptProfanityHit>();
            }

            Write(log, "info", "Review WhisperX profanity hits:", 1);
            foreach (TranscriptProfanityHit hit in editableHits)
                Write(log, "data", DescribeHit(hit), 1);

            Write(log, "question", "Enter comma-separated review numbers to skip, or press Enter to keep all:", 1);
            HashSet<int> excludedNumbers = ParseNumberList((Console.ReadLine() ?? "").Trim());
            foreach (TranscriptProfanityHit hit in editableHits)
            {
                if (excludedNumbers.Contains(hit.ReviewNumber))
                    hit.Approved = false;
            }

            Write(log, "question", "Optional padding overrides as review=beforeMs/afterMs, comma-separated. Press Enter to keep defaults:", 1);
            ApplyPaddingOverrides(editableHits, (Console.ReadLine() ?? "").Trim(), log);

            Write(log, "question", "Optional manual additions as start-end=word[@audioTrackIndexes], separated by semicolons. Press Enter for none:", 1);
            editableHits.AddRange(ParseManualHits((Console.ReadLine() ?? "").Trim(), selectedAudioTracks, log));
            RenumberHits(editableHits);

            List<TranscriptProfanityHit> approved = editableHits.Where(h => h.Approved).ToList();
            Write(log, "success", "Approved " + approved.Count + " of " + editableHits.Count + " WhisperX/manual hit(s).", 2);
            return approved;
        }

        private static void ApplyPaddingOverrides(List<TranscriptProfanityHit> hits, string response, Action<string, string, int> log)
        {
            if (string.IsNullOrWhiteSpace(response))
                return;

            Dictionary<int, TranscriptProfanityHit> byReviewNumber = hits.ToDictionary(h => h.ReviewNumber);
            foreach (string part in response.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] pieces = part.Split('=');
                if (pieces.Length != 2)
                    continue;

                int reviewNumber;
                if (!int.TryParse(pieces[0].Trim(), out reviewNumber) || !byReviewNumber.ContainsKey(reviewNumber))
                    continue;

                string[] padding = pieces[1].Split('/');
                if (padding.Length != 2)
                    continue;

                int beforeMs;
                int afterMs;
                if (int.TryParse(padding[0].Trim(), out beforeMs) && int.TryParse(padding[1].Trim(), out afterMs))
                {
                    byReviewNumber[reviewNumber].PaddingBefore = TimeSpan.FromMilliseconds(Math.Max(0, beforeMs));
                    byReviewNumber[reviewNumber].PaddingAfter = TimeSpan.FromMilliseconds(Math.Max(0, afterMs));
                }
            }
        }

        private static IEnumerable<TranscriptProfanityHit> ParseManualHits(string response, IReadOnlyList<AudioTrackInfo> selectedAudioTracks, Action<string, string, int> log)
        {
            var hits = new List<TranscriptProfanityHit>();
            if (string.IsNullOrWhiteSpace(response))
                return hits;

            List<int> defaultTrackIndexes = (selectedAudioTracks ?? new List<AudioTrackInfo>()).Select(t => t.AudioTrackIndex).ToList();
            foreach (string rawItem in response.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string item = rawItem.Trim();
                string[] assignment = item.Split('=');
                if (assignment.Length != 2)
                    continue;

                string[] range = assignment[0].Split('-');
                if (range.Length != 2)
                    continue;

                TimeSpan start;
                TimeSpan end;
                if (!TryParseReviewTime(range[0].Trim(), out start) || !TryParseReviewTime(range[1].Trim(), out end) || end <= start)
                    continue;

                string word = assignment[1].Trim();
                List<int> trackIndexes = new List<int>(defaultTrackIndexes);
                int atIndex = word.IndexOf('@');
                if (atIndex >= 0)
                {
                    trackIndexes = ParseNumberList(word.Substring(atIndex + 1)).OrderBy(i => i).ToList();
                    word = word.Substring(0, atIndex).Trim();
                }

                if (string.IsNullOrWhiteSpace(word) || trackIndexes.Count == 0)
                    continue;

                var hit = new TranscriptProfanityHit
                {
                    Word = word,
                    DictionaryTerm = word,
                    Start = start,
                    End = end,
                    SourceLabel = "manual",
                    SourceAudioTrackIndex = trackIndexes[0],
                    IsManual = true,
                    Approved = true
                };
                foreach (int trackIndex in trackIndexes)
                    hit.AppliesToAudioTrackIndexes.Add(trackIndex);
                hits.Add(hit);
            }

            if (hits.Count > 0)
                Write(log, "success", "Added " + hits.Count + " manual mute hit(s).", 1);

            return hits;
        }

        private static bool TryParseReviewTime(string value, out TimeSpan time)
        {
            time = TimeSpan.Zero;
            if (string.IsNullOrWhiteSpace(value))
                return false;

            value = value.Trim().Replace(',', '.');
            double seconds;
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out seconds))
            {
                time = TimeSpan.FromSeconds(seconds);
                return seconds >= 0;
            }

            string[] parts = value.Split(':');
            if (parts.Length < 2 || parts.Length > 3)
                return false;

            int hours = 0;
            int minutes;
            decimal secondPart;
            if (parts.Length == 3)
            {
                if (!int.TryParse(parts[0], out hours) || !int.TryParse(parts[1], out minutes) || !decimal.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out secondPart))
                    return false;
            }
            else
            {
                if (!int.TryParse(parts[0], out minutes) || !decimal.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out secondPart))
                    return false;
            }

            time = TimeSpan.FromHours(hours) + TimeSpan.FromMinutes(minutes) + TimeSpan.FromSeconds((double)secondPart);
            return true;
        }

        private static string DescribeHit(TranscriptProfanityHit hit)
        {
            string tracks = hit.AppliesToAudioTrackIndexes.Count == 0
                ? "(none)"
                : string.Join(",", hit.AppliesToAudioTrackIndexes.OrderBy(i => i));
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}. {1} --> {2} | audio {3} | \"{4}\" (dictionary: {5}, padding: {6}ms/{7}ms)",
                hit.ReviewNumber,
                FormatTime(hit.Start),
                FormatTime(hit.End),
                tracks,
                hit.Word,
                hit.DictionaryTerm,
                (int)hit.PaddingBefore.TotalMilliseconds,
                (int)hit.PaddingAfter.TotalMilliseconds);
        }

        private static void WriteSummary(AudioCensorScanSummary summary, Action<string, string, int> log)
        {
            Write(log, "info", "Audio censor processing plan complete.", 1);
            Write(log, "data", "Input MKV: " + summary.MkvPath, 1);
            Write(log, "data", "Dictionary terms: " + summary.DictionaryTermCount, 1);
            Write(log, "data", "Audio streams detected: " + summary.AudioTracks.Count, 1);
            Write(log, "data", "Selected English audio tracks: " + string.Join(", ", summary.SelectedAudioTracks.Select(t => t.AudioTrackIndex)), 1);
            Write(log, "data", "WhisperX source audio track: " + summary.PrimaryTranscriptionTrack.AudioTrackIndex, 1);
            Write(log, "data", "WhisperX JSON: " + GetPrimaryTranscriptPath(summary, true), 1);
            Write(log, "data", "WhisperX SRT: " + GetPrimaryTranscriptPath(summary, false), 1);
            Write(log, "data", "WhisperX profanity hits: " + summary.TranscriptProfanityHits.Count, 1);
            Write(log, "data", "Approved hits: " + summary.ApprovedTranscriptHits.Count, 1);
            Write(log, "data", "English subtitles checked: " + summary.EnglishSubtitleSources.Count, 1);
            Write(log, "data", "Clean output path: " + summary.OutputMkvPath, 1);

            if (summary.CensorPlan != null)
            {
                Write(log, "data", "Clean audio tracks to add: " + summary.CensorPlan.AudioTrackPlans.Count, 1);
                Write(log, "data", "Clean subtitle tracks to add: " + summary.CensorPlan.CensoredSubtitleOutputs.Count, 1);
                if (!string.IsNullOrWhiteSpace(summary.CensorPlan.Message))
                    Write(log, summary.CensorPlan.MuxPlanBuilt ? "success" : "warning", summary.CensorPlan.Message, 1);
            }
        }

        private static string GetPrimaryTranscriptPath(AudioCensorScanSummary summary, bool json)
        {
            WordTranscriptionResult result;
            if (summary.TranscriptionResults.TryGetValue(summary.PrimaryTranscriptionTrack.AudioTrackIndex, out result))
                return json ? result.TranscriptJsonPath : result.TranscriptSrtPath;
            return "";
        }

        private static void WriteLogFile(AudioCensorScanSummary summary, IList<string> commandLog, Action<string, string, int> log)
        {
            var lines = new List<string>
            {
                "BachFlixNfo Audio Censor WhisperX Workflow",
                "Created: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                "",
                "Input MKV: " + summary.MkvPath,
                "Dictionary: " + summary.ProfanityDictionaryPath,
                "Future output MKV: " + summary.OutputMkvPath,
                "WhisperX JSON: " + GetPrimaryTranscriptPath(summary, true),
                "WhisperX SRT: " + GetPrimaryTranscriptPath(summary, false),
                "",
                "FFmpeg: " + DescribeTool(summary.Ffmpeg),
                "FFprobe: " + DescribeTool(summary.Ffprobe),
                "Transcription engine: " + DescribeTool(summary.TranscriptionTool),
                "",
                "Audio tracks:"
            };

            foreach (AudioTrackInfo track in summary.AudioTracks)
                lines.Add("  " + track.Describe());

            lines.Add("");
            lines.Add("Selected English audio tracks: " + string.Join(", ", summary.SelectedAudioTracks.Select(t => t.AudioTrackIndex)));
            lines.Add("WhisperX source audio track: " + summary.PrimaryTranscriptionTrack.AudioTrackIndex);
            lines.Add("");
            lines.Add("Commands executed:");
            if (commandLog == null || commandLog.Count == 0)
                lines.Add("  (none)");
            else
                foreach (string command in commandLog)
                    lines.Add("  " + command);

            lines.Add("");
            lines.Add("WhisperX profanity hits:");
            if (summary.TranscriptProfanityHits.Count == 0)
                lines.Add("  (none)");
            else
                foreach (TranscriptProfanityHit hit in summary.TranscriptProfanityHits)
                    lines.Add("  " + DescribeHit(hit) + " | approved=" + (hit.Approved ? "yes" : "no"));

            lines.Add("");
            lines.Add("Subtitle comparison:");
            if (summary.EnglishSubtitleSources.Count == 0)
                lines.Add("  (none)");
            else
                foreach (SubtitleSourceInfo source in summary.EnglishSubtitleSources)
                    lines.Add("  " + BuildSubtitleCoverageLine(source));

            WritePlanLogLines(summary.CensorPlan, lines);

            string error;
            string logPath = global::BachFlixLog.WriteBachFlixLog(lines, "Audio Censor", "AudioCensor", out error);
            if (!string.IsNullOrWhiteSpace(logPath))
                Write(log, "success", "Audio censor log written: " + logPath, 2);
            else if (!string.IsNullOrWhiteSpace(error))
                Write(log, "warning", "Could not write audio censor log: " + error, 2);
        }

        private static void WritePlanLogLines(AudioCensorPlan plan, List<string> lines)
        {
            lines.Add("");
            lines.Add("Censor plan:");
            if (plan == null)
            {
                lines.Add("  (not built)");
                return;
            }

            lines.Add("  Message: " + plan.Message);
            lines.Add("  Mux plan built: " + (plan.MuxPlanBuilt ? "yes" : "no"));
            lines.Add("  Output MKV: " + plan.OutputMkvPath);
            lines.Add("  Clean audio tracks: " + plan.AudioTrackPlans.Count);
            foreach (AudioTrackCensorPlan trackPlan in plan.AudioTrackPlans)
                lines.Add("    audio " + trackPlan.Track.AudioTrackIndex + ": mute segments=" + trackPlan.MuteSegments.Count + ", filter=" + trackPlan.AudioFilter);

            lines.Add("  Clean subtitle tracks: " + plan.CensoredSubtitleOutputs.Count);
            foreach (CensoredSubtitleOutput subtitle in plan.CensoredSubtitleOutputs)
                lines.Add("    " + subtitle.Path);

            lines.Add("  FFmpeg command:");
            lines.Add(string.IsNullOrWhiteSpace(plan.FfmpegCommandLine) ? "    (none)" : "    " + plan.FfmpegCommandLine);
        }

        private static string BuildSubtitleCoverageLine(SubtitleSourceInfo source)
        {
            if (source == null)
                return "(subtitle source missing)";

            if (!string.IsNullOrWhiteSpace(source.ScanError))
                return source.Describe() + " | comparison unavailable: " + source.ScanError;

            SubtitleCoverageComparison comparison = source.Comparison;
            if (comparison == null)
                return source.Describe() + " | comparison unavailable";

            return string.Format(
                CultureInfo.InvariantCulture,
                "{0} | cues={1}, subtitle profanity={2}, WhisperX covered={3}/{4} ({5:0.0}%), subtitle-only={6}",
                source.Describe(),
                comparison.SubtitleCueCount,
                comparison.SubtitleHitCount,
                comparison.MatchedTranscriptHits.Count,
                comparison.TranscriptHitCount,
                comparison.TranscriptCoveragePercent,
                comparison.ProfanityFoundOnlyInSubtitle.Count);
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
                    continue;
                }

                merged.Add(new AudioMuteSegment
                {
                    Start = segment.Start,
                    End = segment.End,
                    SourceHit = segment.SourceHit
                });
            }

            return merged;
        }

        private static string BuildCleanAudioTitle(AudioTrackInfo track)
        {
            string baseTitle = string.IsNullOrWhiteSpace(track.Title) ? "English" : track.Title;
            return baseTitle + " Clean";
        }

        private static void RenumberHits(IList<TranscriptProfanityHit> hits)
        {
            if (hits == null)
                return;

            List<TranscriptProfanityHit> ordered = hits
                .OrderBy(h => h.Start)
                .ThenBy(h => h.End)
                .ThenBy(h => h.SourceAudioTrackIndex)
                .ToList();

            for (int i = 0; i < ordered.Count; i++)
                ordered[i].ReviewNumber = i + 1;

            List<TranscriptProfanityHit> list = hits as List<TranscriptProfanityHit>;
            if (list != null)
            {
                list.Clear();
                list.AddRange(ordered);
            }
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
                CultureInfo.InvariantCulture,
                "{0:D2}:{1:D2}:{2:D2},{3:D3}",
                (int)value.TotalHours,
                value.Minutes,
                value.Seconds,
                value.Milliseconds);
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
        public string ProfanityDictionaryPath { get; set; }
        public bool DetectExternalTools { get; set; }
        public bool PromptForMissingEnglishAudio { get; set; }
        public List<int> SelectedAudioTrackIndexes { get; set; }
        public string FfmpegPath { get; set; }
        public string FfprobePath { get; set; }
        public string WhisperXPath { get; set; }
        public Action<string, string, int> Log { get; set; }
        public IList<string> CommandLog { get; set; }

        public AudioCensorScanOptions()
        {
            MkvPath = "";
            ProfanityDictionaryPath = "";
            DetectExternalTools = true;
            SelectedAudioTrackIndexes = new List<int>();
            FfmpegPath = "";
            FfprobePath = "";
            WhisperXPath = "";
            CommandLog = new List<string>();
        }
    }

    public sealed class AudioCensorScanSummary
    {
        public string MkvPath { get; set; }
        public string ProfanityDictionaryPath { get; set; }
        public int DictionaryTermCount { get; set; }
        public MediaProbeResult MediaProbe { get; set; }
        public List<AudioTrackInfo> AudioTracks { get; set; }
        public List<AudioTrackInfo> SelectedAudioTracks { get; set; }
        public AudioTrackInfo PrimaryTranscriptionTrack { get; set; }
        public Dictionary<int, WordTranscriptionResult> TranscriptionResults { get; private set; }
        public List<TranscriptProfanityHit> TranscriptProfanityHits { get; set; }
        public List<TranscriptProfanityHit> ApprovedTranscriptHits { get; set; }
        public List<SubtitleSourceInfo> SubtitleSources { get; set; }
        public List<SubtitleSourceInfo> EnglishSubtitleSources { get; set; }
        public string SubtitleWorkDirectory { get; set; }
        public string OutputMkvPath { get; set; }
        public AudioCensorPlan CensorPlan { get; set; }
        public ExternalToolResolution Ffmpeg { get; set; }
        public ExternalToolResolution Ffprobe { get; set; }
        public ExternalToolResolution TranscriptionTool { get; set; }

        public AudioCensorScanSummary()
        {
            MkvPath = "";
            ProfanityDictionaryPath = "";
            MediaProbe = new MediaProbeResult();
            AudioTracks = new List<AudioTrackInfo>();
            SelectedAudioTracks = new List<AudioTrackInfo>();
            PrimaryTranscriptionTrack = new AudioTrackInfo();
            TranscriptionResults = new Dictionary<int, WordTranscriptionResult>();
            TranscriptProfanityHits = new List<TranscriptProfanityHit>();
            ApprovedTranscriptHits = new List<TranscriptProfanityHit>();
            SubtitleSources = new List<SubtitleSourceInfo>();
            EnglishSubtitleSources = new List<SubtitleSourceInfo>();
            SubtitleWorkDirectory = "";
            OutputMkvPath = "";
            CensorPlan = new AudioCensorPlan();
            Ffmpeg = new ExternalToolResolution { ToolName = "ffmpeg" };
            Ffprobe = new ExternalToolResolution { ToolName = "ffprobe" };
            TranscriptionTool = new ExternalToolResolution { ToolName = "whisperx" };
        }
    }

    public sealed class AudioCensorPlan
    {
        public string InputMkvPath { get; set; }
        public string OutputMkvPath { get; set; }
        public bool MuxPlanBuilt { get; set; }
        public string Message { get; set; }
        public List<AudioTrackCensorPlan> AudioTrackPlans { get; private set; }
        public List<CensoredSubtitleOutput> CensoredSubtitleOutputs { get; private set; }
        public FfmpegMuxPlan MuxPlan { get; set; }
        public string FfmpegCommandLine { get; set; }

        public AudioCensorPlan()
        {
            InputMkvPath = "";
            OutputMkvPath = "";
            Message = "";
            AudioTrackPlans = new List<AudioTrackCensorPlan>();
            CensoredSubtitleOutputs = new List<CensoredSubtitleOutput>();
            MuxPlan = new FfmpegMuxPlan();
            FfmpegCommandLine = "";
        }
    }

    public sealed class AudioTrackCensorPlan
    {
        public AudioTrackInfo Track { get; set; }
        public List<AudioMuteSegment> MuteSegments { get; private set; }
        public string AudioFilter { get; set; }

        public AudioTrackCensorPlan()
        {
            Track = new AudioTrackInfo();
            MuteSegments = new List<AudioMuteSegment>();
            AudioFilter = "";
        }
    }
}
