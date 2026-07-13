using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    public static class AudioCensorRunner
    {
        private const string CleanedOutputDirectoryName = "Cleaned";
        private static readonly string[] SupportedVideoExtensions = new[] { ".mkv" };

        public static void RunInteractive(Action<string, string, int> log)
        {
            Write(log, "info", "=== AUDIO CENSOR: WHISPERX WORD TIMESTAMPS ===", 1);
            Write(log, "warning", "WhisperX word timestamps are the primary source. English subtitles are scanned too, and subtitle-only misses are added automatically.", 2);

            string inputPath = PromptForInputPath(log);
            if (string.IsNullOrWhiteSpace(inputPath))
                return;

            RunForInputPath(inputPath, log);
        }

        public static void RunForInputPath(string inputPath, Action<string, string, int> log)
        {
            string cleanInput = CleanConsolePath(inputPath);
            if (string.IsNullOrWhiteSpace(cleanInput))
                return;

            if (File.Exists(cleanInput))
            {
                if (!IsSupportedVideoFile(cleanInput))
                {
                    Write(log, "error", "Unsupported video file for this workflow: " + cleanInput + ". Current clean muxing support is MKV-only.", 1);
                    return;
                }

                if (IsCleanOutputFile(cleanInput))
                {
                    Write(log, "warning", "Skipping an existing clean output file: " + cleanInput, 1);
                    return;
                }

                string mkvPath = Path.GetFullPath(cleanInput);
                string outputPath = new FfmpegMuxService().BuildCleanOutputPath(mkvPath);
                if (File.Exists(outputPath))
                {
                    Write(log, "warning", "Skipping because the clean output already exists: " + outputPath, 2);
                    return;
                }

                string dictionaryPath = ResolveDictionaryPath(log, mkvPath, new List<string> { mkvPath });
                if (string.IsNullOrWhiteSpace(dictionaryPath))
                    return;

                Write(log, "info", "Single-file review mode is on. The scan report will be shown before FFmpeg creates the clean MKV.", 1);
                Write(log, "info", "Using profanity dictionary text file:", 1);
                Write(log, "data", dictionaryPath, 1);
                RunReviewedSingleMkv(mkvPath, dictionaryPath, log);
                return;
            }

            if (Directory.Exists(cleanInput))
            {
                RunAutomatic(cleanInput, log);
                return;
            }

            Write(log, "error", "Input path was not found: " + cleanInput, 1);
        }

        public static AudioCensorBatchSummary RunAutomatic(string inputPath, Action<string, string, int> log)
        {
            var batch = new AudioCensorBatchSummary
            {
                InputPath = CleanConsolePath(inputPath)
            };

            Write(log, "info", "Audio censor autopilot mode is on. Detections will be accepted automatically and clean MKVs will be created when possible.", 1);

            List<string> mkvPaths = ResolveInputMkvFiles(batch.InputPath, log);
            batch.TotalFiles = mkvPaths.Count;
            if (mkvPaths.Count == 0)
            {
                Write(log, "warning", "No supported MKV files were found to clean.", 2);
                return batch;
            }

            string dictionaryPath = ResolveDictionaryPath(log, batch.InputPath, mkvPaths);
            if (string.IsNullOrWhiteSpace(dictionaryPath))
            {
                batch.FailedFiles = mkvPaths.Count;
                batch.FailedInputs.AddRange(mkvPaths);
                return batch;
            }

            batch.ProfanityDictionaryPath = dictionaryPath;
            Write(log, "info", "Using profanity dictionary text file:", 1);
            Write(log, "data", dictionaryPath, 1);
            Write(log, "info", "Edit that text file whenever you want to change what gets censored; otherwise this run will keep moving with it.", 2);

            Write(log, "info", "Found " + mkvPaths.Count + " MKV file(s) to process.", 1);
            for (int i = 0; i < mkvPaths.Count; i++)
            {
                string mkvPath = mkvPaths[i];
                Write(log, "info", "=== AUDIO CENSOR " + (i + 1) + " of " + mkvPaths.Count + " ===", 2);
                Write(log, "data", mkvPath, 1);

                string outputPath = new FfmpegMuxService().BuildCleanOutputPath(mkvPath);
                if (File.Exists(outputPath))
                {
                    batch.SkippedFiles++;
                    batch.SkippedInputs.Add(mkvPath);
                    Write(log, "warning", "Skipping because the clean output already exists: " + outputPath, 2);
                    continue;
                }

                try
                {
                    AudioCensorScanSummary summary = RunAutomaticForSingleMkv(mkvPath, dictionaryPath, log);
                    if (summary != null && summary.CensorPlan != null && summary.CensorPlan.MuxPlanBuilt)
                    {
                        batch.CleanedFiles++;
                        batch.OutputPaths.Add(summary.CensorPlan.OutputMkvPath);
                    }
                    else
                    {
                        batch.SkippedFiles++;
                        batch.SkippedInputs.Add(mkvPath);
                    }
                }
                catch (Exception ex)
                {
                    batch.FailedFiles++;
                    batch.FailedInputs.Add(mkvPath);
                    Write(log, "error", "Audio censor failed for: " + mkvPath, 1);
                    Write(log, "harderror", ex.Message, 2);
                }
            }

            string finalType = batch.FailedFiles == 0 ? "success" : "warning";
            Write(log, finalType, "Audio censor autopilot complete. Cleaned " + batch.CleanedFiles + " of " + batch.TotalFiles + " file(s); skipped " + batch.SkippedFiles + "; failed " + batch.FailedFiles + ".", 2);
            return batch;
        }

        private static AudioCensorScanSummary RunReviewedSingleMkv(string mkvPath, string dictionaryPath, Action<string, string, int> log)
        {
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
                    UseAutomaticAudioTrackFallback = false,
                    Log = log,
                    CommandLog = commandLog
                });
            }
            catch (Exception ex)
            {
                Write(log, "error", "Audio censor scan failed.", 1);
                Write(log, "harderror", ex.Message, 2);
                return null;
            }

            List<TranscriptProfanityHit> reviewHits = BuildAudioMuteReviewHits(summary, log);
            summary.ProposedAudioMuteHits = reviewHits;
            WritePreReviewSummary(summary, reviewHits, log);
            summary.ApprovedTranscriptHits = ReviewAudioMuteHits(reviewHits, summary.SelectedAudioTracks, log);
            summary.CensorPlan = BuildCensorPlan(summary, commandLog, log);

            WriteComparisonReport(summary, log);
            WriteSummary(summary, log);
            WriteLogFile(summary, commandLog, log);

            if (summary.CensorPlan != null && summary.CensorPlan.MuxPlanBuilt)
            {
                Write(log, "question", "Create the _Clean MKV now? Type y to run FFmpeg, or press Enter to leave the report and plan only:", 1);
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
                else
                {
                    Write(log, "warning", "No clean MKV was created. The report and FFmpeg plan were left for review.", 2);
                }
            }
            else
            {
                Write(log, "warning", "No clean MKV was created for this file because no mux plan was available.", 2);
            }

            return summary;
        }

        private static AudioCensorScanSummary RunAutomaticForSingleMkv(string mkvPath, string dictionaryPath, Action<string, string, int> log)
        {
            var commandLog = new List<string>();
            AudioCensorScanSummary summary;

            try
            {
                summary = RunScan(new AudioCensorScanOptions
                {
                    MkvPath = mkvPath,
                    ProfanityDictionaryPath = dictionaryPath,
                    DetectExternalTools = true,
                    PromptForMissingEnglishAudio = false,
                    UseAutomaticAudioTrackFallback = true,
                    Log = log,
                    CommandLog = commandLog
                });
            }
            catch (Exception ex)
            {
                Write(log, "error", "Audio censor scan failed.", 1);
                Write(log, "harderror", ex.Message, 2);
                throw;
            }

            List<TranscriptProfanityHit> reviewHits = BuildAudioMuteReviewHits(summary, log);
            summary.ProposedAudioMuteHits = reviewHits;
            WritePreReviewSummary(summary, reviewHits, log);
            summary.ApprovedTranscriptHits = ApproveAudioMuteHitsAutomatically(reviewHits, log);
            summary.CensorPlan = BuildCensorPlan(summary, commandLog, log);

            WriteComparisonReport(summary, log);
            WriteSummary(summary, log);
            WriteLogFile(summary, commandLog, log);

            if (summary.CensorPlan != null && summary.CensorPlan.MuxPlanBuilt)
            {
                try
                {
                    Write(log, "info", "Creating clean MKV without waiting for confirmation...", 1);
                    new FfmpegMuxService().MuxCensoredMedia(summary.CensorPlan.MuxPlan, log, commandLog);
                    Write(log, "success", "Created clean MKV: " + summary.CensorPlan.OutputMkvPath, 2);
                }
                catch (Exception ex)
                {
                    Write(log, "error", "FFmpeg mux failed.", 1);
                    Write(log, "harderror", ex.Message, 2);
                    throw;
                }
            }
            else
            {
                Write(log, "warning", "No clean MKV was created for this file because no mux plan was available.", 2);
            }

            return summary;
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
            AudioCensorSettings settings = AudioCensorSettings.LoadFromEnvironment();
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
            if (selectedAudioTracks.Count == 0 && options.UseAutomaticAudioTrackFallback)
                selectedAudioTracks = SelectAutomaticAudioTrackFallback(media.AudioTracks, options.Log);

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
                Settings = settings,
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
            ApplyWhisperHitSettings(primaryScan.Hits, settings);
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
                ApplyWhisperHitSettings(scan.Hits, settings);
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

        private static void ApplyWhisperHitSettings(IEnumerable<TranscriptProfanityHit> hits, AudioCensorSettings settings)
        {
            AudioCensorSettings effectiveSettings = settings ?? AudioCensorSettings.LoadFromEnvironment();
            foreach (TranscriptProfanityHit hit in hits ?? Enumerable.Empty<TranscriptProfanityHit>())
            {
                if (hit == null)
                    continue;

                hit.DetectionSource = ProfanityDetectionSource.WHISPERX;
                hit.Confidence = "HIGH";
                hit.PaddingBefore = effectiveSettings.WhisperPrePadding;
                hit.PaddingAfter = effectiveSettings.WhisperPostPadding;
                hit.IsEstimatedTiming = false;
            }
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
            AudioCensorSettings settings = summary.Settings ?? AudioCensorSettings.LoadFromEnvironment();

            List<TranscriptProfanityHit> approvedHits = (summary.ApprovedTranscriptHits ?? new List<TranscriptProfanityHit>())
                .Where(h => h != null && h.Approved)
                .ToList();

            if (approvedHits.Count == 0)
                Write(log, "warning", "No approved audio mute hits. Audio muting will be skipped, but clean subtitles can still use approved subtitle scan hits.", 1);

            var muteBuilder = new AudioMuteFilterBuilder();
            foreach (AudioTrackInfo track in summary.SelectedAudioTracks.OrderBy(t => t.AudioTrackIndex))
            {
                List<TranscriptProfanityHit> hitsForTrack = approvedHits
                    .Where(h => h.AppliesToAudioTrackIndexes.Contains(track.AudioTrackIndex))
                    .ToList();
                List<AudioMuteSegment> segments = MergeSegments(muteBuilder.CreateSegments(hitsForTrack), settings.MergeGap);
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

            var subtitleWriter = new CensoredSubtitleWriter();
            foreach (SubtitleSourceInfo subtitle in summary.EnglishSubtitleSources)
            {
                if (!string.IsNullOrWhiteSpace(subtitle.ScanError))
                    continue;

                try
                {
                    CensoredSubtitleOutput output = subtitleWriter.WriteCensoredSubtitle(subtitle, approvedHits, GetSubtitleCensorOccurrences(subtitle, approvedHits), summary.MkvPath, settings.SubtitleReplacementMode, settings.SubtitleReplacementText);
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

        private static List<AudioTrackInfo> SelectAutomaticAudioTrackFallback(IReadOnlyList<AudioTrackInfo> tracks, Action<string, string, int> log)
        {
            var available = (tracks ?? new List<AudioTrackInfo>())
                .OrderBy(t => t.AudioTrackIndex)
                .ToList();
            if (available.Count == 0)
                return new List<AudioTrackInfo>();

            AudioTrackInfo selected = available
                .Where(t => t.IsDefault)
                .OrderBy(t => t.AudioTrackIndex)
                .FirstOrDefault() ?? available[0];

            Write(log, "warning", "No English audio track could be confidently identified. Autopilot selected " + (selected.IsDefault ? "the default" : "the first") + " audio track instead: " + selected.Describe(), 1);
            return new List<AudioTrackInfo> { selected };
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

        private static void WritePreReviewSummary(AudioCensorScanSummary summary, IReadOnlyList<TranscriptProfanityHit> reviewHits, Action<string, string, int> log)
        {
            if (summary == null)
                return;

            AudioCensorSettings settings = summary.Settings ?? AudioCensorSettings.LoadFromEnvironment();
            int whisperCount = summary.TranscriptProfanityHits == null ? 0 : summary.TranscriptProfanityHits.Count;
            int subtitleCount = GetSubtitleProfanityCount(summary);
            int srtOnlyCount = (reviewHits ?? new List<TranscriptProfanityHit>()).Count(h => h != null && h.DetectionSource == ProfanityDetectionSource.SRT_ONLY_ESTIMATED);
            int proposedIntervalCount = CountMergedProposedIntervals(reviewHits, settings);

            Write(log, "info", "Audio censor detection summary before review:", 1);
            Write(log, "data", "WhisperX profanity count: " + whisperCount, 1);
            Write(log, "data", "Subtitle profanity count: " + subtitleCount, 1);
            Write(log, "data", "SRT-only missing-from-WhisperX count: " + srtOnlyCount, 1);
            Write(log, "data", "Total proposed censor intervals: " + proposedIntervalCount, 1);
        }

        private static int GetSubtitleProfanityCount(AudioCensorScanSummary summary)
        {
            if (summary == null || summary.EnglishSubtitleSources == null)
                return 0;

            return summary.EnglishSubtitleSources
                .Where(s => s != null && string.IsNullOrWhiteSpace(s.ScanError) && s.ScanResult != null && s.ScanResult.Occurrences != null)
                .Sum(s => s.ScanResult.Occurrences.Count);
        }

        private static int CountMergedProposedIntervals(IReadOnlyList<TranscriptProfanityHit> reviewHits, AudioCensorSettings settings)
        {
            var muteBuilder = new AudioMuteFilterBuilder();
            var merger = new CensorIntervalMerger();
            return merger.Merge(muteBuilder.CreateSegments(reviewHits), (settings ?? AudioCensorSettings.LoadFromEnvironment()).MergeGap).Count;
        }

        private static List<TranscriptProfanityHit> ApproveAudioMuteHitsAutomatically(
            IReadOnlyList<TranscriptProfanityHit> hits,
            Action<string, string, int> log)
        {
            var approved = (hits ?? new List<TranscriptProfanityHit>())
                .Where(h => h != null)
                .ToList();

            foreach (TranscriptProfanityHit hit in approved)
                hit.Approved = true;

            RenumberHits(approved);
            Write(log, "success", "Autopilot accepted " + approved.Count + " audio mute hit(s); no profanity review prompt was shown.", 2);
            return approved;
        }

        private static List<TranscriptProfanityHit> ReviewAudioMuteHits(
            IReadOnlyList<TranscriptProfanityHit> hits,
            IReadOnlyList<AudioTrackInfo> selectedAudioTracks,
            Action<string, string, int> log)
        {
            var editableHits = (hits ?? new List<TranscriptProfanityHit>()).ToList();
            if (editableHits.Count == 0)
            {
                Write(log, "success", "Review list is empty. No WhisperX or subtitle-only profanity was detected.", 2);
                return new List<TranscriptProfanityHit>();
            }

            Write(log, "info", "Review audio mute profanity hits:", 1);
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
            Write(log, "success", "Approved " + approved.Count + " of " + editableHits.Count + " audio mute hit(s).", 2);
            return approved;
        }
        private static List<TranscriptProfanityHit> BuildAudioMuteReviewHits(AudioCensorScanSummary summary, Action<string, string, int> log)
        {
            var reviewHits = new List<TranscriptProfanityHit>();
            if (summary == null)
                return reviewHits;

            AudioCensorSettings settings = summary.Settings ?? AudioCensorSettings.LoadFromEnvironment();

            if (summary.TranscriptProfanityHits != null)
                reviewHits.AddRange(summary.TranscriptProfanityHits.Where(h => h != null));

            List<SubtitleOnlyProfanityCandidate> candidates = BuildSubtitleOnlyCandidates(summary);
            List<TimedTranscriptWord> whisperContextWords = GetPrimaryTranscriptWords(summary);
            var estimator = new SrtOnlyTimingEstimator();
            List<TranscriptProfanityHit> subtitleHits = candidates
                .Select(c => ConvertSubtitleOnlyCandidateToHit(c, summary.SelectedAudioTracks, whisperContextWords, settings, estimator))
                .Where(h => h != null)
                .ToList();

            summary.SrtOnlyEstimatedHits = subtitleHits;
            if (subtitleHits.Count > 0)
                Write(log, "warning", "Subtitle comparison added " + subtitleHits.Count + " WhisperX-missed estimated hit(s) to the audio mute review.", 1);

            reviewHits.AddRange(subtitleHits);
            RenumberHits(reviewHits);
            summary.ProposedAudioMuteHits = reviewHits;
            return reviewHits;
        }

        private static List<SubtitleOnlyProfanityCandidate> BuildSubtitleOnlyCandidates(AudioCensorScanSummary summary)
        {
            var candidates = new List<SubtitleOnlyProfanityCandidate>();
            if (summary == null || summary.EnglishSubtitleSources == null)
                return candidates;

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SubtitleSourceInfo source in summary.EnglishSubtitleSources)
            {
                if (source == null || source.Comparison == null || source.Comparison.ProfanityFoundOnlyInSubtitle == null)
                    continue;

                foreach (ProfanityOccurrence occurrence in source.Comparison.ProfanityFoundOnlyInSubtitle)
                {
                    if (occurrence == null || occurrence.Cue == null)
                        continue;

                    string key = BuildSubtitleOnlyCandidateKey(occurrence);
                    if (!seen.Add(key))
                        continue;

                    candidates.Add(new SubtitleOnlyProfanityCandidate
                    {
                        ReviewNumber = candidates.Count + 1,
                        Source = source,
                        Occurrence = occurrence
                    });
                }
            }

            return candidates;
        }

        private static string BuildSubtitleOnlyCandidateKey(ProfanityOccurrence occurrence)
        {
            string term = ProfanityDictionary.NormalizeToken(string.IsNullOrWhiteSpace(occurrence.DictionaryTerm) ? occurrence.Word : occurrence.DictionaryTerm);
            int tokenIndex = occurrence.TokenIndexInCue > 0 ? occurrence.TokenIndexInCue : occurrence.CharacterIndex;
            return term + "|" + occurrence.Cue.Start.Ticks + "|" + occurrence.Cue.End.Ticks + "|" + tokenIndex;
        }

        private static List<TimedTranscriptWord> GetPrimaryTranscriptWords(AudioCensorScanSummary summary)
        {
            if (summary == null || summary.TranscriptionResults == null)
                return new List<TimedTranscriptWord>();

            WordTranscriptionResult primaryResult;
            if (summary.PrimaryTranscriptionTrack != null &&
                summary.TranscriptionResults.TryGetValue(summary.PrimaryTranscriptionTrack.AudioTrackIndex, out primaryResult) &&
                primaryResult != null &&
                primaryResult.Words != null &&
                primaryResult.Words.Count > 0)
            {
                return primaryResult.Words
                    .Where(w => w != null && w.End > w.Start && !string.IsNullOrWhiteSpace(w.NormalizedWord))
                    .OrderBy(w => w.Start)
                    .ThenBy(w => w.End)
                    .ToList();
            }

            return summary.TranscriptionResults.Values
                .Where(r => r != null && r.Words != null)
                .SelectMany(r => r.Words)
                .Where(w => w != null && w.End > w.Start && !string.IsNullOrWhiteSpace(w.NormalizedWord))
                .OrderBy(w => w.Start)
                .ThenBy(w => w.End)
                .GroupBy(w => w.NormalizedWord + "|" + w.Start.Ticks + "|" + w.End.Ticks)
                .Select(g => g.First())
                .ToList();
        }

        private static TranscriptProfanityHit ConvertSubtitleOnlyCandidateToHit(
            SubtitleOnlyProfanityCandidate candidate,
            IReadOnlyList<AudioTrackInfo> selectedAudioTracks,
            IReadOnlyList<TimedTranscriptWord> whisperContextWords,
            AudioCensorSettings settings,
            SrtOnlyTimingEstimator estimator)
        {
            if (candidate == null || candidate.Occurrence == null || candidate.Occurrence.Cue == null)
                return null;

            List<int> trackIndexes = (selectedAudioTracks ?? new List<AudioTrackInfo>())
                .Select(t => t.AudioTrackIndex)
                .Distinct()
                .OrderBy(i => i)
                .ToList();
            if (trackIndexes.Count == 0)
                return null;

            AudioCensorSettings effectiveSettings = settings ?? AudioCensorSettings.LoadFromEnvironment();
            SrtOnlyTimingEstimate estimate = (estimator ?? new SrtOnlyTimingEstimator()).Estimate(candidate.Occurrence, effectiveSettings, whisperContextWords);
            if (estimate.End <= estimate.Start)
                return null;

            var hit = new TranscriptProfanityHit
            {
                Word = candidate.Occurrence.Word,
                DictionaryTerm = candidate.Occurrence.DictionaryTerm,
                Start = estimate.Start,
                End = estimate.End,
                PaddingBefore = effectiveSettings.SrtOnlyPrePadding,
                PaddingAfter = effectiveSettings.SrtOnlyPostPadding,
                SourceAudioTrackIndex = trackIndexes[0],
                SourceLabel = "subtitle-only estimated: " + DescribeSubtitleSourceShort(candidate.Source),
                DetectionSource = ProfanityDetectionSource.SRT_ONLY_ESTIMATED,
                Confidence = string.IsNullOrWhiteSpace(estimate.Confidence) ? "LOWER" : estimate.Confidence,
                SubtitleSourceLabel = DescribeSubtitleSourceShort(candidate.Source),
                SubtitleSourcePath = GetSubtitleSourcePath(candidate.Source),
                SubtitleStreamIndex = candidate.Source == null ? -1 : candidate.Source.StreamIndex,
                SubtitleCueIndex = candidate.Occurrence.CueIndex,
                SubtitleSequenceNumber = candidate.Occurrence.SubtitleSequenceNumber,
                SubtitleCueStart = candidate.Occurrence.Cue.Start,
                SubtitleCueEnd = candidate.Occurrence.Cue.End,
                EstimatedTokenIndexInCue = estimate.TokenIndexInCue,
                EstimatedTokenCountInCue = estimate.TokenCountInCue,
                SubtitleCueText = candidate.Occurrence.Cue.Text ?? "",
                NearbyWhisperWords = estimate.NearbyWhisperWords,
                MatchedLeftAnchorWord = estimate.MatchedLeftAnchorWord,
                MatchedRightAnchorWord = estimate.MatchedRightAnchorWord,
                SrtOnlyDiagnosticReason = estimate.DiagnosticReason,
                UsedTokenPositionFallback = estimate.UsedTokenPositionFallback,
                SrtOnlyFallbackReason = estimate.FallbackReason,
                MinMuteDuration = effectiveSettings.SrtOnlyMinDuration,
                MaxMuteDuration = effectiveSettings.SrtOnlyMaxDuration,
                IsEstimatedTiming = true,
                Approved = true
            };

            foreach (int trackIndex in trackIndexes)
                hit.AppliesToAudioTrackIndexes.Add(trackIndex);

            return hit;
        }

        private static string DescribeSubtitleSourceShort(SubtitleSourceInfo source)
        {
            if (source == null)
                return "subtitle";

            if (source.SourceKind == SubtitleSourceKind.External && !string.IsNullOrWhiteSpace(source.Path))
                return Path.GetFileName(source.Path);

            return "embedded subtitle stream " + source.StreamIndex;
        }

        private static string GetSubtitleSourcePath(SubtitleSourceInfo source)
        {
            if (source == null)
                return "";

            if (!string.IsNullOrWhiteSpace(source.Path))
                return source.Path;

            return source.WorkingSubtitlePath ?? "";
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
                    DetectionSource = ProfanityDetectionSource.MANUAL,
                    Confidence = "MANUAL",
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
            if (hit == null)
                return "(missing hit)";

            string tracks = hit.AppliesToAudioTrackIndexes.Count == 0
                ? "(none)"
                : string.Join(",", hit.AppliesToAudioTrackIndexes.OrderBy(i => i));
            string confidence = string.IsNullOrWhiteSpace(hit.Confidence) ? "UNKNOWN" : hit.Confidence;
            string sourceDetail = string.IsNullOrWhiteSpace(hit.SourceLabel) ? "" : " | source detail=" + hit.SourceLabel;
            string subtitle = string.IsNullOrWhiteSpace(hit.SubtitleSourceLabel) ? "" : " | subtitle=" + hit.SubtitleSourceLabel;
            string estimated = "";
            if (hit.IsEstimatedTiming)
            {
                estimated = string.Format(
                    CultureInfo.InvariantCulture,
                    " | estimated=yes token={0}/{1} cue={2}->{3}",
                    hit.EstimatedTokenIndexInCue,
                    hit.EstimatedTokenCountInCue,
                    FormatTime(hit.SubtitleCueStart),
                    FormatTime(hit.SubtitleCueEnd));

                if (hit.DetectionSource == ProfanityDetectionSource.SRT_ONLY_ESTIMATED)
                {
                    estimated += " | subtitle text=\"" + CleanInline(hit.SubtitleCueText) + "\"";
                    estimated += " | nearby WhisperX=" + CleanInline(hit.NearbyWhisperWords);
                    estimated += " | anchor left=" + EmptyAsNone(hit.MatchedLeftAnchorWord);
                    estimated += " | anchor right=" + EmptyAsNone(hit.MatchedRightAnchorWord);
                    estimated += " | diagnostic=" + hit.SrtOnlyDiagnosticReason;
                    estimated += " | token fallback=" + (hit.UsedTokenPositionFallback ? "yes" : "no");
                    if (hit.UsedTokenPositionFallback && !string.IsNullOrWhiteSpace(hit.SrtOnlyFallbackReason))
                        estimated += " | fallback reason=" + hit.SrtOnlyFallbackReason;
                }
            }

            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}. {1} --> {2} | audio {3} | \"{4}\" | source={5} | confidence={6} | dictionary={7} | padding={8}ms/{9}ms{10}{11}{12}",
                hit.ReviewNumber,
                FormatTime(hit.Start),
                FormatTime(hit.End),
                tracks,
                hit.Word,
                hit.DetectionSource,
                confidence,
                hit.DictionaryTerm,
                (int)hit.PaddingBefore.TotalMilliseconds,
                (int)hit.PaddingAfter.TotalMilliseconds,
                subtitle,
                sourceDetail,
                estimated);
        }

        private static void WriteComparisonReport(AudioCensorScanSummary summary, Action<string, string, int> log)
        {
            try
            {
                string reportPath = new AudioCensorReportWriter().Write(summary);
                summary.ComparisonReportPath = reportPath;
                Write(log, "success", "Audio censor comparison report written: " + reportPath, 2);
            }
            catch (Exception ex)
            {
                Write(log, "warning", "Could not write audio censor comparison report: " + ex.Message, 2);
            }
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
            Write(log, "data", "Subtitle profanity hits: " + GetSubtitleProfanityCount(summary), 1);
            Write(log, "data", "SRT-only estimated hits: " + summary.SrtOnlyEstimatedHits.Count, 1);
            Write(log, "data", "Approved hits: " + summary.ApprovedTranscriptHits.Count, 1);
            Write(log, "data", "English subtitles checked: " + summary.EnglishSubtitleSources.Count, 1);
            Write(log, "data", "SRT-only timing: min=" + summary.Settings.SrtOnlyMinDurationMs + "ms, max=" + summary.Settings.SrtOnlyMaxDurationMs + "ms, padding=" + summary.Settings.SrtOnlyPrePaddingMs + "ms/" + summary.Settings.SrtOnlyPostPaddingMs + "ms, context=" + summary.Settings.SrtOnlyContextWindowBeforeMs + "ms/" + summary.Settings.SrtOnlyContextWindowAfterMs + "ms", 1);
            Write(log, "data", "Merge gap: " + summary.Settings.MergeGapMs + "ms", 1);
            Write(log, "data", "Clean output path: " + summary.OutputMkvPath, 1);
            if (!string.IsNullOrWhiteSpace(summary.ComparisonReportPath))
                Write(log, "data", "Comparison report: " + summary.ComparisonReportPath, 1);

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
                "Comparison report: " + summary.ComparisonReportPath,
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
            lines.Add("Approved audio mute hits:");
            if (summary.ApprovedTranscriptHits.Count == 0)
                lines.Add("  (none)");
            else
                foreach (TranscriptProfanityHit hit in summary.ApprovedTranscriptHits)
                    lines.Add("  " + DescribeHit(hit));

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

        private static List<ProfanityOccurrence> GetSubtitleCensorOccurrences(SubtitleSourceInfo subtitle, IReadOnlyList<TranscriptProfanityHit> approvedHits)
        {
            if (subtitle == null || subtitle.ScanResult == null || subtitle.ScanResult.Occurrences == null)
                return new List<ProfanityOccurrence>();

            List<TranscriptProfanityHit> hits = (approvedHits ?? new List<TranscriptProfanityHit>())
                .Where(h => h != null && h.Approved)
                .ToList();

            return subtitle.ScanResult.Occurrences
                .Where(o => o != null && o.Cue != null && hits.Any(h => OccurrenceMatchesApprovedHit(o, h)))
                .ToList();
        }

        private static bool OccurrenceMatchesApprovedHit(ProfanityOccurrence occurrence, TranscriptProfanityHit hit)
        {
            if (occurrence == null || occurrence.Cue == null || hit == null)
                return false;

            string occurrenceTerm = ProfanityDictionary.NormalizeToken(string.IsNullOrWhiteSpace(occurrence.DictionaryTerm) ? occurrence.Word : occurrence.DictionaryTerm);
            string hitTerm = ProfanityDictionary.NormalizeToken(string.IsNullOrWhiteSpace(hit.DictionaryTerm) ? hit.Word : hit.DictionaryTerm);
            if (string.IsNullOrWhiteSpace(occurrenceTerm) || occurrenceTerm != hitTerm)
                return false;

            if (hit.DetectionSource == ProfanityDetectionSource.SRT_ONLY_ESTIMATED)
            {
                bool cueMatches = Math.Abs((occurrence.Cue.Start - hit.SubtitleCueStart).TotalMilliseconds) <= 50 &&
                                  Math.Abs((occurrence.Cue.End - hit.SubtitleCueEnd).TotalMilliseconds) <= 50;
                bool tokenMatches = hit.EstimatedTokenIndexInCue <= 0 ||
                                    occurrence.TokenIndexInCue <= 0 ||
                                    occurrence.TokenIndexInCue == hit.EstimatedTokenIndexInCue;
                return cueMatches && tokenMatches;
            }

            return TimesOverlap(occurrence.Cue.Start, occurrence.Cue.End, hit.Start, hit.End, TimeSpan.FromSeconds(2));
        }

        private static bool TimesOverlap(TimeSpan startA, TimeSpan endA, TimeSpan startB, TimeSpan endB, TimeSpan tolerance)
        {
            TimeSpan start = startA - tolerance;
            if (start < TimeSpan.Zero)
                start = TimeSpan.Zero;

            TimeSpan end = endA + tolerance;
            return start <= endB && end >= startB;
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

        private static List<AudioMuteSegment> MergeSegments(IEnumerable<AudioMuteSegment> segments, TimeSpan mergeGap)
        {
            return new CensorIntervalMerger().Merge(segments, mergeGap);
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

        private static string PromptForInputPath(Action<string, string, int> log)
        {
            while (true)
            {
                Write(log, "question", "Enter MKV file path or folder path, or 0 to cancel:", 1);
                string input = CleanConsolePath(Console.ReadLine());

                if (input == "0")
                    return null;

                if (File.Exists(input) && IsSupportedVideoFile(input))
                    return input;

                if (Directory.Exists(input))
                    return input;

                Write(log, "error", "Please enter an existing MKV file path or folder path.", 1);
            }
        }

        private static List<string> ResolveInputMkvFiles(string inputPath, Action<string, string, int> log)
        {
            var paths = new List<string>();
            string cleanInput = CleanConsolePath(inputPath);
            if (string.IsNullOrWhiteSpace(cleanInput))
                return paths;

            if (File.Exists(cleanInput))
            {
                if (!IsSupportedVideoFile(cleanInput))
                {
                    Write(log, "error", "Unsupported video file for this workflow: " + cleanInput + ". Current clean muxing support is MKV-only.", 1);
                    return paths;
                }

                if (IsCleanOutputFile(cleanInput))
                {
                    Write(log, "warning", "Skipping an existing clean output file: " + cleanInput, 1);
                    return paths;
                }

                paths.Add(Path.GetFullPath(cleanInput));
                return paths;
            }

            if (Directory.Exists(cleanInput))
                return FindSupportedMkvFiles(Path.GetFullPath(cleanInput), log);

            Write(log, "error", "Input path was not found: " + cleanInput, 1);
            return paths;
        }

        private static List<string> FindSupportedMkvFiles(string rootDirectory, Action<string, string, int> log)
        {
            var results = new List<string>();
            var pending = new Stack<string>();
            pending.Push(rootDirectory);

            while (pending.Count > 0)
            {
                string directory = pending.Pop();
                if (IsCleanedDirectory(directory))
                    continue;

                string[] files;
                try
                {
                    files = Directory.GetFiles(directory);
                }
                catch (Exception ex)
                {
                    Write(log, "warning", "Could not read files in: " + directory + " | " + ex.Message, 1);
                    continue;
                }

                foreach (string file in files)
                {
                    if (IsSupportedVideoFile(file) && !IsCleanOutputFile(file))
                        results.Add(Path.GetFullPath(file));
                }

                string[] subdirectories;
                try
                {
                    subdirectories = Directory.GetDirectories(directory);
                }
                catch (Exception ex)
                {
                    Write(log, "warning", "Could not read folders in: " + directory + " | " + ex.Message, 1);
                    continue;
                }

                foreach (string subdirectory in subdirectories)
                {
                    if (!IsCleanedDirectory(subdirectory) && !IsReparsePoint(subdirectory))
                        pending.Push(subdirectory);
                }
            }

            return results
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool IsSupportedVideoFile(string path)
        {
            string extension = Path.GetExtension(path);
            return SupportedVideoExtensions.Any(e => e.Equals(extension, StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsCleanOutputFile(string path)
        {
            if (IsInCleanedDirectory(path))
                return true;

            string name = Path.GetFileNameWithoutExtension(path) ?? "";
            return name.EndsWith("_Clean", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCleanedDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            string trimmed = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string name = Path.GetFileName(trimmed) ?? "";
            return name.Equals(CleanedOutputDirectoryName, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsInCleanedDirectory(string path)
        {
            string directory = Path.GetDirectoryName(path);
            while (!string.IsNullOrWhiteSpace(directory))
            {
                if (IsCleanedDirectory(directory))
                    return true;

                DirectoryInfo parent = Directory.GetParent(directory);
                directory = parent == null ? null : parent.FullName;
            }

            return false;
        }

        private static bool IsReparsePoint(string path)
        {
            try
            {
                return (File.GetAttributes(path) & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint;
            }
            catch
            {
                return true;
            }
        }

        private static string ResolveDictionaryPath(Action<string, string, int> log, string inputPath, IReadOnlyList<string> mkvPaths)
        {
            var suggestions = new List<string>();
            string cleanInput = CleanConsolePath(inputPath);

            if (Directory.Exists(cleanInput))
                suggestions.Add(Path.Combine(cleanInput, "profanity.txt"));

            if (File.Exists(cleanInput))
                suggestions.AddRange(GetDictionarySuggestions(cleanInput));
            else if (mkvPaths != null && mkvPaths.Count > 0)
                suggestions.AddRange(GetDictionarySuggestions(mkvPaths[0]));

            suggestions.Add(Path.Combine(Environment.CurrentDirectory, "profanity.txt"));
            suggestions.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "profanity.txt"));

            List<string> distinctSuggestions = suggestions
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Select(p => p.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            string found = distinctSuggestions.FirstOrDefault(File.Exists);
            if (!string.IsNullOrWhiteSpace(found))
                return Path.GetFullPath(found);

            Write(log, "error", "No profanity dictionary text file was found. Create or edit one of these paths, then run again:", 1);
            foreach (string suggestion in distinctSuggestions)
                Write(log, "data", suggestion, 1);
            return null;
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

        private static string CleanInline(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? "(none)"
                : value.Replace("\r", " ").Replace("\n", " ").Trim();
        }

        private static string EmptyAsNone(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "(none)" : CleanInline(value);
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

    public sealed class SrtOnlyTimingEstimator
    {
        private static readonly Regex WordRegex = new Regex(@"[\p{L}\p{Nd}]+(?:['\u2019-][\p{L}\p{Nd}]+)*", RegexOptions.Compiled);
        private const int NearbyWordReportLimit = 30;

        public SrtOnlyTimingEstimate Estimate(ProfanityOccurrence occurrence, AudioCensorSettings settings)
        {
            return Estimate(occurrence, settings, null);
        }

        public SrtOnlyTimingEstimate Estimate(
            ProfanityOccurrence occurrence,
            AudioCensorSettings settings,
            IEnumerable<TimedTranscriptWord> whisperWords)
        {
            var estimate = new SrtOnlyTimingEstimate
            {
                Confidence = "LOWER",
                DiagnosticReason = SrtOnlyDiagnosticReason.UNKNOWN,
                UsedTokenPositionFallback = true,
                FallbackReason = "UNKNOWN"
            };

            if (occurrence == null || occurrence.Cue == null || occurrence.Cue.End <= occurrence.Cue.Start)
                return estimate;

            AudioCensorSettings effectiveSettings = settings ?? AudioCensorSettings.LoadFromEnvironment();
            List<SrtOnlyTextToken> cueTokens = TokenizeCueText(occurrence.Cue.Text);
            int tokenCount = occurrence.TokenCountInCue > 0 ? occurrence.TokenCountInCue : Math.Max(1, cueTokens.Count);
            int tokenIndex = ResolveTokenIndex(occurrence, cueTokens, tokenCount);
            int profanityTokenIndex = cueTokens.Count == 0 ? tokenIndex - 1 : Math.Max(0, Math.Min(cueTokens.Count - 1, tokenIndex - 1));

            TimeSpan contextStart = occurrence.Cue.Start - effectiveSettings.SrtOnlyContextWindowBefore;
            if (contextStart < TimeSpan.Zero)
                contextStart = TimeSpan.Zero;
            TimeSpan contextEnd = occurrence.Cue.End + effectiveSettings.SrtOnlyContextWindowAfter;

            List<SrtOnlyWhisperWord> nearbyWords = GetNearbyWhisperWords(whisperWords, contextStart, contextEnd);
            List<SrtOnlyWhisperWord> wordsInCue = nearbyWords
                .Where(w => w.End >= occurrence.Cue.Start && w.Start <= occurrence.Cue.End)
                .ToList();
            SrtOnlyAnchorPair anchors = FindAnchors(cueTokens, profanityTokenIndex, nearbyWords);
            SrtOnlyWhisperWord similarWord = FindSimilarSoundingWord(occurrence, wordsInCue.Count > 0 ? wordsInCue : nearbyWords);
            bool exactProfanityInCue = ContainsProfanityWord(occurrence, wordsInCue);
            bool exactProfanityNearby = ContainsProfanityWord(occurrence, nearbyWords);
            SrtOnlyDiagnosticReason diagnostic = DetermineDiagnosticReason(
                wordsInCue,
                anchors,
                similarWord,
                exactProfanityInCue,
                exactProfanityNearby);

            estimate.TokenIndexInCue = tokenIndex;
            estimate.TokenCountInCue = tokenCount;
            estimate.NearbyWhisperWords = FormatNearbyWords(nearbyWords);
            estimate.MatchedLeftAnchorWord = FormatAnchor(anchors.Left);
            estimate.MatchedRightAnchorWord = FormatAnchor(anchors.Right);
            estimate.DiagnosticReason = diagnostic;

            TimeSpan start;
            TimeSpan end;
            if (TryBuildAnchoredRange(anchors, nearbyWords, contextStart, contextEnd, effectiveSettings, out start, out end))
            {
                estimate.Start = start;
                estimate.End = end;
                estimate.Confidence = anchors.Left != null && anchors.Right != null ? "MEDIUM" : "LOWER";
                estimate.UsedTokenPositionFallback = false;
                estimate.FallbackReason = "";
                return estimate;
            }

            SrtOnlyTimingEstimate fallback = EstimateByTokenPosition(occurrence, effectiveSettings, tokenIndex, tokenCount);
            fallback.NearbyWhisperWords = estimate.NearbyWhisperWords;
            fallback.MatchedLeftAnchorWord = estimate.MatchedLeftAnchorWord;
            fallback.MatchedRightAnchorWord = estimate.MatchedRightAnchorWord;
            fallback.DiagnosticReason = diagnostic;
            fallback.Confidence = "LOWER";
            fallback.UsedTokenPositionFallback = true;
            fallback.FallbackReason = diagnostic.ToString();
            return fallback;
        }

        private static SrtOnlyTimingEstimate EstimateByTokenPosition(
            ProfanityOccurrence occurrence,
            AudioCensorSettings settings,
            int tokenIndex,
            int tokenCount)
        {
            var estimate = new SrtOnlyTimingEstimate
            {
                TokenIndexInCue = tokenIndex,
                TokenCountInCue = tokenCount,
                Confidence = "LOWER",
                UsedTokenPositionFallback = true
            };

            TimeSpan cueDuration = occurrence.Cue.End - occurrence.Cue.Start;
            tokenCount = Math.Max(1, tokenCount);
            tokenIndex = Math.Max(1, Math.Min(tokenCount, tokenIndex));

            double midpointFraction = (tokenIndex - 0.5) / tokenCount;
            TimeSpan midpoint = occurrence.Cue.Start + TimeSpan.FromTicks((long)(cueDuration.Ticks * midpointFraction));
            TimeSpan tokenBucketDuration = TimeSpan.FromTicks(Math.Max(1, cueDuration.Ticks / tokenCount));
            TimeSpan duration = GetTargetDuration(settings);
            if (duration > tokenBucketDuration)
                duration = tokenBucketDuration;
            if (duration > cueDuration)
                duration = cueDuration;
            if (duration <= TimeSpan.Zero)
                duration = TimeSpan.FromMilliseconds(50);

            TimeSpan halfDuration = TimeSpan.FromTicks(duration.Ticks / 2);
            TimeSpan start = midpoint - halfDuration;
            TimeSpan end = start + duration;

            if (start < occurrence.Cue.Start)
            {
                start = occurrence.Cue.Start;
                end = start + duration;
            }

            if (end > occurrence.Cue.End)
            {
                end = occurrence.Cue.End;
                start = end - duration;
                if (start < occurrence.Cue.Start)
                    start = occurrence.Cue.Start;
            }

            estimate.Start = start;
            estimate.End = end;
            return estimate;
        }

        private static bool TryBuildAnchoredRange(
            SrtOnlyAnchorPair anchors,
            List<SrtOnlyWhisperWord> nearbyWords,
            TimeSpan contextStart,
            TimeSpan contextEnd,
            AudioCensorSettings settings,
            out TimeSpan start,
            out TimeSpan end)
        {
            start = TimeSpan.Zero;
            end = TimeSpan.Zero;
            if (anchors == null || !anchors.HasAny)
                return false;

            TimeSpan duration = GetTargetDuration(settings);
            TimeSpan midpoint;

            if (anchors.Left != null && anchors.Right != null)
            {
                TimeSpan leftBoundary = anchors.Left.WhisperWord.End;
                TimeSpan rightBoundary = anchors.Right.WhisperWord.Start;
                if (rightBoundary >= leftBoundary)
                {
                    midpoint = Midpoint(leftBoundary, rightBoundary);
                    TimeSpan gap = rightBoundary - leftBoundary;
                    if (gap > TimeSpan.Zero)
                        duration = ClampDuration(gap, settings);
                }
                else
                {
                    midpoint = Midpoint(anchors.Left.WhisperWord.Start, anchors.Right.WhisperWord.End);
                }
            }
            else if (anchors.Left != null)
            {
                SrtOnlyWhisperWord nextWord = nearbyWords.FirstOrDefault(w => w.Index > anchors.Left.WhisperWord.Index);
                TimeSpan leftBoundary = anchors.Left.WhisperWord.End;
                TimeSpan rightBoundary = nextWord != null && nextWord.Start > leftBoundary
                    ? nextWord.Start
                    : leftBoundary + duration;
                midpoint = Midpoint(leftBoundary, rightBoundary);
                if (rightBoundary > leftBoundary)
                    duration = ClampDuration(rightBoundary - leftBoundary, settings);
            }
            else
            {
                SrtOnlyWhisperWord previousWord = nearbyWords.LastOrDefault(w => w.Index < anchors.Right.WhisperWord.Index);
                TimeSpan rightBoundary = anchors.Right.WhisperWord.Start;
                TimeSpan leftBoundary = previousWord != null && previousWord.End < rightBoundary
                    ? previousWord.End
                    : rightBoundary - duration;
                if (leftBoundary < TimeSpan.Zero)
                    leftBoundary = TimeSpan.Zero;
                midpoint = Midpoint(leftBoundary, rightBoundary);
                if (rightBoundary > leftBoundary)
                    duration = ClampDuration(rightBoundary - leftBoundary, settings);
            }

            BuildRangeAroundMidpoint(midpoint, duration, contextStart, contextEnd, out start, out end);
            return end > start;
        }

        private static void BuildRangeAroundMidpoint(
            TimeSpan midpoint,
            TimeSpan duration,
            TimeSpan lowerBound,
            TimeSpan upperBound,
            out TimeSpan start,
            out TimeSpan end)
        {
            if (duration <= TimeSpan.Zero)
                duration = TimeSpan.FromMilliseconds(50);

            TimeSpan halfDuration = TimeSpan.FromTicks(duration.Ticks / 2);
            start = midpoint - halfDuration;
            end = start + duration;

            if (start < lowerBound)
            {
                start = lowerBound;
                end = start + duration;
            }

            if (upperBound > lowerBound && end > upperBound)
            {
                end = upperBound;
                start = end - duration;
                if (start < lowerBound)
                    start = lowerBound;
            }
        }

        private static TimeSpan GetTargetDuration(AudioCensorSettings settings)
        {
            TimeSpan minimum = settings.SrtOnlyMinDuration;
            TimeSpan maximum = settings.SrtOnlyMaxDuration;
            TimeSpan target = settings.SrtOnlyDefaultMuteDuration;
            if (target <= TimeSpan.Zero)
                target = minimum;

            return ClampDuration(target, settings);
        }

        private static TimeSpan ClampDuration(TimeSpan value, AudioCensorSettings settings)
        {
            TimeSpan minimum = settings.SrtOnlyMinDuration;
            TimeSpan maximum = settings.SrtOnlyMaxDuration;
            if (minimum <= TimeSpan.Zero)
                minimum = TimeSpan.FromMilliseconds(50);
            if (maximum < minimum)
                maximum = minimum;
            if (value < minimum)
                return minimum;
            if (value > maximum)
                return maximum;
            return value;
        }

        private static SrtOnlyDiagnosticReason DetermineDiagnosticReason(
            List<SrtOnlyWhisperWord> wordsInCue,
            SrtOnlyAnchorPair anchors,
            SrtOnlyWhisperWord similarWord,
            bool exactProfanityInCue,
            bool exactProfanityNearby)
        {
            if (exactProfanityNearby && !exactProfanityInCue)
                return SrtOnlyDiagnosticReason.WHISPER_TIMING_SHIFTED;

            if (similarWord != null)
                return SrtOnlyDiagnosticReason.WHISPER_WORD_NEARBY_SIMILAR_SOUND;

            if (wordsInCue == null || wordsInCue.Count == 0)
                return anchors != null && anchors.HasAny
                    ? SrtOnlyDiagnosticReason.POSSIBLE_SRT_TIMING_OFFSET
                    : SrtOnlyDiagnosticReason.NO_WHISPER_WORDS_IN_CUE;

            if (!exactProfanityInCue)
                return SrtOnlyDiagnosticReason.WHISPER_WORDS_EXIST_BUT_NO_PROFANITY;

            return SrtOnlyDiagnosticReason.UNKNOWN;
        }

        private static SrtOnlyAnchorPair FindAnchors(
            List<SrtOnlyTextToken> cueTokens,
            int profanityTokenIndex,
            List<SrtOnlyWhisperWord> nearbyWords)
        {
            var pair = new SrtOnlyAnchorPair();
            if (cueTokens == null || cueTokens.Count == 0 || nearbyWords == null || nearbyWords.Count == 0)
                return pair;

            List<SrtOnlyAnchorMatch> matches = BuildOrderedMatches(cueTokens, profanityTokenIndex, nearbyWords);
            pair.Left = matches
                .Where(m => m.TextToken.Index < profanityTokenIndex)
                .OrderByDescending(m => m.TextToken.Index)
                .FirstOrDefault();
            pair.Right = matches
                .Where(m => m.TextToken.Index > profanityTokenIndex)
                .OrderBy(m => m.TextToken.Index)
                .FirstOrDefault();
            return pair;
        }

        private static List<SrtOnlyAnchorMatch> BuildOrderedMatches(
            List<SrtOnlyTextToken> cueTokens,
            int profanityTokenIndex,
            List<SrtOnlyWhisperWord> nearbyWords)
        {
            int tokenCount = cueTokens.Count;
            int wordCount = nearbyWords.Count;
            int[,] score = new int[tokenCount + 1, wordCount + 1];

            for (int tokenIndex = tokenCount - 1; tokenIndex >= 0; tokenIndex--)
            {
                for (int wordIndex = wordCount - 1; wordIndex >= 0; wordIndex--)
                {
                    int best = Math.Max(score[tokenIndex + 1, wordIndex], score[tokenIndex, wordIndex + 1]);
                    if (CanUseAsAnchor(cueTokens[tokenIndex], tokenIndex, profanityTokenIndex) &&
                        cueTokens[tokenIndex].Normalized == nearbyWords[wordIndex].Normalized)
                    {
                        best = Math.Max(best, 1 + score[tokenIndex + 1, wordIndex + 1]);
                    }

                    score[tokenIndex, wordIndex] = best;
                }
            }

            var matches = new List<SrtOnlyAnchorMatch>();
            int i = 0;
            int j = 0;
            while (i < tokenCount && j < wordCount)
            {
                if (CanUseAsAnchor(cueTokens[i], i, profanityTokenIndex) &&
                    cueTokens[i].Normalized == nearbyWords[j].Normalized &&
                    score[i, j] == 1 + score[i + 1, j + 1])
                {
                    matches.Add(new SrtOnlyAnchorMatch
                    {
                        TextToken = cueTokens[i],
                        WhisperWord = nearbyWords[j]
                    });
                    i++;
                    j++;
                }
                else if (score[i + 1, j] >= score[i, j + 1])
                {
                    i++;
                }
                else
                {
                    j++;
                }
            }

            return matches;
        }

        private static bool CanUseAsAnchor(SrtOnlyTextToken token, int tokenIndex, int profanityTokenIndex)
        {
            return token != null &&
                   tokenIndex != profanityTokenIndex &&
                   !string.IsNullOrWhiteSpace(token.Normalized);
        }

        private static List<SrtOnlyTextToken> TokenizeCueText(string text)
        {
            var tokens = new List<SrtOnlyTextToken>();
            foreach (Match match in WordRegex.Matches(text ?? ""))
            {
                string normalized = ProfanityDictionary.NormalizeToken(match.Value);
                if (string.IsNullOrWhiteSpace(normalized))
                    continue;

                tokens.Add(new SrtOnlyTextToken
                {
                    Index = tokens.Count,
                    Text = match.Value,
                    Normalized = normalized
                });
            }

            return tokens;
        }

        private static int ResolveTokenIndex(ProfanityOccurrence occurrence, List<SrtOnlyTextToken> cueTokens, int tokenCount)
        {
            if (occurrence.TokenIndexInCue > 0)
                return Math.Max(1, Math.Min(Math.Max(1, tokenCount), occurrence.TokenIndexInCue));

            HashSet<string> profanityKeys = GetProfanityKeys(occurrence);
            for (int i = 0; i < cueTokens.Count; i++)
            {
                if (profanityKeys.Contains(cueTokens[i].Normalized))
                    return i + 1;
            }

            return 1;
        }

        private static List<SrtOnlyWhisperWord> GetNearbyWhisperWords(
            IEnumerable<TimedTranscriptWord> whisperWords,
            TimeSpan contextStart,
            TimeSpan contextEnd)
        {
            var words = new List<SrtOnlyWhisperWord>();
            foreach (TimedTranscriptWord word in whisperWords ?? Enumerable.Empty<TimedTranscriptWord>())
            {
                if (word == null || word.End <= word.Start || word.End < contextStart || word.Start > contextEnd)
                    continue;

                string normalized = string.IsNullOrWhiteSpace(word.NormalizedWord)
                    ? ProfanityDictionary.NormalizeToken(word.Word)
                    : word.NormalizedWord;
                if (string.IsNullOrWhiteSpace(normalized))
                    continue;

                words.Add(new SrtOnlyWhisperWord
                {
                    Index = words.Count,
                    Text = string.IsNullOrWhiteSpace(word.Word) ? normalized : word.Word.Trim(),
                    Normalized = normalized,
                    Start = word.Start,
                    End = word.End
                });
            }

            return words
                .OrderBy(w => w.Start)
                .ThenBy(w => w.End)
                .Select((w, i) =>
                {
                    w.Index = i;
                    return w;
                })
                .ToList();
        }

        private static bool ContainsProfanityWord(ProfanityOccurrence occurrence, IEnumerable<SrtOnlyWhisperWord> words)
        {
            HashSet<string> profanityKeys = GetProfanityKeys(occurrence);
            return (words ?? Enumerable.Empty<SrtOnlyWhisperWord>()).Any(w => w != null && profanityKeys.Contains(w.Normalized));
        }

        private static SrtOnlyWhisperWord FindSimilarSoundingWord(ProfanityOccurrence occurrence, IEnumerable<SrtOnlyWhisperWord> words)
        {
            HashSet<string> profanityKeys = GetProfanityKeys(occurrence);
            foreach (SrtOnlyWhisperWord word in words ?? Enumerable.Empty<SrtOnlyWhisperWord>())
            {
                if (word == null || profanityKeys.Contains(word.Normalized))
                    continue;

                foreach (string key in profanityKeys)
                {
                    if (IsSimilarSound(key, word.Normalized))
                        return word;
                }
            }

            return null;
        }

        private static bool IsSimilarSound(string expected, string actual)
        {
            if (string.IsNullOrWhiteSpace(expected) || string.IsNullOrWhiteSpace(actual) || expected == actual)
                return false;

            string left = CollapseRepeatedLetters(expected);
            string right = CollapseRepeatedLetters(actual);
            if (left == right)
                return true;

            int lengthDelta = Math.Abs(left.Length - right.Length);
            if (lengthDelta > 2)
                return false;

            int distance = LevenshteinDistance(left, right);
            if (Math.Min(left.Length, right.Length) <= 4)
                return distance <= 1;

            return distance <= 2 && left[0] == right[0];
        }

        private static string CollapseRepeatedLetters(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            var chars = new List<char>();
            char previous = '\0';
            foreach (char c in value)
            {
                if (c == previous)
                    continue;

                chars.Add(c);
                previous = c;
            }

            return new string(chars.ToArray());
        }

        private static int LevenshteinDistance(string left, string right)
        {
            int[,] distances = new int[left.Length + 1, right.Length + 1];
            for (int i = 0; i <= left.Length; i++)
                distances[i, 0] = i;
            for (int j = 0; j <= right.Length; j++)
                distances[0, j] = j;

            for (int i = 1; i <= left.Length; i++)
            {
                for (int j = 1; j <= right.Length; j++)
                {
                    int cost = left[i - 1] == right[j - 1] ? 0 : 1;
                    distances[i, j] = Math.Min(
                        Math.Min(distances[i - 1, j] + 1, distances[i, j - 1] + 1),
                        distances[i - 1, j - 1] + cost);
                }
            }

            return distances[left.Length, right.Length];
        }

        private static HashSet<string> GetProfanityKeys(ProfanityOccurrence occurrence)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AddNormalizedKey(keys, occurrence == null ? "" : occurrence.Word);
            AddNormalizedKey(keys, occurrence == null ? "" : occurrence.DictionaryTerm);
            return keys;
        }

        private static void AddNormalizedKey(HashSet<string> keys, string value)
        {
            string normalized = ProfanityDictionary.NormalizeToken(value);
            if (!string.IsNullOrWhiteSpace(normalized))
                keys.Add(normalized);
        }

        private static string FormatNearbyWords(IReadOnlyList<SrtOnlyWhisperWord> words)
        {
            if (words == null || words.Count == 0)
                return "(none)";

            string text = string.Join("; ", words
                .Take(NearbyWordReportLimit)
                .Select(w => w.Text + "@" + FormatTime(w.Start) + "-" + FormatTime(w.End))
                .ToArray());
            if (words.Count > NearbyWordReportLimit)
                text += "; +" + (words.Count - NearbyWordReportLimit) + " more";

            return text;
        }

        private static string FormatAnchor(SrtOnlyAnchorMatch anchor)
        {
            if (anchor == null || anchor.TextToken == null || anchor.WhisperWord == null)
                return "";

            return anchor.TextToken.Text + "=" + anchor.WhisperWord.Text + "@" + FormatTime(anchor.WhisperWord.Start) + "-" + FormatTime(anchor.WhisperWord.End);
        }

        private static TimeSpan Midpoint(TimeSpan start, TimeSpan end)
        {
            return TimeSpan.FromTicks((start.Ticks + end.Ticks) / 2);
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

        private sealed class SrtOnlyTextToken
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public string Normalized { get; set; }
        }

        private sealed class SrtOnlyWhisperWord
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public string Normalized { get; set; }
            public TimeSpan Start { get; set; }
            public TimeSpan End { get; set; }
        }

        private sealed class SrtOnlyAnchorMatch
        {
            public SrtOnlyTextToken TextToken { get; set; }
            public SrtOnlyWhisperWord WhisperWord { get; set; }
        }

        private sealed class SrtOnlyAnchorPair
        {
            public SrtOnlyAnchorMatch Left { get; set; }
            public SrtOnlyAnchorMatch Right { get; set; }

            public bool HasAny
            {
                get { return Left != null || Right != null; }
            }
        }
    }

    public sealed class SrtOnlyTimingEstimate
    {
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public int TokenIndexInCue { get; set; }
        public int TokenCountInCue { get; set; }
        public string NearbyWhisperWords { get; set; }
        public string MatchedLeftAnchorWord { get; set; }
        public string MatchedRightAnchorWord { get; set; }
        public SrtOnlyDiagnosticReason DiagnosticReason { get; set; }
        public string Confidence { get; set; }
        public bool UsedTokenPositionFallback { get; set; }
        public string FallbackReason { get; set; }

        public SrtOnlyTimingEstimate()
        {
            NearbyWhisperWords = "";
            MatchedLeftAnchorWord = "";
            MatchedRightAnchorWord = "";
            DiagnosticReason = SrtOnlyDiagnosticReason.UNKNOWN;
            Confidence = "LOWER";
            FallbackReason = "";
        }
    }
    public sealed class AudioCensorReportWriter
    {
        public string Write(AudioCensorScanSummary summary)
        {
            if (summary == null)
                throw new ArgumentNullException(nameof(summary));

            string path = BuildReportPath(summary);
            var lines = new List<string>();
            lines.Add("BachFlixNfo Audio Censor Hybrid Detection Report");
            lines.Add("Created: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            lines.Add("");
            lines.Add("Input MKV: " + summary.MkvPath);
            lines.Add("WhisperX JSON: " + GetPrimaryTranscriptPath(summary));
            lines.Add("WhisperX hits: " + Count(summary.TranscriptProfanityHits));
            lines.Add("Subtitle hits: " + CountSubtitleHits(summary));
            lines.Add("SRT-only hits: " + Count(summary.SrtOnlyEstimatedHits));
            lines.Add("Approved mute hits: " + Count(summary.ApprovedTranscriptHits));
            lines.Add("");
            AddHitSection(lines, "WhisperX hits", summary.TranscriptProfanityHits);
            AddSubtitleHitSection(lines, summary);
            AddHitSection(lines, "SRT-only hits", summary.SrtOnlyEstimatedHits);
            AddSkippedSubtitleSection(lines, summary);
            AddFinalIntervalSection(lines, summary);

            string directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            File.WriteAllLines(path, lines);
            return path;
        }

        private static void AddHitSection(List<string> lines, string title, IReadOnlyList<TranscriptProfanityHit> hits)
        {
            lines.Add(title + ":");
            if (hits == null || hits.Count == 0)
            {
                lines.Add("  (none)");
                lines.Add("");
                return;
            }

            foreach (TranscriptProfanityHit hit in hits)
                lines.Add("  " + DescribeHit(hit));
            lines.Add("");
        }

        private static void AddSubtitleHitSection(List<string> lines, AudioCensorScanSummary summary)
        {
            lines.Add("Subtitle hits:");
            bool any = false;
            foreach (SubtitleSourceInfo source in summary.EnglishSubtitleSources ?? new List<SubtitleSourceInfo>())
            {
                if (source == null || source.ScanResult == null || source.ScanResult.Occurrences == null || source.ScanResult.Occurrences.Count == 0)
                    continue;

                any = true;
                foreach (ProfanityOccurrence occurrence in source.ScanResult.Occurrences)
                    lines.Add("  " + DescribeOccurrence(source, occurrence));
            }

            if (!any)
                lines.Add("  (none)");
            lines.Add("");
        }

        private static void AddSkippedSubtitleSection(List<string> lines, AudioCensorScanSummary summary)
        {
            lines.Add("Skipped subtitle hits:");
            bool any = false;
            foreach (SubtitleSourceInfo source in summary.EnglishSubtitleSources ?? new List<SubtitleSourceInfo>())
            {
                if (source == null || source.ScanResult == null || source.ScanResult.Occurrences == null)
                    continue;

                HashSet<ProfanityOccurrence> srtOnly = new HashSet<ProfanityOccurrence>();
                if (source.Comparison != null && source.Comparison.ProfanityFoundOnlyInSubtitle != null)
                {
                    foreach (ProfanityOccurrence occurrence in source.Comparison.ProfanityFoundOnlyInSubtitle)
                        srtOnly.Add(occurrence);
                }

                foreach (ProfanityOccurrence occurrence in source.ScanResult.Occurrences)
                {
                    if (srtOnly.Contains(occurrence))
                        continue;

                    any = true;
                    lines.Add("  covered by WHISPERX: " + DescribeOccurrence(source, occurrence));
                }
            }

            if (!any)
                lines.Add("  (none)");
            lines.Add("");
        }

        private static void AddFinalIntervalSection(List<string> lines, AudioCensorScanSummary summary)
        {
            lines.Add("Final approved mute intervals:");
            if (summary.CensorPlan == null || summary.CensorPlan.AudioTrackPlans == null || summary.CensorPlan.AudioTrackPlans.Count == 0)
            {
                lines.Add("  (none)");
                lines.Add("");
                return;
            }

            foreach (AudioTrackCensorPlan trackPlan in summary.CensorPlan.AudioTrackPlans)
            {
                foreach (AudioMuteSegment segment in trackPlan.MuteSegments)
                {
                    string sources = string.Join(", ", segment.SourceHits
                        .Where(h => h != null)
                        .Select(h => h.DetectionSource + ":" + h.Word)
                        .Distinct()
                        .ToArray());
                    if (string.IsNullOrWhiteSpace(sources) && segment.SourceHit != null)
                        sources = segment.SourceHit.DetectionSource + ":" + segment.SourceHit.Word;

                    lines.Add("  audio " + trackPlan.Track.AudioTrackIndex + " " + FormatTime(segment.Start) + " --> " + FormatTime(segment.End) + " | sources=" + sources);
                }
            }
            lines.Add("");
        }

        private static string DescribeHit(TranscriptProfanityHit hit)
        {
            if (hit == null)
                return "(missing hit)";

            string subtitle = string.IsNullOrWhiteSpace(hit.SubtitleSourceLabel) ? "" : " | subtitle=" + hit.SubtitleSourceLabel;
            string estimated = hit.IsEstimatedTiming
                ? " | estimated token=" + hit.EstimatedTokenIndexInCue + "/" + hit.EstimatedTokenCountInCue + " cue=" + FormatTime(hit.SubtitleCueStart) + "->" + FormatTime(hit.SubtitleCueEnd)
                : "";
            if (hit.DetectionSource == ProfanityDetectionSource.SRT_ONLY_ESTIMATED)
            {
                estimated += " | subtitle text=\"" + CleanInline(hit.SubtitleCueText) + "\"";
                estimated += " | nearby WhisperX=" + CleanInline(hit.NearbyWhisperWords);
                estimated += " | anchor left=" + EmptyAsNone(hit.MatchedLeftAnchorWord);
                estimated += " | anchor right=" + EmptyAsNone(hit.MatchedRightAnchorWord);
                estimated += " | diagnostic=" + hit.SrtOnlyDiagnosticReason;
                estimated += " | token fallback=" + (hit.UsedTokenPositionFallback ? "yes" : "no");
                if (hit.UsedTokenPositionFallback && !string.IsNullOrWhiteSpace(hit.SrtOnlyFallbackReason))
                    estimated += " | fallback reason=" + hit.SrtOnlyFallbackReason;
            }

            return hit.DetectionSource + " " + FormatTime(hit.Start) + " --> " + FormatTime(hit.End) + " | word=" + hit.Word + " | dictionary=" + hit.DictionaryTerm + " | confidence=" + hit.Confidence + " | padding=" + (int)hit.PaddingBefore.TotalMilliseconds + "ms/" + (int)hit.PaddingAfter.TotalMilliseconds + "ms" + subtitle + estimated + " | approved=" + (hit.Approved ? "yes" : "no");
        }

        private static string DescribeOccurrence(SubtitleSourceInfo source, ProfanityOccurrence occurrence)
        {
            if (occurrence == null || occurrence.Cue == null)
                return "(missing subtitle occurrence)";

            string sourceLabel = source == null ? "subtitle" : source.Describe();
            return sourceLabel + " | cue=" + occurrence.SubtitleSequenceNumber + " " + FormatTime(occurrence.Cue.Start) + " --> " + FormatTime(occurrence.Cue.End) + " | token=" + occurrence.TokenIndexInCue + "/" + occurrence.TokenCountInCue + " | word=" + occurrence.Word + " | dictionary=" + occurrence.DictionaryTerm;
        }

        private static string BuildReportPath(AudioCensorScanSummary summary)
        {
            string transcriptPath = GetPrimaryTranscriptPath(summary);
            string directory = string.IsNullOrWhiteSpace(transcriptPath) ? "" : Path.GetDirectoryName(transcriptPath);
            string baseName = string.IsNullOrWhiteSpace(transcriptPath) ? "" : Path.GetFileNameWithoutExtension(transcriptPath);

            if (string.IsNullOrWhiteSpace(directory))
                directory = Path.GetDirectoryName(summary.MkvPath) ?? "";
            if (string.IsNullOrWhiteSpace(baseName))
                baseName = Path.GetFileNameWithoutExtension(summary.MkvPath);
            if (string.IsNullOrWhiteSpace(baseName))
                baseName = "AudioCensor";

            return Path.Combine(directory, baseName + ".audio-censor-report.txt");
        }

        private static string GetPrimaryTranscriptPath(AudioCensorScanSummary summary)
        {
            if (summary == null || summary.PrimaryTranscriptionTrack == null)
                return "";

            WordTranscriptionResult result;
            if (summary.TranscriptionResults != null && summary.TranscriptionResults.TryGetValue(summary.PrimaryTranscriptionTrack.AudioTrackIndex, out result))
                return result.TranscriptJsonPath;

            return "";
        }

        private static int Count<T>(IReadOnlyList<T> values)
        {
            return values == null ? 0 : values.Count;
        }

        private static int CountSubtitleHits(AudioCensorScanSummary summary)
        {
            if (summary == null || summary.EnglishSubtitleSources == null)
                return 0;

            return summary.EnglishSubtitleSources
                .Where(s => s != null && s.ScanResult != null && s.ScanResult.Occurrences != null)
                .Sum(s => s.ScanResult.Occurrences.Count);
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

        private static string CleanInline(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? "(none)"
                : value.Replace("\r", " ").Replace("\n", " ").Trim();
        }

        private static string EmptyAsNone(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "(none)" : CleanInline(value);
        }
    }
    internal sealed class SubtitleOnlyProfanityCandidate
    {
        public int ReviewNumber { get; set; }
        public SubtitleSourceInfo Source { get; set; }
        public ProfanityOccurrence Occurrence { get; set; }
    }
    public sealed class AudioCensorBatchSummary
    {
        public string InputPath { get; set; }
        public string ProfanityDictionaryPath { get; set; }
        public int TotalFiles { get; set; }
        public int CleanedFiles { get; set; }
        public int SkippedFiles { get; set; }
        public int FailedFiles { get; set; }
        public List<string> OutputPaths { get; private set; }
        public List<string> SkippedInputs { get; private set; }
        public List<string> FailedInputs { get; private set; }

        public AudioCensorBatchSummary()
        {
            InputPath = "";
            ProfanityDictionaryPath = "";
            OutputPaths = new List<string>();
            SkippedInputs = new List<string>();
            FailedInputs = new List<string>();
        }
    }
    public sealed class AudioCensorScanOptions
    {
        public string MkvPath { get; set; }
        public string ProfanityDictionaryPath { get; set; }
        public bool DetectExternalTools { get; set; }
        public bool PromptForMissingEnglishAudio { get; set; }
        public bool UseAutomaticAudioTrackFallback { get; set; }
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
        public List<TranscriptProfanityHit> ProposedAudioMuteHits { get; set; }
        public List<TranscriptProfanityHit> SrtOnlyEstimatedHits { get; set; }
        public List<TranscriptProfanityHit> ApprovedTranscriptHits { get; set; }
        public List<SubtitleSourceInfo> SubtitleSources { get; set; }
        public List<SubtitleSourceInfo> EnglishSubtitleSources { get; set; }
        public string SubtitleWorkDirectory { get; set; }
        public string OutputMkvPath { get; set; }
        public string ComparisonReportPath { get; set; }
        public AudioCensorSettings Settings { get; set; }
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
            ProposedAudioMuteHits = new List<TranscriptProfanityHit>();
            SrtOnlyEstimatedHits = new List<TranscriptProfanityHit>();
            ApprovedTranscriptHits = new List<TranscriptProfanityHit>();
            SubtitleSources = new List<SubtitleSourceInfo>();
            EnglishSubtitleSources = new List<SubtitleSourceInfo>();
            SubtitleWorkDirectory = "";
            OutputMkvPath = "";
            ComparisonReportPath = "";
            Settings = AudioCensorSettings.LoadFromEnvironment();
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
