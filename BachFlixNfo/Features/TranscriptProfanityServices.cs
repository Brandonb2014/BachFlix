using System;
using System.Collections.Generic;

namespace BachFlixNfo.Features
{
    public enum ProfanityDetectionSource
    {
        WHISPERX,
        SRT_ONLY_ESTIMATED,
        MANUAL
    }

    public enum SrtOnlyDiagnosticReason
    {
        NO_WHISPER_WORDS_IN_CUE,
        WHISPER_WORDS_EXIST_BUT_NO_PROFANITY,
        WHISPER_WORD_NEARBY_SIMILAR_SOUND,
        WHISPER_TIMING_SHIFTED,
        POSSIBLE_SRT_TIMING_OFFSET,
        UNKNOWN
    }

    public enum SubtitleReplacementMode
    {
        AsterisksSameLength,
        FixedText
    }

    public sealed class AudioCensorSettings
    {
        public int SrtOnlyDefaultMuteDurationMs { get; set; }
        public int SrtOnlyMinDurationMs { get; set; }
        public int SrtOnlyMaxDurationMs { get; set; }
        public int SrtOnlyPrePaddingMs { get; set; }
        public int SrtOnlyPostPaddingMs { get; set; }
        public int SrtOnlyContextWindowBeforeMs { get; set; }
        public int SrtOnlyContextWindowAfterMs { get; set; }
        public int WhisperPrePaddingMs { get; set; }
        public int WhisperPostPaddingMs { get; set; }
        public int MergeGapMs { get; set; }
        public SubtitleReplacementMode SubtitleReplacementMode { get; set; }
        public string SubtitleReplacementText { get; set; }

        public TimeSpan SrtOnlyDefaultMuteDuration
        {
            get { return TimeSpan.FromMilliseconds(SrtOnlyDefaultMuteDurationMs); }
        }

        public TimeSpan SrtOnlyMinDuration
        {
            get { return TimeSpan.FromMilliseconds(SrtOnlyMinDurationMs); }
        }

        public TimeSpan SrtOnlyMaxDuration
        {
            get { return TimeSpan.FromMilliseconds(SrtOnlyMaxDurationMs); }
        }

        public TimeSpan SrtOnlyPrePadding
        {
            get { return TimeSpan.FromMilliseconds(SrtOnlyPrePaddingMs); }
        }

        public TimeSpan SrtOnlyPostPadding
        {
            get { return TimeSpan.FromMilliseconds(SrtOnlyPostPaddingMs); }
        }

        public TimeSpan SrtOnlyContextWindowBefore
        {
            get { return TimeSpan.FromMilliseconds(SrtOnlyContextWindowBeforeMs); }
        }

        public TimeSpan SrtOnlyContextWindowAfter
        {
            get { return TimeSpan.FromMilliseconds(SrtOnlyContextWindowAfterMs); }
        }

        public TimeSpan WhisperPrePadding
        {
            get { return TimeSpan.FromMilliseconds(WhisperPrePaddingMs); }
        }

        public TimeSpan WhisperPostPadding
        {
            get { return TimeSpan.FromMilliseconds(WhisperPostPaddingMs); }
        }

        public TimeSpan MergeGap
        {
            get { return TimeSpan.FromMilliseconds(MergeGapMs); }
        }

        public AudioCensorSettings()
        {
            SrtOnlyDefaultMuteDurationMs = 180;
            SrtOnlyMinDurationMs = 180;
            SrtOnlyMaxDurationMs = 450;
            SrtOnlyPrePaddingMs = 80;
            SrtOnlyPostPaddingMs = 120;
            SrtOnlyContextWindowBeforeMs = 2000;
            SrtOnlyContextWindowAfterMs = 2000;
            WhisperPrePaddingMs = (int)AudioMuteFilterBuilder.DefaultPaddingBefore.TotalMilliseconds;
            WhisperPostPaddingMs = (int)AudioMuteFilterBuilder.DefaultPaddingAfter.TotalMilliseconds;
            MergeGapMs = 60;
            SubtitleReplacementMode = SubtitleReplacementMode.AsterisksSameLength;
            SubtitleReplacementText = "[censored]";
        }

        public static AudioCensorSettings LoadFromEnvironment()
        {
            var settings = new AudioCensorSettings();
            settings.SrtOnlyMinDurationMs = GetIntSetting("AUDIO_CENSOR_SRT_ONLY_MIN_DURATION_MS", "SrtOnlyMinDurationMs", settings.SrtOnlyMinDurationMs);
            settings.SrtOnlyMaxDurationMs = GetIntSetting("AUDIO_CENSOR_SRT_ONLY_MAX_DURATION_MS", "SrtOnlyMaxDurationMs", settings.SrtOnlyMaxDurationMs);
            settings.SrtOnlyDefaultMuteDurationMs = GetIntSetting("AUDIO_CENSOR_SRT_ONLY_DEFAULT_MUTE_DURATION_MS", "SrtOnlyDefaultMuteDurationMs", settings.SrtOnlyDefaultMuteDurationMs);
            settings.SrtOnlyPrePaddingMs = GetIntSetting("AUDIO_CENSOR_SRT_ONLY_PRE_PADDING_MS", "SrtOnlyPrePaddingMs", settings.SrtOnlyPrePaddingMs);
            settings.SrtOnlyPostPaddingMs = GetIntSetting("AUDIO_CENSOR_SRT_ONLY_POST_PADDING_MS", "SrtOnlyPostPaddingMs", settings.SrtOnlyPostPaddingMs);
            settings.SrtOnlyContextWindowBeforeMs = GetIntSetting("AUDIO_CENSOR_SRT_ONLY_CONTEXT_WINDOW_BEFORE_MS", "SrtOnlyContextWindowBeforeMs", settings.SrtOnlyContextWindowBeforeMs);
            settings.SrtOnlyContextWindowAfterMs = GetIntSetting("AUDIO_CENSOR_SRT_ONLY_CONTEXT_WINDOW_AFTER_MS", "SrtOnlyContextWindowAfterMs", settings.SrtOnlyContextWindowAfterMs);
            settings.WhisperPrePaddingMs = GetIntSetting("AUDIO_CENSOR_WHISPER_PRE_PADDING_MS", "WhisperPrePaddingMs", settings.WhisperPrePaddingMs);
            settings.WhisperPostPaddingMs = GetIntSetting("AUDIO_CENSOR_WHISPER_POST_PADDING_MS", "WhisperPostPaddingMs", settings.WhisperPostPaddingMs);
            settings.MergeGapMs = GetIntSetting("AUDIO_CENSOR_MERGE_GAP_MS", "MergeGapMs", settings.MergeGapMs);

            string mode = GetSetting("AUDIO_CENSOR_SUBTITLE_REPLACEMENT_MODE", "SubtitleReplacementMode");
            string text = GetSetting("AUDIO_CENSOR_SUBTITLE_REPLACEMENT_TEXT", "SubtitleReplacementText");
            string legacyReplacement = Environment.GetEnvironmentVariable("AUDIO_CENSOR_SUBTITLE_REPLACEMENT");

            if (string.IsNullOrWhiteSpace(mode) && !string.IsNullOrWhiteSpace(legacyReplacement))
            {
                if (legacyReplacement.Equals("asterisks", StringComparison.OrdinalIgnoreCase) || legacyReplacement == "*")
                {
                    mode = "AsterisksSameLength";
                }
                else
                {
                    mode = "FixedText";
                    if (string.IsNullOrWhiteSpace(text))
                        text = legacyReplacement;
                }
            }

            if (mode.Equals("FixedText", StringComparison.OrdinalIgnoreCase))
                settings.SubtitleReplacementMode = SubtitleReplacementMode.FixedText;
            else
                settings.SubtitleReplacementMode = SubtitleReplacementMode.AsterisksSameLength;

            if (!string.IsNullOrWhiteSpace(text))
                settings.SubtitleReplacementText = text.Trim();

            settings.SrtOnlyMinDurationMs = Math.Max(50, settings.SrtOnlyMinDurationMs);
            settings.SrtOnlyMaxDurationMs = Math.Max(settings.SrtOnlyMinDurationMs, settings.SrtOnlyMaxDurationMs);
            settings.SrtOnlyDefaultMuteDurationMs = Math.Max(settings.SrtOnlyMinDurationMs, Math.Min(settings.SrtOnlyMaxDurationMs, settings.SrtOnlyDefaultMuteDurationMs));
            settings.SrtOnlyPrePaddingMs = Math.Max(0, settings.SrtOnlyPrePaddingMs);
            settings.SrtOnlyPostPaddingMs = Math.Max(0, settings.SrtOnlyPostPaddingMs);
            settings.SrtOnlyContextWindowBeforeMs = Math.Max(0, settings.SrtOnlyContextWindowBeforeMs);
            settings.SrtOnlyContextWindowAfterMs = Math.Max(0, settings.SrtOnlyContextWindowAfterMs);
            settings.WhisperPrePaddingMs = Math.Max(0, settings.WhisperPrePaddingMs);
            settings.WhisperPostPaddingMs = Math.Max(0, settings.WhisperPostPaddingMs);
            settings.MergeGapMs = Math.Max(0, settings.MergeGapMs);
            return settings;
        }

        private static int GetIntSetting(string environmentName, string alternateName, int fallback)
        {
            string value = GetSetting(environmentName, alternateName);
            int parsed;
            return int.TryParse(value, out parsed) ? parsed : fallback;
        }

        private static string GetSetting(string environmentName, string alternateName)
        {
            string value = Environment.GetEnvironmentVariable(environmentName);
            if (string.IsNullOrWhiteSpace(value))
                value = Environment.GetEnvironmentVariable(alternateName);

            return string.IsNullOrWhiteSpace(value) ? "" : value.Trim();
        }
    }

    public sealed class TranscriptProfanityScanner
    {
        public TranscriptProfanityScanResult Scan(
            WordTranscriptionResult transcript,
            ProfanityDictionary dictionary,
            int audioTrackIndex,
            string sourceLabel)
        {
            if (transcript == null)
                throw new ArgumentNullException(nameof(transcript));

            if (dictionary == null)
                throw new ArgumentNullException(nameof(dictionary));

            var result = new TranscriptProfanityScanResult
            {
                SourceAudioTrackIndex = audioTrackIndex,
                SourceLabel = sourceLabel ?? ""
            };

            foreach (TimedTranscriptWord word in transcript.Words)
            {
                string dictionaryTerm;
                if (!dictionary.TryMatch(word.Word, out dictionaryTerm))
                    continue;

                var hit = new TranscriptProfanityHit
                {
                    ReviewNumber = result.Hits.Count + 1,
                    Word = word.Word,
                    DictionaryTerm = dictionaryTerm,
                    Start = word.Start,
                    End = word.End,
                    SourceAudioTrackIndex = audioTrackIndex,
                    SourceLabel = sourceLabel ?? "",
                    TranscriptJsonPath = transcript.TranscriptJsonPath,
                    DetectionSource = ProfanityDetectionSource.WHISPERX,
                    Confidence = "HIGH",
                    Approved = true
                };
                hit.AppliesToAudioTrackIndexes.Add(audioTrackIndex);
                result.Hits.Add(hit);
            }

            return result;
        }
    }

    public sealed class TranscriptProfanityScanResult
    {
        public int SourceAudioTrackIndex { get; set; }
        public string SourceLabel { get; set; }
        public List<TranscriptProfanityHit> Hits { get; private set; }

        public TranscriptProfanityScanResult()
        {
            SourceLabel = "";
            Hits = new List<TranscriptProfanityHit>();
        }
    }

    public sealed class TranscriptProfanityHit
    {
        public int ReviewNumber { get; set; }
        public string Word { get; set; }
        public string DictionaryTerm { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public TimeSpan PaddingBefore { get; set; }
        public TimeSpan PaddingAfter { get; set; }
        public int SourceAudioTrackIndex { get; set; }
        public string SourceLabel { get; set; }
        public string TranscriptJsonPath { get; set; }
        public ProfanityDetectionSource DetectionSource { get; set; }
        public string Confidence { get; set; }
        public string SubtitleSourceLabel { get; set; }
        public string SubtitleSourcePath { get; set; }
        public int SubtitleStreamIndex { get; set; }
        public int SubtitleCueIndex { get; set; }
        public int SubtitleSequenceNumber { get; set; }
        public TimeSpan SubtitleCueStart { get; set; }
        public TimeSpan SubtitleCueEnd { get; set; }
        public int EstimatedTokenIndexInCue { get; set; }
        public int EstimatedTokenCountInCue { get; set; }
        public string SubtitleCueText { get; set; }
        public string NearbyWhisperWords { get; set; }
        public string MatchedLeftAnchorWord { get; set; }
        public string MatchedRightAnchorWord { get; set; }
        public SrtOnlyDiagnosticReason SrtOnlyDiagnosticReason { get; set; }
        public bool UsedTokenPositionFallback { get; set; }
        public string SrtOnlyFallbackReason { get; set; }
        public TimeSpan MinMuteDuration { get; set; }
        public TimeSpan MaxMuteDuration { get; set; }
        public bool IsEstimatedTiming { get; set; }
        public bool Approved { get; set; }
        public bool IsManual { get; set; }
        public List<int> AppliesToAudioTrackIndexes { get; private set; }

        public TranscriptProfanityHit()
        {
            Word = "";
            DictionaryTerm = "";
            SourceLabel = "";
            TranscriptJsonPath = "";
            PaddingBefore = AudioMuteFilterBuilder.DefaultPaddingBefore;
            PaddingAfter = AudioMuteFilterBuilder.DefaultPaddingAfter;
            DetectionSource = ProfanityDetectionSource.WHISPERX;
            Confidence = "HIGH";
            SubtitleSourceLabel = "";
            SubtitleSourcePath = "";
            SubtitleStreamIndex = -1;
            SubtitleCueText = "";
            NearbyWhisperWords = "";
            MatchedLeftAnchorWord = "";
            MatchedRightAnchorWord = "";
            SrtOnlyDiagnosticReason = SrtOnlyDiagnosticReason.UNKNOWN;
            SrtOnlyFallbackReason = "";
            Approved = true;
            AppliesToAudioTrackIndexes = new List<int>();
        }
    }
}
