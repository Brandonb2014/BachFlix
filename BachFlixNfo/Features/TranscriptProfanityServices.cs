using System;
using System.Collections.Generic;

namespace BachFlixNfo.Features
{
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
            Approved = true;
            AppliesToAudioTrackIndexes = new List<int>();
        }
    }
}
