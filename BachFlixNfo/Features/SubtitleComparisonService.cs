using System;
using System.Collections.Generic;
using System.Linq;

namespace BachFlixNfo.Features
{
    public sealed class SubtitleComparisonService
    {
        private static readonly TimeSpan MatchTolerance = TimeSpan.FromSeconds(2);

        public SubtitleCoverageComparison Compare(
            SubtitleSourceInfo source,
            IReadOnlyList<TranscriptProfanityHit> transcriptHits,
            SubtitleScanResult subtitleScan)
        {
            var comparison = new SubtitleCoverageComparison { Source = source };
            List<TranscriptProfanityHit> hits = (transcriptHits ?? new List<TranscriptProfanityHit>()).ToList();
            List<ProfanityOccurrence> occurrences = subtitleScan == null ? new List<ProfanityOccurrence>() : subtitleScan.Occurrences;

            foreach (TranscriptProfanityHit hit in hits)
            {
                if (occurrences.Any(o => OccurrenceMatchesHit(o, hit)))
                    comparison.MatchedTranscriptHits.Add(hit);
                else
                    comparison.ProfanityMissingFromSubtitle.Add(hit);
            }

            foreach (ProfanityOccurrence occurrence in occurrences)
            {
                if (!hits.Any(h => OccurrenceMatchesHit(occurrence, h)))
                    comparison.ProfanityFoundOnlyInSubtitle.Add(occurrence);
            }

            comparison.TranscriptHitCount = hits.Count;
            comparison.SubtitleHitCount = occurrences.Count;
            comparison.SubtitleCueCount = subtitleScan == null ? 0 : subtitleScan.Cues.Count;
            comparison.TranscriptCoveragePercent = hits.Count == 0 ? 100 : (comparison.MatchedTranscriptHits.Count * 100.0 / hits.Count);
            comparison.SubtitleAgreementPercent = occurrences.Count == 0 ? 100 : ((occurrences.Count - comparison.ProfanityFoundOnlyInSubtitle.Count) * 100.0 / occurrences.Count);
            return comparison;
        }

        private static bool OccurrenceMatchesHit(ProfanityOccurrence occurrence, TranscriptProfanityHit hit)
        {
            if (occurrence == null || occurrence.Cue == null || hit == null)
                return false;

            string occurrenceTerm = ProfanityDictionary.NormalizeToken(string.IsNullOrWhiteSpace(occurrence.DictionaryTerm) ? occurrence.Word : occurrence.DictionaryTerm);
            string hitTerm = ProfanityDictionary.NormalizeToken(string.IsNullOrWhiteSpace(hit.DictionaryTerm) ? hit.Word : hit.DictionaryTerm);
            if (string.IsNullOrWhiteSpace(occurrenceTerm) || occurrenceTerm != hitTerm)
                return false;

            TimeSpan cueStart = occurrence.Cue.Start - MatchTolerance;
            if (cueStart < TimeSpan.Zero)
                cueStart = TimeSpan.Zero;
            TimeSpan cueEnd = occurrence.Cue.End + MatchTolerance;
            return cueStart <= hit.End && cueEnd >= hit.Start;
        }
    }
}
