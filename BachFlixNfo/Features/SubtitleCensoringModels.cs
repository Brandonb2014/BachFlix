using System;
using System.Collections.Generic;

namespace BachFlixNfo.Features
{
    public enum SubtitleSourceKind
    {
        Embedded,
        External
    }

    public sealed class SubtitleSourceInfo
    {
        public SubtitleSourceKind SourceKind { get; set; }
        public int StreamIndex { get; set; }
        public int SubtitleTrackIndex { get; set; }
        public string CodecName { get; set; }
        public string Language { get; set; }
        public string Title { get; set; }
        public string Path { get; set; }
        public string WorkingSubtitlePath { get; set; }
        public string CensoredSubtitlePath { get; set; }
        public bool HasUnknownLanguage { get; set; }
        public bool IsEnglish { get; set; }
        public bool IsAssumedEnglish { get; set; }
        public bool IsDefault { get; set; }
        public string ScanError { get; set; }
        public SubtitleScanResult ScanResult { get; set; }
        public SubtitleCoverageComparison Comparison { get; set; }

        public SubtitleSourceInfo()
        {
            CodecName = "";
            Language = "";
            Title = "";
            Path = "";
            WorkingSubtitlePath = "";
            CensoredSubtitlePath = "";
            ScanError = "";
            ScanResult = new SubtitleScanResult();
            Comparison = new SubtitleCoverageComparison();
        }

        public string Describe()
        {
            string source = SourceKind == SubtitleSourceKind.Embedded
                ? "embedded stream=" + StreamIndex + ", subtitleTrack=" + SubtitleTrackIndex
                : "external " + Path;
            string language = string.IsNullOrWhiteSpace(Language) ? "(none)" : Language;
            string assumed = IsAssumedEnglish ? ", assumed English" : "";
            string title = string.IsNullOrWhiteSpace(Title) ? "" : ", title=" + Title;
            return source + ", codec=" + (string.IsNullOrWhiteSpace(CodecName) ? "(unknown)" : CodecName) + ", language=" + language + assumed + title;
        }
    }

    public sealed class SubtitleCoverageComparison
    {
        public SubtitleSourceInfo Source { get; set; }
        public int TranscriptHitCount { get; set; }
        public int SubtitleHitCount { get; set; }
        public int SubtitleCueCount { get; set; }
        public double TranscriptCoveragePercent { get; set; }
        public double SubtitleAgreementPercent { get; set; }
        public List<TranscriptProfanityHit> MatchedTranscriptHits { get; private set; }
        public List<TranscriptProfanityHit> ProfanityMissingFromSubtitle { get; private set; }
        public List<ProfanityOccurrence> ProfanityFoundOnlyInSubtitle { get; private set; }

        public SubtitleCoverageComparison()
        {
            MatchedTranscriptHits = new List<TranscriptProfanityHit>();
            ProfanityMissingFromSubtitle = new List<TranscriptProfanityHit>();
            ProfanityFoundOnlyInSubtitle = new List<ProfanityOccurrence>();
        }
    }

    public sealed class CensoredSubtitleOutput
    {
        public string Path { get; set; }
        public SubtitleSourceInfo Source { get; set; }
        public string Codec { get; set; }
        public string Language { get; set; }
        public string Title { get; set; }
        public bool IsDefault { get; set; }

        public CensoredSubtitleOutput()
        {
            Path = "";
            Codec = "srt";
            Language = "eng";
            Title = "English Clean";
            IsDefault = true;
        }
    }
}
