using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    public sealed class SubtitleScanner
    {
        private static readonly Regex TimeLineRegex =
            new Regex(@"^\s*(\d{1,3}):(\d{2}):(\d{2}),(\d{3})\s*-->\s*(\d{1,3}):(\d{2}):(\d{2}),(\d{3})",
                RegexOptions.Compiled);

        private static readonly Regex WordRegex =
            new Regex(@"[\p{L}\p{Nd}]+(?:['’-][\p{L}\p{Nd}]+)*", RegexOptions.Compiled);

        private static readonly Regex HtmlTagRegex =
            new Regex("<[^>]+>", RegexOptions.Compiled);

        public SubtitleScanResult ScanFile(string subtitlePath, ProfanityDictionary profanityDictionary)
        {
            if (string.IsNullOrWhiteSpace(subtitlePath))
                throw new ArgumentException("A subtitle path is required.", nameof(subtitlePath));

            if (!File.Exists(subtitlePath))
                throw new FileNotFoundException("Subtitle file was not found.", subtitlePath);

            if (profanityDictionary == null)
                throw new ArgumentNullException(nameof(profanityDictionary));

            List<SubtitleCue> cues = ParseSrtFile(subtitlePath);
            var occurrences = new List<ProfanityOccurrence>();

            foreach (SubtitleCue cue in cues)
            {
                string scanText = BuildScannableText(cue.Text);
                int cueOccurrenceNumber = 0;

                foreach (Match match in WordRegex.Matches(scanText))
                {
                    string dictionaryTerm;
                    if (!profanityDictionary.TryMatch(match.Value, out dictionaryTerm))
                        continue;

                    cueOccurrenceNumber++;
                    occurrences.Add(new ProfanityOccurrence
                    {
                        ReviewNumber = occurrences.Count + 1,
                        Cue = cue,
                        SubtitleSequenceNumber = cue.SequenceNumber,
                        CueIndex = cue.CueIndex,
                        Word = match.Value,
                        DictionaryTerm = dictionaryTerm,
                        CharacterIndex = match.Index,
                        CharacterLength = match.Length,
                        OccurrenceInCue = cueOccurrenceNumber
                    });
                }
            }

            return new SubtitleScanResult
            {
                SubtitlePath = subtitlePath,
                Cues = cues,
                Occurrences = occurrences
            };
        }

        public string FindMatchingSrtForVideo(string videoPath)
        {
            if (string.IsNullOrWhiteSpace(videoPath) || !File.Exists(videoPath))
                return null;

            string directory = Path.GetDirectoryName(videoPath);
            string baseName = Path.GetFileNameWithoutExtension(videoPath);

            if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(baseName))
                return null;

            string[] candidates;
            try
            {
                candidates = Directory.GetFiles(directory, baseName + "*.srt", SearchOption.TopDirectoryOnly);
            }
            catch
            {
                return null;
            }

            if (candidates.Length == 0)
                return null;

            return candidates
                .OrderByDescending(path => ScoreSubtitleCandidate(path, baseName))
                .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
        }

        private static int ScoreSubtitleCandidate(string path, string videoBaseName)
        {
            string fileName = Path.GetFileName(path) ?? "";
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(path) ?? "";
            string lower = fileName.ToLowerInvariant();

            int score = 0;

            if (string.Equals(nameWithoutExtension, videoBaseName, StringComparison.OrdinalIgnoreCase))
                score += 100;

            if (lower.Contains(".eng") || lower.Contains(".english"))
                score += 30;

            if (lower.Contains(".en."))
                score += 20;

            if (lower.Contains(".cc") || lower.Contains(".sdh"))
                score += 5;

            if (lower.Contains(".forced"))
                score -= 25;

            score -= Math.Abs(nameWithoutExtension.Length - videoBaseName.Length);

            return score;
        }

        private static List<SubtitleCue> ParseSrtFile(string subtitlePath)
        {
            var cues = new List<SubtitleCue>();
            string[] lines = File.ReadAllLines(subtitlePath);

            for (int i = 0; i < lines.Length; i++)
            {
                Match match = TimeLineRegex.Match(lines[i]);
                if (!match.Success)
                    continue;

                TimeSpan start = ParseSrtTime(match, 1);
                TimeSpan end = ParseSrtTime(match, 5);
                if (end <= start)
                    continue;

                int sequenceNumber = cues.Count + 1;
                if (i > 0)
                {
                    int parsedSequence;
                    if (int.TryParse(lines[i - 1].Trim(), out parsedSequence))
                        sequenceNumber = parsedSequence;
                }

                var textLines = new List<string>();
                int textLineIndex = i + 1;

                while (textLineIndex < lines.Length && !string.IsNullOrWhiteSpace(lines[textLineIndex]))
                {
                    textLines.Add(lines[textLineIndex]);
                    textLineIndex++;
                }

                cues.Add(new SubtitleCue
                {
                    CueIndex = cues.Count + 1,
                    SequenceNumber = sequenceNumber,
                    SourceLineNumber = i + 1,
                    Start = start,
                    End = end,
                    Text = string.Join(" ", textLines).Trim()
                });

                i = textLineIndex;
            }

            return cues;
        }

        private static TimeSpan ParseSrtTime(Match match, int groupOffset)
        {
            int hours = int.Parse(match.Groups[groupOffset].Value);
            int minutes = int.Parse(match.Groups[groupOffset + 1].Value);
            int seconds = int.Parse(match.Groups[groupOffset + 2].Value);
            int milliseconds = int.Parse(match.Groups[groupOffset + 3].Value);

            return new TimeSpan(0, hours, minutes, seconds, milliseconds);
        }

        private static string BuildScannableText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            string withoutTags = HtmlTagRegex.Replace(text, " ");
            return WebUtility.HtmlDecode(withoutTags);
        }
    }

    public sealed class SubtitleScanResult
    {
        public string SubtitlePath { get; set; }
        public List<SubtitleCue> Cues { get; set; }
        public List<ProfanityOccurrence> Occurrences { get; set; }

        public SubtitleScanResult()
        {
            Cues = new List<SubtitleCue>();
            Occurrences = new List<ProfanityOccurrence>();
        }
    }

    public sealed class SubtitleCue
    {
        public int CueIndex { get; set; }
        public int SequenceNumber { get; set; }
        public int SourceLineNumber { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public string Text { get; set; }

        public SubtitleCue()
        {
            Text = "";
        }
    }

    public sealed class ProfanityOccurrence
    {
        public int ReviewNumber { get; set; }
        public SubtitleCue Cue { get; set; }
        public int CueIndex { get; set; }
        public int SubtitleSequenceNumber { get; set; }
        public string Word { get; set; }
        public string DictionaryTerm { get; set; }
        public int CharacterIndex { get; set; }
        public int CharacterLength { get; set; }
        public int OccurrenceInCue { get; set; }

        public ProfanityOccurrence()
        {
            Word = "";
            DictionaryTerm = "";
        }
    }
}
