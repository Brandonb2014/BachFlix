using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    public sealed class SubtitleProfanityService
    {
        private static readonly Regex WordRegex = new Regex(@"[\p{L}\p{Nd}]+(?:['\u2019-][\p{L}\p{Nd}]+)*", RegexOptions.Compiled);
        private static readonly Regex AssOverrideRegex = new Regex(@"\{[^}]*\}", RegexOptions.Compiled);
        private static readonly Regex VttTimeLineRegex = new Regex(@"^\s*((?:\d{1,3}:)?\d{2}:\d{2}[\.,]\d{3})\s*-->\s*((?:\d{1,3}:)?\d{2}:\d{2}[\.,]\d{3})", RegexOptions.Compiled);

        public SubtitleScanResult Scan(SubtitleSourceInfo source, ProfanityDictionary dictionary)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            string path = source.WorkingSubtitlePath;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new FileNotFoundException("Subtitle source is not available for scanning.", path ?? "");

            string extension = Path.GetExtension(path).ToLowerInvariant();
            if (extension == ".srt")
                return new SubtitleScanner().ScanFile(path, dictionary);

            if (extension == ".ass" || extension == ".ssa")
                return ScanCues(path, ParseAssCues(path), dictionary);

            if (extension == ".vtt")
                return ScanCues(path, ParseVttCues(path), dictionary);

            throw new NotSupportedException("Subtitle format is not supported for profanity comparison: " + extension);
        }

        private static SubtitleScanResult ScanCues(string path, List<SubtitleCue> cues, ProfanityDictionary dictionary)
        {
            var occurrences = new List<ProfanityOccurrence>();

            foreach (SubtitleCue cue in cues)
            {
                int cueOccurrenceNumber = 0;
                foreach (Match match in WordRegex.Matches(cue.Text ?? ""))
                {
                    string dictionaryTerm;
                    if (!dictionary.TryMatch(match.Value, out dictionaryTerm))
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
                SubtitlePath = path,
                Cues = cues,
                Occurrences = occurrences
            };
        }

        private static List<SubtitleCue> ParseVttCues(string path)
        {
            var cues = new List<SubtitleCue>();
            string[] lines = File.ReadAllLines(path);

            for (int i = 0; i < lines.Length; i++)
            {
                Match match = VttTimeLineRegex.Match(lines[i]);
                if (!match.Success)
                    continue;

                TimeSpan start = ParseLooseSubtitleTime(match.Groups[1].Value);
                TimeSpan end = ParseLooseSubtitleTime(match.Groups[2].Value);
                if (end <= start)
                    continue;

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
                    SequenceNumber = cues.Count + 1,
                    SourceLineNumber = i + 1,
                    Start = start,
                    End = end,
                    Text = string.Join(" ", textLines).Trim()
                });

                i = textLineIndex;
            }

            return cues;
        }

        private static List<SubtitleCue> ParseAssCues(string path)
        {
            var cues = new List<SubtitleCue>();
            string[] lines = File.ReadAllLines(path);
            bool inEvents = false;
            List<string> fields = new List<string>();

            for (int i = 0; i < lines.Length; i++)
            {
                string trimmed = lines[i].Trim();
                if (trimmed.Equals("[Events]", StringComparison.OrdinalIgnoreCase))
                {
                    inEvents = true;
                    continue;
                }

                if (inEvents && trimmed.StartsWith("[", StringComparison.Ordinal) && trimmed.EndsWith("]", StringComparison.Ordinal))
                {
                    inEvents = false;
                    continue;
                }

                if (!inEvents)
                    continue;

                if (trimmed.StartsWith("Format:", StringComparison.OrdinalIgnoreCase))
                {
                    fields = trimmed.Substring("Format:".Length)
                        .Split(',')
                        .Select(f => f.Trim().ToLowerInvariant())
                        .ToList();
                    continue;
                }

                if (!trimmed.StartsWith("Dialogue:", StringComparison.OrdinalIgnoreCase) || fields.Count == 0)
                    continue;

                int startIndex = fields.IndexOf("start");
                int endIndex = fields.IndexOf("end");
                int textIndex = fields.IndexOf("text");
                if (startIndex < 0 || endIndex < 0 || textIndex < 0)
                    continue;

                List<string> values = SplitAssFields(trimmed.Substring("Dialogue:".Length), fields.Count);
                if (values.Count <= Math.Max(textIndex, Math.Max(startIndex, endIndex)))
                    continue;

                TimeSpan start = ParseAssTime(values[startIndex].Trim());
                TimeSpan end = ParseAssTime(values[endIndex].Trim());
                if (end <= start)
                    continue;

                cues.Add(new SubtitleCue
                {
                    CueIndex = cues.Count + 1,
                    SequenceNumber = cues.Count + 1,
                    SourceLineNumber = i + 1,
                    Start = start,
                    End = end,
                    Text = CleanAssText(values[textIndex])
                });
            }

            return cues;
        }

        private static List<string> SplitAssFields(string value, int fieldCount)
        {
            string[] pieces = value.Split(new[] { ',' }, fieldCount);
            return pieces.Select(p => p.Trim()).ToList();
        }

        private static string CleanAssText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            string cleaned = AssOverrideRegex.Replace(text, " ");
            cleaned = cleaned.Replace(@"\N", " ").Replace(@"\n", " ").Replace(@"\h", " ");
            return cleaned.Trim();
        }

        private static TimeSpan ParseAssTime(string value)
        {
            string[] parts = (value ?? "").Split(':');
            if (parts.Length != 3)
                return TimeSpan.Zero;

            int hours;
            int minutes;
            decimal seconds;
            if (!int.TryParse(parts[0], out hours) || !int.TryParse(parts[1], out minutes) || !decimal.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out seconds))
                return TimeSpan.Zero;

            return TimeSpan.FromHours(hours) + TimeSpan.FromMinutes(minutes) + TimeSpan.FromSeconds((double)seconds);
        }

        private static TimeSpan ParseLooseSubtitleTime(string value)
        {
            string normalized = (value ?? "").Replace(',', '.');
            string[] parts = normalized.Split(':');
            int hours = 0;
            int minutes = 0;
            decimal seconds = 0;

            if (parts.Length == 3)
            {
                int.TryParse(parts[0], out hours);
                int.TryParse(parts[1], out minutes);
                decimal.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out seconds);
            }
            else if (parts.Length == 2)
            {
                int.TryParse(parts[0], out minutes);
                decimal.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out seconds);
            }

            return TimeSpan.FromHours(hours) + TimeSpan.FromMinutes(minutes) + TimeSpan.FromSeconds((double)seconds);
        }
    }
}
