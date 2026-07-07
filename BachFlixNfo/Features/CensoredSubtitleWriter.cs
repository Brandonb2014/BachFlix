using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    public sealed class CensoredSubtitleWriter
    {
        private static readonly Regex WordRegex = new Regex(@"[\p{L}\p{Nd}]+(?:['\u2019-][\p{L}\p{Nd}]+)*", RegexOptions.Compiled);
        private static readonly Regex SubwordRegex = new Regex(@"[\p{L}\p{Nd}]+(?:['\u2019][\p{L}\p{Nd}]+)*", RegexOptions.Compiled);
        private static readonly Regex SrtVttTimeLineRegex = new Regex(@"^\s*((?:\d{1,3}:)?\d{2}:\d{2}[\.,]\d{3})\s*-->\s*((?:\d{1,3}:)?\d{2}:\d{2}[\.,]\d{3})", RegexOptions.Compiled);

        public CensoredSubtitleOutput WriteCensoredSubtitle(
            SubtitleSourceInfo source,
            IEnumerable<TranscriptProfanityHit> approvedHits,
            IEnumerable<ProfanityOccurrence> approvedSubtitleOccurrences,
            string mkvPath,
            string replacement)
        {
            SubtitleReplacementMode mode = SubtitleReplacementMode.AsterisksSameLength;
            string replacementText = replacement;
            if (!string.IsNullOrWhiteSpace(replacement) &&
                !replacement.Equals("asterisks", StringComparison.OrdinalIgnoreCase) &&
                replacement != "*")
            {
                mode = SubtitleReplacementMode.FixedText;
            }

            return WriteCensoredSubtitle(source, approvedHits, approvedSubtitleOccurrences, mkvPath, mode, replacementText);
        }

        public CensoredSubtitleOutput WriteCensoredSubtitle(
            SubtitleSourceInfo source,
            IEnumerable<TranscriptProfanityHit> approvedHits,
            IEnumerable<ProfanityOccurrence> approvedSubtitleOccurrences,
            string mkvPath,
            SubtitleReplacementMode replacementMode,
            string replacementText)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            string inputPath = source.WorkingSubtitlePath;
            if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
                throw new FileNotFoundException("Subtitle source is not available for censoring.", inputPath ?? "");

            List<TranscriptProfanityHit> hits = (approvedHits ?? Enumerable.Empty<TranscriptProfanityHit>())
                .Where(h => h != null && h.Approved)
                .ToList();

            List<ProfanityOccurrence> subtitleOccurrences = (approvedSubtitleOccurrences ?? Enumerable.Empty<ProfanityOccurrence>())
                .Where(o => o != null && o.Cue != null)
                .ToList();

            if (hits.Count == 0 && subtitleOccurrences.Count == 0)
                return null;

            string extension = Path.GetExtension(inputPath).ToLowerInvariant();
            string outputPath = BuildOutputPath(source, mkvPath, extension);

            if (extension == ".ass" || extension == ".ssa")
                WriteAssSubtitle(inputPath, outputPath, hits, subtitleOccurrences, replacementMode, replacementText);
            else if (extension == ".srt" || extension == ".vtt")
                WriteTimedTextSubtitle(inputPath, outputPath, hits, subtitleOccurrences, replacementMode, replacementText);
            else
                throw new NotSupportedException("Subtitle format is not supported for censor output: " + extension);

            source.CensoredSubtitlePath = outputPath;
            return new CensoredSubtitleOutput
            {
                Path = outputPath,
                Source = source,
                Language = "eng",
                Title = BuildCensoredSubtitleTitle(source),
                Codec = SubtitleCodecForExtension(extension),
                IsDefault = true
            };
        }

        private static void WriteTimedTextSubtitle(string inputPath, string outputPath, List<TranscriptProfanityHit> hits, List<ProfanityOccurrence> subtitleOccurrences, SubtitleReplacementMode replacementMode, string replacementText)
        {
            string[] lines = File.ReadAllLines(inputPath);
            var output = new List<string>(lines.Length);
            TimeSpan cueStart = TimeSpan.Zero;
            TimeSpan cueEnd = TimeSpan.Zero;
            bool inCue = false;

            foreach (string line in lines)
            {
                Match timeMatch = SrtVttTimeLineRegex.Match(line);
                if (timeMatch.Success)
                {
                    cueStart = ParseLooseSubtitleTime(timeMatch.Groups[1].Value);
                    cueEnd = ParseLooseSubtitleTime(timeMatch.Groups[2].Value);
                    inCue = true;
                    output.Add(line);
                    continue;
                }

                if (string.IsNullOrWhiteSpace(line))
                {
                    inCue = false;
                    output.Add(line);
                    continue;
                }

                output.Add(inCue ? ReplaceWords(line, cueStart, cueEnd, hits, subtitleOccurrences, replacementMode, replacementText) : line);
            }

            File.WriteAllLines(outputPath, output);
        }

        private static void WriteAssSubtitle(string inputPath, string outputPath, List<TranscriptProfanityHit> hits, List<ProfanityOccurrence> subtitleOccurrences, SubtitleReplacementMode replacementMode, string replacementText)
        {
            string[] lines = File.ReadAllLines(inputPath);
            var output = new List<string>(lines.Length);
            bool inEvents = false;
            List<string> fields = new List<string>();

            foreach (string line in lines)
            {
                string trimmed = line.Trim();
                if (trimmed.Equals("[Events]", StringComparison.OrdinalIgnoreCase))
                {
                    inEvents = true;
                    output.Add(line);
                    continue;
                }

                if (inEvents && trimmed.StartsWith("[", StringComparison.Ordinal) && trimmed.EndsWith("]", StringComparison.Ordinal))
                {
                    inEvents = false;
                    output.Add(line);
                    continue;
                }

                if (inEvents && trimmed.StartsWith("Format:", StringComparison.OrdinalIgnoreCase))
                {
                    fields = trimmed.Substring("Format:".Length)
                        .Split(',')
                        .Select(f => f.Trim().ToLowerInvariant())
                        .ToList();
                    output.Add(line);
                    continue;
                }

                if (inEvents && trimmed.StartsWith("Dialogue:", StringComparison.OrdinalIgnoreCase) && fields.Count > 0)
                {
                    int startIndex = fields.IndexOf("start");
                    int endIndex = fields.IndexOf("end");
                    int textIndex = fields.IndexOf("text");
                    if (startIndex >= 0 && endIndex >= 0 && textIndex >= 0)
                    {
                        List<string> values = SplitAssFields(trimmed.Substring("Dialogue:".Length), fields.Count);
                        if (values.Count > Math.Max(textIndex, Math.Max(startIndex, endIndex)))
                        {
                            TimeSpan start = ParseAssTime(values[startIndex]);
                            TimeSpan end = ParseAssTime(values[endIndex]);
                            values[textIndex] = ReplaceWords(values[textIndex], start, end, hits, subtitleOccurrences, replacementMode, replacementText);
                            output.Add("Dialogue: " + string.Join(",", values));
                            continue;
                        }
                    }
                }

                output.Add(line);
            }

            File.WriteAllLines(outputPath, output);
        }

        private static string ReplaceWords(string text, TimeSpan cueStart, TimeSpan cueEnd, List<TranscriptProfanityHit> hits, List<ProfanityOccurrence> subtitleOccurrences, SubtitleReplacementMode replacementMode, string replacementText)
        {
            List<TranscriptProfanityHit> cueHits = hits
                .Where(h => TimesOverlap(cueStart, cueEnd, h.Start, h.End, TimeSpan.FromSeconds(2)))
                .ToList();

            List<ProfanityOccurrence> cueOccurrences = subtitleOccurrences
                .Where(o => o.Cue != null && TimesOverlap(cueStart, cueEnd, o.Cue.Start, o.Cue.End, TimeSpan.FromMilliseconds(50)))
                .ToList();

            if (cueHits.Count == 0 && cueOccurrences.Count == 0)
                return text;

            HashSet<string> censorTargets = BuildCensorTargets(cueHits, cueOccurrences);
            return WordRegex.Replace(text, match => ReplaceToken(match.Value, censorTargets, replacementMode, replacementText));
        }

        private static HashSet<string> BuildCensorTargets(List<TranscriptProfanityHit> cueHits, List<ProfanityOccurrence> cueOccurrences)
        {
            var targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (TranscriptProfanityHit hit in cueHits ?? new List<TranscriptProfanityHit>())
            {
                AddCensorTarget(targets, hit.Word);
                AddCensorTarget(targets, hit.DictionaryTerm);
            }

            foreach (ProfanityOccurrence occurrence in cueOccurrences ?? new List<ProfanityOccurrence>())
            {
                AddCensorTarget(targets, occurrence.Word);
                AddCensorTarget(targets, occurrence.DictionaryTerm);
            }

            return targets;
        }

        private static void AddCensorTarget(HashSet<string> targets, string value)
        {
            string normalized = ProfanityDictionary.NormalizeToken(value);
            if (!string.IsNullOrWhiteSpace(normalized))
                targets.Add(normalized);
        }

        private static string ReplaceToken(string token, HashSet<string> censorTargets, SubtitleReplacementMode replacementMode, string replacementText)
        {
            if (string.IsNullOrEmpty(token) || censorTargets == null || censorTargets.Count == 0)
                return token;

            string normalized = ProfanityDictionary.NormalizeToken(token);
            if (censorTargets.Contains(normalized))
                return BuildReplacement(token, replacementMode, replacementText);

            if (token.IndexOf('-') < 0)
                return token;

            bool changed = false;
            string replaced = SubwordRegex.Replace(token, match =>
            {
                string subwordNormalized = ProfanityDictionary.NormalizeToken(match.Value);
                if (!censorTargets.Contains(subwordNormalized))
                    return match.Value;

                changed = true;
                return BuildReplacement(match.Value, replacementMode, replacementText);
            });

            return changed ? replaced : token;
        }

        private static string BuildReplacement(string value, SubtitleReplacementMode replacementMode, string replacementText)
        {
            if (replacementMode == SubtitleReplacementMode.AsterisksSameLength)
                return new string('*', value.Length);

            return string.IsNullOrEmpty(replacementText) ? "***" : replacementText;
        }

        private static bool TimesOverlap(TimeSpan cueStart, TimeSpan cueEnd, TimeSpan hitStart, TimeSpan hitEnd, TimeSpan tolerance)
        {
            TimeSpan start = cueStart - tolerance;
            if (start < TimeSpan.Zero)
                start = TimeSpan.Zero;

            TimeSpan end = cueEnd + tolerance;
            return start <= hitEnd && end >= hitStart;
        }

        private static string BuildOutputPath(SubtitleSourceInfo source, string mkvPath, string extension)
        {
            string directory;
            string baseName;

            if (source.SourceKind == SubtitleSourceKind.External && !string.IsNullOrWhiteSpace(source.Path))
            {
                directory = Path.GetDirectoryName(source.Path);
                baseName = Path.GetFileNameWithoutExtension(source.Path);
            }
            else
            {
                directory = Path.GetDirectoryName(mkvPath);
                baseName = Path.GetFileNameWithoutExtension(mkvPath) + ".stream-" + source.StreamIndex;
            }

            if (string.IsNullOrWhiteSpace(directory))
                directory = Path.GetDirectoryName(mkvPath) ?? "";

            string outputDirectory = Path.Combine(directory, "Cleaned");
            Directory.CreateDirectory(outputDirectory);
            return Path.Combine(outputDirectory, baseName + extension);
        }

        private static string UniquePath(string path)
        {
            if (!File.Exists(path))
                return path;

            string directory = Path.GetDirectoryName(path) ?? "";
            string baseName = Path.GetFileNameWithoutExtension(path);
            string extension = Path.GetExtension(path);
            int counter = 1;
            while (true)
            {
                string candidate = Path.Combine(directory, baseName + "." + counter + extension);
                if (!File.Exists(candidate))
                    return candidate;
                counter++;
            }
        }

        private static string BuildCensoredSubtitleTitle(SubtitleSourceInfo source)
        {
            string baseTitle = string.IsNullOrWhiteSpace(source.Title) ? "English" : source.Title;
            return baseTitle + " Clean";
        }

        private static string SubtitleCodecForExtension(string extension)
        {
            string value = (extension ?? "").Trim().ToLowerInvariant();
            if (value == ".ass" || value == ".ssa")
                return "ass";
            if (value == ".vtt")
                return "webvtt";
            return "srt";
        }

        private static TimeSpan ParseAssTime(string value)
        {
            string[] parts = (value ?? "").Trim().Split(':');
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

        private static List<string> SplitAssFields(string value, int fieldCount)
        {
            string[] pieces = value.Split(new[] { ',' }, fieldCount);
            return pieces.Select(p => p.Trim()).ToList();
        }
    }
}
