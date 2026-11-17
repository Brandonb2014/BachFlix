using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    // Matches the user's FileBot format: English, then forced > CC > SDH.
    internal static class FileBotSubtitleNaming
    {
        private static readonly HashSet<string> LanguageTokens = BuildLanguageTokens();
        private static readonly HashSet<string> VideoExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { ".mkv", ".mp4", ".avi", ".m4v", ".mov", ".wmv", ".mpg", ".mpeg", ".ts" };
        private static readonly Regex Token = new Regex(@"[._-](?<token>[^._-]+)$", RegexOptions.Compiled);
        private static readonly Regex CaptionToken = new Regex(@"[._-](closed[._ ]?captions?|hard[._ ]of[._ ]hearing)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        internal static string RenameTarget(string path, IEnumerable<string> siblingFiles = null)
        {
            if (!string.Equals(Path.GetExtension(path), ".srt", StringComparison.OrdinalIgnoreCase)) return null;
            string stem = Path.GetFileNameWithoutExtension(path);
            if (string.IsNullOrWhiteSpace(stem)) return null;
            string original = stem;
            var videoNames = (siblingFiles ?? Enumerable.Empty<string>())
                .Where(file => VideoExtensions.Contains(Path.GetExtension(file)))
                .Select(Path.GetFileNameWithoutExtension)
                .ToArray();
            string videoBase = videoNames
                .Where(name => string.Equals(stem, name, StringComparison.OrdinalIgnoreCase)
                    || stem.StartsWith(name + ".", StringComparison.OrdinalIgnoreCase)
                    || stem.StartsWith(name + "_", StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(name => name.Length).FirstOrDefault();
            if (videoBase == null)
            {
                var episode = Regex.Match(stem, @"\bS\d{1,2}E\d{1,3}(?:E\d{1,3})*\b", RegexOptions.IgnoreCase);
                if (episode.Success)
                {
                    var matches = videoNames.Where(name => string.Equals(
                        Regex.Match(name, @"\bS\d{1,2}E\d{1,3}(?:E\d{1,3})*\b", RegexOptions.IgnoreCase).Value,
                        episode.Value, StringComparison.OrdinalIgnoreCase)).ToArray();
                    if (matches.Length == 1) videoBase = matches[0];
                }
            }

            if (videoBase != null) stem = videoBase;
            else
            {
                // Strip recognized trailing metadata only. Dots in episode titles stay intact.
                while (true)
                {
                    Match caption = CaptionToken.Match(stem);
                    if (caption.Success) { stem = stem.Substring(0, caption.Index); continue; }
                    Match match = Token.Match(stem);
                    if (!match.Success) break;
                    string token = match.Groups["token"].Value;
                    if (!LanguageTokens.Contains(token) && !IsFlag(token)
                        && !IsEnglishLocaleSuffix(stem, match, token)) break;
                    stem = stem.Substring(0, match.Index);
                }
            }
            if (string.IsNullOrWhiteSpace(stem)) return null;
            string lower = original.ToLowerInvariant();
            string kind = lower.Contains("forced") ? ".forced"
                : lower.Contains("closedcaptions") || lower.Contains("closed_captions") || lower.Contains("closed.captions")
                    || Regex.IsMatch(lower, @"[._-]cc(?:[._-]|$)") ? ".cc"
                : lower.Contains("sdh") || lower.Contains("hoh") ? ".sdh" : "";
            string target = Path.Combine(Path.GetDirectoryName(path) ?? "", stem + ".eng" + kind + ".srt");
            return string.Equals(path, target, StringComparison.Ordinal) ? null : target;
        }

        private static bool IsFlag(string token)
        {
            return new[] { "forced", "sdh", "hoh", "cc" }.Contains(token, StringComparer.OrdinalIgnoreCase);
        }

        internal static bool TryRename(string source, string target)
        {
            if (string.Equals(source, target, StringComparison.Ordinal)) return true;
            bool caseOnly = string.Equals(source, target, StringComparison.OrdinalIgnoreCase);
            if (!caseOnly && File.Exists(target)) return false;
            if (!caseOnly) { File.Move(source, target); return true; }
            // Windows paths ignore case; use an intermediate name to normalize case reliably.
            string temporary = source + ".bachflix-rename-" + Guid.NewGuid().ToString("N");
            File.Move(source, temporary);
            try { File.Move(temporary, target); }
            catch { File.Move(temporary, source); throw; }
            return true;
        }

        private static bool IsEnglishLocaleSuffix(string stem, Match match, string token)
        {
            return (token.Equals("us", StringComparison.OrdinalIgnoreCase) || token.Equals("gb", StringComparison.OrdinalIgnoreCase))
                && Regex.IsMatch(stem.Substring(0, match.Index), @"[._-](en|eng)$", RegexOptions.IgnoreCase);
        }

        private static HashSet<string> BuildLanguageTokens()
        {
            var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "eng", "en", "english", "som", "ger", "bre", "nor", "hat", "ht", "so", "fre", "dut", "chi" };
            foreach (var culture in CultureInfo.GetCultures(CultureTypes.NeutralCultures))
            {
                if (culture.Name.Length == 0) continue;
                tokens.Add(culture.TwoLetterISOLanguageName);
                tokens.Add(culture.ThreeLetterISOLanguageName);
            }
            return tokens;
        }
    }
}
