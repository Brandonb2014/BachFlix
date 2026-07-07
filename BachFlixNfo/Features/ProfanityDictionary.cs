using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BachFlixNfo.Features
{
    public sealed class ProfanityDictionary
    {
        private readonly Dictionary<string, string> _termsByNormalizedValue;

        private ProfanityDictionary(Dictionary<string, string> termsByNormalizedValue)
        {
            _termsByNormalizedValue = termsByNormalizedValue;
        }

        public int Count
        {
            get { return _termsByNormalizedValue.Count; }
        }

        public IReadOnlyCollection<string> Terms
        {
            get
            {
                return _termsByNormalizedValue
                    .Values
                    .OrderBy(t => t, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
        }

        public static ProfanityDictionary LoadFromFile(string dictionaryPath)
        {
            if (string.IsNullOrWhiteSpace(dictionaryPath))
                throw new ArgumentException("A profanity dictionary path is required.", nameof(dictionaryPath));

            if (!File.Exists(dictionaryPath))
                throw new FileNotFoundException("Profanity dictionary file was not found.", dictionaryPath);

            var terms = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (string rawLine in File.ReadAllLines(dictionaryPath))
            {
                string term = CleanDictionaryLine(rawLine);
                if (string.IsNullOrWhiteSpace(term))
                    continue;

                string normalized = NormalizeToken(term);
                if (string.IsNullOrWhiteSpace(normalized))
                    continue;

                if (!terms.ContainsKey(normalized))
                    terms.Add(normalized, term);
            }

            return new ProfanityDictionary(terms);
        }

        public bool TryMatch(string subtitleWord, out string dictionaryTerm)
        {
            dictionaryTerm = null;

            if (string.IsNullOrWhiteSpace(subtitleWord))
                return false;

            string normalized = NormalizeToken(subtitleWord);
            if (string.IsNullOrWhiteSpace(normalized))
                return false;

            if (_termsByNormalizedValue.TryGetValue(normalized, out dictionaryTerm))
                return true;

            string possessiveBase = GetPossessiveBase(normalized);
            return !string.IsNullOrWhiteSpace(possessiveBase) &&
                   _termsByNormalizedValue.TryGetValue(possessiveBase, out dictionaryTerm);
        }

        public bool Contains(string subtitleWord)
        {
            string dictionaryTerm;
            return TryMatch(subtitleWord, out dictionaryTerm);
        }

        public static string NormalizeToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            string trimmed = value.Trim();
            var chars = new List<char>(trimmed.Length);

            for (int i = 0; i < trimmed.Length; i++)
            {
                char c = trimmed[i];
                if (char.IsLetterOrDigit(c))
                {
                    chars.Add(char.ToLowerInvariant(c));
                    continue;
                }

                if (IsApostrophe(c) &&
                    i > 0 &&
                    i < trimmed.Length - 1 &&
                    char.IsLetterOrDigit(trimmed[i - 1]) &&
                    char.IsLetterOrDigit(trimmed[i + 1]))
                {
                    chars.Add('\'');
                }
            }

            return new string(chars.ToArray());
        }

        private static string GetPossessiveBase(string normalized)
        {
            if (string.IsNullOrWhiteSpace(normalized))
                return string.Empty;

            if (normalized.EndsWith("'s", StringComparison.OrdinalIgnoreCase) && normalized.Length > 2)
                return normalized.Substring(0, normalized.Length - 2);

            return string.Empty;
        }

        private static bool IsApostrophe(char value)
        {
            return value == '\'' || value == '\u2019';
        }

        private static string CleanDictionaryLine(string rawLine)
        {
            if (rawLine == null)
                return string.Empty;

            string line = rawLine.Trim();
            if (line.Length == 0)
                return string.Empty;

            if (line.StartsWith("#", StringComparison.Ordinal) ||
                line.StartsWith("//", StringComparison.Ordinal) ||
                line.StartsWith(";", StringComparison.Ordinal))
            {
                return string.Empty;
            }

            return line;
        }
    }
}
