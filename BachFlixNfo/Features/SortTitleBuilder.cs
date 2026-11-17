using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Builds the Movies sheet "Sort Title": drops a leading "The " and pads a leading number with
    /// zeroes to five digits, so Plex and Emby sort "1 Night in San Diego" before "13 Moons" before "1780".
    /// </summary>
    internal static class SortTitleBuilder
    {
        /// <summary>Five digits covers every title number up to 99,999.</summary>
        internal const int PaddedDigits = 5;

        private static readonly Regex LeadingNumber = new Regex(@"^\d+", RegexOptions.Compiled);

        internal static string Build(string title)
        {
            if (string.IsNullOrEmpty(title)) return title;

            string sortTitle = title.StartsWith("The ") ? title.Substring(4) : title;
            return PadLeadingNumber(sortTitle);
        }

        /// <summary>"13 Moons (2002)" -> "00013 Moons (2002)". Titles that don't start with a digit are unchanged.</summary>
        internal static string PadLeadingNumber(string title)
        {
            if (string.IsNullOrEmpty(title)) return title;
            return LeadingNumber.Replace(title, m => m.Value.PadLeft(PaddedDigits, '0'), 1);
        }
    }
}
