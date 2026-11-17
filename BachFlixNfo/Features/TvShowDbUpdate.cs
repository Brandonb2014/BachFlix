using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    // What 55r did to one show in this run.
    internal sealed class ShowRunFacts
    {
        public string ShowName = "";
        public string LibraryFolder = "";
        public bool NewShow;                 // the library had no folder for this show before this run
        public int Added, Replaced;          // new episodes / library episodes that were upgraded
        public List<string> RecordSources = new List<string>();   // Working Area service folder of each moved episode
        public string Resolution = "";       // "1080p", "Mixed", or "" (unsupported, e.g. all 2160p)
        public string Format = "";           // "MKV", "MP4", "Mixed"
        public bool ResolutionKnown;         // false when the library scan failed: Resolution/Format are left alone
        public DateTime Today = DateTime.Today;

        public string Kind { get { return NewShow ? "new show" : Replaced > 0 ? "upgrade" : "new episodes"; } }
    }

    internal sealed class DbCellChange
    {
        public string Column = "", Old = "", New = "";
    }

    // The user's manual DB sheet routine after adding TV episodes, as rules:
    //   every show:            Size cleared (option 24 refills it + Found Locations), Resolution, Format,
    //                          Found Locations cleared if it no longer points at the library folder.
    //   new show / upgrade:    Record Source from the Working Area folder, StreamFab "R - x" -> "Y - x",
    //                          PlayOn-era columns cleared, a dated Note appended.
    //   new episodes only:     no Note, StreamFab and PlayOn columns untouched.
    internal static class TvShowDbUpdate
    {
        public static readonly string[] PlayOnColumns = { "PlayonPal", "Removed Splashes", "Subtitles", "Verify Subtitles line up correctly", "Full Credits" };
        private static readonly string[] ResolutionOrder = { "480p", "576p", "720p", "1080p", "1440p", "2160p" };

        public static List<DbCellChange> Changes(IDictionary<string, string> current, ShowRunFacts facts, string noteOverride = null)
        {
            var changes = new List<DbCellChange>();
            Action<string, string> set = (column, value) =>
            {
                string old;
                if (!current.TryGetValue(column, out old)) return;   // column not in this sheet
                old = old ?? "";
                if (old != value) changes.Add(new DbCellChange { Column = column, Old = old, New = value });
            };
            Func<string, string> get = column => { string v; return current.TryGetValue(column, out v) ? (v ?? "").Trim() : ""; };

            set("Size", "");
            if (facts.ResolutionKnown)
            {
                set("Resolution", facts.Resolution);
                if (facts.Format.Length > 0) set("Format", facts.Format);
            }
            if (!LocationsInclude(get("Found Locations"), facts.LibraryFolder)) set("Found Locations", "");

            if (facts.NewShow || facts.Replaced > 0)
            {
                string source = MostCommon(facts.RecordSources);
                // "Prime, PlutoTV" already names PlutoTV: keep the user's list instead of collapsing it.
                bool alreadyListed = get("Record Source").Split(new[] { ',', ';', '/', '&' }, StringSplitOptions.RemoveEmptyEntries)
                    .Any(s => string.Equals(s.Trim(), source, StringComparison.OrdinalIgnoreCase));
                if (source.Length > 0 && !alreadyListed) set("Record Source", source);

                string streamFab = get("StreamFab");
                var recording = Regex.Match(streamFab, @"^R\b(?<rest>.*)$", RegexOptions.IgnoreCase);
                if (recording.Success) set("StreamFab", "Y" + recording.Groups["rest"].Value);

                foreach (string column in PlayOnColumns) set(column, "");

                string note = noteOverride ?? NoteText(get("Resolution"), get("Format"), facts);
                if (!string.IsNullOrWhiteSpace(note))
                    set("Note", AppendNote(get("Note"), facts.Today.ToString("M/d/yy", CultureInfo.InvariantCulture) + " - " + note.Trim()));
            }
            return changes;
        }

        // "Added in 1080p MKV format." / "Upgraded to 1080p." / "Re-added in MKV format." / "Re-added from Hulu."
        public static string NoteText(string oldResolution, string oldFormat, ShowRunFacts facts)
        {
            string res = Rank(facts.Resolution) >= 0 ? facts.Resolution : "";
            string fmt = facts.Format == "MKV" || facts.Format == "MP4" ? facts.Format : "";
            if (facts.NewShow)
            {
                string what = string.Join(" ", new[] { res, fmt }.Where(s => s.Length > 0));
                return what.Length == 0 ? "Added." : fmt.Length > 0 ? "Added in " + what + " format." : "Added in " + what + ".";
            }
            var parts = new List<string>();
            if (res.Length > 0 && Rank(res) > Rank(oldResolution)) parts.Add("Upgraded to " + res + ".");
            if (fmt == "MKV" && !string.Equals(oldFormat, "MKV", StringComparison.OrdinalIgnoreCase)) parts.Add("Re-added in MKV format.");
            if (parts.Count == 0)
            {
                string source = MostCommon(facts.RecordSources);
                parts.Add(source.Length > 0 ? "Re-added from " + source + "." : "Re-added.");
            }
            return string.Join(" ", parts);
        }

        public static string AppendNote(string existing, string addition)
        {
            existing = (existing ?? "").Trim();
            if (existing.Length == 0) return addition;
            return existing + (existing.EndsWith(".") || existing.EndsWith("!") || existing.EndsWith("?") ? " " : ". ") + addition;
        }

        private static int Rank(string resolution)
        {
            return Array.IndexOf(ResolutionOrder, (resolution ?? "").Trim().ToLowerInvariant());
        }

        private static string MostCommon(IEnumerable<string> values)
        {
            return values.Where(v => !string.IsNullOrWhiteSpace(v))
                .GroupBy(v => v, StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(g => g.Count()).Select(g => g.Key).FirstOrDefault() ?? "";
        }

        // Found Locations can list several paths (comma, semicolon or newline separated).
        public static bool LocationsInclude(string locations, string folder)
        {
            if (string.IsNullOrWhiteSpace(folder)) return true;
            string want = Canonical(folder);
            return SplitLocations(locations).Any(p => string.Equals(Canonical(p), want, StringComparison.OrdinalIgnoreCase));
        }

        // "\\QUADPLEX\E-I\TV Shows\Gabby's Dollhouse (2021),\\QUADPLEX\Other\The Bus\TV Shows\Gabby's Dollhouse (2021)"
        // -> both paths. Only splits on a comma that starts a new path, so commas inside folder names are kept.
        public static List<string> SplitLocations(string locations)
        {
            return Regex.Split(locations ?? "", @"[;\r\n]+|,\s*(?=""?(?:[a-zA-Z]:[\\/]|\\\\))")
                .Select(p => p.Trim().Trim('"').Trim())
                .Where(p => p.Length > 0)
                .ToList();
        }

        private static string Canonical(string path)
        {
            try { return Path.GetFullPath(path).TrimEnd('\\', '/'); } catch { return path.TrimEnd('\\', '/'); }
        }

        // The Working Area service folder: ...\Working Area\<Service>\<Show>\...
        public static string RecordSourceFromPath(string path)
        {
            var parts = (path ?? "").Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
            int i = Array.FindIndex(parts, p => p.Equals("Working Area", StringComparison.OrdinalIgnoreCase));
            return i >= 0 && i + 1 < parts.Length ? parts[i + 1] : "";
        }
    }
}
