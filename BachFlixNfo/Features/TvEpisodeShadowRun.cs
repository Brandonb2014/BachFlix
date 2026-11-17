using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace BachFlixNfo.Features
{
    // A show as the DB sheet knows it.
    internal sealed class KnownTvShow
    {
        public string TvdbId = "";
        public string Name = "";   // "Clean Name with Year", used in the new file name
    }

    internal sealed class ShadowEpisodeRow
    {
        public string ShowFolder = "";
        public string ShowName = "";
        public string TvdbId = "";
        public string SourceFile = "";
        public RecordedEpisodeName Parsed;
        public EpisodeMatchResult Match;
        public string ProposedPath = "";   // relative to the show folder, e.g. "Season 01\Show (2020) - S01E02 - Title.mkv"
    }

    internal sealed class ShadowRunSummary
    {
        public int Shows, UnknownShows, Files, Green, Yellow, Red, Corrected;
        public string ReportPath;
        public readonly List<ShadowEpisodeRow> Rows = new List<ShadowEpisodeRow>();
    }

    // SHADOW MODE: works out what each recorded episode WOULD be renamed to and writes a report.
    // It never renames, moves or deletes anything.
    internal static class TvEpisodeShadowRun
    {
        internal static readonly string[] VideoExtensions = { ".mkv", ".mp4", ".m4v", ".ts", ".avi", ".mov", ".wmv" };

        public static ShadowRunSummary Run(
            IEnumerable<string> showFolders,
            Func<string, KnownTvShow> findShow,
            Func<string, IList<JObject>> getTvdbEpisodes,
            Action<string, string> log,   // (kind, message): kind is info / success / warning / error / data
            string reportDirectory)
        {
            var summary = new ShadowRunSummary();
            var episodeCache = new Dictionary<string, IList<TvdbEpisodeInfo>>(StringComparer.OrdinalIgnoreCase);

            foreach (string showFolder in showFolders.Where(f => !string.IsNullOrWhiteSpace(f)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                string folderName = Path.GetFileName(showFolder.TrimEnd('\\', '/'));
                summary.Shows++;
                List<string> files;
                try
                {
                    files = Directory.GetFiles(showFolder, "*.*", SearchOption.AllDirectories)
                        .Where(f => VideoExtensions.Contains(Path.GetExtension(f).ToLowerInvariant()))
                        .Where(f => Path.GetFileNameWithoutExtension(f).IndexOf("trailer", StringComparison.OrdinalIgnoreCase) < 0)
                        .OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList();
                }
                catch (Exception ex)
                {
                    log("error", "[TV check] Could not read " + showFolder + " | " + ex.Message);
                    continue;
                }
                if (files.Count == 0) continue;

                KnownTvShow show = null;
                try { show = findShow(folderName); } catch (Exception ex) { log("warning", "[TV check] Show lookup failed for " + folderName + " | " + ex.Message); }

                IList<TvdbEpisodeInfo> tvdb = null;
                string showProblem = null;
                if (show == null || string.IsNullOrWhiteSpace(show.TvdbId))
                {
                    summary.UnknownShows++;
                    showProblem = "show not found in DB sheet (would ask you to pick it on TVDB once)";
                }
                else if (!episodeCache.TryGetValue(show.TvdbId, out tvdb))
                {
                    try
                    {
                        tvdb = ToEpisodes(getTvdbEpisodes(show.TvdbId));
                        episodeCache[show.TvdbId] = tvdb;
                    }
                    catch (Exception ex)
                    {
                        showProblem = "could not load TVDB episodes: " + ex.Message;
                    }
                }

                log("info", "[TV check] " + folderName + " | " + files.Count + " file(s)" +
                    (show != null && !string.IsNullOrWhiteSpace(show.TvdbId) ? " | TVDB " + show.TvdbId : ""));

                var showRows = new List<ShadowEpisodeRow>();
                foreach (string file in files)
                {
                    var row = new ShadowEpisodeRow
                    {
                        ShowFolder = showFolder,
                        ShowName = show != null && !string.IsNullOrWhiteSpace(show.Name) ? show.Name : folderName,
                        TvdbId = show != null ? show.TvdbId ?? "" : "",
                        SourceFile = file,
                        Parsed = TvEpisodeMatcher.Parse(Path.GetFileNameWithoutExtension(file), show != null && !string.IsNullOrWhiteSpace(show.Name) ? show.Name : folderName)
                    };
                    row.Match = showProblem != null
                        ? new EpisodeMatchResult { Confidence = EpisodeConfidence.Red, Reason = showProblem }
                        : TvEpisodeMatcher.Match(row.Parsed, tvdb);
                    showRows.Add(row);
                }

                if (showProblem == null)
                    TvEpisodeMatcher.AcceptNumberedSeasons(showRows.Select(r => r.Parsed).ToList(), showRows.Select(r => r.Match).ToList(), tvdb);
                TvEpisodeMatcher.FlagCollisions(showRows.Select(r => r.Match).ToList());

                foreach (var row in showRows)
                {
                    if (row.Match.Episodes.Count > 0)
                        row.ProposedPath = Path.Combine(TvEpisodeMatcher.SeasonFolder(row.Match),
                            TvEpisodeMatcher.BuildFileName(row.ShowName, row.Match) + Path.GetExtension(row.SourceFile));
                    Count(summary, row);
                    LogRow(log, row);
                }
                summary.Rows.AddRange(showRows);
            }

            summary.ReportPath = WriteReport(summary, reportDirectory, log);
            return summary;
        }

        internal static IList<TvdbEpisodeInfo> ToEpisodes(IEnumerable<JObject> raw)
        {
            var list = new List<TvdbEpisodeInfo>();
            foreach (var e in raw ?? Enumerable.Empty<JObject>())
            {
                int season, number;
                if (!int.TryParse((string)e["seasonNumber"], out season) || !int.TryParse((string)e["number"], out number)) continue;
                if (number <= 0) continue;
                list.Add(new TvdbEpisodeInfo { Season = season, Number = number, Title = ((string)e["name"] ?? "").Trim() });
            }
            return list;
        }

        private static void Count(ShadowRunSummary s, ShadowEpisodeRow row)
        {
            s.Files++;
            if (row.Match.Confidence == EpisodeConfidence.Green) s.Green++;
            else if (row.Match.Confidence == EpisodeConfidence.Yellow) s.Yellow++;
            else s.Red++;
            if (row.Match.Episodes.Count > 0 && row.Parsed.Episode.HasValue && row.Match.Code != row.Parsed.Code) s.Corrected++;
        }

        private static void LogRow(Action<string, string> log, ShadowEpisodeRow row)
        {
            string kind = row.Match.Confidence == EpisodeConfidence.Green ? "success"
                : row.Match.Confidence == EpisodeConfidence.Yellow ? "warning" : "error";
            string label = row.Match.Confidence.ToString().ToUpperInvariant();
            log(kind, "  [" + label + "] " + Path.GetFileName(row.SourceFile));
            if (row.ProposedPath.Length > 0) log("data", "      -> " + row.ProposedPath);
            log("data", "      why: " + row.Match.Reason);
            if (row.Match.Alternatives.Count > 0)
                log("data", "      also possible: " + string.Join("; ", row.Match.Alternatives.Select(a => a.Code + " '" + a.Title + "'")));
        }

        private static string WriteReport(ShadowRunSummary summary, string directory, Action<string, string> log)
        {
            if (summary.Rows.Count == 0 || string.IsNullOrWhiteSpace(directory)) return null;
            try
            {
                Directory.CreateDirectory(directory);
                string path = Path.Combine(directory, "TV Episode Check_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");
                var sb = new StringBuilder();
                sb.AppendLine("Confidence,Show,TVDB ID,Current File,File Says,File Title,Matched,Would Rename To,Why,Also Possible,Full Source Path");
                foreach (var r in summary.Rows)
                {
                    sb.AppendLine(string.Join(",", new[]
                    {
                        r.Match.Confidence.ToString(), r.ShowName, r.TvdbId, Path.GetFileName(r.SourceFile),
                        r.Parsed.Code, r.Parsed.Title, r.Match.Code, r.ProposedPath, r.Match.Reason,
                        string.Join("; ", r.Match.Alternatives.Select(a => a.Code + " " + a.Title)), r.SourceFile
                    }.Select(Csv)));
                }
                File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
                return path;
            }
            catch (Exception ex)
            {
                log("warning", "[TV check] Could not write report: " + ex.Message);
                return null;
            }
        }

        private static string Csv(string value)
        {
            value = value ?? "";
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
