using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    // One row of the "TV Episode Review" sheet.
    internal sealed class ReviewRow
    {
        public int SheetRow;   // 1-based row number in the sheet
        public string Skip = "", Confidence = "", OriginalTitle = "", NewTitle = "", Why = "", Show = "", FullPath = "", Result = "";
    }

    internal sealed class PlannedFileMove
    {
        public string Source = "", Target = "";
    }

    // An episode already in the library that the new file replaces (deleted before the move).
    internal sealed class ReplacedEpisode
    {
        public string Video = "";
        public readonly List<string> Sidecars = new List<string>();
        public int Season, FirstEpisode, LastEpisode;
        public IEnumerable<string> AllFiles { get { return new[] { Video }.Concat(Sidecars); } }
    }

    internal sealed class ReviewPlan
    {
        public ReviewRow Row;
        public string ShowFolder = "";
        public int Season, FirstEpisode, LastEpisode;
        public readonly List<PlannedFileMove> Moves = new List<PlannedFileMove>();
        public readonly List<ReplacedEpisode> Replaces = new List<ReplacedEpisode>();
        public readonly List<string> LeftBehind = new List<string>();   // duplicate subtitles not moved (e.g. en-gb when en-us exists)
        public string Problem;         // set when the row can't be applied
        public bool SourceMissing;      // the recorded file is gone (moved or renamed by hand)
        public bool Ok { get { return Problem == null; } }
    }

    // The review sheet flow:
    //   55m writes every checked episode to the sheet (yellow/red pre-marked Skip = X).
    //   The user edits Skip / New Title, then the apply menu option renames + moves every row
    //   without an X into the TV library and deletes the rows that succeeded.
    internal static class TvEpisodeReview
    {
        public const string SheetName = "TV Episode Review";
        public static readonly string[] Headers = { "Skip", "Confidence", "Original Title", "New Title", "Why", "Show", "Full Path", "Result" };

        private static readonly Regex NewTitleShape = new Regex(@"^(?<show>.+?)\s+-\s+S(?<s>\d{1,2})E\d{1,3}", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // ------------------------------------------------------------------
        // 55m -> sheet: rows for files that aren't listed yet
        // ------------------------------------------------------------------
        public static List<IList<object>> NewRows(IEnumerable<ShadowEpisodeRow> checkedRows, ICollection<string> alreadyListed)
        {
            var listed = new HashSet<string>(alreadyListed.Select(NormalizePath), StringComparer.OrdinalIgnoreCase);
            var result = new List<IList<object>>();
            foreach (var r in checkedRows
                .OrderBy(x => x.ShowName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(x => x.SourceFile, StringComparer.OrdinalIgnoreCase))
            {
                string path = NormalizePath(r.SourceFile);
                if (!listed.Add(path)) continue;
                string newTitle = r.ProposedPath.Length > 0 ? Path.GetFileNameWithoutExtension(r.ProposedPath) : "";
                string why = r.Match.Reason;
                if (r.Match.Alternatives.Count > 0)
                    why += " | also possible: " + string.Join("; ", r.Match.Alternatives.Select(a => a.Code + " '" + a.Title + "'"));
                result.Add(new List<object>
                {
                    r.Match.Confidence == EpisodeConfidence.Green ? "" : "X",   // only greens are approved by default
                    r.Match.Confidence.ToString(),
                    Path.GetFileName(r.SourceFile),
                    newTitle,
                    why,
                    r.ShowName,
                    path,
                    ""
                });
            }
            return result;
        }

        // ------------------------------------------------------------------
        // Sheet -> plan. Pure logic so it can be tested without files or Google.
        // ------------------------------------------------------------------
        public static ReviewPlan Plan(ReviewRow row,
            Func<string, string> libraryShowFolder,         // show folder name -> full library path for it
            Func<string, bool> fileExists,
            Func<string, IEnumerable<string>> filesInFolder,
            Func<string, IEnumerable<string>> filesUnderFolder = null)   // recursive listing of a library show folder
        {
            var plan = new ReviewPlan { Row = row };
            string source = (row.FullPath ?? "").Trim();
            if (source.Length == 0) { plan.Problem = "no Full Path"; return plan; }
            if (!fileExists(source)) { plan.Problem = "file not found"; plan.SourceMissing = true; return plan; }

            string newStem = (row.NewTitle ?? "").Trim();
            string ext = Path.GetExtension(source);
            if (newStem.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) newStem = newStem.Substring(0, newStem.Length - ext.Length);
            newStem = TvEpisodeMatcher.SafeFileName(newStem);
            if (newStem.Length == 0) { plan.Problem = "no New Title"; return plan; }

            var shape = NewTitleShape.Match(newStem);
            if (!shape.Success) { plan.Problem = "New Title must look like 'Show (Year) - S01E02 - Title'"; return plan; }
            string showFolder = libraryShowFolder(shape.Groups["show"].Value.Trim());
            if (string.IsNullOrWhiteSpace(showFolder)) { plan.Problem = "couldn't work out the library folder"; return plan; }
            string seasonFolder = Path.Combine(showFolder, "Season " + int.Parse(shape.Groups["s"].Value).ToString("00"));
            var newCode = TvEpisodeMatcher.Parse(newStem, null);
            if (!newCode.Season.HasValue || !newCode.Episode.HasValue) { plan.Problem = "couldn't read SxxEyy from New Title"; return plan; }
            plan.ShowFolder = showFolder;
            plan.Season = newCode.Season.Value;
            plan.FirstEpisode = newCode.Episode.Value;
            plan.LastEpisode = newCode.LastEpisode ?? newCode.Episode.Value;

            plan.Moves.Add(new PlannedFileMove { Source = source, Target = Path.Combine(seasonFolder, newStem + ext) });

            // Sidecars: files next to the video that start with its name, e.g. ".en.srt", ".en.cc.srt".
            string folder = Path.GetDirectoryName(source) ?? "";
            string oldStem = Path.GetFileNameWithoutExtension(source);
            var siblings = (filesInFolder(folder) ?? Enumerable.Empty<string>()).ToList();
            foreach (string sidecar in siblings)
            {
                string name = Path.GetFileName(sidecar);
                if (string.Equals(sidecar, source, StringComparison.OrdinalIgnoreCase)) continue;
                if (!name.StartsWith(oldStem + ".", StringComparison.OrdinalIgnoreCase)) continue;
                if (TvEpisodeShadowRun.VideoExtensions.Contains(Path.GetExtension(name).ToLowerInvariant())) continue;
                // Subtitles get the FileBot language tag (.en.srt -> .eng.srt, .en.cc.srt -> .eng.cc.srt).
                string normalized = FileBotSubtitleNaming.RenameTarget(sidecar, siblings) ?? sidecar;
                string normalizedName = Path.GetFileName(normalized);
                string suffix = normalizedName.StartsWith(oldStem, StringComparison.OrdinalIgnoreCase)
                    ? normalizedName.Substring(oldStem.Length) : name.Substring(oldStem.Length);
                plan.Moves.Add(new PlannedFileMove { Source = sidecar, Target = Path.Combine(seasonFolder, newStem + suffix) });
            }

            // Regional English subtitles (en-us + en-gb) collapse to the same ".eng" name.
            // Keep the preferred one and leave the rest in the source folder.
            foreach (var group in plan.Moves.Skip(1)
                .Where(m => string.Equals(Path.GetExtension(m.Source), ".srt", StringComparison.OrdinalIgnoreCase))
                .GroupBy(m => m.Target, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1).ToList())
            {
                var ranked = group.OrderBy(m => EnglishSubtitleRank(m.Source))
                    .ThenBy(m => m.Source, StringComparer.OrdinalIgnoreCase).ToList();
                foreach (var drop in ranked.Skip(1))
                {
                    plan.Moves.Remove(drop);
                    plan.LeftBehind.Add(drop.Source);
                }
            }

            var dupes = plan.Moves.GroupBy(m => m.Target, StringComparer.OrdinalIgnoreCase).FirstOrDefault(g => g.Count() > 1);
            if (dupes != null) { plan.Problem = "two files would get the same name: " + Path.GetFileName(dupes.Key); return plan; }

            // Episodes already in the library with an overlapping SxxEyy get replaced (an upgrade).
            if (filesUnderFolder != null)
                plan.Replaces.AddRange(FindReplaced(filesUnderFolder(showFolder) ?? Enumerable.Empty<string>(), plan.Season, plan.FirstEpisode, plan.LastEpisode));
            var replacedFiles = new HashSet<string>(plan.Replaces.SelectMany(r => r.AllFiles), StringComparer.OrdinalIgnoreCase);

            var taken = plan.Moves.FirstOrDefault(m => fileExists(m.Target) && !replacedFiles.Contains(m.Target));
            if (taken != null) { plan.Problem = "already in library: " + Path.GetFileName(taken.Target); return plan; }
            return plan;
        }

        // Which English subtitle wins when several map to the same name: en-us, then plain en/eng, then en-gb, then other regions.
        internal static int EnglishSubtitleRank(string path)
        {
            var region = Regex.Match(Path.GetFileNameWithoutExtension(path) ?? "", @"[._-]en(?:g)?[-_](?<r>[a-z]{2})(?=[._-]|$)", RegexOptions.IgnoreCase);
            if (!region.Success) return 1;
            string r = region.Groups["r"].Value.ToLowerInvariant();
            return r == "us" ? 0 : r == "gb" ? 2 : 3;
        }

        // Library videos whose SxxEyy overlaps the new episode range, plus their sidecars
        // (subtitles, -thumb.jpg, -320-10.bif, .nfo...) that share the old file name.
        public static List<ReplacedEpisode> FindReplaced(IEnumerable<string> libraryFiles, int season, int first, int last)
        {
            var files = libraryFiles.ToList();
            var replaced = new List<ReplacedEpisode>();
            foreach (string video in files.Where(f => TvEpisodeShadowRun.VideoExtensions.Contains(Path.GetExtension(f).ToLowerInvariant())))
            {
                var code = TvEpisodeMatcher.Parse(Path.GetFileNameWithoutExtension(video), null);
                if (!code.Season.HasValue || !code.Episode.HasValue || code.Season.Value != season) continue;
                int oldFirst = code.Episode.Value, oldLast = code.LastEpisode ?? code.Episode.Value;
                if (oldLast < first || oldFirst > last) continue;
                string folder = Path.GetDirectoryName(video), stem = Path.GetFileNameWithoutExtension(video);
                var r = new ReplacedEpisode { Video = video, Season = season, FirstEpisode = oldFirst, LastEpisode = oldLast };
                r.Sidecars.AddRange(files.Where(f => !string.Equals(f, video, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(Path.GetDirectoryName(f), folder, StringComparison.OrdinalIgnoreCase)
                    && !TvEpisodeShadowRun.VideoExtensions.Contains(Path.GetExtension(f).ToLowerInvariant())
                    && (Path.GetFileName(f).StartsWith(stem + ".", StringComparison.OrdinalIgnoreCase)
                        || Path.GetFileName(f).StartsWith(stem + "-", StringComparison.OrdinalIgnoreCase))));
                replaced.Add(r);
            }
            return replaced;
        }

        // An old library file is only deleted when this run brings in EVERY episode it holds.
        // e.g. the library has "S01E01-E02" as one file but only a new S01E01 is approved: hold that row.
        public static void CheckReplacementCoverage(IList<ReviewPlan> readyPlans)
        {
            // Repeat until stable: holding one row can leave another row's replacement uncovered.
            bool changed = true;
            while (changed)
            {
                changed = false;
                foreach (var show in readyPlans.Where(p => p.Ok).GroupBy(p => p.ShowFolder, StringComparer.OrdinalIgnoreCase).ToList())
                {
                    var covered = new HashSet<string>(show.SelectMany(p =>
                        Enumerable.Range(p.FirstEpisode, p.LastEpisode - p.FirstEpisode + 1).Select(e => p.Season + "x" + e)));
                    foreach (var p in show)
                    {
                        var partial = p.Replaces.FirstOrDefault(r =>
                            Enumerable.Range(r.FirstEpisode, r.LastEpisode - r.FirstEpisode + 1).Any(e => !covered.Contains(r.Season + "x" + e)));
                        if (partial == null) continue;
                        p.Problem = "library file " + Path.GetFileName(partial.Video) + " holds more episodes than this run replaces; handle by hand";
                        changed = true;
                    }
                }
            }
        }

        private static readonly Regex TrailingYear = new Regex(@"\s*\(\d{4}\)\s*$", RegexOptions.Compiled);

        // Pick an existing library show folder whose name matches, ignoring punctuation and a leading "The".
        // Older library folders have no year ("PJ Masks"), so when exactly one side has no year,
        // the names are compared without it. Two different years never match ("Scrubs (2001)" vs "Scrubs (2026)").
        public static string FindShowFolder(string showName, IEnumerable<string> existingShowFolders)
        {
            string key = TvEpisodeMatcher.NormalizeTitle(showName);
            if (key.Length == 0) return null;
            var folders = existingShowFolders.ToList();
            Func<string, string> leaf = f => Path.GetFileName(f.TrimEnd('\\', '/'));

            var hits = folders.Where(f => TvEpisodeMatcher.NormalizeTitle(leaf(f)) == key).ToList();
            if (hits.Count == 1) return hits[0];
            if (hits.Count > 1) return null;

            bool showHasYear = TrailingYear.IsMatch(showName);
            string keyNoYear = TvEpisodeMatcher.NormalizeTitle(TrailingYear.Replace(showName, ""));
            hits = folders.Where(f =>
            {
                bool folderHasYear = TrailingYear.IsMatch(leaf(f));
                if (showHasYear == folderHasYear) return false;
                return TvEpisodeMatcher.NormalizeTitle(TrailingYear.Replace(leaf(f), "")) == keyNoYear;
            }).ToList();
            return hits.Count == 1 ? hits[0] : null;
        }

        public static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return "";
            try { return Path.GetFullPath(path.Trim()); } catch { return path.Trim(); }
        }

        // ------------------------------------------------------------------
        // Google Sheets I/O
        // ------------------------------------------------------------------
        public static List<ReviewRow> ReadSheet(SheetsService sheets, string spreadsheetId)
        {
            EnsureSheet(sheets, spreadsheetId);
            var values = sheets.Spreadsheets.Values.Get(spreadsheetId, "'" + SheetName + "'!A1:Z").Execute().Values
                         ?? new List<IList<object>>();
            var rows = new List<ReviewRow>();
            if (values.Count == 0) return rows;

            var header = values[0].Select(v => Convert.ToString(v).Trim()).ToList();
            Func<IList<object>, string, string> cell = (row, name) =>
            {
                int i = header.FindIndex(h => h.Equals(name, StringComparison.OrdinalIgnoreCase));
                return i >= 0 && i < row.Count ? Convert.ToString(row[i]).Trim() : "";
            };
            for (int i = 1; i < values.Count; i++)
            {
                var v = values[i];
                if (v.All(x => string.IsNullOrWhiteSpace(Convert.ToString(x)))) continue;
                rows.Add(new ReviewRow
                {
                    SheetRow = i + 1,
                    Skip = cell(v, "Skip"),
                    Confidence = cell(v, "Confidence"),
                    OriginalTitle = cell(v, "Original Title"),
                    NewTitle = cell(v, "New Title"),
                    Why = cell(v, "Why"),
                    Show = cell(v, "Show"),
                    FullPath = cell(v, "Full Path"),
                    Result = cell(v, "Result"),
                });
            }
            return rows;
        }

        // Adds rows for files not already on the sheet. Existing rows (and your edits) are left alone.
        public static int AddResults(SheetsService sheets, string spreadsheetId, IEnumerable<ShadowEpisodeRow> checkedRows)
        {
            var existing = ReadSheet(sheets, spreadsheetId);
            var newRows = NewRows(checkedRows, existing.Select(r => r.FullPath).ToList());
            for (int i = 0; i < newRows.Count; i += 500)
            {
                var body = new ValueRange { Values = newRows.Skip(i).Take(500).ToList() };
                var append = sheets.Spreadsheets.Values.Append(body, spreadsheetId, "'" + SheetName + "'!A1");
                append.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                append.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
                append.Execute();
            }
            return newRows.Count;
        }

        // Writes Result notes first, then deletes finished rows bottom-up so row numbers stay valid.
        public static void WriteBack(SheetsService sheets, string spreadsheetId, IDictionary<int, string> results, IEnumerable<int> rowsToDelete)
        {
            int resultCol = Array.IndexOf(Headers, "Result");
            var header = sheets.Spreadsheets.Values.Get(spreadsheetId, "'" + SheetName + "'!A1:Z1").Execute().Values;
            if (header != null && header.Count > 0)
            {
                int found = header[0].Select(v => Convert.ToString(v).Trim()).ToList()
                    .FindIndex(h => h.Equals("Result", StringComparison.OrdinalIgnoreCase));
                if (found >= 0) resultCol = found;
            }
            var updates = results.Select(kv => new ValueRange
            {
                Range = "'" + SheetName + "'!" + SheetHeaderMapping.ColumnLetter(resultCol) + kv.Key,
                Values = new List<IList<object>> { new List<object> { kv.Value } }
            }).ToList();
            for (int i = 0; i < updates.Count; i += 500)
                sheets.Spreadsheets.Values.BatchUpdate(new BatchUpdateValuesRequest { ValueInputOption = "RAW", Data = updates.Skip(i).Take(500).ToList() }, spreadsheetId).Execute();

            var deletes = rowsToDelete.Distinct().OrderByDescending(r => r).ToList();
            if (deletes.Count == 0) return;
            int sheetId = GetSheetId(sheets, spreadsheetId).Value;
            var requests = deletes.Select(r => new Request
            {
                DeleteDimension = new DeleteDimensionRequest
                {
                    Range = new DimensionRange { SheetId = sheetId, Dimension = "ROWS", StartIndex = r - 1, EndIndex = r }
                }
            }).ToList();
            for (int i = 0; i < requests.Count; i += 500)
                sheets.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests.Skip(i).Take(500).ToList() }, spreadsheetId).Execute();
        }

        private static void EnsureSheet(SheetsService sheets, string spreadsheetId)
        {
            if (GetSheetId(sheets, spreadsheetId).HasValue)
            {
                var first = sheets.Spreadsheets.Values.Get(spreadsheetId, "'" + SheetName + "'!A1:H1").Execute().Values;
                if (first != null && first.Count > 0 && first[0].Count > 0) return;
            }
            else
            {
                sheets.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request>
                    {
                        new Request { AddSheet = new AddSheetRequest { Properties = new SheetProperties
                        {
                            Title = SheetName,
                            GridProperties = new GridProperties { FrozenRowCount = 1 }
                        } } }
                    }
                }, spreadsheetId).Execute();
            }
            var update = sheets.Spreadsheets.Values.Update(new ValueRange { Values = new List<IList<object>> { Headers.Cast<object>().ToList() } },
                spreadsheetId, "'" + SheetName + "'!A1");
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            update.Execute();
        }

        private static int? GetSheetId(SheetsService sheets, string spreadsheetId)
        {
            var spreadsheet = sheets.Spreadsheets.Get(spreadsheetId).Execute();
            var sheet = spreadsheet.Sheets?.FirstOrDefault(s => s.Properties != null &&
                s.Properties.Title.Equals(SheetName, StringComparison.OrdinalIgnoreCase));
            return sheet?.Properties?.SheetId;
        }
    }
}
