using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace BachFlixNfo.Features
{
    internal sealed class TvMediaShow
    {
        internal int Row, ResolutionColumn, FormatColumn;
        internal string CleanName, NameWithYear, Year, TvdbId, Drive, Locations;
        internal string Resolution, Format;
        internal bool ResolutionIsFormula, FormatIsFormula;
        internal string Title { get { return string.IsNullOrWhiteSpace(NameWithYear) ? CleanName : NameWithYear; } }
        internal string Key
        {
            get
            {
                long id;
                return long.TryParse(TvdbId, out id) && id > 0 ? "tvdb:" + id :
                    "name:" + (Title ?? "").Trim().ToUpperInvariant() + "|" + (Year ?? "").Trim();
            }
        }
    }

    internal sealed class TvMediaCellChange
    {
        internal string Range, Value, Previous;
        internal int Row;
        internal bool IsResolution;
        internal string Label { get { return IsResolution ? "Resolution" : "Format"; } }
    }

    /// <summary>Only validated single cells in the header-mapped Resolution/Format columns reach the injected batch writer.</summary>
    internal sealed class TvMediaSheetBatch
    {
        private readonly Action<IList<TvMediaCellChange>> write;
        private readonly int resolutionColumn, formatColumn;
        internal TvMediaSheetBatch(Action<IList<TvMediaCellChange>> write, int resolutionColumn, int formatColumn)
        {
            if (resolutionColumn < 0 || formatColumn < 0 || resolutionColumn == formatColumn)
                throw new InvalidOperationException("Resolution and Format must map to two distinct DB columns.");
            this.write = write; this.resolutionColumn = resolutionColumn; this.formatColumn = formatColumn;
        }
        internal int Write(IList<TvMediaCellChange> changes)
        {
            if (changes.Count == 0) return 0;
            foreach (var change in changes)
            {
                // The range must be exactly one data-row cell in the column the headers mapped for this field.
                string expected = "DB!" + SheetHeaderMapping.ColumnLetter(change.IsResolution ? resolutionColumn : formatColumn) + change.Row;
                if (change.Row < 3 || change.Range != expected)
                    throw new InvalidOperationException("Refusing to update an unintended DB cell: " + change.Range);
                bool resolution = change.IsResolution;
                if (!(resolution ? change.Value == "" || change.Value == "Mixed" || TvEpisodeMediaScanner.IsSupportedResolution(change.Value) : new[] { "MP4", "MKV", "Mixed" }.Contains(change.Value)))
                    throw new InvalidOperationException("Invalid TV media sheet value: " + change.Value);
            }
            write(changes);
            return changes.Count;
        }
    }

    internal static class TvMediaSheet
    {
        internal static Dictionary<string, int> ValidateHeaders(IList<IList<object>> headers)
        {
            if (headers == null || headers.Count != 1 || headers[0].Count == 0)
                throw new InvalidOperationException("DB header row 2 is unavailable.");
            var map = new Dictionary<string, int>();
            foreach (string name in new[] { "Resolution", "Format", "Clean Name", "Clean Name with Year", "Hard Drive Letter", "Found Locations", "TVDB ID", "Year" })
            {
                if (headers[0].Count(h => Convert.ToString(h) == name) > 1)
                    throw new InvalidOperationException("Duplicate DB header: " + name);
                map[name] = -1;
            }
            SheetHeaderMapping.Update(map, headers);
            foreach (string name in new[] { "Resolution", "Format", "Clean Name", "Clean Name with Year", "Hard Drive Letter", "Found Locations", "TVDB ID" })
                if (map[name] < 0) throw new InvalidOperationException("Missing DB row 2 header: " + name);
            // Output columns come from the row 2 headers on every run, so adding/removing columns needs no code change.
            return map;
        }

        internal static List<TvMediaShow> ReadShows(IList<IList<object>> headers, IList<IList<object>> rows)
        {
            var map = ValidateHeaders(headers);
            var shows = new List<TvMediaShow>();
            int rowNumber = 2;
            foreach (var row in rows ?? new List<IList<object>>())
            {
                rowNumber++;
                Func<string, string> cell = key => map[key] >= 0 && map[key] < row.Count ? Convert.ToString(row[map[key]]) : "";
                string name = cell("Clean Name").Trim(), withYear = cell("Clean Name with Year").Trim();
                long id;
                bool hasId = long.TryParse(cell("TVDB ID"), out id) && id > 0;
                // Blank/summary rows lack both a show title and a usable show identity/location.
                if ((name.Length == 0 && withYear.Length == 0) ||
                    (!hasId && string.IsNullOrWhiteSpace(cell("Found Locations")) && string.IsNullOrWhiteSpace(cell("Hard Drive Letter")))) continue;
                shows.Add(new TvMediaShow {
                    Row = rowNumber, ResolutionColumn = map["Resolution"], FormatColumn = map["Format"],
                    CleanName = name, NameWithYear = withYear, Year = cell("Year").Trim(), TvdbId = cell("TVDB ID").Trim(),
                    Drive = cell("Hard Drive Letter").Trim(), Locations = cell("Found Locations").Trim(),
                    Resolution = cell("Resolution"), Format = cell("Format")
                });
            }
            return shows;
        }

        internal static List<TvMediaCellChange> Changes(TvMediaShow show, TvMediaScanResult result)
        {
            var changes = new List<TvMediaCellChange>();
            if (!result.Success) return changes;
            if (!show.ResolutionIsFormula && show.Resolution != result.Resolution)
                changes.Add(new TvMediaCellChange { Range = "DB!" + SheetHeaderMapping.ColumnLetter(show.ResolutionColumn) + show.Row, Row = show.Row, IsResolution = true, Value = result.Resolution, Previous = show.Resolution });
            if (!show.FormatIsFormula && show.Format != result.Format)
                changes.Add(new TvMediaCellChange { Range = "DB!" + SheetHeaderMapping.ColumnLetter(show.FormatColumn) + show.Row, Row = show.Row, IsResolution = false, Value = result.Format, Previous = show.Format });
            return changes;
        }
    }

    internal sealed class TvEpisodeFile
    {
        public string FullPath { get; set; }
        public long SizeBytes { get; set; }
        public DateTime LastWriteTimeUtc { get; set; }
    }

    internal sealed class TvEpisodeMedia
    {
        public string FullPath { get; set; }
        public long SizeBytes { get; set; }
        public DateTime LastWriteTimeUtc { get; set; }
        // Nonempty unsupported resolutions are successful detections, not failures.
        public string Resolution { get; set; }
        public string Format { get; set; }
        internal bool Matches(TvEpisodeFile file)
        {
            return string.Equals(FullPath, file.FullPath, StringComparison.OrdinalIgnoreCase) &&
                SizeBytes == file.SizeBytes && LastWriteTimeUtc == file.LastWriteTimeUtc &&
                !string.IsNullOrEmpty(Resolution) && !string.IsNullOrEmpty(Format);
        }
    }

    internal sealed class TvMediaShowCache
    {
        public int SchemaVersion { get; set; } = TvMediaCache.SchemaVersion;
        public string ShowKey { get; set; }
        public Dictionary<string, TvEpisodeMedia> Episodes { get; set; } = new Dictionary<string, TvEpisodeMedia>(StringComparer.OrdinalIgnoreCase);
    }

    internal sealed class TvMediaCacheDocument
    {
        public int SchemaVersion { get; set; } = TvMediaCache.SchemaVersion;
        public Dictionary<string, TvMediaShowCache> Shows { get; set; } = new Dictionary<string, TvMediaShowCache>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Atomic per-show checkpoints keep the initial scan resumable without rewriting
    /// the entire 68k-file cache for every checkpoint. Compact to the main JSON at exit.
    /// A process lock prevents concurrent scans from overwriting each other's cache.
    /// </summary>
    internal sealed class TvMediaCache : IDisposable
    {
        internal const int SchemaVersion = 1;
        internal readonly string PathName;
        private readonly string checkpointDirectory;
        private readonly FileStream cacheLock;
        private readonly Action<string> log;
        private TvMediaCacheDocument document = new TvMediaCacheDocument();
        internal TvMediaCache(string path, Action<string> log)
        {
            PathName = Path.GetFullPath(path);
            this.log = log;
            checkpointDirectory = PathName + ".checkpoints";
            Directory.CreateDirectory(Path.GetDirectoryName(PathName));
            cacheLock = new FileStream(PathName + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            try
            {
                if (File.Exists(PathName))
                {
                    try
                    {
                        var loaded = JsonConvert.DeserializeObject<TvMediaCacheDocument>(File.ReadAllText(PathName));
                        if (loaded != null && loaded.SchemaVersion == SchemaVersion && loaded.Shows != null) document = loaded;
                        else log("Cache schema changed; rebuilding media metadata.");
                    }
                    catch (JsonException ex) { log("Unreadable cache; rebuilding (" + ex.Message + ")."); }
                }
                document.Shows = new Dictionary<string, TvMediaShowCache>(document.Shows, StringComparer.OrdinalIgnoreCase);
                foreach (string key in document.Shows.Keys.ToArray())
                    if (!Normalize(document.Shows[key])) document.Shows.Remove(key);
                if (Directory.Exists(checkpointDirectory))
                    foreach (string file in Directory.EnumerateFiles(checkpointDirectory, "*.json"))
                    {
                        try
                        {
                            var checkpoint = JsonConvert.DeserializeObject<TvMediaShowCache>(File.ReadAllText(file));
                            if (Normalize(checkpoint) && !string.IsNullOrWhiteSpace(checkpoint.ShowKey)) document.Shows[checkpoint.ShowKey] = checkpoint;
                        }
                        catch (JsonException ex) { log("Unreadable checkpoint " + file + ": " + ex.Message); }
                    }
            }
            catch { cacheLock.Dispose(); throw; }
        }
        private static bool Normalize(TvMediaShowCache show)
        {
            if (show == null || show.SchemaVersion != SchemaVersion || show.Episodes == null) return false;
            show.Episodes = new Dictionary<string, TvEpisodeMedia>(show.Episodes, StringComparer.OrdinalIgnoreCase);
            foreach (string path in show.Episodes.Keys.ToArray())
            {
                var episode = show.Episodes[path];
                if (episode == null || !string.Equals(path, episode.FullPath, StringComparison.OrdinalIgnoreCase) ||
                    episode.SizeBytes < 0 || episode.LastWriteTimeUtc.Kind != DateTimeKind.Utc ||
                    string.IsNullOrEmpty(episode.Resolution) || string.IsNullOrEmpty(episode.Format)) show.Episodes.Remove(path);
                else episode.Resolution = TvEpisodeMediaScanner.ReclassifyCachedResolution(episode.Resolution);
            }
            return true;
        }
        internal TvMediaShowCache GetShow(string key)
        {
            TvMediaShowCache show;
            if (!document.Shows.TryGetValue(key, out show)) document.Shows[key] = show = new TvMediaShowCache { ShowKey = key };
            return show;
        }
        internal void SaveShow(TvMediaShowCache show)
        {
            Directory.CreateDirectory(checkpointDirectory);
            string fileName;
            using (var sha = SHA256.Create()) fileName = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(show.ShowKey))).Replace("-", "");
            AtomicWrite(Path.Combine(checkpointDirectory, fileName + ".json"), show);
        }
        internal void Compact()
        {
            AtomicWrite(PathName, document);
            if (Directory.Exists(checkpointDirectory))
                foreach (string file in Directory.EnumerateFiles(checkpointDirectory, "*.json")) File.Delete(file);
        }
        private static void AtomicWrite(string path, object value)
        {
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    using (var writer = new StreamWriter(file, new UTF8Encoding(false), 4096, true))
                    using (var json = new JsonTextWriter(writer))
                    {
                        JsonSerializer.CreateDefault().Serialize(json, value);
                        json.Flush(); writer.Flush();
                    }
                    file.Flush(true);
                }
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
        public void Dispose() { cacheLock.Dispose(); }
    }

    internal interface ITvMediaFileSystem
    {
        string CanonicalPath(string path);
        IList<string> Directories(string path);
        IList<string> Files(string path);
        TvEpisodeFile Metadata(string path);
    }

    internal sealed class TvMediaDirectoryResolver
    {
        private readonly ITvMediaFileSystem files;
        internal TvMediaDirectoryResolver(ITvMediaFileSystem files) { this.files = files; }
        internal IList<string> Resolve(TvMediaShow show)
        {
            var locations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            // Existing DB discovery joins absolute paths with commas. A show title
            // itself can contain commas, so split only before another absolute path.
            foreach (string location in Regex.Split(show.Locations ?? "", @"[;\r\n]+|,\s*(?=""?(?:[a-zA-Z]:[\\/]|\\\\))"))
                if (!string.IsNullOrWhiteSpace(location)) locations.Add(files.CanonicalPath(location.Trim().Trim('"')));
            // The sheet inventory is authoritative. Blank locations deliberately
            // return no directories, without searching drives or network shares.
            return locations.OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray();
        }
    }

    internal static class TvEpisodeSelection
    {
        private static readonly HashSet<string> VideoExtensions = new HashSet<string>(new[] {
            ".mp4", ".mkv", ".avi", ".m4v", ".mov", ".wmv", ".mpg", ".mpeg", ".ts", ".m2ts", ".mts", ".vob", ".webm", ".flv", ".ogv", ".divx", ".asf", ".3gp", ".3g2", ".m2v", ".mxf", ".ogm", ".rm", ".rmvb", ".f4v", ".mpe", ".m2p" }, StringComparer.OrdinalIgnoreCase);
        private static readonly Regex Extra = new Regex(@"(?:^|[\s._\-\[\]()])(?:trailers?|samples?|featurettes?|extras?|bonus(?:[\s._-]+features?)?|behind[\s._-]+the[\s._-]+scenes|deleted[\s._-]+scenes|interviews?|backstage|bloopers?|outtakes?|shorts?|scenes|other)(?:$|[\s._\-\[\]()])", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex EpisodeNumber = new Regex(@"(?:s\d{1,3}[\s._-]*e\d{1,4}|\b\d{1,3}x\d{1,4}\b)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        internal static bool ExcludedDirectory(string name)
        {
            // Match directory labels exactly, so a season called "Specials" is included.
            return new[] { "trailers", "trailer", "samples", "sample", "featurettes", "featurette", "extras", "extra", "bonus", "bonus features", "behind the scenes", "deleted scenes", "interviews", "interview", "backstage", "bloopers", "outtakes", "shorts", "scenes", "other", ".actors", "@eaDir" }
                .Contains(name, StringComparer.OrdinalIgnoreCase);
        }
        internal static bool IsEpisode(string path)
        {
            if (!VideoExtensions.Contains(Path.GetExtension(path))) return false;
            string name = Path.GetFileNameWithoutExtension(path);
            // Do not reject episode titles such as "The Interview" or "Extras".
            // Standard episode numbering wins, except explicit sample/trailer suffixes.
            if (Regex.IsMatch(name, @"(?:^|[\s._-])(?:sample|trailer)(?:$|[\s._-])", RegexOptions.IgnoreCase)) return false;
            return EpisodeNumber.IsMatch(name) || !Extra.IsMatch(name);
        }
    }

    internal sealed class TvMediaScanResult
    {
        internal bool Success;
        internal string Resolution, Format, Error;
        internal int Discovered, Reused, Probed, New, Modified, Deleted;
    }

    internal sealed class TvEpisodeMediaScanner
    {
        private readonly ITvMediaFileSystem files;
        private readonly Func<string, string> probeResolution;
        private readonly TvMediaCache cache;
        private readonly Action<string> log;
        internal TvEpisodeMediaScanner(ITvMediaFileSystem files, Func<string, string> probeResolution, TvMediaCache cache, Action<string> log)
        { this.files = files; this.probeResolution = probeResolution; this.cache = cache; this.log = log; }

        internal TvMediaScanResult Scan(TvMediaShow show, IList<string> directories, bool force, CancellationToken cancellation)
        {
            var result = new TvMediaScanResult();
            var inventory = new Dictionary<string, TvEpisodeFile>(StringComparer.OrdinalIgnoreCase);
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var pending = new Stack<string>(directories.Reverse());
            try
            {
                while (pending.Count > 0)
                {
                    cancellation.ThrowIfCancellationRequested();
                    string directory = files.CanonicalPath(pending.Pop());
                    if (!visited.Add(directory)) continue;
                    foreach (string path in files.Files(directory))
                    {
                        cancellation.ThrowIfCancellationRequested();
                        if (!TvEpisodeSelection.IsEpisode(path)) continue;
                        string canonical = files.CanonicalPath(path);
                        if (!inventory.ContainsKey(canonical)) inventory.Add(canonical, files.Metadata(canonical));
                    }
                    foreach (string child in files.Directories(directory))
                        if (!TvEpisodeSelection.ExcludedDirectory(Path.GetFileName(child))) pending.Push(child);
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) { result.Error = "Incomplete directory/file inventory: " + ex.Message; return result; }
            result.Discovered = inventory.Count;
            var saved = cache.GetShow(show.Key);
            bool dirty = false, failed = false;
            int sinceCheckpoint = 0;
            var checkpointTimer = Stopwatch.StartNew();
            var media = new List<TvEpisodeMedia>();
            try
            {
                foreach (var file in inventory.Values.OrderBy(f => f.FullPath, StringComparer.OrdinalIgnoreCase))
                {
                    cancellation.ThrowIfCancellationRequested();
                    TvEpisodeMedia previous;
                    bool exists = saved.Episodes.TryGetValue(file.FullPath, out previous);
                    bool matches = exists && previous.Matches(file);
                    if (!force && matches) { result.Reused++; media.Add(previous); continue; }
                    if (!exists) result.New++;
                    else if (!matches) result.Modified++;
                    // Failed forced scans must be retried even if the metadata is unchanged.
                    saved.Episodes.Remove(file.FullPath); dirty = true;
                    try
                    {
                        result.Probed++;
                        string resolution = probeResolution(file.FullPath);
                        if (string.IsNullOrWhiteSpace(resolution)) throw new InvalidOperationException("FFProbe returned no usable video dimensions.");
                        resolution = ReclassifyCachedResolution(resolution);
                        var after = files.Metadata(file.FullPath);
                        if (after.SizeBytes != file.SizeBytes || after.LastWriteTimeUtc != file.LastWriteTimeUtc)
                            throw new IOException("Episode changed during FFProbe; retry on next scan.");
                        var detected = new TvEpisodeMedia { FullPath = file.FullPath, SizeBytes = file.SizeBytes,
                            LastWriteTimeUtc = file.LastWriteTimeUtc, Resolution = resolution, Format = Path.GetExtension(file.FullPath).TrimStart('.').ToUpperInvariant() };
                        saved.Episodes[file.FullPath] = detected; media.Add(detected);
                        if (detected.Format != "MP4" && detected.Format != "MKV") log("Unsupported format " + detected.Format + ": " + file.FullPath);
                        if (!IsSupportedResolution(detected.Resolution)) log("Unsupported resolution " + detected.Resolution.Replace("unsupported:", "") + ": " + file.FullPath);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex) { failed = true; log("FFProbe/file metadata failure for " + file.FullPath + ": " + ShortProbeError(ex.Message)); }
                    sinceCheckpoint++;
                    if (sinceCheckpoint >= 25 || checkpointTimer.Elapsed.TotalSeconds >= 10)
                    { cache.SaveShow(saved); dirty = false; sinceCheckpoint = 0; checkpointTimer.Restart(); }
                    if (result.Probed % 25 == 0) log("  Episode progress: " + (result.Reused + result.Probed) + "/" + inventory.Count + " | FFProbe: " + result.Probed);
                }
                cancellation.ThrowIfCancellationRequested();
                if (failed) { result.Error = "One or more episode probes failed; existing sheet values preserved."; return result; }
                // Prune ONLY after a complete successful inventory and probe pass.
                foreach (string path in saved.Episodes.Keys.Where(p => !inventory.ContainsKey(p)).ToArray())
                { saved.Episodes.Remove(path); result.Deleted++; dirty = true; }
                if (inventory.Count == 0) { result.Error = "No episode files found; existing sheet values preserved."; return result; }
                Aggregate(media, out result.Resolution, out result.Format);
                // Report unsupported cached records too, so each run remains diagnosable.
                var unsupported = media.Where(m => m.Format != "MP4" && m.Format != "MKV").Select(m => m.Format).Distinct().ToArray();
                if (unsupported.Length > 0) log("  Unsupported formats encountered: " + string.Join(", ", unsupported));
                result.Success = true;
                return result;
            }
            finally { if (dirty) cache.SaveShow(saved); }
        }

        internal static void Aggregate(IEnumerable<TvEpisodeMedia> episodes, out string resolution, out string format)
        {
            var records = episodes.ToArray();
            var resolutions = records.Select(e => e.Resolution).Distinct(StringComparer.Ordinal).ToArray();
            // Episodes that don't all share one resolution are "Mixed" (same idea as Format).
            // A single unsupported resolution (e.g. every episode 2160p) still clears the cell.
            resolution = resolutions.Length > 1 ? "Mixed"
                : resolutions.Length == 1 && IsSupportedResolution(resolutions[0]) ? resolutions[0] : "";
            var formats = records.Select(e => e.Format).Distinct(StringComparer.Ordinal).ToArray();
            format = formats.Length == 1 && (formats[0] == "MKV" || formats[0] == "MP4") ? formats[0] : "Mixed";
        }

        // FFProbe errors start with its whole version/build banner; keep only the lines that say what went wrong.
        internal static string ShortProbeError(string message)
        {
            if (string.IsNullOrEmpty(message) || message.IndexOf("ffprobe version", StringComparison.OrdinalIgnoreCase) < 0) return message;
            var useful = message.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(l => !l.StartsWith("ffprobe version", StringComparison.OrdinalIgnoreCase)
                            && (!char.IsWhiteSpace(l[0]) || l.TrimStart().StartsWith("(exit code", StringComparison.OrdinalIgnoreCase))
                            && !l.StartsWith("["))
                .Select(l => l.Trim())
                .ToList();
            return useful.Count > 0 ? string.Join(" ", useful) : message;
        }

        internal static bool IsSupportedResolution(string resolution)
        {
            return resolution == "480p" || resolution == "576p" || resolution == "720p" || resolution == "1080p";
        }

        internal static string ReclassifyCachedResolution(string resolution)
        {
            // Version 1 already saved dimensions for unsupported detections. Reuse
            // that evidence when adding support; don't discard an expensive scan.
            var dimensions = Regex.Match(resolution, @"^unsupported:(\d+)x(\d+)$");
            int width, height;
            if (dimensions.Success && int.TryParse(dimensions.Groups[1].Value, out width) && width > 0 &&
                int.TryParse(dimensions.Groups[2].Value, out height) && height > 0)
                return ClassifyResolution(width, height);
            return resolution;
        }

        internal static string ClassifyResolution(int width, int height)
        {
            if (width <= 0 || height <= 0) throw new InvalidOperationException("Missing video dimensions.");
            // Keep large/ultrawide sources unsupported, including cropped 4K,
            // instead of collapsing them into a lower allowed sheet resolution.
            if (width > 1920 || height >= 1440) return "unsupported:" + width + "x" + height;
            // PAL SD includes 720x576 and square-pixel 1024x576 encodes. Allow
            // up to two vertically cropped pixels, including the observed 718x574.
            if (height >= 574 && height <= 576 && width <= 1024) return "576p";
            // Reuse the established width overrides and height ranges. For example,
            // 720x544 maps to 480p; slight crops no longer need exact dimensions.
            return VideoResolutionClassifier.Classify(width, height);
        }
    }
}
