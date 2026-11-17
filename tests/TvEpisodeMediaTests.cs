using BachFlixNfo.Features;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

// No real media paths, network shares, FFProbe, Google credentials or Sheets calls.
internal static class TvEpisodeMediaTests
{
    private static int passed;
    private static readonly DateTime Stamp = new DateTime(2026, 9, 20, 0, 0, 0, DateTimeKind.Utc);
    private const string ShowDirectory = @"C:\FakeLibrary\TV Shows\Show (2020)";
    private static void Check(bool condition, string message = "Assertion failed") { if (!condition) throw new Exception(message); }
    private static void Test(string name, Action action) { action(); passed++; Console.WriteLine("PASS " + name); }
    private static void Fails<T>(Action action) where T : Exception
    { try { action(); } catch (T) { return; } throw new Exception("Expected " + typeof(T).Name); }
    private static TvMediaShow Show()
    { return new TvMediaShow { Row = 3, ResolutionColumn = 30, FormatColumn = 32, CleanName = "Show", NameWithYear = "Show (2020)", Year = "2020", TvdbId = "42", Locations = ShowDirectory, Resolution = "720p", Format = "Mixed" }; }
    private static IList<IList<object>> Headers()
    {
        var h = Enumerable.Repeat((object)"", 40).ToList();
        h[0] = "Clean Name"; h[1] = "Clean Name with Year"; h[2] = "Hard Drive Letter";
        h[3] = "Found Locations"; h[4] = "TVDB ID"; h[5] = "Year"; h[30] = "Resolution"; h[32] = "Format";
        return new List<IList<object>> { h };
    }
    private static IList<object> Row(string name, string id, string location = "")
    {
        var row = Enumerable.Repeat((object)"", 40).ToList();
        row[0] = name; row[1] = name == "" ? "" : name + " (2020)"; row[3] = location; row[4] = id; row[5] = "2020";
        row[30] = "720p"; row[32] = "Mixed";
        return row;
    }
    private sealed class FakeFiles : ITvMediaFileSystem
    {
        internal readonly Dictionary<string, TvEpisodeFile> Inventory = new Dictionary<string, TvEpisodeFile>(StringComparer.OrdinalIgnoreCase);
        internal readonly HashSet<string> DirectorySet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal readonly HashSet<string> Broken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal readonly Dictionary<string, int> DirectoryReads = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        internal readonly Dictionary<string, string> Aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string CanonicalPath(string path)
        {
            string full = Path.GetFullPath(path).TrimEnd('\\');
            foreach (var alias in Aliases)
                if (full.Equals(alias.Key, StringComparison.OrdinalIgnoreCase) || full.StartsWith(alias.Key + "\\", StringComparison.OrdinalIgnoreCase))
                    return alias.Value + full.Substring(alias.Key.Length);
            return full;
        }
        internal void Directory(string path)
        {
            path = CanonicalPath(path);
            if (!DirectorySet.Add(path)) return;
            var parent = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(parent)) Directory(parent);
        }
        internal string Add(string name, long size = 100, DateTime? stamp = null, string directory = ShowDirectory)
        {
            Directory(directory);
            string path = CanonicalPath(Path.Combine(directory, name));
            Inventory[path] = new TvEpisodeFile { FullPath = path, SizeBytes = size, LastWriteTimeUtc = stamp ?? Stamp };
            return path;
        }
        private void Verify(string path)
        {
            if (Broken.Contains(path)) throw new IOException("Unavailable fake network directory: " + path);
            if (!DirectorySet.Contains(path)) throw new DirectoryNotFoundException(path);
        }
        public IList<string> Directories(string path)
        {
            path = CanonicalPath(path); Verify(path);
            if (!DirectoryReads.ContainsKey(path)) DirectoryReads[path] = 0;
            DirectoryReads[path]++;
            return DirectorySet.Where(p => string.Equals(Path.GetDirectoryName(p), path, StringComparison.OrdinalIgnoreCase)).ToArray();
        }
        public IList<string> Files(string path)
        {
            path = CanonicalPath(path); Verify(path);
            return Inventory.Keys.Where(p => string.Equals(Path.GetDirectoryName(p), path, StringComparison.OrdinalIgnoreCase)).ToArray();
        }
        public TvEpisodeFile Metadata(string path)
        {
            var file = Inventory[CanonicalPath(path)];
            return new TvEpisodeFile { FullPath = file.FullPath, SizeBytes = file.SizeBytes, LastWriteTimeUtc = file.LastWriteTimeUtc };
        }
    }
    private sealed class Fixture : IDisposable
    {
        internal readonly string DirectoryName;
        internal readonly string CachePath;
        internal TvMediaCache Cache;
        internal readonly FakeFiles Files = new FakeFiles();
        internal readonly Dictionary<string, string> Resolutions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        internal readonly HashSet<string> FailedProbes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal readonly List<string> Logs = new List<string>();
        internal readonly List<string> ProbeCalls = new List<string>();
        internal Action<string> ProbeAction;
        internal Fixture()
        {
            DirectoryName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tv-media-test-" + Guid.NewGuid().ToString("N"));
            CachePath = Path.Combine(DirectoryName, "TvEpisodeMediaCache.json");
            Cache = new TvMediaCache(CachePath, Logs.Add);
            Files.Directory(ShowDirectory);
        }
        internal string Add(string name, string resolution = "1080p", long size = 100, DateTime? stamp = null, string directory = ShowDirectory)
        {
            string path = Files.Add(name, size, stamp, directory); Resolutions[path] = resolution; return path;
        }
        internal TvMediaScanResult Scan(TvMediaShow show = null, bool force = false, IList<string> roots = null, CancellationToken cancellation = default(CancellationToken))
        {
            var scanner = new TvEpisodeMediaScanner(Files, path => {
                ProbeCalls.Add(path); ProbeAction?.Invoke(path);
                if (FailedProbes.Contains(path)) throw new IOException("fake FFProbe error");
                return Resolutions[path];
            }, Cache, Logs.Add);
            return scanner.Scan(show ?? Show(), roots ?? new[] { ShowDirectory }, force, cancellation);
        }
        internal void Restart(bool compact = true)
        {
            if (compact) Cache.Compact();
            Cache.Dispose(); Cache = new TvMediaCache(CachePath, Logs.Add); ProbeCalls.Clear();
        }
        internal void SaveLegacyResolution(string path, string resolution)
        {
            var show = Cache.GetShow(Show().Key);
            show.Episodes[path].Resolution = resolution;
            Cache.SaveShow(show);
        }
        public void Dispose()
        {
            Cache.Dispose();
            // This is the exact unique test-owned directory beneath tests/bin.
            System.IO.Directory.Delete(DirectoryName, true);
        }
    }

    private static void Main()
    {
        Test("FFProbe banner is trimmed from probe errors", () => {
            string probe = "ffprobe version 5.1.2 Copyright (c) 2007-2022\n  built with gcc 12.1.0\n  configuration: --enable-gpl\n  libavutil 57. 28.100\n[avi @ 0000021a] Format avi detected only with low score\n\\\\QUADPLEX\\x\\a.avi: Invalid data found when processing input\n (exit code: 1)";
            string got = TvEpisodeMediaScanner.ShortProbeError(probe);
            Check(got == "\\\\QUADPLEX\\x\\a.avi: Invalid data found when processing input (exit code: 1)", got);
            Check(TvEpisodeMediaScanner.ShortProbeError("Access denied") == "Access denied");
        });
        foreach (string resolution in new[] { "1080p", "720p", "576p", "480p" })
            Test("all episodes " + resolution, () => {
                using (var f = new Fixture()) {
                    f.Add("S01E01.mkv", resolution); f.Add("S01E02.mkv", resolution);
                    var r = f.Scan(); Check(r.Success && r.Resolution == resolution && r.Probed == 2);
                }
            });
        Test("one lower-resolution episode among 99 HD episodes makes resolution Mixed", () => {
            using (var f = new Fixture()) {
                for (int i = 1; i <= 99; i++) f.Add("S01E" + i.ToString("D3") + ".mkv");
                f.Add("S01E100.mkv", "720p");
                var r = f.Scan(); Check(r.Success && r.Resolution == "Mixed" && r.Discovered == 100);
                Check(TvMediaSheet.Changes(Show(), r).Any(c => c.Range == "DB!AE3" && c.Value == "Mixed"));
            }
        });
        Test("480p/720p mixture produces Mixed resolution", () => {
            using (var f = new Fixture()) { f.Add("S01E01.mp4", "480p"); f.Add("S01E02.mp4", "720p"); Check(f.Scan().Resolution == "Mixed"); }
        });
        Test("576p mixed with another supported resolution is Mixed", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mp4", "576p"); f.Add("S01E02.mp4", "480p");
                var result = f.Scan(); Check(result.Success && result.Resolution == "Mixed");
            }
        });
        Test("576p is permitted in the Resolution batch writer", () => {
            var writes = new List<IList<TvMediaCellChange>>();
            var changes = TvMediaSheet.Changes(Show(), new TvMediaScanResult { Success = true, Resolution = "576p", Format = "Mixed" });
            Check(new TvMediaSheetBatch(writes.Add, 30, 32).Write(changes) == 1 && writes[0].Single().Range == "DB!AE3" && writes[0].Single().Value == "576p");
        });
        Test("cropped 576p episodes aggregate uniformly and are not logged as unsupported", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mp4", TvEpisodeMediaScanner.ClassifyResolution(718, 576));
                f.Add("S01E02.mp4", TvEpisodeMediaScanner.ClassifyResolution(718, 574));
                var result = f.Scan(); Check(result.Success && result.Resolution == "576p" && !f.Logs.Any(l => l.Contains("Unsupported resolution")));
            }
        });
        Test("main cache reclassifies previous 576p dimensions without FFProbe", () => {
            using (var f = new Fixture()) {
                string first = f.Add("S01E01.mp4", "unsupported:718x576"), second = f.Add("S01E02.mp4", "unsupported:718x574");
                f.Scan(); f.SaveLegacyResolution(first, "unsupported:718x576"); f.SaveLegacyResolution(second, "unsupported:718x574"); f.Restart();
                var result = f.Scan(); Check(result.Success && result.Resolution == "576p" && result.Reused == 2 && result.Probed == 0 && f.ProbeCalls.Count == 0);
                Check(TvMediaSheet.Changes(Show(), result).Any(c => c.Value == "576p"));
                f.Restart(); Check(f.Cache.GetShow(Show().Key).Episodes.Values.All(e => e.Resolution == "576p"));
            }
        });
        Test("partial-scan checkpoints normalize 576p and 480p while their mixture stays blank", () => {
            using (var f = new Fixture()) {
                string first = f.Add("S01E01.mp4", "unsupported:718x574"), second = f.Add("S01E02.mp4", "unsupported:960x540");
                f.Scan(); f.SaveLegacyResolution(first, "unsupported:718x574"); f.SaveLegacyResolution(second, "unsupported:960x540"); f.Restart(false);
                var result = f.Scan(); Check(result.Success && result.Resolution == "Mixed" && result.Reused == 2 && result.Probed == 0);
                Check(f.Cache.GetShow(Show().Key).Episodes.Values.Count(e => e.Resolution == "576p") == 1);
                Check(f.Cache.GetShow(Show().Key).Episodes.Values.Count(e => e.Resolution == "480p") == 1);
            }
        });
        foreach (string resolution in new[] { "2160p", "unsupported:3840x1600" })
            Test("successfully detected unsupported " + resolution + " clears resolution", () => {
                using (var f = new Fixture()) {
                    f.Add("S01E01.mkv", resolution); var r = f.Scan(); Check(r.Success && r.Resolution == "");
                    f.ProbeCalls.Clear(); Check(f.Scan().Reused == 1 && f.ProbeCalls.Count == 0);
                }
            });
        foreach (string extension in new[] { "mkv", "mp4" })
            Test("uniform " + extension + " explicitly populates Format", () => {
                using (var f = new Fixture()) { f.Add("S01E01." + extension); f.Add("S01E02." + extension); Check(f.Scan().Format == extension.ToUpperInvariant()); }
            });
        Test("mixed MP4/MKV and unsupported formats produce Mixed", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); f.Add("S01E02.mp4"); Check(f.Scan().Format == "Mixed");
                f.Add("S01E03.avi"); Check(f.Scan().Format == "Mixed"); Check(f.Logs.Any(l => l.Contains("AVI")));
            }
        });
        Test("only unsupported format produces Mixed and logs extension", () => {
            using (var f = new Fixture()) { f.Add("S01E01.webm"); var r = f.Scan(); Check(r.Success && r.Format == "Mixed" && f.Logs.Any(l => l.Contains("WEBM"))); }
        });
        Test("unchanged show reuses each file without FFProbe", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); f.Add("S01E02.mkv"); f.Scan(); f.ProbeCalls.Clear();
                var r = f.Scan(); Check(r.Success && r.Reused == 2 && r.Probed == 0 && f.ProbeCalls.Count == 0);
            }
        });
        Test("adding an episode probes only the new file", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); f.Scan(); string newPath = f.Add("S01E02.mp4"); f.ProbeCalls.Clear();
                var r = f.Scan(); Check(r.New == 1 && r.Reused == 1 && r.Probed == 1 && f.ProbeCalls.Single() == newPath && r.Format == "Mixed");
            }
        });
        Test("replacement detected despite same count and same newest timestamp", () => {
            using (var f = new Fixture()) {
                string first = f.Add("S01E01.mkv", "720p", stamp: Stamp.AddDays(-19));
                f.Add("S01E02.mkv", stamp: Stamp.AddDays(-10)); f.Add("S01E03.mkv"); f.Scan();
                f.Add("S01E01.mkv", "1080p", stamp: Stamp.AddDays(-5)); f.ProbeCalls.Clear();
                var r = f.Scan(); Check(r.Discovered == 3 && r.Modified == 1 && r.Probed == 1 && r.Reused == 2 && r.Resolution == "1080p" && f.ProbeCalls.Single() == first);
            }
        });
        Test("size-only change invalidates episode cache", () => {
            using (var f = new Fixture()) { f.Add("S01E01.mkv"); f.Scan(); f.Add("S01E01.mkv", size: 101); var r = f.Scan(); Check(r.Modified == 1 && r.Probed == 1); }
        });
        Test("filename/extension change is a new file and removes old record", () => {
            using (var f = new Fixture()) {
                string old = f.Add("S01E01.mp4"); f.Scan(); f.Files.Inventory.Remove(old); f.Add("S01E01.mkv");
                var r = f.Scan(); Check(r.New == 1 && r.Deleted == 1 && r.Probed == 1 && r.Format == "MKV" && !f.Cache.GetShow(Show().Key).Episodes.ContainsKey(old));
            }
        });
        Test("removing an episode recalculates results and prunes cached inventory", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); string removed = f.Add("S01E02.mp4", "720p"); f.Scan(); f.Files.Inventory.Remove(removed);
                var r = f.Scan(); Check(r.Success && r.Deleted == 1 && r.Probed == 0 && r.Resolution == "1080p" && r.Format == "MKV" && f.Cache.GetShow(Show().Key).Episodes.Count == 1);
            }
        });
        Test("missing network path preserves sheet values and cached inventory", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); f.Scan(); f.Files.Broken.Add(ShowDirectory);
                var r = f.Scan(); Check(!r.Success && TvMediaSheet.Changes(Show(), r).Count == 0 && f.Cache.GetShow(Show().Key).Episodes.Count == 1);
            }
        });
        Test("inaccessible second location never produces incomplete updates", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); var r = f.Scan(roots: new[] { ShowDirectory, @"C:\OfflineShare\Show" });
                Check(!r.Success && r.Probed == 0 && TvMediaSheet.Changes(Show(), r).Count == 0);
            }
        });
        Test("FFProbe failure preserves both existing cells and is not cached as success", () => {
            using (var f = new Fixture()) {
                string broken = f.Add("S01E01.mkv"); f.FailedProbes.Add(broken);
                var r = f.Scan(); Check(!r.Success && r.Probed == 1 && TvMediaSheet.Changes(Show(), r).Count == 0 && f.Cache.GetShow(Show().Key).Episodes.Count == 0);
                f.FailedProbes.Clear(); Check(f.Scan().Probed == 1);
            }
        });
        Test("successful episodes survive a failed show and retry only failures", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); string broken = f.Add("S01E02.mkv"); f.FailedProbes.Add(broken);
                Check(!f.Scan().Success); f.Restart(false); f.FailedProbes.Clear();
                var r = f.Scan(); Check(r.Success && r.Reused == 1 && r.Probed == 1);
            }
        });
        Test("incorrect sheet values corrected using valid cache without probing", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mp4"); f.Scan(); f.ProbeCalls.Clear(); var show = Show(); show.Resolution = "480p"; show.Format = "";
                var changes = TvMediaSheet.Changes(show, f.Scan(show));
                Check(f.ProbeCalls.Count == 0 && changes.Count == 2 && changes[0].Value == "1080p" && changes[1].Value == "MP4");
            }
        });
        Test("matching sheet values produce no writes", () => {
            using (var f = new Fixture()) { f.Add("S01E01.mp4"); var s = Show(); s.Resolution = "1080p"; s.Format = "MP4"; Check(TvMediaSheet.Changes(s, f.Scan(s)).Count == 0); }
        });
        Test("cache persists across restarts and compacted checkpoints disappear", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); f.Scan(); f.Restart();
                Check(File.Exists(f.CachePath) && System.IO.Directory.GetFiles(f.CachePath + ".checkpoints", "*.json").Length == 0);
                Check(f.Scan().Reused == 1 && f.ProbeCalls.Count == 0);
            }
        });
        Test("partial first run resumes from checkpoint without a main cache", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); f.Add("S01E02.mkv"); f.Add("S01E03.mkv");
                using (var cancel = new CancellationTokenSource()) {
                    f.ProbeAction = path => { if (f.ProbeCalls.Count == 2) { cancel.Cancel(); cancel.Token.ThrowIfCancellationRequested(); } };
                    Fails<OperationCanceledException>(() => f.Scan(cancellation: cancel.Token));
                }
                Check(!File.Exists(f.CachePath)); f.ProbeAction = null; f.Restart(false);
                var r = f.Scan(); Check(r.Success && r.Reused == 1 && r.Probed == 2);
            }
        });
        Test("periodic atomic checkpoint persists during a long first show", () => {
            using (var f = new Fixture()) {
                for (int i = 1; i <= 30; i++) f.Add("S01E" + i.ToString("D2") + ".mkv");
                f.ProbeAction = _ => {
                    if (f.ProbeCalls.Count == 26) {
                        var file = System.IO.Directory.GetFiles(f.CachePath + ".checkpoints", "*.json").Single();
                        var saved = JsonConvert.DeserializeObject<TvMediaShowCache>(File.ReadAllText(file));
                        Check(saved.Episodes.Count == 25 && saved.SchemaVersion == TvMediaCache.SchemaVersion);
                    }
                };
                Check(f.Scan().Success);
            }
        });
        Test("force scan reprobes unchanged metadata and catches content-only changes", () => {
            using (var f = new Fixture()) {
                string path = f.Add("S01E01.mkv", "720p"); f.Scan(); f.Resolutions[path] = "1080p";
                Check(f.Scan().Resolution == "720p"); var r = f.Scan(force: true); Check(r.Probed == 1 && r.Reused == 0 && r.Resolution == "1080p");
            }
        });
        Test("failed force scan invalidates old success and normal scan retries", () => {
            using (var f = new Fixture()) {
                string path = f.Add("S01E01.mkv"); f.Scan(); f.FailedProbes.Add(path); Check(!f.Scan(force: true).Success);
                f.Restart(false); f.FailedProbes.Clear(); Check(f.Scan().Probed == 1);
            }
        });
        Test("only intended single DB cells reach one batch request", () => {
            var writes = new List<IList<TvMediaCellChange>>(); var batch = new TvMediaSheetBatch(writes.Add, 30, 32);
            var changes = TvMediaSheet.Changes(Show(), new TvMediaScanResult { Success = true, Resolution = "1080p", Format = "MKV" });
            Check(batch.Write(changes) == 2 && writes.Count == 1 && writes[0].Select(c => c.Range).SequenceEqual(new[] { "DB!AE3", "DB!AG3" }));
            Check(batch.Write(new List<TvMediaCellChange>()) == 0 && writes.Count == 1);
            foreach (var bad in new[] { Tuple.Create("DB!AF3", 3), Tuple.Create("DB!AE2", 2), Tuple.Create("Movies!AE3", 3), Tuple.Create("DB!AE3:AG3", 3),
                                        Tuple.Create("DB!AE1", 1), Tuple.Create("DB!AE0", 0), Tuple.Create("DB!AE4", 3) })
                Fails<InvalidOperationException>(() => batch.Write(new[] { new TvMediaCellChange { Range = bad.Item1, Row = bad.Item2, IsResolution = true, Value = "1080p" } }));
            Fails<InvalidOperationException>(() => batch.Write(new[] { new TvMediaCellChange { Range = "DB!AG3", Row = 3, IsResolution = true, Value = "1080p" } }));
            Fails<InvalidOperationException>(() => batch.Write(new[] { new TvMediaCellChange { Range = "DB!AE3", Row = 3, IsResolution = true, Value = "2160p" } }));
            Fails<InvalidOperationException>(() => new TvMediaSheetBatch(writes.Add, 30, 30));
            Check(writes.Count == 1);
        });
        Test("Sheets failure leaves cache reusable on restart", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mp4"); var changes = TvMediaSheet.Changes(Show(), f.Scan());
                var batch = new TvMediaSheetBatch(_ => { throw new IOException("fake Sheets outage"); }, 30, 32);
                Fails<IOException>(() => batch.Write(changes)); f.Restart(false);
                Check(f.Scan().Reused == 1 && f.ProbeCalls.Count == 0);
            }
        });
        Test("formulas in intended output cells are preserved", () => {
            var show = Show(); show.ResolutionIsFormula = true;
            var result = new TvMediaScanResult { Success = true, Resolution = "1080p", Format = "MP4" };
            Check(TvMediaSheet.Changes(show, result).Single().Range == "DB!AG3");
            show.FormatIsFormula = true; Check(TvMediaSheet.Changes(show, result).Count == 0);
        });
        Test("headers validated before rows; reject missing or duplicate columns", () => {
            var h = Headers(); Check(TvMediaSheet.ReadShows(h, null).Count == 0);
            h[0][30] = ""; Fails<InvalidOperationException>(() => TvMediaSheet.ReadShows(h, null));
            h = Headers(); h[0][31] = "Resolution"; Fails<InvalidOperationException>(() => TvMediaSheet.ReadShows(h, null));
            Fails<InvalidOperationException>(() => TvMediaSheet.ReadShows(null, null));
        });
        Test("moved Resolution/Format columns are mapped from headers and written there", () => {
            var h = Headers(); h[0][30] = ""; h[0][32] = ""; h[0][31] = "Resolution"; h[0][35] = "Format";
            var map = TvMediaSheet.ValidateHeaders(h); Check(map["Resolution"] == 31 && map["Format"] == 35);
            var row = Row("Show", "42"); row[31] = "720p"; row[35] = "Mixed";
            var show = TvMediaSheet.ReadShows(h, new List<IList<object>> { row }).Single();
            Check(show.ResolutionColumn == 31 && show.FormatColumn == 35 && show.Resolution == "720p" && show.Format == "Mixed");
            var changes = TvMediaSheet.Changes(show, new TvMediaScanResult { Success = true, Resolution = "1080p", Format = "MKV" });
            Check(changes.Select(c => c.Range).SequenceEqual(new[] { "DB!AF3", "DB!AJ3" }));
            var writes = new List<IList<TvMediaCellChange>>();
            Check(new TvMediaSheetBatch(writes.Add, 31, 35).Write(changes) == 2);
            Fails<InvalidOperationException>(() => new TvMediaSheetBatch(writes.Add, 30, 32).Write(changes));
        });
        Test("physical row numbers retained; blank and summary rows untouched; name/year fallback works", () => {
            var rows = new List<IList<object>> { Row("Show", "42"), Row("", ""), Row("Total", "68761"), Row("Fallback", "", ShowDirectory) };
            // A summary count isn't a TVDB ID if it has no show title in the clean-name columns.
            rows[2][0] = ""; rows[2][1] = "";
            var shows = TvMediaSheet.ReadShows(Headers(), rows);
            Check(shows.Count == 2 && shows[0].Row == 3 && shows[1].Row == 6 && shows[0].Key == "tvdb:42" && shows[1].Key.Contains("FALLBACK (2020)"));
            var a = Show(); var b = Show(); a.TvdbId = ""; b.TvdbId = ""; b.Year = "2021"; Check(a.Key != b.Key);
        });
        Test("overlapping locations and mapped-drive aliases never duplicate an episode", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01-E02.mkv", directory: ShowDirectory + @"\Season 1");
                f.Files.Aliases[@"X:\Show"] = ShowDirectory;
                var r = f.Scan(roots: new[] { ShowDirectory, ShowDirectory + @"\Season 1", @"X:\Show" });
                Check(r.Discovered == 1 && r.Probed == 1);
            }
        });
        Test("specials and season zero included; trailers/samples/extras/non-video excluded", () => {
            using (var f = new Fixture()) {
                f.Add("Show S00E01.mkv", directory: ShowDirectory + @"\Season 0");
                f.Add("Special.mp4", directory: ShowDirectory + @"\Specials");
                f.Add("Show S01E01 - The Interview.mkv", directory: ShowDirectory + @"\Season 1");
                f.Add("Show S01E02.sample.mkv"); f.Add("Show-trailer.mp4"); f.Add("Featurette.mkv");
                f.Add("S01E02.mkv", directory: ShowDirectory + @"\Extras");
                f.Add("S01E03.mkv", directory: ShowDirectory + @"\Bonus Features");
                f.Add("S01E01.nfo"); f.Add("S01E01.srt"); f.Add("poster.jpg");
                var r = f.Scan(); Check(r.Discovered == 3 && r.Probed == 3);
            }
        });
        Test("empty complete show leaves existing cells untouched", () => {
            using (var f = new Fixture()) {
                var r = f.Scan(); Check(!r.Success && TvMediaSheet.Changes(Show(), r).Count == 0);
                var path = f.Add("S01E01.mkv"); f.Scan(); f.Files.Inventory.Remove(path); r = f.Scan();
                Check(!r.Success && r.Deleted == 1 && f.Cache.GetShow(Show().Key).Episodes.Count == 0);
            }
        });
        Test("file changing during FFProbe is rejected and not cached", () => {
            using (var f = new Fixture()) {
                string path = f.Add("S01E01.mkv"); f.ProbeAction = _ => f.Files.Inventory[path].SizeBytes++;
                var r = f.Scan(); Check(!r.Success && TvMediaSheet.Changes(Show(), r).Count == 0 && f.Cache.GetShow(Show().Key).Episodes.Count == 0);
            }
        });
        Test("shared FindResolution rules normalize nonstandard HD/SD; large sources stay unsupported", () => {
            Check(TvEpisodeMediaScanner.ClassifyResolution(1920, 800) == "1080p");
            Check(TvEpisodeMediaScanner.ClassifyResolution(1280, 544) == "720p");
            Check(TvEpisodeMediaScanner.ClassifyResolution(720, 304) == "480p");
            foreach (var d in new[] { new[] { 720, 576 }, new[] { 718, 576 }, new[] { 718, 574 }, new[] { 1024, 576 } })
                Check(TvEpisodeMediaScanner.ClassifyResolution(d[0], d[1]) == "576p");
            foreach (var d in new[] { new[] { 720, 544 }, new[] { 640, 360 }, new[] { 960, 540 }, new[] { 718, 572 } })
                Check(TvEpisodeMediaScanner.ClassifyResolution(d[0], d[1]) == "480p");
            foreach (var d in new[] { new[] { 1918, 1078 }, new[] { 1280, 544 }, new[] { 1000, 768 }, new[] { 720, 544 }, new[] { 640, 360 } })
                Check(TvEpisodeMediaScanner.ClassifyResolution(d[0], d[1]) == VideoResolutionClassifier.Classify(d[0], d[1]));
            foreach (var d in new[] { new[] { 3840, 2160 }, new[] { 3840, 1600 }, new[] { 3840, 1080 }, new[] { 2560, 1440 } })
                Check(TvEpisodeMediaScanner.ClassifyResolution(d[0], d[1]).StartsWith("unsupported:"));
            Fails<InvalidOperationException>(() => TvEpisodeMediaScanner.ClassifyResolution(0, 1080));
        });
        Test("legacy 720x544 checkpoints become uniform 480p without probing again", () => {
            using (var f = new Fixture()) {
                string first = f.Add("S01E01.mkv", "unsupported:720x544"); f.Add("S01E02.mkv", "480p");
                f.Scan(); f.SaveLegacyResolution(first, "unsupported:720x544"); f.Restart(false);
                var result = f.Scan();
                Check(result.Success && result.Resolution == "480p" && result.Reused == 2 && result.Probed == 0 && f.ProbeCalls.Count == 0);
                Check(TvMediaSheet.Changes(Show(), result).Any(c => c.Range == "DB!AE3" && c.Value == "480p"));
            }
        });
        Test("new 720x544 metadata is normalized without unsupported warnings", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv", TvEpisodeMediaScanner.ClassifyResolution(720, 544));
                var result = f.Scan(); Check(result.Success && result.Resolution == "480p" && !f.Logs.Any(l => l.Contains("Unsupported resolution")));
            }
        });
        Test("empty Found Locations skips without discovering matching directories", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv");
                var resolver = new TvMediaDirectoryResolver(f.Files);
                foreach (string value in new[] { null, "", " \t\r\n" }) {
                    var show = Show(); show.Locations = value; show.Drive = @"P:\";
                    Check(resolver.Resolve(show).Count == 0);
                    Check(f.Files.DirectoryReads.Count == 0 && f.ProbeCalls.Count == 0);
                    Check(show.Resolution == "720p" && show.Format == "Mixed");
                }
            }
        });
        Test("known comma-separated locations include all season files", () => {
            using (var f = new Fixture()) {
                string other = @"C:\KnownLocation\Show";
                f.Add("S01E01.mkv"); f.Add("S02E01.mp4", "720p", directory: other);
                var show = Show(); show.Locations += "," + other;
                var roots = new TvMediaDirectoryResolver(f.Files).Resolve(show);
                var r = f.Scan(show, roots: roots); Check(r.Success && r.Discovered == 2 && r.Resolution == "Mixed" && r.Format == "Mixed");
            }
        });
        Test("commas in show folder names are retained when parsing known locations", () => {
            using (var f = new Fixture()) {
                string other = @"C:\KnownLocation\Love, Death & Robots (2019)";
                f.Add("S01E01.mkv"); f.Add("S01E02.mkv", directory: other);
                var show = Show(); show.Locations += ",\"" + other + "\"";
                var roots = new TvMediaDirectoryResolver(f.Files).Resolve(show);
                Check(roots.Count == 2 && roots.Contains(other) && f.Scan(show, roots: roots).Discovered == 2);
            }
        });
        Test("Found Locations scans successfully despite unavailable P drive and unrelated shares", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv");
                var show = Show(); show.Drive = @"P:\";
                var resolver = new TvMediaDirectoryResolver(f.Files);
                var roots = resolver.Resolve(show);
                Check(roots.SequenceEqual(new[] { ShowDirectory }) && f.Files.DirectoryReads.Count == 0);
                var result = f.Scan(show, roots: roots);
                Check(result.Success && result.Discovered == 1 && TvMediaSheet.Changes(show, result).Count == 2);
            }
        });
        Test("all listed UNC locations are scanned directly without discovery", () => {
            using (var f = new Fixture()) {
                string first = @"\\QUADPLEX\J-M\TV Shows\King of the Hill (1997)";
                string second = @"\\QUADPLEX\Other\The Bus\TV Shows\King of the Hill (1997)";
                f.Add("S01E01.mkv", directory: first);
                f.Add("S02E01.mp4", "720p", directory: second);
                var show = Show(); show.Locations = first + "," + second; show.Drive = @"P:\";
                var roots = new TvMediaDirectoryResolver(f.Files).Resolve(show);
                Check(roots.Count == 2 && f.Files.DirectoryReads.Count == 0);
                var result = f.Scan(show, roots: roots);
                Check(result.Success && result.Discovered == 2 && result.Format == "Mixed" && result.Resolution == "Mixed");
            }
        });
        Test("unreadable listed location preserves cells without falling back to another directory", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv");
                var show = Show(); show.Locations += @",C:\MissingListedLocation";
                var roots = new TvMediaDirectoryResolver(f.Files).Resolve(show);
                Check(roots.Count == 2 && f.Files.DirectoryReads.Count == 0);
                var result = f.Scan(show, roots: roots);
                Check(!result.Success && result.Probed == 0 && TvMediaSheet.Changes(show, result).Count == 0);
            }
        });
        Test("incompatible schema invalidates media cache", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); f.Scan(); f.Cache.Compact(); f.Cache.Dispose();
                var doc = JsonConvert.DeserializeObject<TvMediaCacheDocument>(File.ReadAllText(f.CachePath)); doc.SchemaVersion = 0;
                File.WriteAllText(f.CachePath, JsonConvert.SerializeObject(doc)); f.Cache = new TvMediaCache(f.CachePath, f.Logs.Add);
                Check(f.Scan().Probed == 1);
            }
        });
        Test("corrupt cache safely rebuilds from valid per-show checkpoints", () => {
            using (var f = new Fixture()) {
                f.Add("S01E01.mkv"); f.Scan(); f.Cache.Dispose(); File.WriteAllText(f.CachePath, "{broken json");
                f.Cache = new TvMediaCache(f.CachePath, f.Logs.Add);
                Check(f.Scan().Reused == 1 && f.Logs.Any(l => l.Contains("Unreadable cache")));
            }
        });
        Test("concurrent cache writer rejected without overwriting active cache", () => {
            using (var f = new Fixture()) Fails<IOException>(() => { using (new TvMediaCache(f.CachePath, f.Logs.Add)) { } });
        });
        Test("skips are yellow and errors are red, including formula write failures", () => {
            Check(TvMediaScanReport.Color("  SKIPPED: Found Locations is empty.") == "Yellow");
            Check(TvMediaScanReport.Color("TV Shows skipped: 68") == "Yellow");
            Check(TvMediaScanReport.Color("TV Shows with errors: 2") == "Red");
            Check(TvMediaScanReport.Color("Scan errors: 1") == "Red");
            Check(TvMediaScanReport.Color("ERROR: formula validation failed") == "Red");
            Check(TvMediaScanReport.Color("FFProbe/file metadata failure for episode.mkv") == "Red");
            Check(TvMediaScanReport.Color("TV Shows with errors: 0") == "DarkGray");
        });
        Test("errors-only report includes show and row context and excludes skips", () => {
            var report = new TvMediaScanReport();
            report.Record("  SKIPPED: Found Locations is empty.", Show());
            report.Record("Unsupported format AVI", Show());
            Check(report.ErrorDetails.Count == 0);
            report.Record("FFProbe/file metadata failure for C:\\Show\\S01E01.mkv: timeout", Show());
            report.Record("  ERROR: Incomplete directory/file inventory: missing path", Show());
            report.Record("Scan stopped (IOException): write failed", null);
            Check(report.ErrorDetails.Count == 3);
            Check(report.ErrorDetails[0].Contains("Show (2020) (DB row 3)") && report.ErrorDetails[0].Contains("S01E01.mkv"));
            Check(report.ErrorDetails[1].Contains("missing path"));
            Check(report.ErrorDetails[2].StartsWith("Scan: Scan stopped"));
        });
        Test("normal skips stay separate from true errors in final counters", () => {
            var report = new TvMediaScanReport { Examined = 1046, Unchanged = 349, Updated = 629, Skipped = 68 };
            report.CompleteUnfinished(false);
            Check(report.Skipped == 68 && report.Errors == 0);
        });
        Test("failed pending writes count as errors without counting completed shows twice", () => {
            var report = new TvMediaScanReport { Examined = 10, Unchanged = 2, Updated = 3, Skipped = 1, Errors = 1 };
            report.CompleteUnfinished(false);
            Check(report.Errors == 4 && report.Skipped == 1);
            report.CompleteUnfinished(false);
            Check(report.Errors == 4);
        });
        Test("canceled pending writes are skips and preserve existing errors", () => {
            var report = new TvMediaScanReport { Examined = 10, Unchanged = 2, Updated = 3, Skipped = 1, Errors = 1 };
            report.CompleteUnfinished(true);
            Check(report.Skipped == 4 && report.Errors == 1);
        });
        Console.WriteLine("Passed " + passed + " TV episode media tests.");
    }
}
