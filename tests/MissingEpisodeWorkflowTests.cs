using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using BachFlixNfo;
using BachFlixNfo.Features;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using TvdbApiCall;

internal static class MissingEpisodeWorkflowTests
{
    static int passed;
    static readonly DateTime Today = new DateTime(2026, 10, 17);
    static void Check(bool value, string message = "Assertion failed") { if (!value) throw new Exception(message); }
    static void Test(string name, Action body) { body(); passed++; Console.WriteLine("PASS " + name); }
    static void Throws<T>(Action body) where T : Exception { try { body(); } catch (T) { return; } throw new Exception("Expected " + typeof(T).Name); }
    static JObject Owned(int episode = 1) { return JObject.Parse("{\"ParentIndexNumber\":1,\"IndexNumber\":" + episode + ",\"MediaSources\":[{\"Protocol\":\"File\",\"Type\":\"Default\",\"IsRemote\":false,\"Path\":\"C:/Fake/ep.mkv\",\"Size\":1234,\"SupportsDirectPlay\":true}]}"); }
    static JObject Episode(int number = 1, string date = "2020-01-01") { return new JObject { ["seasonNumber"] = 1, ["number"] = number, ["aired"] = date }; }
    static JObject Series() { return JObject.Parse("{\"Id\":\"emby-42\",\"Name\":\"Show\",\"ProductionYear\":2020,\"ProviderIds\":{\"Tvdb\":\"42\",\"Imdb\":\"tt42\"}}"); }
    static MissingEpisodeShow Show() { return new MissingEpisodeShow { Row = 3, OutputColumn = 15, Title = "Show", TvdbId = "42", Year = "2020" }; }
    static MissingEpisodeResult Compare(IList<JObject> owned, IList<JObject> episodes, MissingEpisodeShow show = null, JObject series = null)
    { return new MissingEpisodePreview(new[] { series ?? Series() }, _ => owned, _ => episodes).Check(show ?? Show(), Today); }
    static void NoValue(MissingEpisodeResult r) { Check(!r.ProposedValue.HasValue && r.Status != MissingEpisodeStatus.Success); }

    sealed class FakeSheet : IMissingEpisodeSheet
    {
        internal readonly List<MissingEpisodeShow> State = new List<MissingEpisodeShow>();
        internal readonly List<int> BatchSizes = new List<int>();
        internal int FailAfter = -1, WriteCalls;
        internal bool FailRead;
        internal FakeSheet(int count = 1)
        { for (int i = 0; i < count; i++) State.Add(new MissingEpisodeShow { Row = i + 3, OutputColumn = 15, TvdbId = (i + 42).ToString(), Title = "Show " + i }); }
        public IList<MissingEpisodeShow> Read()
        {
            if (FailRead) throw new Exception("read unavailable");
            return State.Select(s => new MissingEpisodeShow { Row = s.Row, OutputColumn = s.OutputColumn, TvdbId = s.TvdbId, Title = s.Title, ExistingValue = s.ExistingValue, OutputIsFormula = s.OutputIsFormula }).ToList();
        }
        public void Write(IList<MissingEpisodeCellChange> changes)
        {
            WriteCalls++; BatchSizes.Add(changes.Count);
            for (int i = 0; i < changes.Count; i++)
            {
                if (FailAfter == i) throw new Exception("lost acknowledgement");
                State.Single(s => s.Row == changes[i].Result.Show.Row).ExistingValue = changes[i].Value;
            }
            if (FailAfter == changes.Count) throw new Exception("lost acknowledgement after commit");
        }
    }
    static MissingEpisodeResult Success(MissingEpisodeShow show, int count = 0)
    { return new MissingEpisodeResult { Show = show, Status = MissingEpisodeStatus.Success, MissingEpisodes = Enumerable.Range(1, count).Select(n => new EpisodePair(1, n)).ToList() }; }
    static MissingEpisodeRun Scan(FakeSheet sheet, bool update = true, Func<MissingEpisodeShow, MissingEpisodeResult> check = null, CancellationToken token = default(CancellationToken), IList<MissingEpisodeShow> initial = null)
    {
        var shows = initial ?? sheet.Read();
        var run = new MissingEpisodeRun { UpdateMode = update, TotalShows = shows.Count };
        MissingEpisodeWorkflow.Execute(run, shows, check ?? (s => Success(s)), update ? sheet : null, token, null);
        return run;
    }
    sealed class Handler : HttpMessageHandler
    {
        internal Func<HttpRequestMessage, HttpResponseMessage> Reply;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage r, CancellationToken token) { return Task.FromResult(Reply(r)); }
    }
    static IRestResponse Response(string body, int status = 200) { return new RestResponse { Content = body, StatusCode = (HttpStatusCode)status, ResponseStatus = ResponseStatus.Completed }; }
    static string Page(int id, string next = null, int? total = null, int series = 42)
    { return new JObject { ["status"] = "success", ["data"] = new JObject { ["series"] = new JObject { ["id"] = series }, ["episodes"] = new JArray(new JObject { ["id"] = id, ["seriesId"] = series }) }, ["links"] = new JObject { ["next"] = next, ["total_items"] = total } }.ToString(); }
    static TvdbClient Tvdb(Func<RestRequest, IRestResponse> request, Action<TimeSpan> pause = null)
    { return new TvdbClient(() => "fake-key", r => r.Resource == "login" ? Response("{\"status\":\"success\",\"data\":{\"token\":\"fake-token\"}}") : request(r), pause ?? (_ => { })); }

    internal static int Run()
    {
        Test("empty, malformed and specials-only TVDB inventories never propose zero", () => {
            foreach (var episodes in new[] { new JObject[0], new[] { new JObject() }, new[] { JObject.Parse("{\"seasonNumber\":0}") } }) NoValue(Compare(new[] { Owned() }, episodes));
        });
        Test("unreadable dates and invalid TVDB numbering block all proposals", () => {
            foreach (var ep in new[] { Episode(2, "bad"), Episode(0), JObject.Parse("{\"number\":2,\"aired\":\"2020-01-01\"}") }) NoValue(Compare(new[] { Owned() }, new[] { Episode(), ep }));
        });
        Test("undated episode after the last aired one is not aired yet", () => {
            var r = Compare(new[] { Owned() }, new[] { Episode(), Episode(2, null), Episode(3, "") });
            Check(r.ProposedValue == 0 && r.ProposedText == MissingEpisodeFormat.NoneMissing, "Upcoming undated episodes counted as missing");
        });
        Test("undated episode before an aired one counts as aired, with a warning", () => {
            var r = Compare(new[] { Owned() }, new[] { Episode(), Episode(2, null), Episode(3, "2021-01-01") });
            Check(r.ProposedValue == 2 && r.MissingEpisodes.Contains(new EpisodePair(1, 2)), "Old undated episode not counted");
            Check(r.Warnings.Any(w => w.Contains("no air date")), "Missing undated warning");
        });
        Test("undated episode before only future-dated ones is not aired yet", () => {
            var r = Compare(new[] { Owned() }, new[] { Episode(), Episode(2, null), Episode(3, "2027-01-01") });
            Check(r.ProposedValue == 0, "Undated episode before an unaired one counted as missing");
        });
        Test("owned undated episode is not a numbering conflict", () => {
            var r = Compare(new[] { Owned(), Owned(2) }, new[] { Episode(), Episode(2, null), Episode(3, "2021-01-01") });
            Check(r.Status == MissingEpisodeStatus.Success && r.ProposedValue == 1, "Owned undated episode rejected");
        });
        Test("inventory with no air dates at all still blocks proposals", () => NoValue(Compare(new[] { Owned() }, new[] { Episode(1, null), Episode(2, null) })));
        Test("conflicting duplicate official dates use the earliest date, with a warning", () => {
            var r = Compare(new JObject[0], new[] { Episode(1, "2027-01-01"), Episode(1, "2021-01-01") });
            Check(r.ProposedValue == 1 && r.Warnings.Any(w => w.Contains("earliest")), "Earliest duplicate date not used");
            var later = Compare(new JObject[0], new[] { Episode(1, "2021-01-01"), Episode(1, "2027-01-01") });
            Check(later.ProposedValue == 1, "Duplicate order changed the result");
        });
        Test("fully dated future inventory is a verified empty aired set", () => Check(Compare(new JObject[0], new[] { Episode(1, "2027-01-01") }).ProposedValue == 0));
        Test("mixed invalid owned numbering including special cannot yield a count", () => {
            var invalid = Owned(2); invalid.Remove("IndexNumber");
            var special = Owned(); special["ParentIndexNumber"] = 0;
            NoValue(Compare(new[] { Owned(), invalid }, new[] { Episode(), Episode(2) }));
            NoValue(Compare(new[] { special, invalid }, new[] { Episode(), Episode(2) }));
        });
        Test("remote, virtual, placeholders and empty media sources are not owned", () => {
            var remote = Owned(); remote["MediaSources"][0]["Protocol"] = "Http"; remote["MediaSources"][0]["IsRemote"] = true;
            var virtualItem = new JObject { ["IsVirtualItem"] = true };
            var placeholder = Owned(); placeholder["MediaSources"][0]["Type"] = "Placeholder";
            var strm = Owned(); strm["MediaSources"][0]["Path"] = "C:/Fake/reference.strm";
            var empty = new JObject { ["MediaSources"] = new JArray() };
            Check(Compare(new[] { remote, virtualItem, placeholder, strm, empty }, new[] { Episode() }).ProposedValue == 1);
        });
        Test("missing ownership evidence fails; LocationType is not required for a real source", () => {
            NoValue(Compare(new[] { JObject.Parse("{\"ParentIndexNumber\":1,\"IndexNumber\":1}") }, new[] { Episode() }));
            Check(Compare(new[] { Owned() }, new[] { Episode() }).ProposedValue == 0);
        });
        Test("unverifiable local source cannot silently count as absent", () => {
            foreach (string field in new[] { "Path", "Size", "IsRemote", "Protocol", "SupportsDirectPlay" }) {
                var item = Owned(); ((JObject)item["MediaSources"][0]).Remove(field);
                NoValue(Compare(new[] { item }, new[] { Episode() }));
            }
        });
        Test("owned official-order mismatch blocks tally but a known future owned pair does not", () => {
            NoValue(Compare(new[] { Owned(), Owned(2) }, new[] { Episode() }));
            Check(Compare(new[] { Owned(), Owned(2) }, new[] { Episode(), Episode(2, "2027-01-01") }).ProposedValue == 0);
        });
        Test("harmless accents, punctuation and year suffix do not reject TVDB identity", () => {
            var s = Series(); s["Name"] = "Pokemon"; s["ProductionYear"] = 1997;
            var show = Show(); show.Title = "Pokémon (1997)"; show.Year = "1997";
            Check(Compare(new[] { Owned() }, new[] { Episode() }, show, s).ProposedValue == 0);
        });
        Test("alias alone and year alone warn; combined conflict blocks", () => {
            var alias = Show(); alias.Title = "Translated alias";
            Check(Compare(new[] { Owned() }, new[] { Episode() }, alias).ProposedValue == 0);
            var year = Show(); year.Year = "1990";
            Check(Compare(new[] { Owned() }, new[] { Episode() }, year).Warnings.Count > 0);
            alias.Year = "1990"; NoValue(Compare(new[] { Owned() }, new[] { Episode() }, alias));
        });
        Test("conflicting alternate provider ID blocks a matching TVDB ID", () => {
            var show = Show(); show.ImdbId = "ttOTHER"; NoValue(Compare(new[] { Owned() }, new[] { Episode() }, show));
        });
        Test("continuing and ended labels use identical aired inventory rules", () => {
            foreach (string status in new[] { "Yes", "No", "Ended", "Continuing" }) { var show = Show(); show.Continuing = status; Check(Compare(new[] { Owned() }, new[] { Episode(), Episode(2), Episode(3, "2027-01-01") }, show).ProposedValue == 1); }
        });
        Test("duplicate normalized Sheet IDs rejected", () => Throws<MissingEpisodeSheetException>(() => MissingEpisodePreview.ReadShows(
            new List<IList<object>> { new object[] { "Series Name", "TVDB ID", "Missing Episodes" } },
            new List<IList<object>> { new object[] { "A", "42" }, new object[] { "B", "0042" } })));
        Test("preview cannot write even when changes and valid zero are proposed", () => {
            var sheet = new FakeSheet(2); var run = Scan(sheet, false, s => Success(s, s.Row == 3 ? 0 : 2));
            Check(sheet.WriteCalls == 0 && sheet.State.All(s => s.ExistingValue == null) && run.Updated == 0 && run.ProposedChanges == 2 && run.TotalMissing == 2);
        });
        Test("complete and missing episode lists write to exact mapped cells", () => {
            var sheet = new FakeSheet(2); var run = Scan(sheet, true, s => Success(s, s.Row == 3 ? 0 : 2));
            Check(run.Updated == 2 && (string)sheet.State[0].ExistingValue == "None" && (string)sheet.State[1].ExistingValue == "S01E01 & 02");
        });
        Test("old numeric counts are replaced by episode lists", () => {
            var sheet = new FakeSheet(2); sheet.State[0].ExistingValue = 0; sheet.State[1].ExistingValue = "2";
            var initial = sheet.Read();
            Check(Scan(sheet, true, s => Success(s, s.Row == 3 ? 0 : 2), initial: initial).Updated == 2);
        });
        Test("episode list formatting collapses runs and groups seasons", () => {
            Func<int, int[], IEnumerable<EpisodePair>> s = (season, eps) => eps.Select(e => new EpisodePair(season, e));
            Check(MissingEpisodeFormat.Describe(new EpisodePair[0]) == "None");
            Check(MissingEpisodeFormat.Describe(s(1, new[] { 4 })) == "S01E04");
            Check(MissingEpisodeFormat.Describe(s(1, new[] { 13, 4, 5, 6, 7, 8, 10 })) == "S01E04-08, 10 & 13");
            Check(MissingEpisodeFormat.Describe(s(2, new[] { 3, 4 }).Concat(s(3, new[] { 2, 3, 4, 5 }))) == "S02E03 & 04, S03E02-05");
            Check(MissingEpisodeFormat.Describe(s(1, new[] { 99, 100, 101, 105 })) == "S01E99-101 & 105");
        });
        Test("unchanged cells and reruns issue no writes", () => {
            var sheet = new FakeSheet(2); sheet.State[0].ExistingValue = "None"; sheet.State[1].ExistingValue = " None ";
            Check(Scan(sheet).Unchanged == 2 && sheet.WriteCalls == 0);
            sheet.State[0].ExistingValue = null; Check(Scan(sheet).Updated == 1);
            int calls = sheet.WriteCalls; Check(Scan(sheet).Unchanged == 2 && sheet.WriteCalls == calls);
        });
        Test("failed checks preserve existing values without affecting successful rows", () => {
            var sheet = new FakeSheet(2); sheet.State[0].ExistingValue = 9;
            var run = Scan(sheet, true, s => s.Row == 3 ? new MissingEpisodeResult { Show = s, Status = MissingEpisodeStatus.TvdbUnavailable } : Success(s));
            Check(run.Failed == 1 && run.Updated == 1 && (int)sheet.State[0].ExistingValue == 9);
        });
        Test("writes bounded to 50 rows", () => { var sheet = new FakeSheet(121); var run = Scan(sheet); Check(run.Updated == 121 && string.Join(",", sheet.BatchSizes) == "50,50,21"); });
        Test("row swap, deletion and output-column move preserve destinations", () => {
            foreach (int kind in new[] { 0, 1, 2 }) {
                var sheet = new FakeSheet(2); var initial = sheet.Read();
                if (kind == 0) { string id = sheet.State[0].TvdbId; sheet.State[0].TvdbId = sheet.State[1].TvdbId; sheet.State[1].TvdbId = id; }
                if (kind == 1) sheet.State.Clear();
                if (kind == 2) foreach (var s in sheet.State) s.OutputColumn++;
                var run = Scan(sheet, initial: initial); Check(run.SheetFailures == 2 && sheet.WriteCalls == 0);
            }
        });
        Test("concurrent cell edit and newly inserted formula are preserved", () => {
            var sheet = new FakeSheet(2); var initial = sheet.Read();
            sheet.State[0].ExistingValue = 5; sheet.State[1].OutputIsFormula = true;
            var run = Scan(sheet, initial: initial); Check(run.SheetFailures == 2 && sheet.WriteCalls == 0);
        });
        Test("pre-existing formulas are preserved in preview and update", () => {
            var sheet = new FakeSheet(); sheet.State[0].OutputIsFormula = true;
            Check(Scan(sheet).SheetFailures == 1 && Scan(sheet, false).SheetFailures == 1 && sheet.WriteCalls == 0);
        });
        Test("preflight failure aborts without writes", () => {
            var sheet = new FakeSheet(); var initial = sheet.Read(); sheet.FailRead = true;
            Check(Scan(sheet, initial: initial).Aborted && sheet.WriteCalls == 0);
        });
        Test("partial write failure is read back and rerun repairs only unconfirmed cells", () => {
            var sheet = new FakeSheet(2); sheet.FailAfter = 1;
            var run = Scan(sheet); Check(run.Aborted && run.Updated == 1 && run.SheetFailures == 1 && sheet.WriteCalls == 1);
            sheet.FailAfter = -1; var rerun = Scan(sheet); Check(rerun.Updated == 1 && rerun.Unchanged == 1 && sheet.WriteCalls == 2);
        });
        Test("lost acknowledgement after full commit is confirmed without write retry", () => {
            var sheet = new FakeSheet(2); sheet.FailAfter = 2; var run = Scan(sheet);
            Check(run.Aborted && run.Updated == 2 && sheet.WriteCalls == 1 && run.Results.All(r => r.SheetOutcome == "ConfirmedAfterWriteError"));
        });
        Test("cancellation preserves pending batch and produces a partial report", () => {
            var sheet = new FakeSheet(2); var source = new CancellationTokenSource();
            var run = Scan(sheet, true, s => { source.Cancel(); return Success(s); }, source.Token);
            Check(run.Aborted && sheet.WriteCalls == 0 && run.Results.Count == 1 && run.Results[0].SheetOutcome == "NotAttempted");
            Check(JsonConvert.SerializeObject(run.Report(Today)).Contains("NotChecked"));
        });
        Test("systemic failure does not flush pending verified rows", () => {
            var sheet = new FakeSheet(2); var run = Scan(sheet, true, s => s.Row == 3 ? Success(s) : new MissingEpisodeResult { Show = s, Status = MissingEpisodeStatus.TvdbAuthentication, Abort = true });
            Check(run.Aborted && sheet.WriteCalls == 0 && run.Results[0].SheetOutcome == "NotAttempted");
        });
        Test("report contains missing numbers, null failures and required totals", () => {
            var sheet = new FakeSheet(3); var run = Scan(sheet, false, s => s.Row == 5 ? new MissingEpisodeResult { Show = s, Status = MissingEpisodeStatus.NoEmbyMatch } : Success(s, s.Row == 3 ? 2 : 0));
            string report = JsonConvert.SerializeObject(run.Report(Today));
            Check(run.TotalShows == 3 && run.SuccessfullyChecked == 2 && run.WithMissing == 1 && run.ZeroMissing == 1 && run.TotalMissing == 2 && run.NotInEmby == 1 && run.Failed == 0);
            Check(run.Results.Single(r => r.Show.Row == 5).ProposedText == MissingEpisodeFormat.NotInEmby && run.Results.Single(r => r.Show.Row == 5).SheetOutcome == "WouldUpdate");
            Check(report.Contains("S01E02") && report.Contains("\"ProposedValue\":null") && report.Contains("RequiresAttention"));
        });
        Test("CLI preview stays read-only; explicit update and row selection parsed", () => {
            Check(!MissingEpisodeOptions.Parse(new[] { "--missing-episodes-preview", "3" }).Update);
            var options = MissingEpisodeOptions.Parse(new[] { "--missing-episodes-update", "--row", "742" }); Check(options.Update && options.Row == 742);
            foreach (var args in new[] { new[] { "--missing-episodes-preview", "0" }, new[] { "--missing-episodes-update", "--row", "2" }, new[] { "--missing-episodes-preview", "3", "extra" } }) Throws<ArgumentException>(() => MissingEpisodeOptions.Parse(args));
        });
        Test("TVDB 429 Retry-After and 503 retry then recover", () => {
            int calls = 0; var waits = new List<TimeSpan>();
            var c = Tvdb(r => { calls++; var response = calls == 1 ? Response("", 429) : calls == 2 ? Response("", 503) : Response(Page(1)); if (calls == 1) response.Headers.Add(new Parameter("Retry-After", "4", ParameterType.HttpHeader)); return response; }, waits.Add);
            Check(c.GetOfficialEpisodes("42").Count == 1 && calls == 3 && waits[0] == TimeSpan.FromSeconds(4));
        });
        Test("TVDB exhausted transient errors stop at three attempts", () => {
            int calls = 0; var c = Tvdb(r => { calls++; return Response("private-secret", 503); });
            Throws<TvdbException>(() => c.GetOfficialEpisodes("42")); Check(calls == 3);
        });
        Test("TVDB repeated IDs, series mismatch and changing totals rejected", () => {
            foreach (int kind in new[] { 0, 1, 2, 3 }) {
                int calls = 0;
                var c = Tvdb(r => { calls++; return Response(kind == 1 ? Page(1, series: 99) : calls == 1 ? Page(1, "?page=1", 2) : Page(kind == 0 ? 1 : 2, total: kind == 2 ? 3 : kind == 3 ? (int?)null : 2)); });
                if (kind == 3) Check(c.GetOfficialEpisodes("42").Count == 2); else Throws<TvdbException>(() => c.GetOfficialEpisodes("42"));
            }
        });
        Test("Emby transient GET retries but shared refresh POST never retries", () => {
            int calls = 0; var handler = new Handler { Reply = r => { calls++; return new HttpResponseMessage(calls < 3 ? HttpStatusCode.ServiceUnavailable : HttpStatusCode.OK) { Content = new StringContent("{\"Items\":[],\"TotalRecordCount\":0}") }; } };
            using (var c = new EmbyClient("http://fake", "fake-key", handler, _ => { })) Check(c.GetSeries().Count == 0 && calls == 3);
            calls = 0;
            using (var c = new EmbyClient("http://fake", "fake-key", new Handler { Reply = r => { calls++; return new HttpResponseMessage(HttpStatusCode.ServiceUnavailable); } }, _ => { })) Throws<EmbyException>(() => c.RefreshMovies("3"));
            Check(calls == 1);
        });
        Test("long Retry-After never causes an early retry", () => Throws<InvalidOperationException>(() => ApiRetry.Pause(0, TimeSpan.FromMinutes(2), _ => { throw new Exception("must not sleep"); })));
        Test("Emby episode parent/type mismatch and changing pagination totals fail safely", () => {
            foreach (int kind in new[] { 0, 1, 2 }) {
                int calls = 0;
                var handler = new Handler { Reply = r => {
                    calls++;
                    var item = new JObject { ["Id"] = calls.ToString(), ["Type"] = kind == 1 ? "Movie" : "Episode", ["SeriesId"] = kind == 0 ? "other" : "e1" };
                    return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(new JObject { ["Items"] = new JArray(item), ["TotalRecordCount"] = kind == 2 && calls > 1 ? 3 : 2 }.ToString()) };
                }};
                using (var client = new EmbyClient("http://fake", "fake-key", handler, _ => { })) Throws<EmbyException>(() => client.GetEpisodes("e1"));
            }
        });
        Test("Emby network errors retry three times without leaking transport text", () => {
            int calls = 0;
            using (var client = new EmbyClient("http://fake", "fake-key", new Handler { Reply = r => { calls++; throw new HttpRequestException("private-secret"); } }, _ => { })) {
                try { client.GetSeries(); throw new Exception("Expected failure"); }
                catch (EmbyException ex) { Check(calls == 3 && !ex.ToString().Contains("private-secret")); }
            }
        });
        Test("normalized cache evaluates air dates using each check's date", () => {
            int calls = 0;
            var preview = new MissingEpisodePreview(new[] { Series() }, _ => new JObject[0], _ => { calls++; return new[] { Episode(1, "2026-10-18") }; });
            Check(preview.Check(Show(), Today).ProposedValue == 0);
            Check(preview.Check(Show(), Today.AddDays(1)).ProposedValue == 1 && calls == 1);
        });
        return passed;
    }
}
