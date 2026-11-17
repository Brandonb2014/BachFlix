using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using BachFlixNfo;
using BachFlixNfo.Features;
using Newtonsoft.Json.Linq;
using RestSharp;
using TvdbApiCall;

class MissingEpisodeTests
{
    static int passed;
    static void Check(bool b, string message = "Assertion failed") { if (!b) throw new Exception(message); }
    static void Test(string name, Action action) { action(); passed++; Console.WriteLine("PASS " + name); }
    static void Fails<T>(Action action) where T : Exception
    { try { action(); } catch (T) { return; } throw new Exception("Expected " + typeof(T).Name); }
    static JObject J(string json) { return JObject.Parse(json); }
    static JObject Series(string id = "e1", string tvdb = "42", string title = "Show", int year = 2020)
    { return new JObject { ["Id"] = id, ["Name"] = title, ["ProductionYear"] = year, ["ProviderIds"] = new JObject { ["Tvdb"] = tvdb, ["Imdb"] = "tt001" } }; }
    static JObject Owned(int season, int first, int? last = null)
    { return new JObject { ["ParentIndexNumber"] = season, ["IndexNumber"] = first, ["IndexNumberEnd"] = last,
        ["MediaSources"] = new JArray(new JObject { ["Protocol"] = "File", ["Type"] = "Default", ["IsRemote"] = false,
            ["Path"] = "C:/Fake/episode.mkv", ["Size"] = 1000, ["SupportsDirectPlay"] = true }) }; }
    static JObject Aired(int season, int episode, string date = "2020-01-01")
    { return new JObject { ["seasonNumber"] = season, ["number"] = episode, ["aired"] = date }; }
    static MissingEpisodeShow Show(string tvdb = "42", string title = "Show")
    { return new MissingEpisodeShow { Row = 3, OutputColumn = 4, Title = title, TvdbId = tvdb, Year = "2020" }; }
    static readonly DateTime Today = new DateTime(2026, 9, 23);
    static MissingEpisodePreview Preview(IList<JObject> series = null, IList<JObject> owned = null, IList<JObject> aired = null)
    { return new MissingEpisodePreview(series ?? new[] { Series() }, _ => owned ?? new[] { Owned(1, 1) }, _ => aired ?? new[] { Aired(1, 1) }); }
    static IList<IList<object>> Rows(params object[][] rows) { return rows.Select(r => (IList<object>)r.ToList()).ToList(); }

    sealed class Handler : HttpMessageHandler
    {
        internal Func<HttpRequestMessage, HttpResponseMessage> Reply;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage r, CancellationToken c) { return Task.FromResult(Reply(r)); }
    }
    static HttpResponseMessage Http(string body, int status = 200)
    { return new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent(body) }; }
    static IRestResponse TvResponse(string body, int status = 200)
    { return new RestResponse { Content = body, StatusCode = (HttpStatusCode)status, ResponseStatus = ResponseStatus.Completed }; }
    static string Page(int number, string next)
    { var episode = Aired(1, number); episode["id"] = number; episode["seriesId"] = 42;
        return new JObject { ["status"] = "success", ["data"] = new JObject { ["series"] = new JObject { ["id"] = 42 }, ["episodes"] = new JArray(episode) }, ["links"] = new JObject { ["next"] = next } }.ToString(); }

    static void Main()
    {
        Test("single, inclusive multi-episode, duplicates and sorted missing pairs", () => {
            var r = Preview(owned: new[] { Owned(2, 3, 4), Owned(2, 4), Owned(1, 1) },
                aired: new[] { Aired(2, 5), Aired(2, 4), Aired(2, 3), Aired(1, 2), Aired(1, 1), Aired(1, 2) }).Check(Show(), Today);
            Check(r.Status == MissingEpisodeStatus.Success && r.OwnedEpisodeCount == 3 && r.AiredEpisodeCount == 5 && r.ProposedValue == 2);
            Check(string.Join(",", r.MissingEpisodes) == "S01E02,S02E05");
        });
        Test("complete show explicitly proposes numeric zero", () => Check(Preview().Check(Show(), Today).ProposedValue == 0));
        Test("owned outside aired explains count discrepancy without changing missing tally", () => {
            var r = Preview(owned: new[] { Owned(1, 1), Owned(1, 3), Owned(1, 3) },
                aired: new[] { Aired(1, 1), Aired(1, 2), Aired(1, 3, "2027-01-01") }).Check(Show(), Today);
            Check(r.OwnedEpisodeCount == 2 && r.AiredEpisodeCount == 2 && r.MissingCount == 1);
            Check(string.Join(",", r.OwnedOutsideAired) == "S01E03");
        });
        Test("conflicting ID reports exact title and alternate ID candidates without fetching episodes", () => {
            var show = Show("99"); show.ImdbId = "tt001";
            var p = new MissingEpisodePreview(new[] { Series(), Series("e2", "77", "Other") },
                _ => { throw new Exception("Unexpected fetch"); }, _ => { throw new Exception("Unexpected fetch"); });
            var r = p.Check(show, Today);
            Check(r.Status == MissingEpisodeStatus.NoEmbyMatch && r.ProposedValue == null);
            Check(r.MatchAttempt == "Tvdb ID: 99" && r.MatchCandidates.Count == 2);
            Check(r.MatchCandidates[0].Contains("TVDB=42"));
        });
        Test("today included, future, unknown date, season zero and invalid numbers excluded", () => {
            var pairs = MissingEpisodePreview.Aired(new[] { Aired(1, 1, "2026-09-23"), Aired(1, 2, "2026-09-24"), Aired(0, 1), Aired(1, 3, null), Aired(1, 4, "bad"), Aired(1, 0) }, Today);
            Check(pairs.Count == 1 && pairs.Contains(new EpisodePair(1, 1)));
        });
        Test("invalid Emby numbering ignored, specials and virtual episodes excluded", () => {
            int ignored;
            var virtualItem = Owned(1, 8); virtualItem["LocationType"] = "Virtual";
            var pairs = MissingEpisodePreview.Owned(new[] { Owned(1, 1), Owned(0, 2), Owned(1, 4, 3), J("{}"), Owned(-1, 1), virtualItem }, out ignored);
            Check(pairs.Count == 1 && ignored == 3);
        });
        Test("all unusable items fail instead of proposing an inaccurate count", () => {
            var r = Preview(owned: new[] { J("{}") }).Check(Show(), Today);
            Check(r.Status == MissingEpisodeStatus.EmbyEpisodesUnavailable && r.ProposedValue == null && !r.Abort);
        });
        Test("empty valid Emby inventory means all aired episodes missing", () => Check(Preview(owned: new JObject[0]).Check(Show(), Today).ProposedValue == 1));
        Test("TVDB ID takes precedence over conflicting title and IMDb", () => {
            var r = Preview(new[] { Series(title: "Different Title"), Series("e2", "99", "Show") }).Check(Show(), Today);
            Check(r.EmbySeries == "Different Title" && r.ProposedValue == 0);
        });
        Test("unmatched ID never falls back to an unrelated title", () => {
            var r = Preview().Check(Show("99"), Today);
            Check(r.Status == MissingEpisodeStatus.NoEmbyMatch && r.ProposedValue == null && r.ProposedText == MissingEpisodeFormat.NotInEmby);
        });
        Test("same provider ID in two Emby series combines their episodes", () => {
            var p = new MissingEpisodePreview(new[] { Series(), Series("e2") },
                id => id == "e1" ? new[] { Owned(1, 1) } : new[] { Owned(1, 2) },
                _ => new[] { Aired(1, 1), Aired(1, 2), Aired(1, 3) });
            var r = p.Check(Show(), Today);
            Check(r.Status == MissingEpisodeStatus.Success && r.OwnedEpisodeCount == 2 && r.ProposedValue == 1);
            Check(r.MissingEpisodes.Single().Equals(new EpisodePair(1, 3)) && r.Warnings.Any(w => w.Contains("combined")));
        });
        Test("identical Emby copies count each episode once", () => {
            var r = Preview(new[] { Series(), Series("e2") }).Check(Show(), Today);
            Check(r.Status == MissingEpisodeStatus.Success && r.OwnedEpisodeCount == 1 && r.ProposedValue == 0);
        });
        Test("uncertain items in either Emby copy preserve the existing value", () => {
            var bad = Owned(1, 1); bad.Remove("MediaSources");
            var p = new MissingEpisodePreview(new[] { Series(), Series("e2") },
                id => id == "e1" ? new[] { Owned(1, 1) } : new[] { bad }, _ => new[] { Aired(1, 1) });
            Check(p.Check(Show(), Today).Status == MissingEpisodeStatus.EmbyEpisodesUnavailable);
        });
        Test("a conflicting IMDb ID on any Emby copy is an identity conflict", () => {
            var other = Series("e2"); other["ProviderIds"]["Imdb"] = "tt999";
            var show = Show(); show.ImdbId = "tt001";
            Check(Preview(new[] { Series(), other }).Check(show, Today).Status == MissingEpisodeStatus.IdentityConflict);
        });
        Test("title-only matches to two Emby series stay ambiguous", () => {
            var show = Show("");
            var r = Preview(new[] { Series(), Series("e2", "77") }).Check(show, Today);
            Check(r.Status == MissingEpisodeStatus.AmbiguousEmbyMatch && r.ProposedText == null);
        });
        Test("IMDb fallback resolves TVDB from matched Emby series", () => {
            var show = Show(""); show.ImdbId = "tt001";
            Check(Preview().Check(show, Today).TvdbId == "42");
        });
        Test("TMDB fallback when higher priority IDs absent", () => {
            var show = Show(""); show.TmdbId = "55";
            var series = Series(); series["ProviderIds"]["Tmdb"] = "55";
            Check(Preview(new[] { series }).Check(show, Today).ProposedValue == 0);
        });
        Test("title fallback uses year and rejects ambiguity", () => {
            var series = new[] { Series(), Series("e2", "99", year: 2021) };
            Check(Preview(series).Check(Show(""), Today).TvdbId == "42");
            var show = Show(""); show.Year = "";
            Check(Preview(series).Check(show, Today).Status == MissingEpisodeStatus.AmbiguousEmbyMatch);
        });
        Test("no TVDB ID fails without searching titles", () => {
            var r = Preview(new[] { Series(tvdb: "") }).Check(Show(""), Today);
            Check(r.Status == MissingEpisodeStatus.MissingTvdbId && r.ProposedValue == null);
        });
        Test("preview maps moved headers and physical rows without modifying sheet input", () => {
            var headers = Rows(new object[] { "Series Name", "Missing Episodes", "TVDB ID", "RowNum" });
            var rows = Rows(new object[] { "Show", 98, "42", 999 }, new object[0], new object[] { "Missing", 7, "99", 1 });
            var shows = MissingEpisodePreview.ReadShows(headers, rows);
            var preview = Preview(); var before = Newtonsoft.Json.JsonConvert.SerializeObject(rows);
            Check(shows[0].OutputCell == "DB!B3" && shows[1].OutputCell == "DB!B5");
            Check(preview.Check(shows[0], Today).ProposedValue == 0 && preview.Check(shows[1], Today).ProposedValue == null);
            Check(before == Newtonsoft.Json.JsonConvert.SerializeObject(rows));
            Check(SheetHeaderMapping.ColumnLetter(26) == "AA" && SheetHeaderMapping.ColumnLetter(702) == "AAA");
        });
        Test("missing and duplicate headers stop before scanning", () => {
            Fails<MissingEpisodeSheetException>(() => MissingEpisodePreview.ReadShows(Rows(new object[] { "TVDB ID", "Series Name" }), null));
            Fails<MissingEpisodeSheetException>(() => MissingEpisodePreview.ReadShows(Rows(new object[] { "Missing Episodes", "Missing Episodes", "TVDB ID", "Series Name" }), null));
        });
        Test("per-run Emby and TVDB cache avoids repeat fetches", () => {
            int e = 0, t = 0;
            var p = new MissingEpisodePreview(new[] { Series() }, _ => { e++; return new[] { Owned(1, 1) }; }, _ => { t++; return new[] { Aired(1, 1) }; });
            p.Check(Show(), Today); p.Check(Show(), Today); Check(e == 1 && t == 1);
        });
        Test("unmatched series does not fetch episodes", () => {
            var p = new MissingEpisodePreview(new JObject[0], _ => { throw new Exception(); }, _ => { throw new Exception(); });
            Check(p.Check(Show(), Today).ProposedValue == null);
        });
        Test("TVDB authentication and API failures abort without proposed values; 404 is per show", () => {
            foreach (var failure in new[] { TvdbFailure.Authentication, TvdbFailure.Api, TvdbFailure.SeriesNotFound }) {
                var p = new MissingEpisodePreview(new[] { Series() }, _ => new[] { Owned(1, 1) }, _ => { throw new TvdbException(failure, "safe failure"); });
                var r = p.Check(Show(), Today); Check(r.ProposedValue == null && r.Abort == (failure != TvdbFailure.SeriesNotFound));
            }
        });
        Test("Emby systemic failures abort, show-level failures can continue", () => {
            foreach (bool systemic in new[] { true, false }) {
                var p = new MissingEpisodePreview(new[] { Series() }, _ => { throw new EmbyException("safe failure", systemic); }, _ => new JObject[0]);
                var r = p.Check(Show(), Today); Check(r.ProposedValue == null && r.Abort == systemic && r.Status == MissingEpisodeStatus.EmbyEpisodesUnavailable);
            }
        });
        Test("Emby GET pagination, provider fields and header authentication", () => {
            int calls = 0;
            var handler = new Handler { Reply = request => {
                Check(request.Method == HttpMethod.Get && request.Headers.GetValues("X-Emby-Token").Single() == "fake-key");
                Check(!request.RequestUri.ToString().Contains("fake-key") && request.RequestUri.Query.Contains("Fields=ProviderIds"));
                Check(request.RequestUri.Query.Contains("StartIndex=" + calls)); calls++;
                return Http("{\"Items\":[{\"Id\":\"" + calls + "\"}],\"TotalRecordCount\":2}");
            }};
            using (var c = new EmbyClient("http://test/emby", "fake-key", handler)) Check(c.GetSeries().Count == 2);
            Check(calls == 2);
        });
        Test("Emby episode query requests media evidence and a stable sort", () => {
            var h = new Handler { Reply = r => { string q = r.RequestUri.Query;
            Check(q.Contains("ParentId=e1") && q.Contains("IncludeItemTypes=Episode") && q.Contains("Recursive=true") && q.Contains("MediaSources") && q.Contains("Path") && q.Contains("SortBy=SortName"));
                return Http("{\"Items\":[],\"TotalRecordCount\":0}"); } };
            using (var c = new EmbyClient("http://test", "fake-key", h)) Check(c.GetEpisodes("e1").Count == 0);
        });
        Test("Emby malformed, partial and repeated pages fail safely", () => {
            foreach (string body in new[] { "bad-secret", "{}", "{\"Items\":[],\"TotalRecordCount\":2}", "{\"Items\":[{\"Id\":\"1\"}],\"TotalRecordCount\":2}" }) {
                using (var c = new EmbyClient("http://test", "fake-key", new Handler { Reply = _ => Http(body) })) Fails<EmbyException>(() => c.GetSeries());
            }
        });
        Test("Emby error redacts response body", () => {
            using (var c = new EmbyClient("http://test", "fake-key", new Handler { Reply = _ => Http("private-secret", 401) })) {
                try { c.GetSeries(); throw new Exception("Expected failure"); }
                catch (EmbyException ex) { Check(ex.Systemic && !ex.ToString().Contains("private-secret") && !ex.ToString().Contains("fake-key")); }
            }
        });
        Test("shared movie refresh preserves POST refresh parameters", () => {
            using (var c = new EmbyClient("http://test", "fake-key", new Handler { Reply = r => {
                Check(r.Method == HttpMethod.Post && r.RequestUri.AbsolutePath == "/emby/Items/3/Refresh");
                Check(r.RequestUri.Query.Contains("ImageRefreshMode=FullRefresh") && r.RequestUri.Query.Contains("ReplaceAllMetadata=false")); return Http("");
            }})) c.RefreshMovies("3");
        });
        Test("TVDB complete inventory follows pages using one login", () => {
            int calls = 0;
            var c = new TvdbClient(() => "fake-key", r => {
                calls++;
                if (calls == 1) return TvResponse("{\"status\":\"success\",\"data\":{\"token\":\"fake-token\"}}");
                Check(new RestClient(TvdbClient.BaseUrl).BuildUri(r).AbsoluteUri == TvdbClient.BaseUrl + "/series/42/episodes/official?page=" + (calls - 2));
                return TvResponse(Page(calls - 1, calls == 2 ? "https://api4.thetvdb.com/v4/series/42/episodes/official?page=1" : null));
            });
            Check(c.GetOfficialEpisodes("42").Count == 2 && calls == 3);
        });
        Test("TVDB English inventory uses the /eng endpoint", () => {
            int calls = 0;
            var c = new TvdbClient(() => "fake-key", r => {
                calls++;
                if (calls == 1) return TvResponse("{\"status\":\"success\",\"data\":{\"token\":\"fake-token\"}}");
                Check(new RestClient(TvdbClient.BaseUrl).BuildUri(r).AbsoluteUri == TvdbClient.BaseUrl + "/series/42/episodes/official/eng?page=" + (calls - 2));
                return TvResponse(Page(calls - 1, calls == 2 ? "https://api4.thetvdb.com/v4/series/42/episodes/official/eng?page=1" : null));
            });
            Check(c.GetOfficialEpisodes("42", "eng").Count == 2 && calls == 3);
            Fails<ArgumentException>(() => c.GetOfficialEpisodes("42", "../x"));
        });
        Test("TVDB missing pagination, looping pagination and incomplete totals fail", () => {
            foreach (string body in new[] { "{\"status\":\"success\",\"data\":{\"episodes\":[]}}", Page(1, "?page=0"), "{\"status\":\"success\",\"data\":{\"episodes\":[]},\"links\":{\"next\":null,\"total_items\":2}}" }) {
                var c = new TvdbClient(() => "fake-key", r => TvResponse(r.Resource == "login" ? "{\"status\":\"success\",\"data\":{\"token\":\"fake-token\"}}" : body));
                Fails<TvdbException>(() => c.GetOfficialEpisodes("42"));
            }
        });
        passed += MissingEpisodeWorkflowTests.Run();
        Console.WriteLine(passed + " missing-episode tests passed.");
    }
}
