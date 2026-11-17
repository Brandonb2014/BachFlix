using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using BachFlixNfo.Features;
using Google.Apis.Http;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json.Linq;
using SheetsQuickstart;

// Compiled with the real runner/adapter, but without the application's credential loader.
namespace SheetsQuickstart
{
    partial class Program
    {
        private const string SPREADSHEET_ID = "offline-test";
        private static SheetsService CreateSheetsService() { throw new Exception("Tests must never load credentials."); }
        private static string GetEmbyServerUrl() { return ""; }
        private static string GetEmbyApiKey() { return ""; }
        private static void Type(string text, int speed = 0, int pause = 0, int lines = 1, string color = "") { }
        private static void DisplayMessage(string level, string text) { }
    }
}

internal static class MissingEpisodeSheetAdapterTests
{
    static int passed;
    static void Check(bool condition) { if (!condition) throw new Exception("Assertion failed"); }
    static void Test(string name, Action action) { action(); passed++; Console.WriteLine("PASS " + name); }
    static void Throws<T>(Action action) where T : Exception { try { action(); } catch (T) { return; } throw new Exception("Expected " + typeof(T).Name); }
    sealed class Handler : HttpMessageHandler
    {
        internal Func<HttpRequestMessage, HttpResponseMessage> Reply;
        internal int Posts, Gets;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellation)
        {
            if (request.Method == HttpMethod.Post) Posts++; else if (request.Method == HttpMethod.Get) Gets++; else throw new Exception("Unexpected HTTP method");
            // All HTTP is intercepted; this handler has no network-capable inner handler.
            return Task.FromResult(Reply(request));
        }
    }
    sealed class Factory : IHttpClientFactory
    {
        internal Handler Handler;
        public ConfigurableHttpClient CreateHttpClient(CreateHttpClientArgs args)
        { return new ConfigurableHttpClient(new ConfigurableMessageHandler(Handler)); }
    }
    static SheetsService Service(Handler handler)
    { return new SheetsService(new BaseClientService.Initializer { HttpClientFactory = new Factory { Handler = handler }, ApplicationName = "MissingEpisodeOfflineTests", GZipEnabled = false }); }
    static HttpResponseMessage Response(string json, int status = 200)
    { return new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json") }; }
    static MissingEpisodeCellChange Change(int row = 3, int count = 0)
    { return new MissingEpisodeCellChange { Result = new MissingEpisodeResult { Show = new MissingEpisodeShow { Row = row, OutputColumn = 15, TvdbId = "42" }, Status = MissingEpisodeStatus.Success, MissingEpisodes = Enumerable.Range(1, count).Select(n => new EpisodePair(1, n)).ToList() } }; }
    static int Main()
    {
        try { RunTests(); return 0; }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    static void RunTests()
    {
        Test("real Sheets adapter sends RAW episode text and exact single-cell ranges", () => {
            var handler = new Handler { Reply = request => {
                Check(request.Method == HttpMethod.Post && request.RequestUri.AbsolutePath.EndsWith("/values:batchUpdate"));
                var body = JObject.Parse(request.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                Check((string)body["valueInputOption"] == "RAW" && (string)body["data"][0]["range"] == "DB!P3");
                Check(body["data"][0]["values"][0][0].Type == JTokenType.String && (string)body["data"][0]["values"][0][0] == "None");
                Check((string)body["data"][1]["range"] == "DB!P4" && (string)body["data"][1]["values"][0][0] == "S01E01 & 02");
                return Response("{\"totalUpdatedCells\":2}");
            }};
            using (var service = Service(handler)) new GoogleMissingEpisodeSheet(service, "offline", true).Write(new[] { Change(), Change(4, 2) });
            Check(handler.Posts == 1);
        });
        Test("real preview adapter rejects writes before any HTTP", () => {
            var handler = new Handler { Reply = r => { throw new Exception("No HTTP expected"); } };
            using (var service = Service(handler)) Throws<InvalidOperationException>(() => new GoogleMissingEpisodeSheet(service, "offline", false).Write(new[] { Change() }));
            Check(handler.Posts == 0 && handler.Gets == 0);
        });
        Test("real adapter refuses failed results, duplicate destinations and oversized batches", () => {
            var handler = new Handler { Reply = r => { throw new Exception("No HTTP expected"); } };
            using (var service = Service(handler)) {
                var adapter = new GoogleMissingEpisodeSheet(service, "offline", true);
                var bad = Change(); bad.Result.Status = MissingEpisodeStatus.TvdbUnavailable;
                Throws<MissingEpisodeSheetException>(() => adapter.Write(new[] { bad }));
                Throws<MissingEpisodeSheetException>(() => adapter.Write(new[] { Change(), Change() }));
                Throws<MissingEpisodeSheetException>(() => adapter.Write(Enumerable.Range(3, 51).Select(n => Change(n)).ToList()));
            }
            Check(handler.Posts == 0);
        });
        Test("incomplete acknowledgement throws for workflow readback", () => {
            var handler = new Handler { Reply = r => Response("{\"totalUpdatedCells\":0}") };
            using (var service = Service(handler)) Throws<MissingEpisodeSheetException>(() => new GoogleMissingEpisodeSheet(service, "offline", true).Write(new[] { Change() }));
            Check(handler.Posts == 1);
        });
        Test("real adapter never retries an ambiguous 503 write", () => {
            var handler = new Handler { Reply = r => Response("{\"error\":{\"code\":503,\"message\":\"synthetic failure\"}}", 503) };
            using (var service = Service(handler)) Throws<Google.GoogleApiException>(() => new GoogleMissingEpisodeSheet(service, "offline", true).Write(new[] { Change() }));
            Check(handler.Posts == 1);
        });
        Test("real read adapter resolves moved column, physical blank rows and output formulas", () => {
            var handler = new Handler { Reply = r => {
                Check(r.Method == HttpMethod.Get);
                string uri = Uri.UnescapeDataString(r.RequestUri.ToString());
                if (uri.Contains("DB!A2:2")) return Response("{\"values\":[[\"Series Name\",\"TVDB ID\",\"Missing Episodes\"]]}");
                if (uri.Contains("valueRenderOption=FORMULA")) return Response("{\"values\":[[\"=0\"],[],[\"\"]]}");
                Check(uri.Contains("DB!A3:C") && uri.Contains("UNFORMATTED_VALUE"));
                return Response("{\"values\":[[\"First\",42,0],[],[\"Second\",43]]}");
            }};
            using (var service = Service(handler)) {
                var rows = new GoogleMissingEpisodeSheet(service, "offline", false).Read();
                Check(rows.Count == 2 && rows[0].OutputCell == "DB!C3" && rows[0].OutputIsFormula && rows[1].Row == 5 && !rows[1].OutputIsFormula);
            }
            Check(handler.Gets == 4 && handler.Posts == 0);
        });
        Test("missing headers abort real adapter before rows or writes", () => {
            var handler = new Handler { Reply = r => Response("{\"values\":[[\"Series Name\",\"TVDB ID\"]]}") };
            using (var service = Service(handler)) Throws<MissingEpisodeSheetException>(() => new GoogleMissingEpisodeSheet(service, "offline", true).Read());
            Check(handler.Gets == 1 && handler.Posts == 0);
        });
        Test("header mutation during snapshot aborts real adapter", () => {
            int headers = 0;
            var handler = new Handler { Reply = r => {
                string uri = Uri.UnescapeDataString(r.RequestUri.ToString());
                if (uri.Contains("DB!A2:2")) { headers++; return Response(headers == 1 ? "{\"values\":[[\"Series Name\",\"TVDB ID\",\"Missing Episodes\"]]}" : "{\"values\":[[\"Series Name\",\"Missing Episodes\",\"TVDB ID\"]]}"); }
                return Response("{\"values\":[]}");
            }};
            using (var service = Service(handler)) Throws<MissingEpisodeSheetException>(() => new GoogleMissingEpisodeSheet(service, "offline", true).Read());
            Check(handler.Posts == 0);
        });
        Console.WriteLine(passed + " Sheets adapter tests passed.");
    }
}
