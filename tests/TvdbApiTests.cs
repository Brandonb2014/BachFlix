using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using TvdbApiCall;

class TvdbApiTests
{
    static int passed;
    const string Series = "{\"status\":\"success\",\"data\":{\"id\":42,\"name\":\"Example\",\"firstAired\":\"2020-01-01\",\"status\":{\"name\":\"Continuing\"},\"seasons\":[{\"number\":0,\"type\":{\"type\":\"official\"}},{\"number\":1,\"type\":{\"type\":\"official\"}},{\"number\":1,\"type\":{\"type\":\"official\"}},{\"number\":2,\"type\":{\"type\":\"official\"}},{\"number\":3,\"type\":{\"type\":\"dvd\"}}],\"contentRatings\":[{\"country\":\"gbr\",\"name\":\"12\"},{\"country\":\"usa\",\"name\":\"TV-PG\"}]}}";
    const string Episode = "{\"status\":\"success\",\"data\":{\"episodes\":[{\"id\":9,\"seasonNumber\":1,\"number\":2,\"name\":\"Second\",\"overview\":\"Plot\"}]}}";
    static IRestResponse Response(string body, int status = 200) { return new RestResponse { Content = body, StatusCode = (HttpStatusCode)status, ResponseStatus = ResponseStatus.Completed }; }
    static IRestResponse Login(string token = "test-token") { return Response("{\"status\":\"success\",\"data\":{\"token\":\"" + token + "\"}}"); }
    static void Check(bool value, string message) { if (!value) throw new Exception(message); }
    static void Test(string name, Action action) { action(); passed++; Console.WriteLine("PASS " + name); }
    static void Fails(Action action, TvdbFailure expected, string context = null)
    {
        try { action(); } catch (TvdbException ex) {
            Check(ex.Failure == expected, "Wrong failure category: " + ex.Failure);
            if (context != null) Check(ex.Message.Contains(context), "Missing error context");
            Check(!ex.ToString().Contains("private-secret"), "Secret leaked"); return;
        }
        throw new Exception("Expected failure");
    }
    static TvdbClient Client(List<RestRequest> requests, params IRestResponse[] responses)
    {
        var queue = new Queue<IRestResponse>(responses);
        IRestResponse last = null;
        return new TvdbClient(() => "fake-key", r => {
            requests.Add(r);
            if (queue.Count > 0) last = queue.Dequeue();
            else Check(last != null && (last.ResponseStatus != ResponseStatus.Completed || BachFlixNfo.ApiRetry.Transient((int)last.StatusCode)), "Unexpected extra request");
            return last;
        }, _ => { });
    }
    static void Main()
    {
        Test("v4 login body, data.token, and no Authorization", () => {
            var requests = new List<RestRequest>(); var c = Client(requests, Login());
            Check(c.Login() == "test-token", "Token");
            var r = requests.Single(); Check(r.Method == Method.POST && r.Resource == "login", "Login endpoint");
            Check(!r.Parameters.Any(p => p.Name == "Authorization"), "Login auth header");
            var body = JObject.Parse(JsonConvert.SerializeObject(r.Parameters.Single(p => p.Type == ParameterType.RequestBody).Value));
            Check(body.Properties().Count() == 1 && (string)body["apikey"] == "fake-key", "Login must contain only apikey");
        });
        Test("series fields, official distinct seasons excluding specials, USA rating, bearer and token reuse", () => {
            var requests = new List<RestRequest>(); var c = Client(requests, Login(), Response(Series), Response(Series));
            var d = c.GetSeriesDetails("42")["data"]; c.GetSeriesDetails("42");
            Check((string)d["seriesName"] == "Example" && (int)d["season"] == 2 && (string)d["rating"] == "TV-PG" && (string)d["status"] == "Continuing" && (string)d["firstAired"] == "2020-01-01", "Mapping");
            Check(requests.Count == 3 && requests[1].Resource == "series/42/extended", "Cache and endpoint");
            Check((string)requests[1].Parameters.Single(p => p.Name == "Authorization").Value == "Bearer test-token", "Bearer");
        });
        Test("episode fields and official season/episode filters", () => {
            var rs = new List<RestRequest>(); var c = Client(rs, Login(), Response(Episode));
            var d = c.GetEpisodeDetails("42", "01", "02")["data"][0];
            Check((string)d["episodeName"] == "Second" && (string)d["overview"] == "Plot", "Episode mapping");
            Check(new RestClient(TvdbClient.BaseUrl).BuildUri(rs[1]).AbsoluteUri == TvdbClient.BaseUrl + "/series/42/episodes/official?page=0&season=1&episodeNumber=2", "Episode endpoint");
        });
        Test("401 reauth once and return retry result", () => {
            var rs = new List<RestRequest>(); var c = Client(rs, Login(), Response("",401), Login("new-token"), Response(Episode));
            Check((string)c.GetEpisodeDetails("42","1","2")["data"][0]["episodeName"] == "Second", "Retry result");
            Check(rs.Count == 4 && (string)rs[3].Parameters.Single(p=>p.Name=="Authorization").Value == "Bearer new-token", "Bounded retry");
            Check(!rs[2].Parameters.Any(p=>p.Name=="Authorization"), "Reauth header");
        });
        Test("persistent 401 stops and clears token", () => {
            var rs = new List<RestRequest>(); var c = Client(rs, Login(), Response("",401), Login(), Response("private-secret",401));
            Fails(()=>c.GetSeriesDetails("42"), TvdbFailure.Authentication, "HTTP 401"); Check(rs.Count==4 && c.Token==null,"Retry count");
        });
        Test("failed reauthentication stops", () => {
            var rs = new List<RestRequest>(); var c = Client(rs, Login(), Response("",401), Response("private-secret",401));
            Fails(()=>c.GetSeriesDetails("42"),TvdbFailure.Authentication); Check(rs.Count==3,"Requests");
        });
        Test("login missing nested token",()=>Fails(()=>Client(new List<RestRequest>(),Response("{\"status\":\"success\",\"token\":\"private-secret\"}")).Login(),TvdbFailure.Authentication));
        Test("login rejects invalid credentials",()=>Fails(()=>Client(new List<RestRequest>(),Response("private-secret",401)).Login(),TvdbFailure.Authentication,"HTTP 401"));
        Test("slug endpoint used first",()=>{
            var rs=new List<RestRequest>(); var c=Client(rs,Login(),Response("{\"status\":\"success\",\"data\":{\"id\":477319,\"slug\":\"my-show\"}}"));
            Check(c.GetSeriesId("my-show")=="477319","Slug match"); Check(rs.Count==2 && rs[1].Resource=="series/slug/my-show","Slug endpoint");
        });
        Test("search ignores unrelated first result",()=>{
            var rs=new List<RestRequest>(); var c=Client(rs,Login(),Response("",404),Response("{\"status\":\"success\",\"data\":[{\"slug\":\"unrelated\",\"tvdb_id\":\"1\"},{\"slug\":\"my-show\",\"tvdb_id\":\"42\"}]}"));
            Check(c.GetSeriesId("my-show")=="42","Search match"); Check(new RestClient(TvdbClient.BaseUrl).BuildUri(rs[2]).AbsoluteUri==TvdbClient.BaseUrl + "/search?type=series&query=my%20show","Search endpoint");
        });
        Test("search not found",()=>Fails(()=>Client(new List<RestRequest>(),Login(),Response("",404),Response("{\"status\":\"success\",\"data\":[{\"slug\":\"other\",\"tvdb_id\":\"1\"}]}")).GetSeriesId("my-show"),TvdbFailure.SeriesNotFound));
        Test("search ambiguous",()=>Fails(()=>Client(new List<RestRequest>(),Login(),Response("",404),Response("{\"status\":\"success\",\"data\":[{\"slug\":\"same\",\"tvdb_id\":\"1\"},{\"slug\":\"same\",\"tvdb_id\":\"2\"}]}")).GetSeriesId("same"),TvdbFailure.Api));
        Test("numeric ID skips search",()=>{var rs=new List<RestRequest>(); Check(Client(rs,Login(),Response(Series)).GetSeriesId("42")=="42","ID"); Check(rs.Count==2 && rs[1].Resource=="series/42/extended","Direct ID");});
        Test("series 404",()=>Fails(()=>Client(new List<RestRequest>(),Login(),Response("",404)).GetSeriesDetails("42"),TvdbFailure.SeriesNotFound,"HTTP 404"));
        Test("episode 404",()=>Fails(()=>Client(new List<RestRequest>(),Login(),Response("",404)).GetEpisodeDetails("42","1","2"),TvdbFailure.EpisodeNotFound));
        Test("empty episodes",()=>Fails(()=>Client(new List<RestRequest>(),Login(),Response("{\"status\":\"success\",\"data\":{\"episodes\":[]}}")).GetEpisodeDetails("42","1","2"),TvdbFailure.EpisodeNotFound));
        Test("wrong episode is not accepted",()=>Fails(()=>Client(new List<RestRequest>(),Login(),Response(Episode)).GetEpisodeDetails("42","2","2"),TvdbFailure.EpisodeNotFound));
        Test("null episode text is safe for callers",()=>{var c=Client(new List<RestRequest>(),Login(),Response(Episode.Replace("\"Second\"","null").Replace("\"Plot\"","null"))); Check((string)c.GetEpisodeDetails("42","1","2")["data"][0]["overview"]=="","Null text");});
        Test("500 reports status without response secrets",()=>Fails(()=>Client(new List<RestRequest>(),Login(),Response("private-secret",500)).GetSeriesDetails("42"),TvdbFailure.Api,"HTTP 500"));
        Test("network timeout",()=>Fails(()=>Client(new List<RestRequest>(),new RestResponse{ResponseStatus=ResponseStatus.TimedOut,ErrorMessage="private-secret"}).Login(),TvdbFailure.Api,"TimedOut"));
        Test("transport exception redacted",()=>Fails(()=>new TvdbClient(()=>"fake-key",r=>{throw new Exception("private-secret");}, _ => {}).Login(),TvdbFailure.Api));
        Test("invalid JSON redacted",()=>Fails(()=>Client(new List<RestRequest>(),Response("private-secret")).Login(),TvdbFailure.Api,"invalid JSON"));
        Test("API status failure",()=>Fails(()=>Client(new List<RestRequest>(),Response("{\"status\":\"failure\",\"message\":\"private-secret\"}")).Login(),TvdbFailure.Api));
        Test("malformed series data",()=>Fails(()=>Client(new List<RestRequest>(),Login(),Response("{\"status\":\"success\",\"data\":null}")).GetSeriesDetails("42"),TvdbFailure.Api));
        Test("malformed episode data",()=>Fails(()=>Client(new List<RestRequest>(),Login(),Response("{\"status\":\"success\",\"data\":{}}")).GetEpisodeDetails("42","1","2"),TvdbFailure.Api));
        Console.WriteLine(passed + " tests passed.");
    }
}
