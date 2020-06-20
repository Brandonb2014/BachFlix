using Newtonsoft.Json;
using RestSharp;

namespace TvdbApiCall
{
    class TvdbApi
    {
        private const string TVDB_API_KEY = "be8960724f2067d5c1f69b1772b30e74";
        private const string TVDB_JWT_KEY = "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJleHAiOjE1OTE5ODIzMzQsImlkIjoiQmFjaEZsaXhORk8iLCJvcmlnX2lhdCI6MTU5MTM3NzUzNH0.e_az82qbR42hctbfy-FdiA_BWM3WlV_YEIVnpUYlK-sGQQl4GWqOs3PH2LFx19TYdlqLV_QwuIhGPV_6hGKTdHObg9WllXHrJLrNPi65sqdYOOdyzs8h0lu30WKF-v-ia-63oHFsB67pCvDVI7dpmg42MsmTfRNhduJL2Ebq5NV2QH7SvusA8JZjWEkDz21Ml-xCVrMk95ueJMV6LUZFw2CFeVfj8XLO8lBYtoAiuY8PYIYG6SY4irgV2UW7PT3Yinl44t7DXV4FMuYBfC2xXo9FugDMEwS6iG-46IHYUTqa1St2iaWcBBhCjFtpmWV7wR_ehYmWPOxmkxjyVIh21Q";

        /// <summary>
        /// Get the primary information about a movie.
        /// </summary>
        /// <param name="ImdbId"></param>
        /// <returns></returns>
        public static dynamic MoviesGetDetails(string ImdbId)
        {
            string strRestClient = "https://api.themoviedb.org/3/movie/" + ImdbId + "?api_key=" + TVDB_API_KEY + "&language=en-US";

            RestClient client = new RestClient(strRestClient);
            RestRequest request = new RestRequest(Method.GET);
            request.AddParameter("undefined", "{}", ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            dynamic json = JsonConvert.DeserializeObject(response.Content);
            return json;
        }

        //public static dynamic TvEpisodesGetDetails(string TmdbId, string seasonNum, string episodeNum)
        //{
        //    Program.Type("Calling TMDB API with the following data--", 0, 0, 1);
        //    Program.Type("TmdbId: " + TmdbId, 0, 0, 1);
        //    Program.Type("seasonNum: " + seasonNum, 0, 0, 1);
        //    Program.Type("episodeNum: " + episodeNum, 0, 0, 1);
        //    try
        //    {
        //        //string strRestClient = "https://api.themoviedb.org/3/tv/" + TmdbId + "/season/" + seasonNum + "/episode/" + episodeNum + "?api_key=" + TMDB_API_KEY + "&language=en-US";

        //        RestClient client = new RestClient(strRestClient);
        //        RestRequest request = new RestRequest(Method.GET);
        //        request.AddParameter("undefined", "{}", ParameterType.RequestBody);
        //        IRestResponse response = client.Execute(request);
        //        dynamic json = JsonConvert.DeserializeObject(response.Content);
        //        return json;
        //    }
        //    catch (System.Exception e)
        //    {
        //        Program.Type("An error occured!", 0, 0, 1, "Red");
        //        Program.Type(e.Message, 0, 0, 1, "DarkRed");
        //        throw;
        //    }
        //}
    }
}
