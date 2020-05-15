using Newtonsoft.Json;
using RestSharp;

namespace TmdbApiCall
{
    class TmdbApi
    {
        private const string TMDB_API_KEY = "5809fe4e5d491f9514343fba6087cc34";
        private const string TMDB_LIST_ID = "122047";
        private const string TMDB_SESSION_ID = "6d26160352e952a088ccba1004addbb7c12d4ea9";

        /// <summary>
        /// Get the primary information about a movie.
        /// </summary>
        /// <param name="ImdbId"></param>
        /// <returns></returns>
        public static dynamic MoviesGetDetails(string ImdbId)
        {
            string strRestClient = "https://api.themoviedb.org/3/movie/" + ImdbId + "?api_key=" + TMDB_API_KEY + "&language=en-US";

            RestClient client = new RestClient(strRestClient);
            RestRequest request = new RestRequest(Method.GET);
            request.AddParameter("undefined", "{}", ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            dynamic json = JsonConvert.DeserializeObject(response.Content);
            return json;
        }

        /// <summary>
        /// You can use this method to check if a movie has already been added to the list.
        /// NOT CURRENTLY WORKING DUE TO DOUBLE OAUTH.
        /// </summary>
        /// <param name="TmdbId"></param>
        /// <returns></returns>
        public static dynamic ListsCheckItemStatus(string TmdbId)
        {
            string strRestClient = "https://api.themoviedb.org/3/list/" + TMDB_LIST_ID + "/item_status?api_key=" + TMDB_API_KEY + "&movie_id=" + TmdbId;

            RestClient client = new RestClient(strRestClient);
            RestRequest request = new RestRequest(Method.GET);
            request.AddParameter("undefined", "{}", ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            dynamic json = JsonConvert.DeserializeObject(response.Content);
            return json;
        }

        /// <summary>
        /// Remove a movie from a list.
        /// NOT CURRENTLY WORKING DUE TO DOUBLE OAUTH.
        /// </summary>
        /// <param name="TmdbId"></param>
        /// <returns></returns>
        public static dynamic ListsRemoveMovie(string TmdbId)
        {
            string strRestClient = "https://api.themoviedb.org/3/list/" + TMDB_LIST_ID + "/remove_item?api_key=" + TMDB_API_KEY + "&session_id=" + TMDB_SESSION_ID;

            RestClient client = new RestClient(strRestClient);
            RestRequest request = new RestRequest(Method.GET);
            request.AddParameter("media_id", TmdbId, ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            dynamic json = JsonConvert.DeserializeObject(response.Content);
            return json;
        }

        /// <summary>
        /// Create a temporary request token that can be used to validate a TMDb user login.
        /// </summary>
        /// <returns></returns>
        public static dynamic AuthenticationCreateRequestToken()
        {
            string strRestClient = "https://api.themoviedb.org/3/authentication/token/new?api_key=" + TMDB_API_KEY;

            RestClient client = new RestClient(strRestClient);
            RestRequest request = new RestRequest(Method.GET);
            request.AddParameter("undefined", "{}", ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            dynamic json = JsonConvert.DeserializeObject(response.Content);
            return json;
        }

        /// <summary>
        /// Create a temporary request token that can be used to validate a TMDb user login.
        /// </summary>
        /// <returns></returns>
        public static dynamic AuthenticationSendRequestToken(string requestToken)
        {
            string strRestClient = "https://www.themoviedb.org/authenticate/" + requestToken;

            RestClient client = new RestClient(strRestClient);
            RestRequest request = new RestRequest(Method.GET);
            request.AddParameter("undefined", "{}", ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            dynamic json = JsonConvert.DeserializeObject(response.Content);
            return json;
        }
    }
}
