﻿using Newtonsoft.Json;
using RestSharp;
using SheetsQuickstart;
using System.Threading;

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
            Thread.Sleep(250);
            string strRestClient = "https://api.themoviedb.org/3/find/" + ImdbId + "?api_key=" + TMDB_API_KEY + "&language=en-US&external_source=imdb_id&append_to_response=videos";

            RestClient client = new RestClient(strRestClient);
            RestRequest request = new RestRequest(Method.GET);
            IRestResponse response = client.Execute(request);
            if (response.IsSuccessful)
            {
                dynamic json = JsonConvert.DeserializeObject(response.Content);
                return json;
            }
            else
            {
                return "";
            }
        }

        public static dynamic MoviesGetDetailsByTmdbId(string TmdbId)
        {
            try
            {
                string strRestClient = "https://api.themoviedb.org/3/movie/" + TmdbId + "?api_key=" + TMDB_API_KEY + "&language=en-US";
                RestClient client = new RestClient(strRestClient);
                RestRequest request = new RestRequest(Method.GET);
                IRestResponse response = client.Execute(request);
                if (response.IsSuccessful)
                {
                    dynamic json = JsonConvert.DeserializeObject(response.Content);
                    return json;
                }
                else
                {
                    Program.DisplayMessage("error", response.StatusCode.ToString());
                    return "";
                }
            }
            catch (System.Exception e)
            {
                Program.Type("An error occured!", 0, 0, 1, "Red");
                Program.Type(e.Message, 0, 0, 1, "DarkRed");
                throw;
            }
        }

        /// <summary>
        /// Get the credits information about a movie.
        /// </summary>
        /// <param name="ImdbId"></param>
        /// <returns></returns>
        public static dynamic MoviesGetCredits(string ImdbId)
        {
            Thread.Sleep(250);
            string strRestClient = "https://api.themoviedb.org/3/movie/" + ImdbId + "/credits?api_key=" + TMDB_API_KEY + "&language=en-US";

            RestClient client = new RestClient(strRestClient);
            RestRequest request = new RestRequest(Method.GET);
            IRestResponse response = client.Execute(request);
            if (response.IsSuccessful)
            {
                dynamic json = JsonConvert.DeserializeObject(response.Content);
                return json;
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// Get the movie watch providers (Streaming services).
        /// </summary>
        /// <param name="TmdbId"></param>
        /// <returns></returns>
        public static dynamic MoviesGetWatchProviders(string TmdbId)
        {
            Thread.Sleep(250);
            string strRestClient = "https://api.themoviedb.org/3/movie/" + TmdbId + "/watch/providers?api_key=" + TMDB_API_KEY + "&locale=US";

            RestClient client = new RestClient(strRestClient);
            RestRequest request = new RestRequest(Method.GET);
            IRestResponse response = client.Execute(request);
            if (response.IsSuccessful)
            {
                dynamic json = JsonConvert.DeserializeObject(response.Content);
                return json;
            }
            else
            {
                return response.Content;
            }
        }

        /// <summary>
        /// Get the videos for the selected movie.
        /// </summary>
        /// <param name="ImdbId"></param>
        /// <returns></returns>
        public static dynamic MoviesGetVideos(string ImdbId)
        {
            Thread.Sleep(250);
            string strRestClient = "https://api.themoviedb.org/3/movie/" + ImdbId + "/videos?api_key=" + TMDB_API_KEY + "&language=en-US";

            RestClient client = new RestClient(strRestClient);
            RestRequest request = new RestRequest(Method.GET);
            IRestResponse response = client.Execute(request);
            if (response.IsSuccessful)
            {
                dynamic json = JsonConvert.DeserializeObject(response.Content);
                return json;
            }
            else
            {
                return "";
            }
        }

        public static dynamic TvEpisodesGetDetails(string TmdbId, string seasonNum, string episodeNum)
        {
            Program.Type("Calling TMDB API with the following data--", 0, 0, 1);
            Program.Type("TmdbId: " + TmdbId, 0, 0, 1);
            Program.Type("seasonNum: " + seasonNum, 0, 0, 1);
            Program.Type("episodeNum: " + episodeNum, 0, 0, 1);
            try
            {
                string strRestClient = "https://api.themoviedb.org/3/tv/" + TmdbId + "/season/" + seasonNum + "/episode/" + episodeNum + "?api_key=" + TMDB_API_KEY + "&language=en-US";
                RestClient client = new RestClient(strRestClient);
                RestRequest request = new RestRequest(Method.GET);
                request.AddParameter("undefined", "{}", ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                dynamic json = JsonConvert.DeserializeObject(response.Content);
                return json;
            }
            catch (System.Exception e)
            {
                Program.Type("An error occured!", 0, 0, 1, "Red");
                Program.Type(e.Message, 0, 0, 1, "DarkRed");
                throw;
            }
        }

        public static dynamic ActorsGetMovieCredits(string PersonId)
        {
            try
            {
                string strRestClient = "https://api.themoviedb.org/3/person/" + PersonId + "/movie_credits?api_key=" + TMDB_API_KEY + "&language = en-US";
                RestClient client = new RestClient(strRestClient);
                RestRequest request = new RestRequest(Method.GET);
                IRestResponse response = client.Execute(request);
                if (response.IsSuccessful)
                {
                    dynamic json = JsonConvert.DeserializeObject(response.Content);
                    return json;
                }
                else
                {
                    Program.DisplayMessage("error", response.StatusCode.ToString());
                    return "";
                }
            }
            catch (System.Exception e)
            {
                Program.Type("An error occured!", 0, 0, 1, "Red");
                Program.Type(e.Message, 0, 0, 1, "DarkRed");
                throw;
            }
        }




        //// Nothing below here works yet.... and may never work.

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
