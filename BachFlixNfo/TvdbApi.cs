using Newtonsoft.Json;
using RestSharp;
using System;
using SheetsQuickstart;
using System.Collections.Generic;
using System.Threading;

namespace TvdbApiCall
{
    class TvdbApi
    {
        private const string TVDB_API_KEY = "be8960724f2067d5c1f69b1772b30e74";
        private const string TVDB_USERNAME = "qtip888";
        private const string TVDB_USER_KEY = "98XBFB05IQ4CYPE4";

        /// <summary>
        /// Log into TVDB to get a fresh JWT Key.     CURRENTLY NOT WORKING.
        /// </summary>
        /// <returns></returns>
        public static bool LogIntoTvdbAsync(ref string token)
        {
            try
            {
                string strRestClient = "https://api.thetvdb.com/login";
                RestClient client = new RestClient(strRestClient);
                RestRequest request = new RestRequest(Method.POST);
                request.AddParameter("Accept", "application/json", ParameterType.HttpHeader);
                request.AddParameter("Content-Type", "application/json", ParameterType.HttpHeader);
                request.AddParameter("Authorization", "Bearer " + token, ParameterType.HttpHeader);
                request.AddParameter("undefined", "{\n  \"apikey\": \"" + TVDB_API_KEY + "\",\n  \"userkey\": \"" + TVDB_USER_KEY + "\",\n  \"username\": \"" + TVDB_USERNAME + "\"\n}", ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);

                if (response.IsSuccessful)
                {
                    dynamic json = JsonConvert.DeserializeObject(response.Content);
                    token = json.token.ToString();
                    SetTvdbJwtKey(token);
                    return true;
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    Program.DisplayMessage("error", "Error logging into TVDB: " + response.Content);
                    return false;
                }
                else
                {
                    Program.DisplayMessage("error", "Error logging into TVDB: " + response.Content);
                    return false;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }

        } // End LogIntoTvdb()

        public static bool RefreshToken(ref string token)
        {
            try
            {
                string strRestClient = "https://api.thetvdb.com/refresh_token";
                RestClient client = new RestClient(strRestClient);
                RestRequest request = new RestRequest(Method.GET);
                request.AddParameter("Accept", "application/json", ParameterType.HttpHeader);
                request.AddParameter("Content-Type", "application/json", ParameterType.HttpHeader);
                request.AddParameter("Authorization", "Bearer " + token, ParameterType.HttpHeader);
                IRestResponse response = client.Execute(request);
                
                if (response.IsSuccessful)
                {
                    token = response.Content;
                    SetTvdbJwtKey(token);
                    return true;
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    return false;
                }
                else
                {
                    Console.WriteLine("Failure refreshing token: " + response.Content);
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
        } // End RefreshToken()

        /// <summary>
        /// Get the Series ID from the TVDB Slug.
        /// </summary>
        /// <param name="TvdbSlug">The slug of the series from the TVDB URL.</param>
        /// <returns></returns>
        public static dynamic GetSeriesIdAsync(ref string token, string TvdbSlug)
        {
            try
            {
                string strRestClient = "https://api.thetvdb.com/search/series?slug=" + TvdbSlug;
                RestClient client = new RestClient(strRestClient);
                RestRequest request = new RestRequest(Method.GET);
                request.AddParameter("Accept", "application/json", ParameterType.HttpHeader);
                request.AddParameter("Authorization", "Bearer " + token, ParameterType.HttpHeader);
                IRestResponse response = client.Execute(request);

                if (response.IsSuccessful)
                {
                    dynamic json = JsonConvert.DeserializeObject(response.Content);
                    return json.data[0].id.ToString();
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    if (RefreshToken(ref token)) GetSeriesIdAsync(ref token, TvdbSlug);
                    else
                    {
                        LogIntoTvdbAsync(ref token);
                        GetSeriesIdAsync(ref token, TvdbSlug);
                    }
                    return response;
                }
                else
                {
                    return response;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
        } // End GetSeriesIdAsync()

        /// <summary>
        /// Get the Series details from the TVDB ID.
        /// </summary>
        /// <param name="TvdbId">The TVDBs ID of the show.</param>
        /// <returns></returns>
        public static dynamic GetSeriesDetailsAsync(ref string token, string TvdbId)
        {
            try
            {
                string strRestClient = "https://api.thetvdb.com/series/" + TvdbId;
                RestClient client = new RestClient(strRestClient);
                RestRequest request = new RestRequest(Method.GET);
                request.AddParameter("Accept", "application/json", ParameterType.HttpHeader);
                request.AddParameter("Authorization", "Bearer " + token, ParameterType.HttpHeader);
                IRestResponse response = client.Execute(request);

                if (response.IsSuccessful)
                {
                    dynamic json = JsonConvert.DeserializeObject(response.Content);
                    return json;
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    if (RefreshToken(ref token)) GetSeriesDetailsAsync(ref token, TvdbId);
                    else
                    {
                        LogIntoTvdbAsync(ref token);
                        GetSeriesDetailsAsync(ref token, TvdbId);
                    }
                    return response;
                }
                else
                {
                    return response;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
        } // End GetSeriesDetailsAsync()

        public static dynamic GetTvEpisodeDetails(ref string token, string TvdbID, string seasonNum, string episodeNum)
        {
            try
            {
                // Thread.Sleep(600);
                string strRestClient = "https://api.thetvdb.com/series/" + TvdbID + "/episodes/query?airedSeason=" + seasonNum + "&airedEpisode=" + episodeNum;
                RestClient client = new RestClient(strRestClient);
                RestRequest request = new RestRequest(Method.GET);
                request.AddParameter("Accept", "application/json", ParameterType.HttpHeader);
                request.AddParameter("Authorization", "Bearer " + token, ParameterType.HttpHeader);
                IRestResponse response = client.Execute(request);

                if (response.IsSuccessful)
                {
                    dynamic json = JsonConvert.DeserializeObject(response.Content);
                    return json;
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    if (RefreshToken(ref token)) GetTvEpisodeDetails(ref token, TvdbID, seasonNum, episodeNum);
                    else
                    {
                        LogIntoTvdbAsync(ref token);
                        GetTvEpisodeDetails(ref token, TvdbID, seasonNum, episodeNum);
                    }
                    return response;
                }
                else
                {
                    return response;
                }
            }
            catch (System.Exception e)
            {
                Program.Type("An error occured!", 0, 0, 1, "Red");
                Program.Type(e.Message, 0, 0, 1, "DarkRed");
                throw;
            }
        } // End TvEpisodesGetDetails()

        public static string GetTvdbJwtKey()
        {
            Program.DisplayMessage("warning", "Grabbing JWT Token... ", 0);
            IList<IList<Object>> varsData = Program.GetData("VARS!A3");
            Program.DisplayMessage("success", "DONE");

            return varsData[0][0].ToString();
        }
        public static void SetTvdbJwtKey(string token)
        {
            Program.WriteSingleCellToSheet(token, "VARS!A3");
            Program.DisplayMessage("log", "JWT Token saved to sheet.");
        }
    }
}
