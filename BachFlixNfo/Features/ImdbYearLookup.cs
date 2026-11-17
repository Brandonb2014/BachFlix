using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Linq;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Looks up the year IMDb shows for a title, using IMDb's public suggestion endpoint.
    /// IMDb's year is the title's first release anywhere, which can differ from TMDB's release date.
    /// </summary>
    internal static class ImdbYearLookup
    {
        /// <summary>
        /// Returns IMDb's year for the given tt ID, or null when it can't be found.
        /// </summary>
        internal static string GetYear(string imdbId)
        {
            if (string.IsNullOrWhiteSpace(imdbId) || !imdbId.StartsWith("tt", StringComparison.OrdinalIgnoreCase))
                return null;

            try
            {
                var client = new RestClient("https://v3.sg.media-imdb.com/suggestion/x/" + imdbId + ".json");
                IRestResponse response = client.Execute(new RestRequest(Method.GET));
                if (!response.IsSuccessful || string.IsNullOrWhiteSpace(response.Content))
                    return null;

                return ParseYear(response.Content, imdbId);
            }
            catch (Exception)
            {
                return null;
            }
        }

        internal static string ParseYear(string json, string imdbId)
        {
            var results = JObject.Parse(json)["d"] as JArray;
            var match = results?.FirstOrDefault(r => string.Equals((string)r["id"], imdbId, StringComparison.OrdinalIgnoreCase));
            int? year = (int?)match?["y"];
            return year.HasValue && year.Value > 1800 ? year.Value.ToString() : null;
        }
    }
}
