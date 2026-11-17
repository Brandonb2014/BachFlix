using Newtonsoft.Json.Linq;
using RestSharp;
using System;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Reads IMDb's content rating and plot through OMDb (omdbapi.com), since IMDb blocks direct automated requests.
    /// Needs OMDB_API_KEY in BachFlixNfo.local.env; without it, lookups return null and the caller keeps going.
    /// </summary>
    internal static class OmdbApi
    {
        private const string OMDB_API_KEY_ENV = "OMDB_API_KEY";

        internal sealed class ImdbDetails
        {
            public string ContentRating { get; set; }
            public string Plot { get; set; }
            /// <summary>"movie", "series", "episode", ... or null.</summary>
            public string Type { get; set; }
            /// <summary>Comma-separated genres, e.g. "Animation, Short", or null.</summary>
            public string Genre { get; set; }
            public int? RuntimeMinutes { get; set; }
        }

        internal static bool IsConfigured => !string.IsNullOrWhiteSpace(LocalEnvironment.Get(OMDB_API_KEY_ENV));

        /// <summary>
        /// Returns IMDb's content rating and plot for the given tt ID, or null when there's no key or the lookup fails.
        /// </summary>
        internal static ImdbDetails GetDetails(string imdbId)
        {
            string apiKey = LocalEnvironment.Get(OMDB_API_KEY_ENV);
            if (string.IsNullOrWhiteSpace(apiKey) || string.IsNullOrWhiteSpace(imdbId))
                return null;

            try
            {
                var client = new RestClient("https://www.omdbapi.com/");
                var request = new RestRequest(Method.GET);
                request.AddQueryParameter("i", imdbId);
                request.AddQueryParameter("apikey", apiKey);
                IRestResponse response = client.Execute(request);
                if (!response.IsSuccessful || string.IsNullOrWhiteSpace(response.Content))
                    return null;

                return Parse(response.Content);
            }
            catch (Exception)
            {
                return null;
            }
        }

        internal static ImdbDetails Parse(string json)
        {
            var result = JObject.Parse(json);
            if (!string.Equals((string)result["Response"], "True", StringComparison.OrdinalIgnoreCase))
                return null;

            return new ImdbDetails
            {
                ContentRating = Clean((string)result["Rated"]),
                Plot = Clean((string)result["Plot"]),
                Type = Clean((string)result["Type"]),
                Genre = Clean((string)result["Genre"]),
                RuntimeMinutes = MovieCreditFilter.ParseRuntimeMinutes(Clean((string)result["Runtime"]))
            };
        }

        // OMDb uses "N/A" for missing values.
        private static string Clean(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            value = value.Trim();
            return value.Equals("N/A", StringComparison.OrdinalIgnoreCase) ? null : value;
        }
    }
}
