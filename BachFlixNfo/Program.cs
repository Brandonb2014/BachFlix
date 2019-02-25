using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using System.Diagnostics;
using RestSharp;
using Newtonsoft.Json;

namespace SheetsQuickstart
{
    class Program
    {
        // Data ranges for each sheet.
        private const string MOVIES_TITLE_RANGE = "Movies!A2:2";
        private const string MOVIES_DATA_RANGE = "Movies!A3:2001";
        private const string TEMP_MOVIES_TITLE_RANGE = "Temp!A2:2";
        private const string TEMP_MOVIES_DATA_RANGE = "Temp!A3:1001";
        private const string YOUTUBE_TITLE_RANGE = "YouTube!A1:1";
        private const string YOUTUBE_DATA_RANGE = "YouTube!A2:914";
        private const string BONUS_TITLE_RANGE = "Bonus!A1:1";
        private const string BONUS_DATA_RANGE = "Bonus!A2:2036";
        private const string EPISODES_TITLE_RANGE = "Episodes!A1:1";
        private const string EPISODES_DATA_RANGE = "Episodes!A2:1000";

        // If modifying these scopes, delete your previously saved credentials
        // at \BachFlixNfo\bin\Debug\token.json\Google.Apis.Auth.OAuth2.Responses.TokenResponse-user
        static readonly string[] SCOPES = { SheetsService.Scope.Spreadsheets };
        static string APLICATION_NAME = "Google Sheets API .NET Quickstart";
        static readonly string SPREADSHEET_ID = "1LE9Tiz0TgcG60qeul_y9wC4j8qNLQlfKTLnAg5tgBr0";

        static void Main(string[] args)
        {
            Type("Hello, and welcome to the BachFlix NFO Filer 3000!", 0, 0, 1);
            Menu();

        } // End Main

        /// <summary>
        /// Gives the main menu on startup.
        /// </summary>
        static void Menu()
        {
            bool keepAskingForChoice = true;
            do
            {
                string[] choices;
                Type("Please choose from one of the following options..", 0, 0, 1);
                Type("(Or do multiple options by separating the options with a comma. i.e. 1,3)", 0, 0, 1);
                Type("0- Exit", 0, 0, 1);
                Type("--- NFO File Creation ---", 0, 0, 1);
                Type("1- All Movies", 0, 0, 1);
                Type("1q- Selected Movies", 0, 0, 1);
                Type("2- All Temp Movies", 0, 0, 1);
                Type("2q- Selected Temp Movies", 0, 0, 1);
                Type("3- All TV Shows (Coming Soon)", 0, 0, 1);
                Type("3q- Selected TV Shows (Coming Soon)", 0, 0, 1);
                Type("4- All YouTube Videos", 0, 0, 1);
                Type("4q- Selected YouTube Videos", 0, 0, 1);
                Type("-- Convert Files ---", 0, 0, 1);
                Type("5- Movies", 0, 0, 1);
                Type("6- Bonus Features", 0, 0, 1);
                Type("7- TV Shows", 0, 0, 1);
                Type("--- Misc. ---", 0, 0, 1);
                Type("8- Count Files", 0, 0, 1);
                Type("9- Remove the UPC numbers from the folder name.", 0, 0, 1);
                Type("10- Insert movie plots into the Google Sheet.", 0, 0, 1);

                try
                {
                    choices = Console.ReadLine().Split(',');
                    foreach(string choice in choices)
                    {
                        switch (choice.Trim())
                        {
                            case "0": // Exit.
                                Type("Thank you, have a nice day! \\(^.^)/", 7, 100, 1);
                                Type("Press ENTER to exit.", 7, 100, 1);
                                Console.ReadLine();
                                keepAskingForChoice = false;
                                break;
                            case "1": // NFO files for Movies - Default.
                                DefaultMovies();
                                break;
                            case "1q": // NFO files for Movies - Default - Quick.
                                DefaultMoviesQuick();
                                break;
                            case "2": // NFO files for Movies Temp - Default.
                                DefaultTempMovies();
                                break;
                            case "2q": // NFO files for Movies Temp - Default - Quick.
                                DefaultTempMoviesQuick();
                                break;
                            case "3": // NFO files for TV Shows - Default.
                                Type("Sorry, this is not ready. Try another option.", 0, 0, 1);
                                break;
                            case "3q": // NFO files for TV Shows - Default - Quick.
                                Type("Sorry, this is not ready. Try another option.", 0, 0, 1);
                                break;
                            case "4": // NFO files for YouTube - Default.
                                DefaultYoutube();
                                break;
                            case "4q": // NFO files for YouTube - Default - Quick.
                                DefaultYoutubeQuick();
                                break;
                            case "5": // Convert Movies.
                                Type("Convert movies. Let's go!", 7, 100, 1);
                                GetDataToConvertMovies();
                                break;
                            case "6": // Convert Bonus Features.
                                Type("Convert bonus features. Let's go!", 7, 100, 1);
                                GetDataToConvertBonusFeatures();
                                break;
                            case "7": // Convert TV Shows.
                                Type("Convert TV Shows. Let's go!", 7, 100, 1);
                                GetDataToConvertEpisodes();
                                break;
                            case "8": // Count files.
                                CountFiles();
                                break;
                            case "9": // Rename folder name.
                                Type("Rename folders.", 7, 100, 1);
                                GetFolders();
                                break;
                            case "10": // Grab movie plots.
                                Type("Grab the movie plots. Let's go!", 7, 100, 1);
                                InputMoviePlots();
                                break;
                            case "20": // TMDB
                                //RestClient client = new RestClient("https://api.themoviedb.org/3/movie/tt0147800?api_key=5809fe4e5d491f9514343fba6087cc34&language=en-US");
                                //RestRequest request = new RestRequest(Method.GET);
                                //request.AddParameter("undefined", "{}", ParameterType.RequestBody);
                                //IRestResponse response = client.Execute(request);
                                //dynamic json = Newtonsoft.Json.JsonConvert.DeserializeObject(response.Content);
                                //var outerJSON = json["movie_results"];
                                //var plot = outerJSON[0].overview;

                                //foreach (var item in json)
                                //{
                                //    foreach(var item2 in item)
                                //    {
                                //        Console.WriteLine(item2.id);
                                //    }
                                //}
                                break;
                            case "21": // 
                                Type("1", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("2", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("3", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("4", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("5", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("6", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("7", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("8", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("9", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("10", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("11", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("12", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("13", 100, 100, 0);
                                Console.SetCursorPosition(0, Console.CursorTop);
                                Type("14", 100, 100, 0);
                                break;
                            case "22": // 
                                RestClient client = new RestClient("https://api.themoviedb.org/3/movie/tt0147800?api_key=5809fe4e5d491f9514343fba6087cc34&language=en-US");
                                RestRequest request = new RestRequest(Method.GET);
                                request.AddParameter("undefined", "{}", ParameterType.RequestBody);
                                IRestResponse response = client.Execute(request);
                                dynamic json = JsonConvert.DeserializeObject(response.Content);
                                var plot = json.overview;
                                //var plot = outerJSON[0].overview;

                                UserCredential credential;

                                using (var stream =
                                    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                                {
                                    string credPath = "token.json";
                                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                                        GoogleClientSecrets.Load(stream).Secrets,
                                        SCOPES,
                                        "user",
                                        CancellationToken.None,
                                        new FileDataStore(credPath, true)).Result;
                                }

                                SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
                                {
                                    HttpClientInitializer = credential,
                                    ApplicationName = "Google-SheetsSample/0.1",
                                });

                                // The ID of the spreadsheet to update.
                                string spreadsheetId = "1LE9Tiz0TgcG60qeul_y9wC4j8qNLQlfKTLnAg5tgBr0";

                                // The A1 notation of the values to update.
                                string range = "Movies!L3";

                                // How the input data should be interpreted.
                                SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum valueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

                                // TODO: Assign values to desired properties of `requestBody`. All existing
                                // properties will be replaced:
                                ValueRange requestBody = new ValueRange();
                                requestBody.MajorDimension = "COLUMNS"; // "ROWS" / "COLUMNS"
                                var oblist = new List<object>() { plot };
                                requestBody.Values = new List<IList<object>> { oblist };

                                SpreadsheetsResource.ValuesResource.UpdateRequest request2 = sheetsService.Spreadsheets.Values.Update(requestBody, spreadsheetId, range);
                                request2.ValueInputOption = valueInputOption;

                                // To execute asynchronously in an async method, replace `request.Execute()` as shown:
                                UpdateValuesResponse response2 = request2.Execute();
                                // Data.UpdateValuesResponse response = await request.ExecuteAsync();

                                // TODO: Change code below to process the `response` object:
                                Console.WriteLine(JsonConvert.SerializeObject(response2));

                                break;
                            case "23": // 
                                Console.WriteLine("Give me a number");
                                string theirChoice = Console.ReadLine();
                                Console.WriteLine(ColumnNumToLetter(int.Parse(theirChoice)));
                                break;
                            default: // Other.
                                Type("I'm sorry I didn't quite understand " + choice + ".", 14, 100, 1);
                                Type("Please make sure your choice matches an option exactly from the menu.", 14, 100, 1);
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Type(ex.ToString(), 0, 0, 1);
                    Type("I'm sorry I didn't quite get that.", 14, 100, 1);
                    Type("Please make sure your choice matches an option exactly from the menu.", 14, 100, 1);
                }

            } while (keepAskingForChoice);
        } // End Menu()

        protected static void InputMoviePlots()
        {
            // Declare variables.
            int intImdbIdColumn = -1,
                intPlotColumn = -1,
                intImdbTitleColumn = -1,
                intRowNum = 3,
                intPlotsDoneCount = 0,
                intPlotsSkippedCount = 0;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, TEMP_MOVIES_TITLE_RANGE);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;
            if (titleValues != null && titleValues.Count > 0)
            {
                int x = 0;
                foreach (var row in titleValues)
                {
                    do
                    {
                        if (row[x].ToString() == "IMDB ID")
                        {
                            intImdbIdColumn = x;
                        }
                        else if (row[x].ToString() == "Plot")
                        {
                            intPlotColumn = x;
                        }
                        else if (row[x].ToString() == "IMDB Title")
                        {
                            intImdbTitleColumn = x;
                        }
                        x++;

                    } while (x < row.Count);

                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }

            if (intImdbIdColumn == -1 || intPlotColumn == -1 || intImdbTitleColumn == -1)
            {
                Type("It looks like a column name was changed and I can no longer find it.", 1, 100, 1);
                Type("IMDB ID: " + intImdbIdColumn + ". IMDB Title: " + intImdbTitleColumn + ". Plot: " + intPlotColumn + ".", 1, 100, 1);
            }
            else
            {

                SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
                        service.Spreadsheets.Values.Get(SPREADSHEET_ID, TEMP_MOVIES_DATA_RANGE);

                ValueRange dataRowResponse = dataRowRequest.Execute();
                IList<IList<Object>> dataValues = dataRowResponse.Values;
                if (dataValues != null)
                {
                    foreach (var row in dataValues)
                    {
                        if (row.Count > 20) // If it's an empty row then it should have much less than 20.
                        {
                            try
                            {
                                if (row[intPlotColumn].ToString() == "")
                                {
                                    string strRestClient = "https://api.themoviedb.org/3/movie/" + row[intImdbIdColumn].ToString() + "?api_key=5809fe4e5d491f9514343fba6087cc34&language=en-US";
                                    string strCellForPlotSummary = "Temp!" + ColumnNumToLetter(intPlotColumn) + intRowNum;

                                    RestClient client = new RestClient(strRestClient);
                                    RestRequest request = new RestRequest(Method.GET);
                                    request.AddParameter("undefined", "{}", ParameterType.RequestBody);
                                    IRestResponse response = client.Execute(request);
                                    dynamic json = JsonConvert.DeserializeObject(response.Content);
                                    var plot = json.overview;
                                    //var plot = outerJSON[0].overview;

                                    UserCredential credential2;

                                    using (var stream =
                                        new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                                    {
                                        string credPath = "token.json";
                                        credential2 = GoogleWebAuthorizationBroker.AuthorizeAsync(
                                            GoogleClientSecrets.Load(stream).Secrets,
                                            SCOPES,
                                            "user",
                                            CancellationToken.None,
                                            new FileDataStore(credPath, true)).Result;
                                    }

                                    SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
                                    {
                                        HttpClientInitializer = credential2,
                                        ApplicationName = "Google-SheetsSample/0.1",
                                    });

                                    // The ID of the spreadsheet to update.
                                    string spreadsheetId = "1LE9Tiz0TgcG60qeul_y9wC4j8qNLQlfKTLnAg5tgBr0";

                                    // The A1 notation of the values to update.
                                    string range = strCellForPlotSummary;

                                    // How the input data should be interpreted.
                                    SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum valueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

                                    // TODO: Assign values to desired properties of `requestBody`. All existing
                                    // properties will be replaced:
                                    ValueRange requestBody = new ValueRange();
                                    requestBody.MajorDimension = "COLUMNS"; // "ROWS" / "COLUMNS"
                                    var oblist = new List<object>() { plot };
                                    requestBody.Values = new List<IList<object>> { oblist };

                                    SpreadsheetsResource.ValuesResource.UpdateRequest request2 = sheetsService.Spreadsheets.Values.Update(requestBody, spreadsheetId, range);
                                    request2.ValueInputOption = valueInputOption;

                                    // To execute asynchronously in an async method, replace `request.Execute()` as shown:
                                    UpdateValuesResponse response2 = request2.Execute();
                                    // Data.UpdateValuesResponse response = await request.ExecuteAsync();

                                    Type("Plot saved for: " + row[intImdbTitleColumn].ToString(), 0, 0, 1);
                                    intPlotsDoneCount++;
                                }
                                else
                                {
                                    Type("Plot not saved for: " + row[intImdbTitleColumn].ToString(), 0, 0, 1);
                                    intPlotsSkippedCount++;
                                }
                                
                                
                            }
                            catch (Exception e)
                            {
                                Type("Something went wrong..." + e.Message, 3, 100, 1);
                                break;
                            }

                        }
                        intRowNum++;
                    }

                }
                else
                {
                    Console.WriteLine("No data found.");
                }
                Type("It looks like that's the end of it.", 3, 100, 2);
                Type("Plots done: " + intPlotsDoneCount, 3, 100, 2);
                Type("Plots skipped: " + intPlotsSkippedCount, 3, 100, 2);

            }

        } // End InputMoviePlots()

        /// <summary>
        /// Convert the number to the column letter.
        /// i.e. 0 = A
        /// </summary>
        /// <param name="columnNum">The number of the column.</param>
        /// <returns></returns>
        protected static string ColumnNumToLetter(int columnNum)
        {
            if(columnNum < 53)
            {
                string[] myString = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ" };
                
                return myString[columnNum];
            }
            else
            {
                return "";
            }
            
        }

        /// <summary>
        /// This is a test method for writing to the same console line.
        /// Not using it yet.
        /// </summary>
        protected static void ClearCurrentConsoleLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }

        protected static void GetDataToConvertMovies()
        {
            // Declare variables.
            int intInputFolder = 0,
                intOutputFolder = 0,
                intIsoTitleNumber = 0,
                intTotalMoviesCount = 0,
                intImagesCount = 0,
                intAlreadyConvertedFilesCount = 0,
                intNoTitleCount = 0,
                intConvertedFilesCount = 0;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, MOVIES_TITLE_RANGE);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;
            if (titleValues != null && titleValues.Count > 0)
            {
                int x = 0;
                foreach (var row in titleValues)
                {
                    do
                    {
                        if (row[x].ToString() == "ISO Input")
                        {
                            intInputFolder = x;
                        }
                        else if (row[x].ToString() == "MKV Output")
                        {
                            intOutputFolder = x;
                        }
                        else if (row[x].ToString() == "ISO Title #")
                        {
                            intIsoTitleNumber = x;
                        }
                        x++;

                    } while (x < row.Count);

                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }

            SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, MOVIES_DATA_RANGE);

            ValueRange dataRowResponse = dataRowRequest.Execute();
            IList<IList<Object>> dataValues = dataRowResponse.Values;
            if (dataValues != null)
            {
                foreach (var row in dataValues)
                {
                    if (row.Count > 20) // If it's an empty row then it should have much less than 20.
                    {
                        intTotalMoviesCount++;
                        try
                        {
                            string i = row[intInputFolder].ToString(),
                                    o = row[intOutputFolder].ToString(),
                                    title = row[intIsoTitleNumber].ToString();

                            if (File.Exists(i))
                            {
                                Type("We found " + i, 0, 0, 1);
                                intImagesCount++;
                                if (File.Exists(o))
                                {
                                    Type("We found " + o, 0, 0, 1);
                                    Type("We won't have to convert this one.", 0, 0, 1);
                                    intAlreadyConvertedFilesCount++;
                                }
                                else
                                {
                                    Type("We didn't find " + o, 0, 0, 1);

                                    if (title != "")
                                    {
                                        Type("We will use title #" + title, 0, 0, 1);
                                        string strMyConversionString = "HandBrakeCLI -i \"" + i + "\" -o \"" + o + "\" --preset-import-file preset.json -t " + title;
                                        HandBrake(strMyConversionString);
                                        intConvertedFilesCount++;
                                    }
                                    else
                                    {
                                        Type("We don't have a title to go off of.", 0, 0, 1);
                                        intNoTitleCount++;
                                    }
                                }
                            }
                            else
                            {
                                Type("We didn't find " + i, 0, 0, 1);
                                Type("We won't be able to convert this one at this time.", 0, 0, 1);
                            }
                            Type("-------------------------------------------------------------------", 0, 0, 1);

                        }
                        catch (Exception e)
                        {
                            Type("Something went wrong..." + e.Message, 3, 100, 1);
                            break;
                        }
                    }
                } // End foreach
                Type("-----SUMMARY-----", 7, 100, 1);
                Type(intTotalMoviesCount + " Total Movies.", 7, 100, 1);
                Type(intImagesCount + " Images Found.", 7, 100, 1);
                Type(intAlreadyConvertedFilesCount + " Movie Files Found.", 7, 100, 1);
                Type(intConvertedFilesCount + " Movies converted.", 7, 100, 1);
                Type(intNoTitleCount + " Missing Titles to convert.", 7, 100, 1);
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Type("It looks like that's the end of it.", 3, 100, 2);
        } // End GetDataToConvertMovies()

        protected static void GetDataToConvertBonusFeatures()
        {
            // Declare variables.
            int intInputFolder = 0,
                intOutputFolder = 0,
                intIsoTitleNumber = 0,
                intBonusFeatureTitle = 0,
                intTotalMoviesCount = 0,
                intImagesCount = 0,
                intAlreadyConvertedFilesCount = 0,
                intNoTitleCount = 0,
                intConvertedFilesCount = 0;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, BONUS_TITLE_RANGE);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;
            if (titleValues != null && titleValues.Count > 0)
            {
                int x = 0;
                foreach (var row in titleValues)
                {
                    do
                    {
                        if (row[x].ToString() == "ISO Input")
                        {
                            intInputFolder = x;
                        }
                        else if (row[x].ToString() == "MKV Output")
                        {
                            intOutputFolder = x;
                        }
                        else if (row[x].ToString() == "ISO Title #")
                        {
                            intIsoTitleNumber = x;
                        }
                        else if (row[x].ToString() == "Clean Bonus Feature Title")
                        {
                            intBonusFeatureTitle = x;
                        }
                        x++;

                    } while (x < row.Count);

                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }

            SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, BONUS_DATA_RANGE);

            ValueRange dataRowResponse = dataRowRequest.Execute();
            IList<IList<Object>> dataValues = dataRowResponse.Values;
            if (dataValues != null)
            {
                foreach (var row in dataValues)
                {
                    if (row.Count > 5) // If it's an empty row then it should have much less than 20.
                    {
                        intTotalMoviesCount++;
                        try
                        {
                            string i = row[intInputFolder].ToString(),
                                    directory = row[intOutputFolder].ToString(),
                                    o = row[intOutputFolder].ToString() + "\\" + row[intBonusFeatureTitle].ToString() + ".mkv",
                                    title = row[intIsoTitleNumber].ToString();

                            if (File.Exists(i))
                            {
                                Type("We found " + i, 0, 0, 1);
                                intImagesCount++;
                                if (File.Exists(o))
                                {
                                    Type("We found " + o, 0, 0, 1);
                                    Type("We won't have to convert this one.", 0, 0, 1);
                                    intAlreadyConvertedFilesCount++;
                                }
                                else
                                {
                                    Type("We didn't find " + o, 0, 0, 1);

                                    if (title != "")
                                    {
                                        Type("We will use title #" + title, 0, 0, 1);

                                        Directory.CreateDirectory(directory);

                                        string strMyConversionString = "HandBrakeCLI -i \"" + i + "\" -o \"" + o + "\" --preset-import-file preset.json -t " + title;
                                        HandBrake(strMyConversionString);
                                        intConvertedFilesCount++;
                                    }
                                    else
                                    {
                                        Type("We don't have a title to go off of.", 0, 0, 1);
                                        intNoTitleCount++;
                                    }
                                }
                            }
                            else
                            {
                                Type("We didn't find " + i, 0, 0, 1);
                                Type("We won't be able to convert this one at this time.", 0, 0, 1);
                            }
                            Type("-------------------------------------------------------------------", 0, 0, 1);

                        }
                        catch (Exception e)
                        {
                            Type("Something went wrong..." + e.Message, 3, 100, 1);
                            break;
                        }
                    }
                } // End foreach
                Type("-----SUMMARY-----", 7, 100, 1);
                Type(intTotalMoviesCount + " Total Movies.", 7, 100, 1);
                Type(intImagesCount + " Images Found.", 7, 100, 1);
                Type(intAlreadyConvertedFilesCount + " Movie Files Found.", 7, 100, 1);
                Type(intConvertedFilesCount + " Movies converted.", 7, 100, 1);
                Type(intNoTitleCount + " Missing Titles to convert.", 7, 100, 1);
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Type("It looks like that's the end of it.", 3, 100, 2);
        } // End GetDataToConvertBonusFeatures()

        /// <summary>
        /// This method gets the directory from the user then sends each subfolder to the RenameDirectory() method to be renamed.
        /// </summary>
        protected static void GetFolders()
        {
            // A bool to keep checking for the directory if the user inputs an invalid directory.
            bool keepAskingForDirectory = true;
            do
            {
                // Ask for the directory.
                Type("Enter your directory", 7, 100, 1);

                // Add that directory to the directory variable.
                var directory = Console.ReadLine();

                // Now get all directories in the given directory and put them in an array.
                string[] subdirectoryEntries = Directory.GetDirectories(directory);

                // Check that there are some subdirectories.
                if(subdirectoryEntries.Length > 0)
                {
                    // Since this is a valid directory then change our flag.
                    keepAskingForDirectory = false;

                    // Loop through each subdirectory and send them to be renamed.
                    foreach (string subdirectory in subdirectoryEntries)
                    {
                        // Check that the directory is a folder and not a file.
                        if (Directory.Exists(subdirectory))
                        {
                            // This path is a directory so send it to be renamed.
                            string folderName = Path.GetFileName(subdirectory);

                            // Send the original folder name and the folder name itself.
                            RenameDirectory(folderName, subdirectory);
                        }
                        else
                        {
                            // Let the user know that it is invalid.
                            Type(subdirectory + " is not a valid file or directory.", 14, 100, 1);
                        }
                    }
                }
                else
                {
                    // Let the user know that there are no subdirectories in the folder.
                    Type("There are no folders to rename in this directory.", 7, 100, 1);
                }
                
            } while (keepAskingForDirectory);

            // Now finish and let the user know that's it.
            Type("It looks like that's it.", 3, 100, 2);

        } // End RenameFolders()
        
        /// <summary>
        /// Takes the original folder name and figures out what to rename it to.
        /// </summary>
        /// <param name="folderName">The name of the folder to be renamed.</param>
        /// <param name="path">The full path to the folder.</param>
        protected static void RenameDirectory(string folderName, string path)
        {
            // Counts how many dashes are in the folder name.
            int intDashCount = CountCharacter(folderName, '-');

            // If there is only one dash in the name then it likely needs to be renamed.
            if (intDashCount == 1)
            {
                // Split the name on the dash into an array.
                string[] split = folderName.Split('-');

                // If the second split has a '(' then it likely was a movie like:
                // The Twilight Saga Breaking Dawn - Part 1 (2011)
                // and should not be renamed.
                if (split[1].Contains("("))
                {
                    Type(folderName + " was not split because the dash seems to be part of the movie title.", 7, 100, 1);
                }
                // Else it doesn't contain a '(' and is probably fine to replace.
                else
                {
                    // Now replace the original path name with the split name.
                    string replacedName = path.Replace(folderName, split[0]);

                    // Finally, actually rename the folder.
                    Directory.Move(path, replacedName);

                    // Tell the user what happened.
                    Type(folderName + " was split.", 7, 100, 1);
                }
                
            }
            // Else if there is more than one dash, I don't want to rename it.
            else if(intDashCount > 1)
            {
                // Tell the user it wasn't split because of too many dashes.
                // Just rename those manually.
                Type(folderName + " has more than one dash and wasn't split.", 7, 100, 1);

            }
            // Else it doesn't have any dashes and won't be renamed.
            else
            {
                // Tell the user it wasn't split because it has no dashes.
                Type(folderName + " has no dashes and was not split.", 7, 100, 1);

            }
        } // RenameDirectory()

        /// <summary>
        /// Counts the number of asked characters in a sent string and returns the count.
        /// </summary>
        /// <param name="value">The string to check.</param>
        /// <param name="ch">The character to count.</param>
        /// <returns></returns>
        protected static int CountCharacter(string value, char ch)
        {
            // Star with an empty count.
            int count = 0;

            // Loop through each character in the string.
            foreach(char c in value)
            {
                // If the character in the string is equal to the requested character then add one to our count.
                if(c == ch)
                {
                    count++;
                }
            }
            // Return the count.
            return count;
        } // End CountCharacter()

        protected static void HandBrake(string command)
        {
            try
            {
                // create the ProcessStartInfo using "cmd" as the program to be run,
                // and "/c " as the parameters.
                // Incidentally, /c tells cmd that we want it to execute the command that follows,
                // and then exit.
                ProcessStartInfo procStartInfo =
                    new ProcessStartInfo("cmd", "/c " + command);

                // The following commands are needed to redirect the standard output.
                // This means that it will be redirected to the Process.StandardOutput StreamReader.
                procStartInfo.RedirectStandardOutput = true;
                procStartInfo.UseShellExecute = false;
                // Do not create the black window.
                //procStartInfo.CreateNoWindow = true;
                // Now we create a process, assign its ProcessStartInfo and start it
                Process proc = new Process();
                proc.StartInfo = procStartInfo;
                proc.Start();
                // Get the output into a string
                string result = proc.StandardOutput.ReadToEnd();
                // Display the command output.
                Console.WriteLine(result);
            }
            catch (Exception objException)
            {
                Type("Unable to convert file. | " + objException.Message, 7, 100, 1);
            }
        } // End HandBrake()

        /// <summary>
        /// An attempt to dry out the code.
        /// Not fully functioning.
        /// </summary>
        /// <param name="strTitleRowRange"></param>
        /// <returns></returns>
        protected static IList<IList<Object>> GetTitleRowData(string strTitleRowRange)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, strTitleRowRange);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;

            return titleValues;
        } // End GetTitleRowData()

        protected static void DefaultMovies()
        {
            // Declare variables.
            int intLetterColumn = -1,
                intCleanTitleColumn = -1,
                intNfoBodyColumn = -1;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, MOVIES_TITLE_RANGE);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;
            if (titleValues != null && titleValues.Count > 0)
            {
                int x = 0;
                foreach (var row in titleValues)
                {
                    do
                    {
                        if (row[x].ToString() == "Movie Letter")
                        {
                            intLetterColumn = x;
                        }
                        else if (row[x].ToString() == "Clean Title")
                        {
                            intCleanTitleColumn = x;
                        }
                        else if (row[x].ToString() == "NFO Body")
                        {
                            intNfoBodyColumn = x;
                        }
                        x++;

                    } while (x < row.Count);

                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }

            if(intNfoBodyColumn == -1 || intLetterColumn == -1 || intCleanTitleColumn == -1)
            {
                Type("It looks like a column name was changed and I can no longer find it.", 1, 100, 1);
                Type("Letter: " + intLetterColumn + ". Title: " + intCleanTitleColumn + ". NFO: " + intNfoBodyColumn, 1, 100, 1);
            }
            else
            {
                string directory = "F:\\Movies\\";

                SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
                        service.Spreadsheets.Values.Get(SPREADSHEET_ID, MOVIES_DATA_RANGE);

                ValueRange dataRowResponse = dataRowRequest.Execute();
                IList<IList<Object>> dataValues = dataRowResponse.Values;
                if (dataValues != null)
                {
                    foreach (var row in dataValues)
                    {
                        if (row.Count > 20) // If it's an empty row then it should have much less than 20.
                        {
                            try
                            {
                                Directory.CreateDirectory(directory + row[intLetterColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString());

                                string path = directory + row[intLetterColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString() + ".nfo";
                                string fileText = row[intNfoBodyColumn].ToString();

                                File.WriteAllText(path, fileText, Encoding.UTF8);

                                Console.WriteLine("NFO created for: " + row[intCleanTitleColumn].ToString());
                            }
                            catch (Exception e)
                            {
                                Type("Something went wrong..." + e.Message, 3, 100, 1);
                                break;
                            }

                        }
                    }

                }
                else
                {
                    Console.WriteLine("No data found.");
                }
                Type("It looks like that's the end of it.", 3, 100, 2);
            }
            
        } // End DefaultMovies()

        protected static void DefaultMoviesQuick()
        {
            // Declare variables.
            int intLetterColumn = 0,
                intCleanTitleColumn = 0,
                intNfoBodyColumn = 0,
                intQuickCreateColumn = 0;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, MOVIES_TITLE_RANGE);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;
            if (titleValues != null && titleValues.Count > 0)
            {
                int x = 0;
                foreach (var row in titleValues)
                {
                    do
                    {
                        if (row[x].ToString() == "Movie Letter")
                        {
                            intLetterColumn = x;
                        }
                        else if (row[x].ToString() == "Clean Title")
                        {
                            intCleanTitleColumn = x;
                        }
                        else if (row[x].ToString() == "NFO Body")
                        {
                            intNfoBodyColumn = x;
                        }
                        else if (row[x].ToString() == "Quick Create")
                        {
                            intQuickCreateColumn = x;
                        }
                        x++;

                    } while (x < row.Count);

                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }

            string directory = "F:\\Movies\\";

            SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, MOVIES_DATA_RANGE);

            ValueRange dataRowResponse = dataRowRequest.Execute();
            IList<IList<Object>> dataValues = dataRowResponse.Values;
            if (dataValues != null)
            {
                foreach (var row in dataValues)
                {
                    if (row.Count > intQuickCreateColumn) // If it's an empty row then it should have much less than 20.
                    {
                        try
                        {
                            if (row[intQuickCreateColumn].ToString().ToUpper() == "X")
                            {
                                Directory.CreateDirectory(directory + row[intLetterColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString());

                                string path = directory + row[intLetterColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString() + ".nfo";
                                string fileText = row[intNfoBodyColumn].ToString();

                                File.WriteAllText(path, fileText, Encoding.UTF8);

                                Type("NFO created for: " + row[intCleanTitleColumn].ToString(), 3, 100, 1);
                            }
                        }
                        catch (Exception e)
                        {
                            Type("Something went wrong..." + e.Message, 3, 100, 1);
                            break;
                        }

                    }
                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Type("It looks like that's the end of it.", 3, 100, 2);
        } // End DefaultMoviesQuick()

        protected static void DefaultTempMovies()
        {
            // Declare variables.
            string directory = "";

            int intLetterColumn = 0,
                intCleanTitleColumn = 0,
                intNfoBodyColumn = 0,
                intDriveLetter = 0;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, TEMP_MOVIES_TITLE_RANGE);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;
            if (titleValues != null && titleValues.Count > 0)
            {
                int x = 0;
                foreach (var row in titleValues)
                {
                    do
                    {
                        if (row[x].ToString() == "Letter")
                        {
                            intLetterColumn = x;
                        }
                        else if (row[x].ToString() == "Clean Title")
                        {
                            intCleanTitleColumn = x;
                        }
                        else if (row[x].ToString() == "NFO Body")
                        {
                            intNfoBodyColumn = x;
                        }
                        else if (row[x].ToString() == "Drive Letter")
                        {
                            intDriveLetter = x;
                        }
                        x++;

                    } while (x < row.Count);

                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }

            SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, TEMP_MOVIES_DATA_RANGE);

            ValueRange dataRowResponse = dataRowRequest.Execute();
            IList<IList<Object>> dataValues = dataRowResponse.Values;
            if (dataValues != null)
            {
                foreach (var row in dataValues)
                {
                    if (row.Count > 20) // If it's an empty row then it should have much less than 20.
                    {
                        try
                        {
                            if(row[intDriveLetter].ToString() == "C")
                            {
                                directory = "C:\\Plex\\Movies\\";
                            }
                            else if (row[intDriveLetter].ToString() == "G")
                            {
                                directory = "G:\\Movies  - Temp\\";
                            }

                            Directory.CreateDirectory(directory + row[intLetterColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString());

                            string path = directory + row[intLetterColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString() + ".nfo";
                            string fileText = row[intNfoBodyColumn].ToString();

                            File.WriteAllText(path, fileText, Encoding.UTF8);

                            Console.WriteLine(path);
                        }
                        catch (Exception e)
                        {
                            Type("Something went wrong..." + e.Message, 3, 100, 1);
                            break;
                        }

                    }
                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Type("It looks like that's the end of it.", 3, 100, 2);
        } // End DefaultTempMovies()

        protected static void DefaultTempMoviesQuick()
        {
            // Declare variables.
            int intLetterColumn = 0,
                intCleanTitleColumn = 0,
                intNfoBodyColumn = 0,
                intQuickCreateColumn = 0;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, TEMP_MOVIES_TITLE_RANGE);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;
            if (titleValues != null && titleValues.Count > 0)
            {
                int x = 0;
                foreach (var row in titleValues)
                {
                    do
                    {
                        if (row[x].ToString() == "Letter")
                        {
                            intLetterColumn = x;
                        }
                        else if (row[x].ToString() == "Clean Title")
                        {
                            intCleanTitleColumn = x;
                        }
                        else if (row[x].ToString() == "NFO Body")
                        {
                            intNfoBodyColumn = x;
                        }
                        else if (row[x].ToString() == "Quick Create")
                        {
                            intQuickCreateColumn = x;
                        }
                        x++;

                    } while (x < row.Count);

                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }

            string directory = "G:\\Movies  - Temp\\";

            SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, TEMP_MOVIES_DATA_RANGE);

            ValueRange dataRowResponse = dataRowRequest.Execute();
            IList<IList<Object>> dataValues = dataRowResponse.Values;
            if (dataValues != null)
            {
                foreach (var row in dataValues)
                {
                    if (row.Count > intQuickCreateColumn)
                    {
                        try
                        {
                            if (row[intQuickCreateColumn].ToString().ToUpper() == "X")
                            {
                                Directory.CreateDirectory(directory + row[intLetterColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString());

                                string path = directory + row[intLetterColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString() + "\\" + row[intCleanTitleColumn].ToString() + ".nfo";
                                string fileText = row[intNfoBodyColumn].ToString();

                                File.WriteAllText(path, fileText, Encoding.UTF8);

                                Type("NFO created for: " + row[intCleanTitleColumn].ToString(), 3, 100, 1);
                            }
                        }
                        catch (Exception e)
                        {
                            Type("Something went wrong..." + e.Message, 3, 100, 1);
                            break;
                        }

                    }
                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Type("It looks like that's the end of it.", 3, 100, 2);
        } // End DefaultTempMoviesQuick()

        protected static void DefaultYoutube()
        {
            // Declare variables.
            int intTitleColumn = 0,
                intTopFolderColumn = 0,
                intPlaylistColumn = 0,
                intNfoColumn = 0;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, YOUTUBE_TITLE_RANGE);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;
            if (titleValues != null && titleValues.Count > 0)
            {
                int x = 0;
                foreach (var row in titleValues)
                {
                    do
                    {
                        if (row[x].ToString() == "Title")
                        {
                            intTitleColumn = x;
                        }
                        else if (row[x].ToString() == "Top Folder")
                        {
                            intTopFolderColumn = x;
                        }
                        else if (row[x].ToString() == "Playlist")
                        {
                            intPlaylistColumn = x;
                        }
                        else if (row[x].ToString() == "NFO")
                        {
                            intNfoColumn = x;
                        }
                        x++;

                    } while (x < row.Count);

                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }

            string directory = "E:\\Plex\\Youtube\\";

            SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, YOUTUBE_DATA_RANGE);

            ValueRange dataRowResponse = dataRowRequest.Execute();
            IList<IList<Object>> dataValues = dataRowResponse.Values;
            if (dataValues != null)
            {
                foreach (var row in dataValues)
                {
                    if (row.Count > 1)
                    {
                        try
                        {
                            Directory.CreateDirectory(directory + row[intTopFolderColumn].ToString() + "\\" + row[intPlaylistColumn].ToString() + "\\" + row[intTitleColumn].ToString());

                            string path = directory + row[intTopFolderColumn].ToString() + "\\" + row[intPlaylistColumn].ToString() + "\\" + row[intTitleColumn].ToString() + "\\" + row[intTitleColumn].ToString() + ".nfo";
                            string fileText = row[intNfoColumn].ToString();

                            File.WriteAllText(path, fileText, Encoding.UTF8);

                            Type("Directory created for: " + row[intTitleColumn].ToString(), 1, 100, 1);
                        }
                        catch (Exception e)
                        {
                            Type("Something went wrong..." + e.Message, 3, 100, 1);
                            break;
                        }

                    }
                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Type("It looks like that's the end of it.", 3, 100, 2);
        } // End DefaultYoutube()

        protected static void DefaultYoutubeQuick()
        {
            // Declare variables.
            int intTitleColumn = 0,
                intTopFolderColumn = 0,
                intPlaylistColumn = 0,
                intNfoColumn = 0,
                intQuickCreateColumn = 0;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, YOUTUBE_TITLE_RANGE);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;
            if (titleValues != null && titleValues.Count > 0)
            {
                int x = 0;
                foreach (var row in titleValues)
                {
                    do
                    {
                        if (row[x].ToString() == "Title")
                        {
                            intTitleColumn = x;
                        }
                        else if (row[x].ToString() == "Top Folder")
                        {
                            intTopFolderColumn = x;
                        }
                        else if (row[x].ToString() == "Playlist")
                        {
                            intPlaylistColumn = x;
                        }
                        else if (row[x].ToString() == "NFO")
                        {
                            intNfoColumn = x;
                        }
                        else if (row[x].ToString() == "Quick Create")
                        {
                            intQuickCreateColumn = x;
                        }
                        x++;

                    } while (x < row.Count);

                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }

            string directory = "E:\\Plex\\Youtube\\";

            SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, YOUTUBE_DATA_RANGE);

            ValueRange dataRowResponse = dataRowRequest.Execute();
            IList<IList<Object>> dataValues = dataRowResponse.Values;
            if (dataValues != null)
            {
                foreach (var row in dataValues)
                {
                    if (row.Count > intQuickCreateColumn)
                    {
                        try
                        {
                            if(row[intQuickCreateColumn].ToString().ToUpper() == "X")
                            {
                                Directory.CreateDirectory(directory + row[intTopFolderColumn].ToString() + "\\" + row[intPlaylistColumn].ToString() + "\\" + row[intTitleColumn].ToString());

                                string path = directory + row[intTopFolderColumn].ToString() + "\\" + row[intPlaylistColumn].ToString() + "\\" + row[intTitleColumn].ToString() + "\\" + row[intTitleColumn].ToString() + ".nfo";
                                string fileText = row[intNfoColumn].ToString();

                                File.WriteAllText(path, fileText, Encoding.UTF8);

                                Type("Directory created for: " + row[intTitleColumn].ToString(), 1, 100, 1);
                            }
                            
                        }
                        catch (Exception e)
                        {
                            Type("Something went wrong..." + e.Message, 3, 100, 1);
                            break;
                        }

                    }
                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Type("It looks like that's the end of it.", 3, 100, 2);
        } // End DefaultYoutubeQuick()

        protected static void GetDataToConvertEpisodes()
        {
            // Declare variables.
            int intInputFolderColumn = -1,
                intOutputFolderColumn = -1,
                intIsoTitleNumberColumn = -1,
                intChapterNumberColumn = -1,
                intTotalEpisodesCount = 0,
                intImagesCount = 0,
                intAlreadyConvertedFilesCount = 0,
                intNoTitleCount = 0,
                intConvertedFilesCount = 0;

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            SpreadsheetsResource.ValuesResource.GetRequest titleRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, EPISODES_TITLE_RANGE);

            ValueRange titleRowresponse = titleRowRequest.Execute();
            IList<IList<Object>> titleValues = titleRowresponse.Values;
            if (titleValues != null && titleValues.Count > 0)
            {
                int x = 0;
                foreach (var row in titleValues)
                {
                    do
                    {
                        if (row[x].ToString() == "Image Location")
                        {
                            intInputFolderColumn = x;
                        }
                        else if (row[x].ToString() == "Episode Location")
                        {
                            intOutputFolderColumn = x;
                        }
                        else if (row[x].ToString() == "ISO Title #")
                        {
                            intIsoTitleNumberColumn = x;
                        }
                        else if (row[x].ToString() == "Chapter")
                        {
                            intChapterNumberColumn = x;
                        }
                        x++;

                    } while (x < row.Count);

                }

            }
            else
            {
                Console.WriteLine("No data found.");
            }

            SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, EPISODES_DATA_RANGE);

            ValueRange dataRowResponse = dataRowRequest.Execute();
            IList<IList<Object>> dataValues = dataRowResponse.Values;
            if (dataValues != null)
            {
                foreach (var row in dataValues)
                {
                    if (row.Count > 12)
                    {
                        Console.WriteLine("Row count is: " + row.Count);

                        intTotalEpisodesCount++;
                        try
                        {
                            string i = row[intInputFolderColumn].ToString(),
                                    o = row[intOutputFolderColumn].ToString(),
                                    title = row[intIsoTitleNumberColumn].ToString(),
                                    chapter = row[intChapterNumberColumn].ToString();
                                    
                            if (File.Exists(i))
                            {
                                //Type("We found " + i, 0, 0, 1);
                                intImagesCount++;
                                if (File.Exists(o))
                                {
                                    //Type("We found " + o, 0, 0, 1);
                                    //Type("We won't have to convert this one.", 0, 0, 1);
                                    intAlreadyConvertedFilesCount++;
                                }
                                else
                                {
                                    Type("We found " + i, 0, 0, 1);
                                    Type("We didn't find " + o, 0, 0, 1);

                                    // Create the directory if needed.
                                    int lastIndexOf = o.LastIndexOf("\\");
                                    string fileLocation = o.Substring(0, lastIndexOf);
                                    Directory.CreateDirectory(fileLocation);
                                    Type("Directory created at: " + fileLocation, 0, 0, 1);

                                    if (title != "")
                                    {
                                        Type("We will use title #" + title, 0, 0, 1);
                                        if (chapter != "")
                                            Type("And we will use Chapter #" + chapter, 0, 0, 1);

                                        string strMyConversionString = "HandBrakeCLI -i \"" + i + "\" -o \"" + o + "\" --preset-import-file preset.json --two-pass -t " + title;
                                        if (chapter != "")
                                            strMyConversionString += " -c " + chapter;

                                        Type(strMyConversionString, 0, 0, 1);
                                        HandBrake(strMyConversionString);
                                        intConvertedFilesCount++;
                                        Type("-------------------------------------------------------------------", 0, 0, 1);
                                    }
                                    else
                                    {
                                        Type("We don't have a title to go off of.", 0, 0, 1);
                                        intNoTitleCount++;
                                    }
                                }
                            }
                            else
                            {
                                Type("We didn't find " + i, 0, 0, 1);
                                Type("We won't be able to convert this one at this time.", 0, 0, 1);
                            }
                            //Type("-------------------------------------------------------------------", 0, 0, 1);

                        }
                        catch (Exception e)
                        {
                            Type("Something went wrong..." + e.Message, 3, 100, 1);
                            break;
                        }
                    }
                } // End foreach
                Type("-----SUMMARY-----", 7, 100, 1);
                Type(intTotalEpisodesCount + " Total Episodes.", 7, 100, 1);
                Type(intImagesCount + " Images Found.", 7, 100, 1);
                Type(intAlreadyConvertedFilesCount + " Episode Files Found.", 7, 100, 1);
                Type(intConvertedFilesCount + " Episodes converted.", 7, 100, 1);
                Type(intNoTitleCount + " Missing Titles to convert.", 7, 100, 1);
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Type("It looks like that's the end of it.", 3, 100, 2);
        } // End GetDataToConvertEpisodes()

        protected static void CountFiles()
        {
            bool keepAskingForDirectory = true;
            do
            {
                Type("Enter your directory", 7, 100, 1);
                var directory = Console.ReadLine();
                string[] fileEntries = Directory.GetFiles(directory);
                string[] subdirectoryEntries = Directory.GetDirectories(directory);
                Type("The directory: '" + directory + "' contains " + subdirectoryEntries.Length + " sub folders and " + fileEntries.Length + " files.", 0, 100, 1);
                if (File.Exists(directory))
                {
                    // This path is a file
                    ProcessFile(directory);
                    keepAskingForDirectory = false;
                }
                else if (Directory.Exists(directory))
                {
                    // This path is a directory
                    ProcessDirectory(directory);
                    keepAskingForDirectory = false;
                }
                else
                {
                    Type(directory + " is not a valid file or directory.", 14, 100, 1);
                }
            } while (keepAskingForDirectory);

            //int fileCount = Directory.GetFiles(directory).Length;
            //foreach (FileInfo f in directory.GetFiles().Length)
            //{
            //    Type(f.FullName, 100, 1);
            //}
            //int fileCount = Directory.getf
            //Type("Number of files: " + fileCount, 100, 1);
            Type("It looks like that's it.", 3, 100, 2);
        } // End CountFiles()

        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        public static void ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            if (fileEntries.Length > 0)
            {
                //if (fileEntries.Length < 3 & fileEntries.Length > 0)
                //{
                //    Type(fileEntries[0].Replace(".nfo", "").Replace(".mp4", "").Replace(".jpg", ""),100,1);
                //}
                //foreach (string fileName in fileEntries)
                //    ProcessFile(fileName);

                int nfoCount = 0, jpgCount = 0, mp4Count = 0, mkvCount = 0, unidentifiedCount = 0, isoCount = 0;
                foreach (string fileName in fileEntries)
                {
                    if (fileName.ToUpper().Contains(".NFO"))
                        nfoCount++;
                    else if (fileName.ToUpper().Contains(".JPG"))
                        jpgCount++;
                    else if (fileName.ToUpper().Contains(".MP4"))
                        mp4Count++;
                    else if (fileName.ToUpper().Contains(".MKV"))
                        mkvCount++;
                    else if (fileName.ToUpper().Contains(".ISO"))
                        isoCount++;
                    else
                    {
                        //Type("Unidentified file: " + fileName, 0, 0, 1);
                        unidentifiedCount++;
                    }

                }
                Type(nfoCount + " nfo, " + jpgCount + " jpg, " + mp4Count + " mp4, " + mkvCount + " mkv, " + isoCount + " iso, " + unidentifiedCount + " unidentified in " + targetDirectory, 0, 0, 1);
            }
            else
                Type(targetDirectory, 0, 0, 1);
            
            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory);
        }

        // Insert logic for processing found files here.
        public static void ProcessFile(string path)
        {
            Console.WriteLine("Processed file '{0}'.", path);
        }

        /// <summary>
        /// Simply types out the text in a typewriter manner. Then adds the number of new lines.
        /// </summary>
        /// <param name="myString"> The string to type out. </param>
        /// <param name="speed"> The speed to type. </param>
        /// <param name="timeToPauseBeforeNewLine"> The amount of time in Milliseconds to wait before starting the next line. </param>
        /// <param name="numberOfNewLines"> The number of new lines to insert. </param>
        static void Type(string myString, int speed, int timeToPauseBeforeNewLine, int numberOfNewLines)
        {
            for (int i = 0; i < myString.Length; i++)
            {
                Console.Write(myString[i]);
                Thread.Sleep(speed);
            }

            Thread.Sleep(timeToPauseBeforeNewLine);

            while (numberOfNewLines > 0)
            {
                Console.WriteLine();
                numberOfNewLines--;
            }
        } // End Type()
    }
}