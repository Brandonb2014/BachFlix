using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using SheetsQuickstart;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using TmdbApiCall;

namespace BachFlixNfoCall
{
    class BachFlixNfo
    {
        // If modifying these scopes, delete your previously saved credentials
        // at \BachFlixNfo\bin\Debug\token.json\Google.Apis.Auth.OAuth2.Responses.TokenResponse-user
        static readonly string[] SCOPES = { SheetsService.Scope.Spreadsheets };
        static string APLICATION_NAME = "Google Sheets API .NET Quickstart";
        static readonly string SPREADSHEET_ID = "1LE9Tiz0TgcG60qeul_y9wC4j8qNLQlfKTLnAg5tgBr0";

        // Data ranges for each sheet.
        private const string MOVIES_TITLE_RANGE = "Movies!A2:2";
        private const string MOVIES_DATA_RANGE = "Movies!A3:4010";
        private const string FITNESS_VIDEO_TITLE_RANGE = "Fitness Videos!A1:1";
        private const string FITNESS_VIDEO_DATA_RANGE = "Fitness Videos!A2:1000";

        private const bool DEBUG = false;

        public static void InputTvShowPlots(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, int type = 1)
        {
            int intPlotDoneCount = 0, intTmdbIdCorrectedCount = 0, intPlotSkippedCount = 0, intPlotNotFoundCount = 0, intRowNum = 3;

            string tmdbIdValue = "", combinedEpisodeName = "", episode1TitleValue = "", episode2TitleValue = "", episode1PlotValue = "", episode2PlotValue = "", episode1SeasonValue = "", episode1NumValue = "", episode2SeasonValue = "", episode2NumValue = "", strCellToPutData = "";
            int episode1PlotColumnNum = 0, episode2PlotColumnNum = 0; // Used to input the returned plot back into the Google Sheet.
            dynamic tmdbResponse;

            foreach (var row in data)
            {
                if (row.Count > 1)
                {
                    try
                    {
                        bool responseIsBroken = true;
                        combinedEpisodeName = row[Convert.ToInt16(sheetVariables["Combined Episode Name"])].ToString();
                        tmdbIdValue = row[Convert.ToInt16(sheetVariables["TMDB ID"])].ToString();
                        episode1TitleValue = row[Convert.ToInt16(sheetVariables["Episode 1 Title"])].ToString();
                        episode2TitleValue = row[Convert.ToInt16(sheetVariables["Episode 2 Title"])].ToString();
                        episode1PlotValue = row[Convert.ToInt16(sheetVariables["Episode 1 Plot"])].ToString();
                        episode2PlotValue = row[Convert.ToInt16(sheetVariables["Episode 2 Plot"])].ToString();
                        episode1PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 1 Plot"]);
                        episode2PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 2 Plot"]);
                        episode1SeasonValue = row[Convert.ToInt16(sheetVariables["Episode 1 Season"])].ToString();
                        episode1NumValue = row[Convert.ToInt16(sheetVariables["Episode 1 No."])].ToString();
                        episode2SeasonValue = row[Convert.ToInt16(sheetVariables["Episode 2 Season"])].ToString();
                        episode2NumValue = row[Convert.ToInt16(sheetVariables["Episode 2 No."])].ToString();

                        if (combinedEpisodeName != "")
                        {
                            if (type == 1) // Input only missing plots.
                            {
                                if (episode1PlotValue.Equals(""))
                                {
                                    do
                                    {
                                        Thread.Sleep(250);
                                        tmdbResponse = TmdbApi.TvEpisodesGetDetails(tmdbIdValue, episode1SeasonValue, episode1NumValue);

                                        if (tmdbResponse.overview != null)
                                        {
                                            episode1PlotValue = tmdbResponse.overview.ToString();
                                            responseIsBroken = false;
                                        }
                                        else if (tmdbResponse.status_message != null)
                                        {
                                            Program.Type(episode1TitleValue + " errored | " + tmdbResponse.status_message, 0, 0, 1, "Red");
                                            episode1PlotValue = "";
                                            responseIsBroken = false;
                                        }
                                        else
                                        {
                                            Thread.Sleep(5000);
                                        }
                                    } while (responseIsBroken);

                                    strCellToPutData = "Combined Episodes!" + ColumnNumToLetter(episode1PlotColumnNum) + intRowNum;

                                    if (episode1PlotValue != "")
                                    {
                                        if (WriteSingleCellToSheet(episode1PlotValue, strCellToPutData))
                                        {
                                            Program.Type("Plot saved for: " + episode1TitleValue, 0, 0, 1, "Green");
                                            intPlotDoneCount++;
                                        }
                                        else
                                        {
                                            Program.Type("An error occured!", 0, 0, 1, "Red");
                                        }
                                    }
                                    else
                                    {
                                        intPlotNotFoundCount++;
                                    }
                                }
                                else intPlotSkippedCount++;

                                if (episode2PlotValue.Equals(""))
                                {
                                    do
                                    {
                                        Thread.Sleep(250);
                                        tmdbResponse = TmdbApi.TvEpisodesGetDetails(tmdbIdValue, episode2SeasonValue, episode2NumValue);

                                        if (tmdbResponse.overview != null)
                                        {
                                            episode2PlotValue = tmdbResponse.overview.ToString();
                                            responseIsBroken = false;
                                        }
                                        else if (tmdbResponse.status_message != null)
                                        {
                                            Program.Type(episode2TitleValue + " errored | " + tmdbResponse.status_message, 0, 0, 1, "Red");
                                            episode2PlotValue = "";
                                            responseIsBroken = false;
                                        }
                                        else
                                        {
                                            Thread.Sleep(5000);
                                        }
                                    } while (responseIsBroken);

                                    strCellToPutData = "Combined Episodes!" + ColumnNumToLetter(episode2PlotColumnNum) + intRowNum;

                                    if (episode2PlotValue != "")
                                    {
                                        if (WriteSingleCellToSheet(episode2PlotValue, strCellToPutData))
                                        {
                                            Program.Type("Plot saved for: " + episode2TitleValue, 0, 0, 1, "Green");
                                            intPlotDoneCount++;
                                        }
                                        else
                                        {
                                            Program.Type("An error occured!", 0, 0, 1, "Red");
                                        }
                                    }
                                    else
                                    {
                                        intPlotNotFoundCount++;
                                    }
                                }
                                else intPlotSkippedCount++;

                            }
                            else if (type == 2) // Input ALL TMDB IDs including fixing wrong ones.
                            {
                                //do
                                //{
                                //    Thread.Sleep(250);
                                //    tmdbResponse = TmdbApi.MoviesGetDetails(ImdbId);

                                //    if (tmdbResponse.id != null)
                                //    {
                                //        tmdbId = tmdbResponse.id.ToString();
                                //        responseIsBroken = false;
                                //    }
                                //    else if (tmdbResponse.status_message != null)
                                //    {
                                //        Type(ImdbTitle + " | " + tmdbResponse.status_message, 0, 0, 1, "Red");
                                //        tmdbId = "";
                                //        responseIsBroken = false;
                                //    }
                                //    else
                                //    {
                                //        Thread.Sleep(5000);
                                //    }
                                //} while (responseIsBroken);


                                //strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbIdColumnNum) + intRowNum;

                                //if (tmdbId != "")
                                //{
                                //    if (tmdbIdValue.Equals("")) // If the ID is missing insert it.
                                //    {
                                //        if (WriteSingleCellToSheet(tmdbId, strCellToPutData))
                                //        {
                                //            Type("TMDB ID saved for: " + ImdbTitle, 0, 0, 1, "Green");
                                //            intPlotDoneCount++;
                                //        }
                                //        else
                                //        {
                                //            Type("An error occured!", 0, 0, 1, "Red");
                                //        }

                                //    }
                                //    else if (tmdbIdValue != tmdbId) // Or if the new ID doesn't equal the old one overwrite it.
                                //    {
                                //        if (WriteSingleCellToSheet(tmdbId, strCellToPutData))
                                //        {
                                //            Type("TMDB ID corrected for: " + ImdbTitle, 0, 0, 1, "Blue");
                                //            intTmdbIdCorrectedCount++;
                                //        }
                                //        else
                                //        {
                                //            Type("An error occured!", 0, 0, 1, "Red");
                                //        }

                                //    }
                                //    else // Else just skip it.
                                //    {
                                //        intTmdbIdSkippedCount++;
                                //    }

                                //}
                                //else
                                //{
                                //    Type("We didn't find a TMDB ID for: " + ImdbTitle, 0, 0, 1, "Yellow");
                                //}
                            }
                        }

                        intRowNum++;

                    }
                    catch (Exception e)
                    {
                        Program.Type("Something went wrong with " + episode1TitleValue + " | " + e.Message, 3, 100, 1, "Red");
                    }

                }
            }
            Console.WriteLine();
            Program.Type("It looks like theat's the end of it.", 0, 0, 1);
            Program.Type("TMDB IDs inserted: " + intPlotDoneCount, 0, 0, 1, "Green");
            Program.Type("TMDB IDs skipped: " + intPlotSkippedCount, 0, 0, 1, "Yellow");
            Program.Type("TMDB IDs corrected: " + intTmdbIdCorrectedCount, 0, 0, 1, "Blue");
            Program.Type("TMDB IDs not available: " + intPlotNotFoundCount, 0, 0, 1, "Red");
        } // End InputTvShowPlots()

        public static void OverwriteFitnessVideoNfoFiles(IList<IList<Object>> data, Dictionary<string, int> sheetVariables)
        {
            string d = @"C:\Plex";
            string baseDir = d + @"\Fitness Videos";
            Directory.CreateDirectory(d);
            Directory.CreateDirectory(baseDir);
            foreach (var row in data)
            {
                Console.WriteLine("row.Count: " + row.Count);
                if (row.Count > 1)
                {
                    var program = CleanString(row[Convert.ToInt16(sheetVariables["Program"])].ToString());
                    var subfolder = CleanString(row[Convert.ToInt16(sheetVariables["Subfolder"])].ToString());
                    var name = CleanString(row[Convert.ToInt16(sheetVariables["Name"])].ToString());
                    var title = CleanString(row[Convert.ToInt16(sheetVariables["Title"])].ToString());
                    var nfoBody = row[Convert.ToInt16(sheetVariables["NFO Body"])].ToString();

                    string programDir = Path.Combine(baseDir, program);
                    Directory.CreateDirectory(programDir);
                    string subfolderDir = Path.Combine(programDir, subfolder);
                    Directory.CreateDirectory(subfolderDir);
                    string nameDir = Path.Combine(subfolderDir, name);
                    Directory.CreateDirectory(nameDir);
                    string fullPath = Path.Combine(nameDir, name) + ".nfo";

                    WriteNfoFile(fullPath, nfoBody);
                }
            }
        } // End OverwriteFitnessVideoNfoFiles()

        public static void FixRecordedNames(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string directory)
        {
            Program.DisplayMessage("warning", "Searching directory for recorded names...");
            if (DEBUG) Console.WriteLine("data.Count: " + data.Count);
            foreach (var row in data)
            {
                if (DEBUG) Console.WriteLine("row.Count: " + row.Count);
                if (row.Count > 1) 
                { 
                    var recordedName = row[Convert.ToInt16(sheetVariables["Recorded Name"])].ToString();
                    var actualName = row[Convert.ToInt16(sheetVariables["Actual Name"])].ToString();

                    if (DEBUG) Console.WriteLine("recordedName: " + recordedName);
                    if (DEBUG) Console.WriteLine("actualName: " + actualName);

                    if (recordedName != "")
                    {
                        // Since my sheet may or may not have the .mp4 extension with the name we just remove it to be sure and add it manually.
                        recordedName.ToLower().Replace(".mp4", "");
                        var sourceFile = Path.Combine(directory, recordedName) + ".mp4";

                        if (File.Exists(sourceFile))
                        {
                            if (actualName != "")
                            {
                                // Since my sheet may or may not have the .mp4 extension with the name we just remove it to be sure and add it manually.
                                actualName.ToLower().Replace(".mp4", "");
                                var destinationFile = Path.Combine(directory, actualName) + ".mp4";

                                // Now that we have verified we have the correct source and destination files, move them.
                                File.Move(sourceFile, destinationFile);

                                // Display results.
                                Program.DisplayMessage("warning", " " + recordedName);
                                Program.DisplayMessage("info", " changed to");
                                Program.DisplayMessage("success", " " + actualName);
                                Program.DisplayEndOfCurrentProcessLines();
                            }
                            else
                            {
                                Program.DisplayMessage("error", "Missing Actual Name for: " + recordedName);
                            }

                        }
                    }
                    else
                    {
                        Program.DisplayMessage("error", "We noticed a missing Recorded Name so we skipped this line.");
                    }
                }
            }
            Program.DisplayMessage("success", "DONE");
        } // End FixRecordedNames()

        public static string CleanString(string s)
        {
            string pattern = @"[\\/:*?""<>|]";

            Regex regEx = new Regex(pattern);
            return regEx.Replace(s, "");
        } // End CleanString()

        /// <summary>
        /// Simply writes the NFO File to the chosen path.
        /// </summary>
        /// <param name="path">The path to the folder location.</param>
        /// <param name="fileText">The text of the NFO File to be saved.</param>
        protected static void WriteNfoFile(string path, string fileText)
        {
            try
            {
                File.WriteAllText(path, fileText, Encoding.UTF8);
            }
            catch (Exception e)
            {
                Program.Type("Something went wrong writing to path: " + path + " | " + e.Message, 3, 100, 1, "Red");
                Program.Type(e.Message, 3, 100, 2, "DarkRed");
                throw;
            }

        } // End WriteNfoFile()

        /// <summary>
        /// Convert the number to the column letter.
        /// i.e. 0 = A
        /// </summary>
        /// <param name="columnNum">The number of the column.</param>
        /// <returns>The column letter.</returns>
        protected static string ColumnNumToLetter(int columnNum)
        {
            string[] myString = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ" };

            return myString[columnNum];
        }

        protected static bool WriteSingleCellToSheet(string strDataToSave, string strCellToSaveData)
        {
            try
            {
                Thread.Sleep(1000); // Sleep for a second so we don't go over the Google allotted requests.
                // How the input data should be interpreted.
                SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum valueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

                // TODO: Assign values to desired properties of `requestBody`. All existing
                // properties will be replaced:
                ValueRange requestBody = new ValueRange
                {
                    MajorDimension = "COLUMNS" // "ROWS" / "COLUMNS"
                };
                var oblist = new List<object>() { strDataToSave };
                requestBody.Values = new List<IList<object>> { oblist };

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

                SpreadsheetsResource.ValuesResource.UpdateRequest request = sheetsService.Spreadsheets.Values.Update(requestBody, SPREADSHEET_ID, strCellToSaveData);
                request.ValueInputOption = valueInputOption;

                // To execute asynchronously in an async method, replace `request.Execute()` as shown:
                UpdateValuesResponse response = request.Execute();
                // Data.UpdateValuesResponse response = await request.ExecuteAsync();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }

        } // End WriteSingleCellToSheet

    }
}
