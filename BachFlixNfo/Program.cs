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
using System.Text;
using System.Diagnostics;
using TmdbApiCall;
using BachFlixNfoCall;
using System.Collections;
using System.Management;
using Newtonsoft.Json;

namespace SheetsQuickstart
{
    class Program
    {
        // Data ranges for each sheet.
        private const string MOVIES_TITLE_RANGE = "Movies!A2:2";
        private const string MOVIES_DATA_RANGE = "Movies!A3:5200";
        private const string TEMP_MOVIES_TITLE_RANGE = "Temp!A2:2";
        private const string TEMP_MOVIES_DATA_RANGE = "Temp!A3:2001";
        private const string YOUTUBE_TITLE_RANGE = "YouTube!A2:2";
        private const string YOUTUBE_DATA_RANGE = "YouTube!A3:2344";
        private const string FITNESS_VIDEO_TITLE_RANGE = "Fitness Videos!A1:1";
        private const string FITNESS_VIDEO_DATA_RANGE = "Fitness Videos!A2:401";
        private const string BONUS_TITLE_RANGE = "Bonus!A1:1";
        private const string BONUS_DATA_RANGE = "Bonus!A2:2036";
        private const string EPISODES_TITLE_RANGE = "Episodes!A2:2";
        private const string EPISODES_DATA_RANGE = "Episodes!A3:1200";
        private const string TEMP_EPISODES_TITLE_RANGE = "Temp Episodes!A1:1";
        private const string TEMP_EPISODES_DATA_RANGE = "Temp Episodes!A2:1000";
        private const string COMBINED_EPISODES_TITLE_RANGE = "Combined Episodes!A2:2";
        private const string COMBINED_EPISODES_DATA_RANGE = "Combined Episodes!A3:1502";
        private const string RECORDED_NAMES_TITLE_RANGE = "Recorded Names!A2:2";
        private const string RECORDED_NAMES_DATA_RANGE = "Recorded Names!A3:1102";
        private const string DB_TITLE_RANGE = "DB!A2:2";
        private const string DB_DATA_RANGE = "DB!A3:502";

        // The following are the column titles for the Movies sheet. (I guess in case I change the column header I don't have to change it in so many places... but I've yet needed this)
        private const string DIRECTORY = "Directory";
        private const string CLEAN_TITLE = "Clean Title";
        private const string ISO_INPUT = "ISO Input";
        private const string ISO_TITLE_NUM = "ISO Title #";
        private const string ISO_CH_NUM = "ISO Ch #";
        private const string QUICK_CREATE = "Quick Create";
        private const string ADDITIONAL_COMMANDS = "Additional Commands";
        private const string NFO_BODY = "NFO Body";
        private const string STATUS = "Status";

        // TMDB API error values.
        const int AUTHENTICATION_FAILED = 3;
        const int INVALID_API_KEY = 7;
        const int DELETED_SUCCESSFULLY = 13;
        const int RESOURCE_CANT_BE_FOUND = 34;

        // Lists for the CountFiles method.
        static List<string> missingNfo = new List<string>();
        static List<string> missingJpg = new List<string>();
        static List<string> missingMovie = new List<string>();
        static List<string> missingIso = new List<string>();
        static List<string> emptyDirectory = new List<string>();
        static List<string> partFiles = new List<string>();

        // The method in which to input the data into the Google Sheet.
        const int INSERT_MISSING_DATA = 1;
        const int OVERWRITE_DATA = 2;

        // Menu item variables.
        static string exitChoice,
            missingMovieNfoFilesChoice,
            overwriteAllMovieNfoFilesChoice,
            selectedMovieNfoFilesChoice,
            missingCombinedEpisodeNfoFilesChoice,
            overwriteAllTvShowNfoFilesChoice,
            selectedTvShowNfoFilesChoice,
            missingYoutubeNfoFilesChoice,
            overwriteAllYoutubeNfoFilesChoice,
            selectedYoutubeNfoFilesChoice,
            missingFitnessVideoNfoFilesChoice,
            overwriteAllFitnessVideoNfoFilesChoice,
            selectedFitnessVideoNfoFilesChoice,
            convertMoviesChoice,
            convertDirectoryChoice,
            convertMoviesSlowChoice,
            convertBonusFeaturesChoice,
            convertBonusFeaturesSlowChoice,
            convertTvShowsChoice,
            convertTvShowsSlowChoice,
            convertTempTvShowsChoice,
            convertTempTVShowsSlowChoice,
            insertMissingMovieDataChoice,
            updateMovieDataChoice,
            insertMissingTmdbIdsChoice,
            insertAndOverwriteTmdbIdsChoice,
            copyMovieFilesToDestinationChoice,
            deleteMovieFilesAtDestinationChoice,
            removeMetadataChoice,
            createFoldersAndMoveFilesChoice,
            trimTitlesInDirectoryChoice,
            bothTrimAndCreateFoldersChoice,
            findSizeOfVideoFilesInDirectoryChoice,
            fetchTvShowPlotsChoice,
            fixRecordedNamesChoice,
            copyMultipleFilesToOneLocationChoice,
            insertMissingDbDataChoice,
            updateDbSheetChoice,
            insertMissingCombinedEpisodesChoice,
            updateCombinedEpisodesChoice,
            writeToCombinedEpisodesChoice,
            insertMissingEpisodeDataChoice;

        static string fileSize;
        static long fileSizeBytes;
        // runningDifference holds the total amount of savings in this run.
        // runningSessionSavings holds the total amount of savings for as long as the current session has been open.
        // totalSessionSavings holds the total size of files this session.
        // runningFileSize holds the size of the original files as we re-encode them.
        // runningSessionFileSize holds the size of the original files as we re-encode them for as long as this current session has been open.
        static long runningDifference = 0, runningSessionSavings = 0, totalSessionSavings = 0,
                runningFileSize = 0, runningSessionFileSize = 0;

        private const int STARTING_ROW_NUMBER = 3;
        static TimeSpan runningTotalConversionTime = new TimeSpan();
        static TimeSpan sessionDuration = new TimeSpan();

        // If modifying these scopes, delete your previously saved credentials
        // at \BachFlixNfo\bin\Debug\token.json\Google.Apis.Auth.OAuth2.Responses.TokenResponse-user
        static readonly string[] SCOPES = { SheetsService.Scope.Spreadsheets };
        static string APLICATION_NAME = "Google Sheets API .NET Quickstart";
        static readonly string SPREADSHEET_ID = "1LE9Tiz0TgcG60qeul_y9wC4j8qNLQlfKTLnAg5tgBr0";

        static void Main(string[] args)
        {
            Type("Welcome to the BachFlix NFO Filer 3000! v1.5", 0, 0, 1, "blue");

            bool keepAskingForChoice = true;

            do
            {
                string[] choices = Menu();

                if (choices.Length > 0)
                {
                    foreach (string choice in choices)
                    {
                        if (choice.Trim() != "")
                        {
                            keepAskingForChoice = CallSwitch(choice);
                            if (!keepAskingForChoice) break;
                        }
                    }
                }
                else
                {
                    Type("You must input an option.", 0, 0, 1, "Red");
                }

                if (keepAskingForChoice) AskForMenu();

            } while (keepAskingForChoice);

        } // End Main

        /// <summary>
        /// Gives the main menu.
        /// </summary>
        private static string[] Menu()
        {
            Type("Please choose from one of the following options..", 0, 0, 1);
            Type("(Or do multiple options by separating them with a comma. i.e. 1,3)", 0, 0, 1);
            exitChoice = "0";
            Type(exitChoice + "- Exit", 0, 0, 1, "darkgray");
            Type("--- NFO File Creation ---", 0, 0, 1, "darkgreen");
            missingMovieNfoFilesChoice = "n1";
            Type(missingMovieNfoFilesChoice + "- Missing Movie NFO Files", 0, 0, 1, "darkgreen");
            overwriteAllMovieNfoFilesChoice = "n1o";
            Type(overwriteAllMovieNfoFilesChoice + "- Overwrite ALL Movie NFO Files", 0, 0, 1, "darkgreen");
            selectedMovieNfoFilesChoice = "n1s";
            Type(selectedMovieNfoFilesChoice + "- Selected Movie NFO Files", 0, 0, 1, "darkgreen");
            missingCombinedEpisodeNfoFilesChoice = "n2";
            Type(missingCombinedEpisodeNfoFilesChoice + "- Create missing NFO files for Combined TV Show episodes.", 0, 0, 1, "darkgreen");
            overwriteAllTvShowNfoFilesChoice = "n2o";
            Type(overwriteAllTvShowNfoFilesChoice + "- Overwrite ALL TV Show NFO Files (Under Construction)", 0, 0, 1, "darkgreen");
            selectedTvShowNfoFilesChoice = "n2s";
            Type(selectedTvShowNfoFilesChoice + "- Selected TV Show NFO Files (Under Construction)", 0, 0, 1, "darkgreen");
            missingYoutubeNfoFilesChoice = "n3";
            Type(missingYoutubeNfoFilesChoice + "- Missing YouTube NFO Files", 0, 0, 1, "darkgreen");
            overwriteAllYoutubeNfoFilesChoice = "n3o";
            Type(overwriteAllYoutubeNfoFilesChoice + "- Overwrite ALL YouTube NFO Files", 0, 0, 1, "darkgreen");
            selectedYoutubeNfoFilesChoice = "n3s";
            Type(selectedYoutubeNfoFilesChoice + "- Selected YouTube NFO Files", 0, 0, 1, "darkgreen");
            missingFitnessVideoNfoFilesChoice = "n4";
            Type(missingFitnessVideoNfoFilesChoice + "- Missing Fitness Video NFO Files", 0, 0, 1, "darkgreen");
            overwriteAllFitnessVideoNfoFilesChoice = "n4o";
            Type(overwriteAllFitnessVideoNfoFilesChoice + "- Overwrite ALL Fitness Video NFO Files", 0, 0, 1, "darkgreen");
            selectedFitnessVideoNfoFilesChoice = "n4s";
            Type(selectedFitnessVideoNfoFilesChoice + "- Selected Fitness Video NFO Files", 0, 0, 1, "darkgreen");

            Type("-- Convert Files ---", 0, 0, 1, "darkcyan");
            convertMoviesChoice = "5";
            Type(convertMoviesChoice + "- Movies", 0, 0, 1, "darkcyan");
            convertMoviesSlowChoice = "5s";
            Type(convertMoviesSlowChoice + "- Movies (Slow)", 0, 0, 1, "darkcyan");
            convertBonusFeaturesChoice = "6";
            Type(convertBonusFeaturesChoice + "- Bonus Features", 0, 0, 1, "darkcyan");
            convertBonusFeaturesSlowChoice = "6s";
            Type(convertBonusFeaturesSlowChoice + "- Bonus Features (Slow)", 0, 0, 1, "darkcyan");
            convertTvShowsChoice = "7";
            Type(convertTvShowsChoice + "- TV Shows", 0, 0, 1, "darkcyan");
            convertTvShowsSlowChoice = "7s";
            Type(convertTvShowsSlowChoice + "- TV Shows (Slow)", 0, 0, 1, "darkcyan");
            convertTempTvShowsChoice = "7t";
            Type(convertTempTvShowsChoice + "- Temp TV Shows", 0, 0, 1, "darkcyan");
            convertTempTVShowsSlowChoice = "7ts";
            Type(convertTempTVShowsSlowChoice + "- Temp TV Shows (Slow)", 0, 0, 1, "darkcyan");
            convertDirectoryChoice = "19";
            Type(convertDirectoryChoice + "- Convert a selected directory.", 0, 0, 1, "darkcyan");

            Type("--- TMDB Call ---", 0, 0, 1, "green"); 
            insertMissingMovieDataChoice = "10";
            Type(insertMissingMovieDataChoice + "- Insert movie data into the Google Sheet (plot, rating, & TMDB ID).", 0, 0, 1, "green");
            updateMovieDataChoice = "10o";
            Type(updateMovieDataChoice + "- Insert and overwrite movie data into the Google Sheet (plot, rating, & TMDB ID).", 0, 0, 1, "green");
            insertMissingTmdbIdsChoice = "11";
            Type(insertMissingTmdbIdsChoice + "- Insert missing TMDB IDs into the Google Sheet.", 0, 0, 1, "green");
            insertAndOverwriteTmdbIdsChoice = "11o";
            Type(insertAndOverwriteTmdbIdsChoice + "- Insert and override TMDB IDs in the Google Sheet.", 0, 0, 1, "green");
            fetchTvShowPlotsChoice = "25";
            Type(fetchTvShowPlotsChoice + "- Insert TV Show plots into the Combined Episodes sheet.", 0, 0, 1, "green");

            Type("--- Update Google Sheet ---", 0, 0, 1, "Magenta");
            insertMissingDbDataChoice = "28";
            Type(insertMissingDbDataChoice + "- Insert missing data into the DB sheet from TVDB.", 0, 0, 1, "Magenta");
            updateDbSheetChoice = "29";
            Type(updateDbSheetChoice + "- Update the DB sheet with updated info.", 0, 0, 1, "Magenta");
            insertMissingCombinedEpisodesChoice = "30";
            Type(insertMissingCombinedEpisodesChoice + "- Insert missing data into the Combined Episodes sheet from TVDB.", 0, 0, 1, "Magenta");
            updateCombinedEpisodesChoice = "31";
            Type(updateCombinedEpisodesChoice + "- Update data in the Combined Episodes sheet from TVDB.", 0, 0, 1, "Magenta");
            writeToCombinedEpisodesChoice = "32";
            Type(writeToCombinedEpisodesChoice + "- Write video file names to the Combined Episodes sheet.", 0, 0, 1, "Magenta");
            insertMissingEpisodeDataChoice = "33";
            Type(insertMissingEpisodeDataChoice + "- Write episode names to the Episodes sheet.", 0, 0, 1, "Magenta");

            Type("--- Misc. ---", 0, 0, 1, "darkyellow");
            Type("8- Count Files", 0, 0, 1, "darkyellow");
            Type("9- Remove the UPC numbers from the folder name.", 0, 0, 1, "darkyellow");
            Type("12- Move Kids movies.", 0, 0, 1, "darkyellow");
            Type("13- Copy JPG files. (Work in progress)", 0, 0, 1, "darkyellow");
            copyMovieFilesToDestinationChoice = "14c";
            Type(copyMovieFilesToDestinationChoice + "- Copy Movie files from Google Sheet list to chosen hard drive.", 0, 0, 1, "darkyellow");
            deleteMovieFilesAtDestinationChoice = "14d";
            Type(deleteMovieFilesAtDestinationChoice + "- Delete Movie files from Google Sheet list to chosen hard drive.", 0, 0, 1, "darkyellow");
            Type("15- Mark Owned Movies as D=Done || X=Not Done.", 0, 0, 1, "darkyellow");
            Type("16- Remove movies from TMDB List. (Work in progress)", 0, 0, 1, "darkyellow");
            Type("17- Move Movies to new rating directory.", 0, 0, 1, "darkyellow");
            removeMetadataChoice = "18";
            Type(removeMetadataChoice + "- Remove Metadata.", 0, 0, 1, "darkyellow");
            Type("20- Add Comment to file.", 0, 0, 1, "darkyellow");
            createFoldersAndMoveFilesChoice = "21";
            Type(createFoldersAndMoveFilesChoice + "- Create directories and move files into them.", 0, 0, 1, "darkyellow");
            trimTitlesInDirectoryChoice = "22";
            Type(trimTitlesInDirectoryChoice + "- Trim titles in chosen directory.", 0, 0, 1, "darkyellow");
            bothTrimAndCreateFoldersChoice = "23";
            Type(bothTrimAndCreateFoldersChoice + "- Trim the titles AND create directories then move files into directories.", 0, 0, 1, "darkyellow");
            findSizeOfVideoFilesInDirectoryChoice = "24";
            Type(findSizeOfVideoFilesInDirectoryChoice + "- Give the size of video files in a directory.", 0, 0, 1, "darkyellow");
            fixRecordedNamesChoice = "26";
            Type(fixRecordedNamesChoice + "- Fix recorded names.", 0, 0, 1, "darkyellow");
            copyMultipleFilesToOneLocationChoice = "27";
            Type(copyMultipleFilesToOneLocationChoice + "- Copy multiple chosen files to one location.", 0, 0, 1, "darkyellow");

            return Console.ReadLine().Split(',');

        } // End Menu()
        static bool CallSwitch(string choice)
        {
            bool keepAskingForChoice = true;
            try
            {
                Dictionary<string, int> sheetVariables = new Dictionary<string, int> { };
                string titleRowDataRange = "", mainDataRange = "";
                int type = 0;

                if (choice.Trim().Equals(exitChoice))
                {
                    Type("Thank you, have a nice day! \\(^.^)/", 7, 100, 1);
                    keepAskingForChoice = false;

                }
                else if (choice.Trim().Equals(missingMovieNfoFilesChoice)) // NFO files for New Movies - does not overwrite any, just puts in missing NFO files.
                {
                    Type("Insert missing NFO Files. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add(DIRECTORY, -1); // The path to the folder holding the movie.
                    sheetVariables.Add(CLEAN_TITLE, -1); // Concatenate the Clean Title to the Directory to save the NFO File.
                    sheetVariables.Add(NFO_BODY, -1); // The text of the NFO File to save.
                    sheetVariables.Add(STATUS, -1); // The Status of the movie i.e. if the movie should actually be there.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    type = 3;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    CreateNfoFiles(movieData, sheetVariables, type);

                }
                else if (choice.Trim().Equals(overwriteAllMovieNfoFilesChoice)) // NFO files for All Movies - overwrite old NFO files AND put in new ones.
                {
                    Type("Insert missing AND overwrite NFO Files. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add(DIRECTORY, -1); // The path to the folder holding the movie.
                    sheetVariables.Add(CLEAN_TITLE, -1); // Concatenate the Clean Title to the Directory to save the NFO File.
                    sheetVariables.Add(NFO_BODY, -1); // The text of the NFO File to save.
                    sheetVariables.Add(STATUS, -1); // The Status of the movie i.e. if the movie should actually be there.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    type = 1;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    CreateNfoFiles(movieData, sheetVariables, type);
                }
                else if (choice.Trim().Equals(selectedMovieNfoFilesChoice)) // NFO files for Selected Movies - overwrite or put in new ones. if they are selected.
                {
                    Type("Insert selected NFO Files. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add(DIRECTORY, -1); // The path to the folder holding the movie.
                    sheetVariables.Add(CLEAN_TITLE, -1); // Concatenate the Clean Title to the Directory to save the NFO File.
                    sheetVariables.Add(NFO_BODY, -1); // The text of the NFO File to save.
                    sheetVariables.Add(STATUS, -1); // The Status of the movie i.e. if the movie should actually be there.
                    sheetVariables.Add(QUICK_CREATE, -1); // If this column has an 'X' then we write/overwrite the file.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    type = 2;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    CreateNfoFiles(movieData, sheetVariables, type);
                }
                else if (choice.Trim().Equals(missingCombinedEpisodeNfoFilesChoice)) // Create missing NFO files for TV Show episodes.
                {
                    Type("Create missing NFO files for TV Show episodes. Let's go!", 7, 100, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {// A dictionary to hold the columns we need to find.
                        sheetVariables.Add("TVDB ID", -1); // The ID from the TVDB.
                        sheetVariables.Add("Combined Episode Name", -1); // The original combined episode name that needs to be changed.
                        sheetVariables.Add("New Episode Name", -1); // The name to change the episode to.
                        sheetVariables.Add("NFO Body", -1); // The text of the NFO file.

                        titleRowDataRange = COMBINED_EPISODES_TITLE_RANGE;
                        mainDataRange = COMBINED_EPISODES_DATA_RANGE;

                        IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                        CreateMissingCombinedEpisodeNfoFiles(movieData, sheetVariables, directory);
                    }
                }
                else if (choice.Trim().Equals(missingYoutubeNfoFilesChoice)) // NFO files for New videos - does not overwrite any, just puts in missing NFO files.
                {
                    Type("Create missing YouTube NFO Files. Let's go!", 7, 100, 1);

                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add(DIRECTORY, -1); // The path to the folder holding the movie.
                    sheetVariables.Add(CLEAN_TITLE, -1); // Concatenate the Clean Title to the Directory to save the NFO File.
                    sheetVariables.Add(NFO_BODY, -1); // The text of the NFO File to save.
                    sheetVariables.Add(STATUS, -1); // The Status of the movie i.e. if the movie should actually be there.

                    titleRowDataRange = YOUTUBE_TITLE_RANGE;
                    mainDataRange = YOUTUBE_DATA_RANGE;

                    type = 3;

                    IList<IList<Object>> videoData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    CreateNfoFiles(videoData, sheetVariables, type, true);
                }
                else if (choice.Trim().Equals(overwriteAllYoutubeNfoFilesChoice)) // NFO files for All videos - overwrite old NFO files AND put in new ones.
                {
                    Type("Overwrite ALL YouTube NFO Files. Let's go!", 7, 100, 1);

                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add(DIRECTORY, -1); // The path to the folder holding the movie.
                    sheetVariables.Add(CLEAN_TITLE, -1); // Concatenate the Clean Title to the Directory to save the NFO File.
                    sheetVariables.Add(NFO_BODY, -1); // The text of the NFO File to save.
                    sheetVariables.Add(STATUS, -1); // The Status of the movie i.e. if the movie should actually be there.

                    titleRowDataRange = YOUTUBE_TITLE_RANGE;
                    mainDataRange = YOUTUBE_DATA_RANGE;

                    type = 1;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    CreateNfoFiles(movieData, sheetVariables, type, true);
                }
                else if (choice.Trim().Equals(selectedYoutubeNfoFilesChoice)) // NFO files for Selected videos - overwrite or put in new ones. if they are selected.
                {
                    Type("Create/Overwrite selected YouTube NFO Files. Let's go!", 7, 100, 1);

                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add(DIRECTORY, -1); // The path to the folder holding the movie.
                    sheetVariables.Add(CLEAN_TITLE, -1); // Concatenate the Clean Title to the Directory to save the NFO File.
                    sheetVariables.Add(NFO_BODY, -1); // The text of the NFO File to save.
                    sheetVariables.Add(STATUS, -1); // The Status of the movie i.e. if the movie should actually be there.
                    sheetVariables.Add(QUICK_CREATE, -1); // Create/Overwrite selected NFO files.

                    titleRowDataRange = YOUTUBE_TITLE_RANGE;
                    mainDataRange = YOUTUBE_DATA_RANGE;

                    type = 2;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    CreateNfoFiles(movieData, sheetVariables, type, true);
                }
                else if (choice.Trim().Equals(missingFitnessVideoNfoFilesChoice)) // NFO files for New videos - does not overwrite any, just puts in missing NFO files.
                {
                    Type("This method is still in the works, please try another one.", 7, 100, 1, "Yellow");
                    //Type("Create missing Fitness Video NFO Files. Let's go!", 7, 100, 1, "Blue");

                    //// A dictionary to hold the columns we need to find.
                    //sheetVariables.Add("Program", -1);
                    //sheetVariables.Add("Subfolder", -1);
                    //sheetVariables.Add("Name", -1);
                    //sheetVariables.Add("Title", -1);
                    //sheetVariables.Add("NFO Body", -1);

                    //titleRowDataRange = FITNESS_VIDEO_TITLE_RANGE;
                    //mainDataRange = FITNESS_VIDEO_DATA_RANGE;

                    //IList<IList<Object>> videoData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    //BachFlixNfo.MissingFitnessVideoNfoFiles(videoData, sheetVariables);
                }
                else if (choice.Trim().Equals(overwriteAllFitnessVideoNfoFilesChoice)) // NFO files for All videos - overwrite old NFO files AND put in new ones.
                {
                    Type("Overwrite ALL YouTube NFO Files. Let's go!", 7, 100, 1, "Blue");

                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("Program", -1);
                    sheetVariables.Add("Subfolder", -1);
                    sheetVariables.Add("Name", -1);
                    sheetVariables.Add("Title", -1);
                    sheetVariables.Add("NFO Body", -1);

                    titleRowDataRange = FITNESS_VIDEO_TITLE_RANGE;
                    mainDataRange = FITNESS_VIDEO_DATA_RANGE;

                    IList<IList<Object>> videoData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    BachFlixNfo.OverwriteFitnessVideoNfoFiles(videoData, sheetVariables);
                }
                //else if (choice.Trim().Equals(selectedYoutubeNfoFilesChoice)) // NFO files for Selected videos - overwrite or put in new ones. if they are selected.
                //{
                //    Type("Create/Overwrite selected YouTube NFO Files. Let's go!", 7, 100, 1);

                //    // A dictionary to hold the columns we need to find.
                //    sheetVariables.Add(DIRECTORY, -1); // The path to the folder holding the movie.
                //    sheetVariables.Add(CLEAN_TITLE, -1); // Concatenate the Clean Title to the Directory to save the NFO File.
                //    sheetVariables.Add(NFO_BODY, -1); // The text of the NFO File to save.
                //    sheetVariables.Add(STATUS, -1); // The Status of the movie i.e. if the movie should actually be there.
                //    sheetVariables.Add(QUICK_CREATE, -1); // Create/Overwrite selected NFO files.

                //    titleRowDataRange = YOUTUBE_TITLE_RANGE;
                //    mainDataRange = YOUTUBE_DATA_RANGE;

                //    type = 2;

                //    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                //    CreateNfoFiles(movieData, sheetVariables, type, true);
                //}
            else if (choice.Trim().Equals(convertMoviesChoice)) // Convert movies the fast cheap way.
                {
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("Directory", -1); // The path to the folder holding the video.
                    sheetVariables.Add("Clean Title", -1); // Concatenate the Clean Title to the Directory.
                    sheetVariables.Add("ISO Input", -1); // The path to the ISO file.
                    sheetVariables.Add("ISO Title #", -1); // The number of the ISO title to use.
                    sheetVariables.Add("ISO Ch #", -1); // The number of the ISO chapter to use.
                    sheetVariables.Add("Quick Create", -1); // Convert the selected files.
                    sheetVariables.Add("Additional Commands", -1); // Add any additional commands to the convert process.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);
                    ConvertVideo(movieData, sheetVariables, "--preset-import-file MP4_RF22f.json -Z \"MP4 RF22f\"");
                }
                else if (choice.Trim().Equals(convertMoviesSlowChoice)) // Convert movies the LONG slow way.
                {
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("Directory", -1); // The path to the folder holding the video.
                    sheetVariables.Add("Clean Title", -1); // Concatenate the Clean Title to the Directory.
                    sheetVariables.Add("ISO Input", -1); // The path to the ISO file.
                    sheetVariables.Add("ISO Title #", -1); // The number of the ISO title to use.
                    sheetVariables.Add("ISO Ch #", -1); // The number of the ISO chapter to use.
                    sheetVariables.Add("Quick Create", -1); // Convert the selected files.
                    sheetVariables.Add("Additional Commands", -1); // Add any additional commands to the convert process.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);
                    ConvertVideo(movieData, sheetVariables, "--preset-import-file MP4_RF22s.json -Z \"MP4 RF22s\"");
                }
                else if (choice.Trim().Equals(convertTvShowsChoice)) // Convert TV Shows the fast cheap way.
                {
                    DisplayMessage("info", "Convert TV Show Episodes. Let's go!");
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("Directory", -1); // The path to the folder holding the video.
                    sheetVariables.Add("Show", -1); // The show title.
                    sheetVariables.Add("Season #", -1); // The Season #.
                    sheetVariables.Add("Clean Title", -1); // Concatenate the Clean Title to the Directory.
                    sheetVariables.Add("ISO Input", -1); // The path to the ISO file.
                    sheetVariables.Add("ISO Title #", -1); // The number of the ISO title to use.
                    sheetVariables.Add("ISO Ch #", -1); // The number of the ISO chapter to use.
                    sheetVariables.Add("Quick Create", -1); // Convert the selected files.
                    sheetVariables.Add("Additional Commands", -1); // Add any additional commands to the convert process.
                    sheetVariables.Add("Override Show", -1); // Grab this incase the show title nedds to be overwritten.

                    titleRowDataRange = EPISODES_TITLE_RANGE;
                    mainDataRange = EPISODES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);
                    ConvertVideo(movieData, sheetVariables, "--preset-import-file MP4_RF22f.json -Z \"MP4 RF22f\"");
                }
                else if (choice.Trim().Equals(convertDirectoryChoice)) // Convert a directory.
                {
                    DisplayMessage("info", "Convert a directory. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        bool keepLooping = true;
                        do
                        {
                            // Grab all files in the directory.
                            DisplayMessage("warning", "Grabbing all files... ", 0);
                            string[] fileEntries = Directory.GetFiles(directory);
                            DisplayMessage("success", "DONE");

                            // Filter out the files that aren't video files.
                            ArrayList videoFiles = GrabMovieFiles(fileEntries);

                            fileSizeBytes = SizeOfFiles(videoFiles);
                            totalSessionSavings += fileSizeBytes;

                            fileSize = FormatSize(fileSizeBytes);


                            string plural = videoFiles.Count == 1 ? " file " : " files ";

                            DisplayMessage("info", "The size of the " + videoFiles.Count + plural + "is: ", 0);
                            DisplayMessage("data", fileSize);

                            // Send those video files off to be converted.
                            ConvertHandbrakeList(videoFiles);

                            ResetGlobals();

                            // Check for more files.
                            DisplayMessage("warning", "Checking for more files... ");
                            fileEntries = Directory.GetFiles(directory);
                            ArrayList videoFilesCheck = GrabMovieFiles(fileEntries);
                            if (videoFilesCheck.Count == 0) keepLooping = false;
                            else DisplayMessage("info", "More files found. Restarting conversion...");
                        } while (keepLooping);
                    }

                }
                else if (choice.Trim().Equals(findSizeOfVideoFilesInDirectoryChoice)) // Find the size of video files in a directory.
                {
                    Type("Find the size of video files in a directory. Let's go!", 7, 100, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        // Grab all files in the directory.
                        Type("Grabbing all files... ", 10, 0, 0, "Yellow");
                        string[] fileEntries = Directory.GetFiles(directory);
                        Type("DONE", 100, 0, 1, "Green");

                        // Filter out the files that aren't video files.
                        ArrayList videoFiles = GrabMovieFiles(fileEntries);

                        long sizeOfFiles = SizeOfFiles(videoFiles);

                        string sizeInText = FormatSize(sizeOfFiles);

                        string plural = videoFiles.Count == 1 ? " file " : " files ";

                        Type("The size of the " + videoFiles.Count + plural + "is: ", 0, 0, 0, "Blue");
                        Type(sizeInText, 0, 0, 1, "Cyan");
                    }

                }
                else if (choice.Trim().Equals(fetchTvShowPlotsChoice)) // Fetch TV Show episode plots from TVDB.
                {
                    Type("Gather the TV Show episode plots from TVDB. Let's go!", 7, 100, 1);

                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("Combined Episode Name", -1);
                    sheetVariables.Add("TMDB ID", -1);
                    sheetVariables.Add("Episode 1 Title", -1);
                    sheetVariables.Add("Episode 2 Title", -1);
                    sheetVariables.Add("Episode 1 Season", -1);
                    sheetVariables.Add("Episode 1 No.", -1);
                    sheetVariables.Add("Episode 2 Season", -1);
                    sheetVariables.Add("Episode 2 No.", -1);
                    sheetVariables.Add("Episode 1 Plot", -1);
                    sheetVariables.Add("Episode 2 Plot", -1);

                    titleRowDataRange = COMBINED_EPISODES_TITLE_RANGE;
                    mainDataRange = COMBINED_EPISODES_DATA_RANGE;

                    IList<IList<Object>> videoData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    BachFlixNfo.InputTvShowPlots(videoData, sheetVariables);

                }
                else if (choice.Trim().Equals(fixRecordedNamesChoice)) // Fix recorded names.
                {
                    DisplayMessage("info", "Let's go!", 7, 100, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        // A dictionary to hold the columns we need to find.
                        sheetVariables.Add("Recorded Name", -1);
                        sheetVariables.Add("Actual Name", -1);

                        titleRowDataRange = RECORDED_NAMES_TITLE_RANGE;
                        mainDataRange = RECORDED_NAMES_DATA_RANGE;

                        IList<IList<Object>> videoData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                        BachFlixNfo.FixRecordedNames(videoData, sheetVariables, directory);
                    }

                }
                else if (choice.Trim().Equals(copyMultipleFilesToOneLocationChoice)) // Copy an array of files to one location.
                {
                    DisplayMessage("info", "Let's go!", 2);

                    ArrayList videoFiles = AskForFilesToCopy();

                    var directory = AskForDirectory("Now, where are we copying these files to?");

                    if (directory != "0")
                    {
                        try
                        {
                            DisplayMessage("warning", "We will now copy each file...");

                            foreach (var myFile in videoFiles)
                            {
                                DisplayMessage("info", "Copying ", 0);
                                string fileName = Path.GetFileName(myFile.ToString());
                                DisplayMessage("default", fileName + "... ", 0);
                                File.Copy(myFile.ToString(), Path.Combine(directory, fileName));
                                DisplayMessage("success", "DONE");
                            }
                        }
                        catch (Exception e)
                        {
                            DisplayMessage("error", "An error occured copying the files.");
                            DisplayMessage("harderror", e.Message);
                            throw;
                        }
                    }

                }
                else if (choice.Trim().Equals("20")) // Add a comment to a directory.
                {
                    Type("Add a comment to a directory. Let's go!", 7, 100, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        string[] fileEntries = Directory.GetFiles(directory);

                        ArrayList videoFiles = GrabMovieFiles(fileEntries);

                        foreach (var myFile in videoFiles)
                        {
                            DateTime convertedTime = DateTime.Now;

                            AddComment(myFile.ToString(), "Recorded in HD, re-encoded with black bars.\nConverted on: " + convertedTime.ToString("MM/dd/yyyy"));
                        }
                    }
                }
                else if (choice.Trim().Equals(createFoldersAndMoveFilesChoice)) // Create directories that match the names of files in a directory, then move those files into their respective directories.
                {
                    Type("Create directories then move files into directories. Let's go!", 7, 100, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        CreateFoldersAndMoveFiles(directory);
                    }
                }
                else if (choice.Trim().Equals(trimTitlesInDirectoryChoice))
                {
                    Type("Trim the titles in a chosen directory. Let's go!", 7, 100, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        TrimTitlesInDirectory(directory);
                    }
                }
                else if (choice.Trim().Equals(bothTrimAndCreateFoldersChoice))
                {
                    Type("Trim the titles AND create directories then move files into directories. Let's go!", 7, 100, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        TrimTitlesInDirectory(directory);
                        CreateFoldersAndMoveFiles(directory);
                    }
                }
                else if (choice.Trim().Equals(convertTvShowsSlowChoice)) // Convert TV Shows the LONG slow way.
                {
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("Directory", -1); // The path to the folder holding the video.
                    sheetVariables.Add("Clean Title", -1); // Concatenate the Clean Title to the Directory.
                    sheetVariables.Add("ISO Input", -1); // The path to the ISO file.
                    sheetVariables.Add("ISO Title #", -1); // The number of the ISO title to use.
                    sheetVariables.Add("ISO Ch #", -1); // The number of the ISO chapter to use.
                    sheetVariables.Add("Quick Create", -1); // Convert the selected files.
                    sheetVariables.Add("Additional Commands", -1); // Add any additional commands to the convert process.

                    titleRowDataRange = EPISODES_TITLE_RANGE;
                    mainDataRange = EPISODES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);
                    ConvertVideo(movieData, sheetVariables, "--preset-import-file MP4_RF22s.json -Z \"MP4 RF22s\"");
                }
                else if (choice.Trim().Equals(insertMissingMovieDataChoice))
                {
                    Type("Insert missing movie data into the Google Sheet. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("RowNum", -1); // Simply the row number we are on, to better help writing back to the Sheet.
                    sheetVariables.Add("TMDB ID", -1); // The TMDB ID to input into the sheet.
                    sheetVariables.Add("TMDB Rating", -1); // The TMDB Rating to input into the sheet.
                    sheetVariables.Add("Plot", -1); // The movie plot to input into the sheet.
                    sheetVariables.Add("IMDB ID", -1); // The IMDB ID to grab the movie data from the TMDB API.
                    sheetVariables.Add("IMDB Title", -1); // Used to print out if the write was successfull or not.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    DisplayMessage("warning", "Looking through sheet data for missing data... ");
                    InputMovieData(movieData, sheetVariables);
                    DisplayMessage("success", "Done");

                }
                else if (choice.Trim().Equals(updateMovieDataChoice))
                {
                    Type("Insert missing and update movie data into the Google Sheet. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("RowNum", -1); // Simply the row number we are on, to better help writing back to the Sheet.
                    sheetVariables.Add("TMDB ID", -1); // The TMDB ID to input into the sheet.
                    sheetVariables.Add("TMDB Rating", -1); // The TMDB Rating to input into the sheet.
                    sheetVariables.Add("Plot", -1); // The movie plot to input into the sheet.
                    sheetVariables.Add("IMDB ID", -1); // The IMDB ID to grab the movie data from the TMDB API.
                    sheetVariables.Add("IMDB Title", -1); // Used to print out if the write was successfull or not.
                    sheetVariables.Add("Quick Create", -1); // Used to mark which movie plot was updated and needs the NFO file updated.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    DisplayMessage("warning", "Looking through sheet data to update data... ");
                    InputMovieData(movieData, sheetVariables, true);
                    DisplayMessage("success", "Done");

                }
                else if (choice.Trim().Equals(insertMissingTmdbIdsChoice)) // Insert TMDB IDs into the Google Sheet.
                {
                    Type("Insert missing TMDB IDs into the Google Sheet. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("TMDB ID", -1); // The TMDB ID to input into the sheet.
                    sheetVariables.Add("IMDB ID", -1); // The IMDB ID to grab the movie data from the TMDB API.
                    sheetVariables.Add("IMDB Title", -1); // Used to print out if the write was successfull or not.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    InputTmdbId(movieData, sheetVariables, 1);

                }
                else if (choice.Trim().Equals(insertAndOverwriteTmdbIdsChoice)) // Insert TMDB IDs into the Google Sheet.
                {
                    Type("Insert missing AND overwrite TMDB IDs into the Google Sheet. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("TMDB ID", -1); // The TMDB ID to input into the sheet.
                    sheetVariables.Add("IMDB ID", -1); // The IMDB ID to grab the movie data from the TMDB API.
                    sheetVariables.Add("IMDB Title", -1); // Used to print out if the write was successfull or not.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    InputTmdbId(movieData, sheetVariables, 2);

                }
                else if (choice.Trim().Equals("12")) // Move kids movies.
                {
                    Type("Move kids movies. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("Kids", -1); // Check if movie is marked with an "X" to move accordingly.
                    sheetVariables.Add("Directory", -1); // The location of the movie's directory.
                    sheetVariables.Add("Status", -1); // If the first character is an "X" then we don't need to worry about looking for the movie.
                    sheetVariables.Add("Movie Letter", -1); // Use this to replace \Kids Movies\ and vice versa.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    MoveKidsMovies(movieData, sheetVariables);

                }
                else if (choice.Trim().Equals(copyMovieFilesToDestinationChoice)) // Copy movies.
                {
                    Type("Copy movies. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("Cindy", -1); // Check if movie is marked with an "y" to move accordingly.
                    sheetVariables.Add("Dave", -1); // Check if movie is marked with an "y" to move accordingly.
                    sheetVariables.Add("Directory", -1); // The location of the movie's directory.
                    sheetVariables.Add("Clean Title", -1); // Concatenate the Clean Title to the Directory to save the NFO File.
                    sheetVariables.Add("Status", -1); // If the first character is an "X" then we don't need to worry about looking for the movie.
                    sheetVariables.Add("Movie Letter", -1); // The movie letter where the movie will reside on the hard drive we are copying to.e.
                    sheetVariables.Add("Kids", -1); // The movie letter where the movie will reside on the hard drive we are copying to.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    CopyMovieFiles(movieData, sheetVariables);

                }
                else if (choice.Trim().Equals(deleteMovieFilesAtDestinationChoice)) // Delete movies from moms hard drive.
                {
                    Type("Delete movies. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("Cindy", -1); // Check if movie is marked with an "y" to move accordingly.
                    sheetVariables.Add("Dave", -1); // Check if movie is marked with an "y" to move accordingly.
                    sheetVariables.Add("Directory", -1); // The location of the movie's directory.
                    sheetVariables.Add("Clean Title", -1); // Concatenate the Clean Title to the Directory to save the NFO File.
                    sheetVariables.Add("Status", -1); // If the first character is an "X" then we don't need to worry about looking for the movie.
                    sheetVariables.Add("Movie Letter", -1); // The movie letter where the movie will reside on the hard drive we are copying to.e.
                    sheetVariables.Add("Kids", -1); // The movie letter where the movie will reside on the hard drive we are copying to.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    DeleteMovieFiles(movieData, sheetVariables);

                }
                else if (choice.Trim().Equals("16")) // Remove movies from TMDB List.
                {
                    Type("Remove movies from TMDB List. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("TMDB ID", -1); // The ID to check against the list.
                    sheetVariables.Add("Clean Title", -1); // To display the movie we are working with.
                    sheetVariables.Add("Status", -1); // If the movie is marked as done then we can look for the movie in our list and remove it.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    RemoveMoviesFromTmdbList(movieData, sheetVariables);

                }
                else if (choice.Trim().Equals("44"))
                {
                    Type("Getting authorization.", 0, 0, 1, "Blue");
                    dynamic tmdbResponse = TmdbApi.AuthenticationCreateRequestToken();

                    var requestToken = tmdbResponse.request_token.ToString();

                    tmdbResponse = TmdbApi.AuthenticationSendRequestToken(requestToken);

                    //Type(tmdbResponse.request_token.ToString(), 0, 0, 1);
                    Type("Authorization received.", 0, 0, 1, "Green");
                }
                else if (choice.Trim().Equals("17"))
                {
                    Type("Move Movies into new rating directory layout. Let's go!", 7, 100, 1);
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("Directory", -1); // The path to the new location.
                    sheetVariables.Add("Old Directory", -1); // The path to the old location.

                    titleRowDataRange = MOVIES_TITLE_RANGE;
                    mainDataRange = MOVIES_DATA_RANGE;

                    IList<IList<Object>> movieData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    MoveMovies(movieData, sheetVariables);

                }
                else if (choice.Trim().Equals(removeMetadataChoice))
                {
                    Type("Remove Metadata. Let's go!", 7, 100, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        string[] fileEntries = Directory.GetFiles(directory);

                        ArrayList videoFiles = GrabMovieFiles(fileEntries);

                        RemoveMetadata(videoFiles);
                    }
                }
                else if (choice.Trim().Equals(insertMissingDbDataChoice))
                {
                    DisplayMessage("info", "Insert missing data into the DB sheet from TVDB. Let's go!");
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("RowNum", -1); // Simply the row number we are on, to better help writing back to the Sheet.
                    sheetVariables.Add("Series Name", -1); // The name of the series to populate if empty.
                    sheetVariables.Add("Season Count", -1); // The number of seasons for the show.
                    sheetVariables.Add("Continuing?", -1); // If the show is continuing or not.
                    sheetVariables.Add("TVDB Slug", -1); // The show slug to add to the url call.
                    sheetVariables.Add("TVDB ID", -1); // The ID of the TV Show to use for gathering episode data.

                    var jwtToken = TvdbApiCall.TvdbApi.GetTvdbJwtKey();

                    titleRowDataRange = DB_TITLE_RANGE;
                    mainDataRange = DB_DATA_RANGE;

                    IList<IList<Object>> dbData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    InsertMissingDbData(dbData, sheetVariables, jwtToken);

                }
                else if (choice.Trim().Equals(updateDbSheetChoice))
                {
                    DisplayMessage("info", "Update DB sheet info. Let's go!");
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("RowNum", -1); // Simply the row number we are on, to better help writing back to the Sheet.
                    sheetVariables.Add("Series Name", -1); // The name of the series to populate if empty.
                    sheetVariables.Add("Season Count", -1); // The number of seasons for the show.
                    sheetVariables.Add("Continuing?", -1); // If the show is continuing or not.
                    sheetVariables.Add("TVDB Slug", -1); // The show slug to add to the url call.
                    sheetVariables.Add("TVDB ID", -1); // The ID of the TV Show to use for gathering episode data.

                    var jwtToken = TvdbApiCall.TvdbApi.GetTvdbJwtKey();

                    titleRowDataRange = DB_TITLE_RANGE;
                    mainDataRange = DB_DATA_RANGE;

                    IList<IList<Object>> dbData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    UpdateDbData(dbData, sheetVariables, jwtToken);
                }
                else if (choice.Trim().Equals(insertMissingCombinedEpisodesChoice))
                {
                    DisplayMessage("info", "Insert missing data into the Combined Episodes sheet from TVDB. Let's go!");
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("RowNum", -1); // Simply the row number we are on, to better help writing back to the Sheet.
                    sheetVariables.Add("TVDB ID", -1); // The ID of the TV Show to use for gathering episode data.
                    sheetVariables.Add("Episode 1 Season", -1); // The season number for the first episode.
                    sheetVariables.Add("Episode 1 No.", -1); // The episode number for the first episode.
                    sheetVariables.Add("Episode 1 Plot", -1); // The plot for the first episode.
                    sheetVariables.Add("Episode 2 Season", -1); // The season number for the second episode.
                    sheetVariables.Add("Episode 2 No.", -1); // The episode number for the second episode.
                    sheetVariables.Add("Episode 2 Plot", -1); // The plot for the second episode.
                    sheetVariables.Add("Show Title", -1); // Simply the title to the show.

                    var jwtToken = TvdbApiCall.TvdbApi.GetTvdbJwtKey();

                    titleRowDataRange = COMBINED_EPISODES_TITLE_RANGE;
                    mainDataRange = COMBINED_EPISODES_DATA_RANGE;

                    IList<IList<Object>> dbData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    InsertMissingCombinedEpisodeData(dbData, sheetVariables, jwtToken);
                }
                else if (choice.Trim().Equals(insertMissingEpisodeDataChoice))
                {
                    DisplayMessage("info", "Insert missing data into the Episodes sheet from TVDB. Let's go!");
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("RowNum", -1); // Simply the row number we are on, to better help writing back to the Sheet.
                    sheetVariables.Add("TVDB ID", -1); // The ID of the TV Show to use for gathering episode data.
                    sheetVariables.Add("Season #", -1); // The season number for the episode.
                    sheetVariables.Add("Episode #", -1); // The episode number for the episode.
                    sheetVariables.Add("Episode Name", -1); // The name of the episode.
                    sheetVariables.Add("Plot", -1); // The plot of the episode.
                    sheetVariables.Add("Show", -1); // The name of the series.

                    var jwtToken = TvdbApiCall.TvdbApi.GetTvdbJwtKey();

                    titleRowDataRange = EPISODES_TITLE_RANGE;
                    mainDataRange = EPISODES_DATA_RANGE;

                    IList<IList<Object>> dbData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    InsertMissingEpisodeData(dbData, sheetVariables, jwtToken);
                }
                else if (choice.Trim().Equals(updateCombinedEpisodesChoice))
                {
                    DisplayMessage("info", "Update data in the Combined Episodes sheet from TVDB. Let's go!");
                    // A dictionary to hold the columns we need to find.
                    sheetVariables.Add("RowNum", -1); // Simply the row number we are on, to better help writing back to the Sheet.
                    sheetVariables.Add("TVDB ID", -1); // The ID of the TV Show to use for gathering episode data.
                    sheetVariables.Add("Episode 1 Season", -1); // The season number for the first episode.
                    sheetVariables.Add("Episode 1 No.", -1); // The episode number for the first episode.
                    sheetVariables.Add("Episode 1 Plot", -1); // The plot for the first episode.
                    sheetVariables.Add("Lock Plot 1", -1); // Marks if the plot for the first episode is locked (i.e. I manually changed the TVDB plot and don't want it overwritten).
                    sheetVariables.Add("Episode 2 Season", -1); // The season number for the second episode.
                    sheetVariables.Add("Episode 2 No.", -1); // The episode number for the second episode.
                    sheetVariables.Add("Episode 2 Plot", -1); // The plot for the second episode.
                    sheetVariables.Add("Lock Plot 2", -1); // Marks if the plot for the second episode is locked (i.e. I manually changed the TVDB plot and don't want it overwritten).
                    sheetVariables.Add("Show Title", -1); // Simply the title to the show.

                    var jwtToken = TvdbApiCall.TvdbApi.GetTvdbJwtKey();

                    titleRowDataRange = COMBINED_EPISODES_TITLE_RANGE;
                    mainDataRange = COMBINED_EPISODES_DATA_RANGE;

                    IList<IList<Object>> dbData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                    UpdateCombinedEpisodeData(dbData, sheetVariables, jwtToken);
                }
                else if (choice.Trim().Equals(writeToCombinedEpisodesChoice))
                {

                    DisplayMessage("info", "Add video file names to the Combined Episodes sheet. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        // A dictionary to hold the columns we need to find.
                        sheetVariables.Add("RowNum", -1); // Simply the row number we are on, to better help writing back to the Sheet.
                        sheetVariables.Add("Combined Episode Name", -1); // The name of the file when recorded.

                        string[] fileEntries = Directory.GetFiles(directory); // Gather ALL files from the directory.

                        ArrayList videoFiles = GrabMovieFiles(fileEntries); // Now filter out anything that isn't a video file.

                        titleRowDataRange = COMBINED_EPISODES_TITLE_RANGE;
                        mainDataRange = COMBINED_EPISODES_DATA_RANGE;

                        IList<IList<Object>> sheetData = CallGetData(sheetVariables, titleRowDataRange, mainDataRange);

                        WriteToSheetColumn(videoFiles, sheetData, "Combined Episodes", 1);
                    }
                }

                switch (choice.Trim())
                {
                    case "8": // Count files.
                        DisplayMessage("info", "Count the files. Let's go!");
                        CountFiles();
                        break;
                    //case "9": // Rename folder name.
                    //    Type("Rename folders.", 7, 100, 1);
                    //    GetFolders();
                    //    break;
                    //case "11": // Move kids movies.
                    //    Type("Move the kids movies around. Let's go!", 7, 100, 1);
                    //    MoveKidsMovies();
                    //    break;
                    case "13": // Copy JPG files.
                        CopyJpgFiles();
                        break;
                    //case "14": // Copy movie files.
                    //    CopyMovieFiles();
                    //    break;
                    //case "15": // Mark Owned Movies as D=Done || X=Not Done.
                    //    Type("Mark Owned Movies as D=Done || X=Not Done.", 7, 100, 1);
                    //    CheckForMovie("Main");
                    //    break;
                    //case "21": // testing rewriting to the same console line.
                    //    Type("1", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("2", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("3", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("4", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("5", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("6", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("7", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("8", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("9", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("10", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("11", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("12", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("13", 100, 100, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("14", 100, 100, 0);
                    //    break;
                    case "t22": // Simply prints out examples of all the font colors I use.

                        string myColor = "Black";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "Blue";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "Cyan";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "DarkBlue";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "DarkCyan";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "DarkGray";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "DarkGreen";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "DarkMagenta";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "DarkRed";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "DarkYellow";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "Gray";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "Green";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "Magenta";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "Red";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "White";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        myColor = "Yellow";
                        Type("This is " + myColor, 0, 0, 1, myColor);

                        break;
                    case "t23": // Testing time interval.
                        DateTime a = new DateTime(2008, 01, 02, 06, 30, 00);
                        DateTime b = new DateTime(2008, 01, 03, 09, 43, 55);

                        TimeSpan duration = b - a;

                        Console.WriteLine(duration.ToString());
                        //Console.WriteLine("Days: " + duration.TotalDays + ", Hours: " + (duration.TotalHours % 24) + ", Minutes: " + (duration.TotalMinutes % 60) + ", Seconds: " + (duration.TotalSeconds % 60));
                        break;
                    case "t24":
                        DisplayMessage("info", "Testing the countdown for 10 seconds.");
                        Countdown(10);
                        DisplayMessage("success", "Test Complete");
                        break;
                        //default: // Other.
                        //    DidntUnderstand(choice);
                        //    break;
                } // End switch

                return keepAskingForChoice;
            } // End try
            catch (Exception ex)
            {
                Type(ex.ToString(), 0, 0, 1);
                DidntUnderstand(choice);

                return keepAskingForChoice;
            }
        } // End CallSwitch()

        public static void CreateMissingCombinedEpisodeNfoFiles(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string directory)
        {
            try
            {
                string[] fileEntries = Directory.GetFiles(directory);
                var fileCount = fileEntries.Length;
                string plural = fileCount == 1 ? " file " : " files ";
                var count = 1;
                DisplayMessage("warning", "Searching directory for Combined Episodes...");
                DisplayMessage("info", "Found " + fileCount + " " + plural);
                foreach (var row in data)
                {
                    var tvdbId = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                    var oldName = row[Convert.ToInt16(sheetVariables["Combined Episode Name"])].ToString();
                    var newName = row[Convert.ToInt16(sheetVariables["New Episode Name"])].ToString();
                    var nfoBody = row[Convert.ToInt16(sheetVariables["NFO Body"])].ToString();

                    if (!tvdbId.Equals("")) // If the ID is empty then just skip the rest.
                    {
                        var sourceFile = Directory.GetFiles(directory, oldName + ".*");
                        if (sourceFile.Length > 0)
                        {
                            var destinationFile = Path.Combine(directory, newName) + Path.GetExtension(sourceFile[0]);
                            var nfoFile = Path.Combine(directory, newName) + ".nfo";
                            File.Move(sourceFile[0], destinationFile);
                            WriteNfoFile(nfoFile, nfoBody);
                            DisplayMessage("default", "(" + count + " of " + fileCount + ")" + " File renamed and NFO file created for: ", 0);
                            DisplayMessage("success", newName);
                            count++;
                        }
                    }
                }
                DisplayMessage("success", "DONE");
            }
            catch (Exception e)
            {
                DisplayMessage("error", "An error occured in CreateMissingCombinedEpisodeNfoFiles method:");
                DisplayMessage("harderror", e.Message);
                throw;
            }
        } // End FixRecordedNames()

        private static void InsertMissingDbData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string token)
        {
            DisplayMessage("info", "Now inserting info into DB sheet.");

            int intSeriesNamesInsertedCount = 0,
                intSeasonCountsInsertedCount = 0,
                intContinuingsInsertedCount = 0,
                intTvdbIdsInsertedCount = 0;

            string tvdbIdValue = "", // Our current TVDB ID value from the Google Sheet.
                seriesNameValue = "", // Our current TVDB Series Name from the Google Sheet.
                seasonCountValue = "", // Our current Season Count value from the Google Sheet.
                continuingValue = "", // Our current Continuing? value from the Google Sheet.
                tvdbSlugValue = "", // Our current TVDB Slug value from the Google Sheet.
                rowNum = "", // Holds the row number we are on.
                strCellToPutData = ""; // The string of the location to write the data to.

            int tvdbIdColumnNum = 0, // Used to input the returned ID back into the Google Sheet.
                seriesNameColumnNum = 0, // Used to input the returned rating into the Google Sheet.
                seasonCountColumnNum = 0, // Used to input the returned rating into the Google Sheet.
                continuingColumnNum = 0; // Used to input the returned overview into the Google Sheet.

            foreach (var row in data)
            {
                try
                {
                    rowNum = row[Convert.ToInt16(sheetVariables["RowNum"])].ToString();
                    tvdbSlugValue = row[Convert.ToInt16(sheetVariables["TVDB Slug"])].ToString();
                    tvdbIdValue = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                    tvdbIdColumnNum = Convert.ToInt16(sheetVariables["TVDB ID"]);
                    seriesNameValue = row[Convert.ToInt16(sheetVariables["Series Name"])].ToString();
                    seriesNameColumnNum = Convert.ToInt16(sheetVariables["Series Name"]);
                    seasonCountValue = row[Convert.ToInt16(sheetVariables["Season Count"])].ToString();
                    seasonCountColumnNum = Convert.ToInt16(sheetVariables["Season Count"]);
                    continuingValue = row[Convert.ToInt16(sheetVariables["Continuing?"])].ToString();
                    continuingColumnNum = Convert.ToInt16(sheetVariables["Continuing?"]);

                    if (!tvdbSlugValue.Equals("")) // If there is no slug then the row is considered empty and should be skipped.
                    {
                        // First check to see if the id is empty and populate it if it is.
                        if (tvdbIdValue.Equals(""))
                        {
                            var response = TvdbApiCall.TvdbApi.GetSeriesIdAsync(ref token, tvdbSlugValue);
                            tvdbIdValue = response;

                            strCellToPutData = "DB!" + ColumnNumToLetter(tvdbIdColumnNum) + rowNum;

                            WriteSingleCellToSheet(response, strCellToPutData);

                            DisplayMessage("default", "ID saved for ", 0, 10);
                            DisplayMessage("success", tvdbSlugValue, 1, 5, 10);
                            intTvdbIdsInsertedCount++;
                        }
                        if (seriesNameValue.Equals("") || seasonCountValue.Equals("") || continuingValue.Equals(""))
                        {
                            var response = TvdbApiCall.TvdbApi.GetSeriesDetailsAsync(ref token, tvdbIdValue);

                            if (seriesNameValue.Equals(""))
                            {
                                seriesNameValue = response.data.seriesName.ToString();
                                strCellToPutData = "DB!" + ColumnNumToLetter(seriesNameColumnNum) + rowNum;

                                WriteSingleCellToSheet(seriesNameValue, strCellToPutData);
                                DisplayMessage("default", "Series Name saved for ", 0);
                                DisplayMessage("success", seriesNameValue);
                                intSeriesNamesInsertedCount++;
                            }
                            if (seasonCountValue.Equals(""))
                            {
                                seasonCountValue = response.data.season.ToString();
                                strCellToPutData = "DB!" + ColumnNumToLetter(seasonCountColumnNum) + rowNum;

                                WriteSingleCellToSheet(seasonCountValue, strCellToPutData);
                                DisplayMessage("default", "Season Count saved for ", 0);
                                DisplayMessage("success", seriesNameValue);
                                intSeasonCountsInsertedCount++;
                            }
                            if (continuingValue.Equals(""))
                            {
                                if (response.data.status.ToString() == "Continuing")
                                {
                                    continuingValue = "Y";
                                }
                                else
                                {
                                    continuingValue = "N";
                                }

                                strCellToPutData = "DB!" + ColumnNumToLetter(continuingColumnNum) + rowNum;

                                WriteSingleCellToSheet(continuingValue, strCellToPutData);
                                DisplayMessage("default", "Continuing saved for ", 0);
                                DisplayMessage("success", seriesNameValue);
                                intContinuingsInsertedCount++;
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    //Type("Something went wrong while putting in movie data for: " + imdbTitle, 3, 100, 1, "Red");
                    Type(e.Message, 3, 100, 2, "DarkRed");
                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("New TVDB IDs inserted: " + intTvdbIdsInsertedCount, 0, 0, 1, "Green");
            Type("New Series Names inserted: " + intSeriesNamesInsertedCount, 0, 0, 1, "Green");
            Type("New Season Counts inserted: " + intSeasonCountsInsertedCount, 0, 0, 1, "Green");
            Type("New Continuings inserted: " + intContinuingsInsertedCount, 0, 0, 1, "Green");

        } // End InsertMissingDbData()

        private static void UpdateDbData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string token)
        {
            int intSeriesNamesUpdatedCount = 0,
                intSeasonCountsUpdatedCount = 0,
                intContinuingsUpdatedCount = 0,
                intTvdbIdsInsertedCount = 0;

            string tvdbIdValue = "", // Our current TVDB ID value from the Google Sheet.
                seriesNameValue = "", // Our current TVDB Series Name from the Google Sheet.
                seasonCountValue = "", // Our current Season Count value from the Google Sheet.
                continuingValue = "", // Our current Continuing? value from the Google Sheet.
                tvdbSlugValue = "", // Our current TVDB Slug value from the Google Sheet.
                rowNum = "", // Holds the row number we are on.
                strCellToPutData = "", // The string of the location to write the data to.
                seriesNameCall = "", // The series name pulled from the API call.
                seasonCountCall = "", // The Season Count pulled from the API call.
                continuingCall = ""; // The Continuing value pulled from the API call.

            int tvdbIdColumnNum = 0, // Used to input the returned ID back into the Google Sheet.
                seriesNameColumnNum = 0, // Used to input the returned rating into the Google Sheet.
                seasonCountColumnNum = 0, // Used to input the returned rating into the Google Sheet.
                continuingColumnNum = 0; // Used to input the returned overview into the Google Sheet.

            foreach (var row in data)
            {
                try
                {
                    var showUpdated = false;
                    rowNum = row[Convert.ToInt16(sheetVariables["RowNum"])].ToString();
                    tvdbSlugValue = row[Convert.ToInt16(sheetVariables["TVDB Slug"])].ToString();
                    tvdbIdValue = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                    tvdbIdColumnNum = Convert.ToInt16(sheetVariables["TVDB ID"]);
                    seriesNameValue = row[Convert.ToInt16(sheetVariables["Series Name"])].ToString();
                    seriesNameColumnNum = Convert.ToInt16(sheetVariables["Series Name"]);
                    seasonCountValue = row[Convert.ToInt16(sheetVariables["Season Count"])].ToString();
                    seasonCountColumnNum = Convert.ToInt16(sheetVariables["Season Count"]);
                    continuingValue = row[Convert.ToInt16(sheetVariables["Continuing?"])].ToString();
                    continuingColumnNum = Convert.ToInt16(sheetVariables["Continuing?"]);

                    if (!tvdbSlugValue.Equals("")) // If there is no slug then the row is considered empty and should be skipped.
                    {
                        // First check to see if the id is empty and populate it if it is.
                        if (tvdbIdValue.Equals(""))
                        {
                            var idResponse = TvdbApiCall.TvdbApi.GetSeriesIdAsync(ref token, tvdbSlugValue);
                            tvdbIdValue = idResponse;

                            strCellToPutData = "DB!" + ColumnNumToLetter(tvdbIdColumnNum) + rowNum;

                            WriteSingleCellToSheet(idResponse, strCellToPutData);

                            DisplayMessage("default", "ID saved for ", 0);
                            DisplayMessage("success", tvdbSlugValue, 1, 0, 600);
                            intTvdbIdsInsertedCount++;
                            showUpdated = true;
                        }

                        var detailsResponse = TvdbApiCall.TvdbApi.GetSeriesDetailsAsync(ref token, tvdbIdValue);
                        seriesNameCall = detailsResponse.data.seriesName.ToString();
                        seasonCountCall = detailsResponse.data.season.ToString();

                        if (detailsResponse.data.status.ToString() == "Continuing")
                        {
                            continuingCall = "Y";
                        }
                        else
                        {
                            continuingCall = "N";
                        }

                        if (seriesNameValue != seriesNameCall)
                        {
                            strCellToPutData = "DB!" + ColumnNumToLetter(seriesNameColumnNum) + rowNum;

                            WriteSingleCellToSheet(seriesNameCall, strCellToPutData);
                            DisplayMessage("default", "Updated Series Name from '", 0);
                            DisplayMessage("info", seriesNameValue, 0);
                            DisplayMessage("default", "' to '", 0);
                            DisplayMessage("success", seriesNameCall, 0);
                            DisplayMessage("default", "'", 1, 0, 600);
                            intSeriesNamesUpdatedCount++;
                            showUpdated = true;
                        }
                        if (seasonCountValue != seasonCountCall)
                        {
                            strCellToPutData = "DB!" + ColumnNumToLetter(seasonCountColumnNum) + rowNum;

                            WriteSingleCellToSheet(seasonCountCall, strCellToPutData);
                            DisplayMessage("default", "Updated Season Count from '", 0);
                            DisplayMessage("info", seasonCountValue, 0);
                            DisplayMessage("default", "' to '", 0);
                            DisplayMessage("success", seasonCountCall, 0);
                            DisplayMessage("default", "' for '", 0);
                            DisplayMessage("question", seriesNameValue, 0);
                            DisplayMessage("default", "'", 1, 0, 600);
                            intSeasonCountsUpdatedCount++;
                            showUpdated = true;
                        }
                        if (continuingValue != continuingCall)
                        {
                            strCellToPutData = "DB!" + ColumnNumToLetter(continuingColumnNum) + rowNum;

                            WriteSingleCellToSheet(continuingCall, strCellToPutData);
                            DisplayMessage("default", "Updated Continuing status from '", 0);
                            DisplayMessage("info", continuingValue, 0);
                            DisplayMessage("default", "' to '", 0);
                            DisplayMessage("success", continuingCall, 0);
                            DisplayMessage("default", "' for ", 0);
                            DisplayMessage("question", seriesNameValue, 1, 0, 600);
                            intContinuingsUpdatedCount++;
                            showUpdated = true;
                        }

                        if (!showUpdated)
                        {
                            //DisplayMessage("default", "Nothing updated for: ", 0);
                            //DisplayMessage("info", seriesNameValue, 1, 0, 600);
                        }
                    }
                }
                catch (Exception e)
                {
                    Type(e.Message, 3, 100, 2, "DarkRed");
                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("New TVDB IDs inserted: " + intTvdbIdsInsertedCount, 0, 0, 1, "Green");
            Type("New Series Names updated: " + intSeriesNamesUpdatedCount, 0, 0, 1, "Green");
            Type("New Season Counts updated: " + intSeasonCountsUpdatedCount, 0, 0, 1, "Green");
            Type("New Continuings updated: " + intContinuingsUpdatedCount, 0, 0, 1, "Green");

        } // End UpdateDbData()

        private static void InsertMissingCombinedEpisodeData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string token)
        {
            int intPlot1InsertedCount = 0,
                intPlot2InsertedCount = 0,
                intPlot1EmptyCount = 0,
                intPlot2EmptyCount = 0;

            string rowNum = "", // Holds the row number we are on.
                tvdbIdValue = "", // Our current TVDB ID value from the Google Sheet.
                strCellToPutData = "", // The string of the location to write the data to.
                plot1Data = "",
                plot2Data = "",
                episode1SeasonNum = "",
                episode1Num = "",
                episode2SeasonNum = "",
                episode2Num = "",
                showTitle = "";

            int plot1ColumnNum = 0, // Used to input the returned plot into the Google Sheet.
                plot2ColumnNum = 0; // Used to input the returned plot into the Google Sheet.

            foreach (var row in data)
            {
                try
                {
                    if (row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString() != "") // If there is no id then the row is considered empty and should be skipped.
                    {
                        rowNum = row[Convert.ToInt16(sheetVariables["RowNum"])].ToString();
                        tvdbIdValue = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                        plot1Data = row[Convert.ToInt16(sheetVariables["Episode 1 Plot"])].ToString();
                        plot1ColumnNum = Convert.ToInt16(sheetVariables["Episode 1 Plot"]);
                        plot2Data = row[Convert.ToInt16(sheetVariables["Episode 2 Plot"])].ToString();
                        plot2ColumnNum = Convert.ToInt16(sheetVariables["Episode 2 Plot"]);
                        episode1SeasonNum = row[Convert.ToInt16(sheetVariables["Episode 1 Season"])].ToString();
                        episode1Num = row[Convert.ToInt16(sheetVariables["Episode 1 No."])].ToString();
                        episode2SeasonNum = row[Convert.ToInt16(sheetVariables["Episode 2 Season"])].ToString();
                        episode2Num = row[Convert.ToInt16(sheetVariables["Episode 2 No."])].ToString();
                        showTitle = row[Convert.ToInt16(sheetVariables["Show Title"])].ToString();

                        if (plot1Data.Equals(""))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode1SeasonNum, episode1Num);

                            plot1Data = response.data[0].overview.ToString();

                            if (plot1Data.Equals(""))
                            {
                                DisplayMessage("default", "No plot available for ", 0);
                                DisplayMessage("warning", showTitle + " - S" + episode1SeasonNum + "E" + episode1Num);

                                intPlot1EmptyCount++;
                            }
                            else
                            {
                                strCellToPutData = "Combined Episodes!" + ColumnNumToLetter(plot1ColumnNum) + rowNum;

                                WriteSingleCellToSheet(plot1Data, strCellToPutData);
                                DisplayMessage("default", "Plot saved for ", 0);
                                DisplayMessage("success", showTitle + " - S" + episode1SeasonNum + "E" + episode1Num + " - at " + strCellToPutData);

                                intPlot1InsertedCount++;
                            }
                        }
                        if (plot2Data.Equals(""))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode2SeasonNum, episode2Num);

                            plot2Data = response.data[0].overview.ToString();

                            if (plot2Data.Equals(""))
                            {
                                DisplayMessage("default", "No plot available for ", 0);
                                DisplayMessage("warning", showTitle + " - S" + episode2SeasonNum + "E" + episode2Num);

                                intPlot2EmptyCount++;
                            }
                            else
                            {
                                strCellToPutData = "Combined Episodes!" + ColumnNumToLetter(plot2ColumnNum) + rowNum;

                                WriteSingleCellToSheet(plot2Data, strCellToPutData);
                                DisplayMessage("default", "Plot saved for ", 0);
                                DisplayMessage("success", showTitle + " - S" + episode2SeasonNum + "E" + episode2Num + " - at " + strCellToPutData);

                                intPlot2InsertedCount++;
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Type(e.Message, 3, 100, 2, "DarkRed");
                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("New Plots inserted for episode one: " + intPlot1InsertedCount, 0, 0, 1, "Green");
            Type("New Plots inserted for episode two: " + intPlot2InsertedCount, 0, 0, 1, "Green");
            Type("Plots skipped due to no plot for episode one: " + intPlot1EmptyCount, 0, 0, 1, "Yellow");
            Type("Plots skipped due to no plot for episode two: " + intPlot2EmptyCount, 0, 0, 1, "Yellow");

        } // End InsertMissingCombinedEpisodeData()
        
        private static void InsertMissingEpisodeData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string token)
        {
            int intNamesInsertedCount = 0,
                intPlotsInsertedCount = 0,
                intPlotEmptyCount = 0;

            string rowNum = "", // Holds the row number we are on.
                tvdbIdValue = "", // Our current TVDB ID value from the Google Sheet.
                strCellToPutData = "", // The string of the location to write the data to.
                plotData = "",
                episodeSeasonNum = "",
                episodeNum = "",
                episodeName = "",
                showName = "";

            int episodeNameColumnNum = 0, // Used to input the returned name into the Google Sheet.
                plotColumnNum = 0; // Used to input the returned plot into the Google Sheet.
            foreach (var row in data)
            {
                try
                {
                    if (row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString() != "") // If there is no id then the row is considered empty and should be skipped.
                    {
                        rowNum = row[Convert.ToInt16(sheetVariables["RowNum"])].ToString();
                        tvdbIdValue = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                        plotData = row[Convert.ToInt16(sheetVariables["Plot"])].ToString();
                        plotColumnNum = Convert.ToInt16(sheetVariables["Plot"]);
                        episodeName = row[Convert.ToInt16(sheetVariables["Episode Name"])].ToString();
                        episodeNameColumnNum = Convert.ToInt16(sheetVariables["Episode Name"]);
                        episodeNum = row[Convert.ToInt16(sheetVariables["Episode #"])].ToString();
                        episodeSeasonNum = row[Convert.ToInt16(sheetVariables["Season #"])].ToString();
                        showName = row[Convert.ToInt16(sheetVariables["Show"])].ToString();

                        if (episodeName.Equals("") || plotData.Equals(""))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episodeSeasonNum, episodeNum);

                            if (response.Content != null)
                            {
                                string error = response.Content.ToString();

                                if (error.Contains("No results"))
                                {
                                    DisplayMessage("error", "TVDB does not contain data for: ", 0);
                                    DisplayMessage("default", showName + " - S" + episodeSeasonNum + "E" + episodeNum);
                                }
                            } else
                            {
                                string plot = response.data[0].overview.ToString(),
                                    name = response.data[0].episodeName.ToString();

                                if (episodeName.Equals(""))
                                {
                                    strCellToPutData = "Episodes!" + ColumnNumToLetter(episodeNameColumnNum) + rowNum;

                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("success", "Name saved for: ", 0);
                                    DisplayMessage("default", showName + " - S" + episodeSeasonNum + "E" + episodeNum + " - " + name);
                                    intNamesInsertedCount++;
                                }

                                if (plotData.Equals(""))
                                {
                                    if (!plot.Equals(""))
                                    {
                                        strCellToPutData = "Episodes!" + ColumnNumToLetter(plotColumnNum) + rowNum;

                                        WriteSingleCellToSheet(plot, strCellToPutData);
                                        DisplayMessage("success", "Plot saved for: ", 0);
                                        DisplayMessage("default", showName + " - S" + episodeSeasonNum + "E" + episodeNum + " - " + name);
                                        intPlotsInsertedCount++;
                                    }
                                    else
                                    {
                                        DisplayMessage("warning", "No plot available for: ", 0);
                                        DisplayMessage("default", showName + " - S" + episodeSeasonNum + "E" + episodeNum + " - " + name);

                                        intPlotEmptyCount++;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Type(e.Message, 3, 100, 2, "DarkRed");
                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            DisplayMessage("success", "New Names inserted: ", 0);
            DisplayMessage("default", intNamesInsertedCount.ToString()) ;
            DisplayMessage("success", "New Plots inserted: ", 0);
            DisplayMessage("default", intPlotsInsertedCount.ToString());
            DisplayMessage("warning", "Plots skipped due to no plot for the episode: ", 0);
            DisplayMessage("default", intPlotEmptyCount.ToString());

        } // End InsertMissingEpisodeData()
            
        private static void UpdateCombinedEpisodeData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string token)
        {
            int intPlot1InsertedCount = 0,
                intPlot2InsertedCount = 0,
                intPlot1LockedCount = 0,
                intPlot2LockedCount = 0,
                intPlot1EmptyCount = 0,
                intPlot2EmptyCount = 0;

            string tvdbIdValue = "", // Our current TVDB ID value from the Google Sheet.
                rowNum = "", // Holds the row number we are on.
                strCellToPutData = "", // The string of the location to write the data to.
                plot1Data = "", // The plot for episode 1 in the Google Sheet.
                plot2Data = "", // The plot for episode 2 in the Google Sheet.
                plot1Call = "", // The plot for episode 1 from the API call.
                plot2Call = "", // The plot for episode 2 from the API call.
                episode1SeasonNum = "",
                episode1Num = "",
                episode2SeasonNum = "",
                episode2Num = "",
                showTitle = "",
                lockPLot1 = "",
                lockPlot2 = "";

            int plot1ColumnNum = 0, // Used to input the returned plot into the Google Sheet.
                plot2ColumnNum = 0; // Used to input the returned plot into the Google Sheet.

            foreach (var row in data)
            {
                try
                {
                    rowNum = row[Convert.ToInt16(sheetVariables["RowNum"])].ToString();
                    tvdbIdValue = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                    plot1Data = row[Convert.ToInt16(sheetVariables["Episode 1 Plot"])].ToString();
                    plot1ColumnNum = Convert.ToInt16(sheetVariables["Episode 1 Plot"]);
                    plot2Data = row[Convert.ToInt16(sheetVariables["Episode 2 Plot"])].ToString();
                    plot2ColumnNum = Convert.ToInt16(sheetVariables["Episode 2 Plot"]);
                    episode1SeasonNum = row[Convert.ToInt16(sheetVariables["Episode 1 Season"])].ToString();
                    episode1Num = row[Convert.ToInt16(sheetVariables["Episode 1 No."])].ToString();
                    episode2SeasonNum = row[Convert.ToInt16(sheetVariables["Episode 2 Season"])].ToString();
                    episode2Num = row[Convert.ToInt16(sheetVariables["Episode 2 No."])].ToString();
                    showTitle = row[Convert.ToInt16(sheetVariables["Show Title"])].ToString();
                    lockPLot1 = row[Convert.ToInt16(sheetVariables["Lock Plot 1"])].ToString();
                    lockPlot2 = row[Convert.ToInt16(sheetVariables["Lock Plot 2"])].ToString();

                    if (!tvdbIdValue.Equals("")) // If there is no id then the row is considered empty and should be skipped.
                    {
                        var episode1Response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode1SeasonNum, episode1Num);
                        plot1Call = episode1Response.data[0].overview.ToString();

                        if (plot1Call.Equals(""))
                        {
                            DisplayMessage("default", "No plot available for ", 0);
                            DisplayMessage("warning", showTitle + " - S" + episode1SeasonNum + "E" + episode1Num);

                            intPlot1EmptyCount++;
                        }
                        else if (lockPLot1.ToUpper().Equals("X"))
                        {
                            DisplayMessage("default", "Plot locked for ", 0);
                            DisplayMessage("info", showTitle + " - S" + episode1SeasonNum + "E" + episode1Num);

                            intPlot1LockedCount++;
                        }
                        else if (plot1Call != plot1Data)
                        {
                            strCellToPutData = "Combined Episodes!" + ColumnNumToLetter(plot1ColumnNum) + rowNum;

                            WriteSingleCellToSheet(plot1Call, strCellToPutData);
                            DisplayMessage("default", "Plot updated for ", 0);
                            DisplayMessage("success", showTitle + " - S" + episode1SeasonNum + "E" + episode1Num + " - at " + strCellToPutData);

                            intPlot1InsertedCount++;
                        }

                        var episode2Response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode2SeasonNum, episode2Num);
                        plot2Call = episode2Response.data[0].overview.ToString();

                        if (plot2Call.Equals(""))
                        {
                            DisplayMessage("default", "No plot available for ", 0);
                            DisplayMessage("warning", showTitle + " - S" + episode2SeasonNum + "E" + episode2Num);

                            intPlot2EmptyCount++;
                        }
                        else if (lockPlot2.ToUpper().Equals("X"))
                        {
                            DisplayMessage("default", "Plot locked for ", 0);
                            DisplayMessage("info", showTitle + " - S" + episode2SeasonNum + "E" + episode2Num);

                            intPlot2LockedCount++;
                        }
                        else if (plot2Call != plot2Data)
                        {
                            strCellToPutData = "Combined Episodes!" + ColumnNumToLetter(plot2ColumnNum) + rowNum;

                            WriteSingleCellToSheet(plot2Call, strCellToPutData);
                            DisplayMessage("default", "Plot saved for ", 0);
                            DisplayMessage("success", showTitle + " - S" + episode2SeasonNum + "E" + episode2Num + " - at " + strCellToPutData);

                            intPlot2InsertedCount++;
                        }
                    }
                }
                catch (Exception e)
                {
                    Type(e.Message, 3, 100, 2, "DarkRed");
                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("New Plots inserted for episode one: " + intPlot1InsertedCount, 0, 0, 1, "Green");
            Type("New Plots inserted for episode two: " + intPlot2InsertedCount, 0, 0, 1, "Green");
            Type("Plots skipped due to the plot being locked for episode one: " + intPlot1LockedCount, 0, 0, 1, "Blue");
            Type("Plots skipped due to the plot being locked for episode two: " + intPlot2LockedCount, 0, 0, 1, "Blue");
            Type("Plots skipped due to no plot for episode one: " + intPlot1EmptyCount, 0, 0, 1, "Yellow");
            Type("Plots skipped due to no plot for episode two: " + intPlot2EmptyCount, 0, 0, 1, "Yellow");

        } // End UpdateCombinedEpisodeData()

        private static void WriteToSheetColumn(ArrayList videoFilesList, IList<IList<Object>> sheetData, string sheetName, int columnNum)
        {
            var i = 0;
            DisplayMessage("info", "Adding " + videoFilesList.Count + " names to sheet '" + sheetName + "'");
            foreach (var row in sheetData)
            {
                string intRowNum =  row[0].ToString(),
                    columnToWriteTo = row[columnNum].ToString();

                if (columnToWriteTo.Equals("") && i < videoFilesList.Count)
                {
                    string fileName = Path.GetFileNameWithoutExtension(videoFilesList[i].ToString());
                    string strCellToPutData = sheetName + "!" + ColumnNumToLetter(columnNum) + int.Parse(intRowNum);
                    WriteSingleCellToSheet(fileName, strCellToPutData);
                    i++;
                    DisplayMessage("default", i + " of " + videoFilesList.Count, 0);
                    DisplayMessage("success", " - " + fileName, 0);
                    DisplayMessage("default", " - saved to row ", 0);
                    DisplayMessage("info", intRowNum);
                }
            }
        }

        private static void CreateFoldersAndMoveFiles(string directory)
        {
            string[] fileEntries = Directory.GetFiles(directory);
            ArrayList fileNamesWithoutExtensions = new ArrayList();

            foreach (var myFile in fileEntries)
            {
                var sourceFile = myFile;
                var fileName = Path.GetFileName(myFile);
                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(myFile);
                var directoryName = Path.Combine(Path.GetDirectoryName(myFile), fileNameWithoutExtension);
                var destinationFile = Path.Combine(directoryName, fileName);

                try
                {
                    Directory.CreateDirectory(directoryName);
                }
                catch (Exception e)
                {
                    Type("Something went wrong while creating the directory: " + directoryName, 3, 100, 1, "Red");
                    Type(e.Message, 3, 100, 2, "DarkRed");
                    throw;
                }

                try
                {
                    File.Move(sourceFile, destinationFile);
                    Type(fileName, 0, 0, 0);
                    Type(" Moved", 0, 0, 1, "Green");
                }
                catch (Exception e)
                {
                    Type("Something went wrong while moving the file: " + fileName, 3, 100, 1, "Red");
                    Type(e.Message, 3, 100, 2, "DarkRed");
                    throw;
                }
            }
        }

        private static void TrimTitlesInDirectory(string directory)
        {
            DirectoryInfo d = new DirectoryInfo(directory);
            FileInfo[] fileEntries = d.GetFiles();

            if (fileEntries.Length > 0)
            {
                foreach (FileInfo f in fileEntries)
                {
                    try
                    {
                        if (f.Name.Length > 20)
                        {
                            File.Move(f.FullName, Path.Combine(f.DirectoryName, f.Name.Substring(0, 20).Trim()) + f.Extension);
                            Type(f.Name, 0, 0, 0);
                            Type(" Trimmed", 0, 0, 1, "Green");
                        } else
                        {
                            Type(f.Name, 0, 0, 0);
                            Type(" NOT Trimmed", 0, 0, 1, "Blue");
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong while trimming the file: " + f.Name, 3, 100, 1, "Red");
                        Type(e.Message, 3, 100, 2, "DarkRed");
                        throw;
                    }
                    
                }
            }

        }

        private static void ResetGlobals()
        {
            runningTotalConversionTime = new TimeSpan();
            runningDifference = 0;
            runningFileSize = 0;
        }

        private static ArrayList GrabMovieFiles(string[] files)
        {
            Type("Grabbing just the video files... ", 10, 0, 0, "Yellow");
            ArrayList videoFiles = new ArrayList();

            try
            {
                foreach (string file in files)
                {
                    if (file.ToUpper().Contains(".MP4") || file.ToUpper().Contains(".MKV") || file.ToUpper().Contains(".M4V") || file.ToUpper().Contains(".AVI") || file.ToUpper().Contains(".WEBM"))
                    {
                        videoFiles.Add(file);
                    }
                }
                Type("DONE", 100, 0, 1, "Green");
                return videoFiles;
            }
            catch (Exception e)
            {
                Type("An error occured!", 0, 0, 1, "Red");
                Type(e.Message, 0, 0, 1, "DarkRed");
                throw;
            }

        } // End GrabMovieFiles()

        public static ArrayList AskForFilesToCopy()
        {
            ArrayList videoFiles = new ArrayList();

            try
            {
                do
                {
                    DisplayMessage("question", "Give me a file you'd like to copy- (Type 0 when done)");

                    var file = RemoveCharFromString(Console.ReadLine(), '"');

                    if (file != "0")
                    {
                        videoFiles.Add(file);
                    }
                    else
                    {
                        return videoFiles;
                    }
                } while (true);
                
            }
            catch (Exception e)
            {
                DisplayMessage("error", "An error occured while adding your file to my list |");
                DisplayMessage("harderror", e.Message);
                throw;
            }
        }

        public static string RemoveCharFromString(string myString, char c)
        {
            string newString = "";
            for (int i = 0; i < myString.Length; i++)
            {
                if (myString[i] != c)
                {
                    newString += myString[i];
                }
            }
            return newString;
        } // End RemoveCharFromString()

        private static long SizeOfFiles(ArrayList files)
        {
            long fileSize = new long();

            try
            {
                foreach (string file in files)
                {
                    if (File.Exists(file))
                    {
                        fileSize += new FileInfo(file).Length;
                    }
                }

                return fileSize;
            }
            catch (Exception e)
            {
                Type("An error occured!", 0, 0, 1, "Red");
                Type(e.Message, 0, 0, 1, "DarkRed");
                throw;
            }
        }

        static readonly string[] suffixes =
            { "Bytes", "KB", "MB", "GB", "TB", "PB" };
        public static string FormatSize(Int64 bytes)
        {
            int counter = 0;
            decimal number = (decimal)bytes;
            while (Math.Round(number / 1024) >= 1)
            {
                number = number / 1024;
                counter++;
            }
            return string.Format("{0:n1}{1}", number, suffixes[counter]);
        }

        /// <summary>
        /// Takes the sheetVariables Dictionary variable with -1 values and finds the column numbers.
        /// </summary>
        /// <param name="sheetVariables">The Dictionary holding the column names and number.</param>
        /// <param name="titleData">Holds the title row data to fill the sheetVariables Dictionary.</param>
        /// <returns></returns>
        private static Dictionary<string, int> UpdateSheetVariables(Dictionary<string, int> sheetVariables, IList<IList<Object>> titleData)
        {
            int x = 0;
            foreach (var row in titleData)
            {
                do
                {
                    foreach (var variable in sheetVariables.ToList())
                    {
                        if (row[x].ToString() == variable.Key)
                        {
                            sheetVariables[variable.Key] = x;
                        }
                    }
                    x++;

                } while (x < row.Count);

            }

            return sheetVariables;
        } // End UpdateSheetVariables()

        private static void MissingColumn(Dictionary<string, int> NotFoundColumns)
        {
            Type("We didn't find a column that we were looking for...", 1, 100, 1, "Red");
            foreach (KeyValuePair<string, int> variable in NotFoundColumns)
            {
                Type("Missing Column: '" + variable.Key.ToString() + "'", 1, 100, 1, "DarkRed");
            }
            Type("It's likely that the column we are looking for has changed names.", 1, 100, 2, "Red");
            Type("Press ENTER to exit the program.", 1, 100, 1, "DarkRed");
            Console.ReadLine();
            Environment.Exit(0);
        }

        static void AskForMenu()
        {
            Console.WriteLine();
            Type("Press any key to open the menu...", 0, 0, 1, "Magenta");
            Console.ReadKey();
        }

        protected static void RemoveMoviesFromTmdbList(IList<IList<Object>> data, Dictionary<string, int> sheetVariables)
        {
            Type("Going through the list to find movies to remove...", 0, 0, 1, "Blue");
            int intMoviesRemovedCount = 0, intMoviesSkippedCount = 0, intMoviesNotInListCount = 0, intTmdbIdNotFoundCount = 0;

            string tmdbIdValue = "", CleanTitle = "", status = "";
            dynamic tmdbResponse;
            bool responseIsBroken = false, movieIsInList = false;

            foreach (var row in data)
            {
                if (row.Count > 20) // If it's an empty row then it should have less than this.
                {
                    try
                    {
                        tmdbIdValue = row[Convert.ToInt16(sheetVariables["TMDB ID"])].ToString();
                        CleanTitle = row[Convert.ToInt16(sheetVariables["Clean Title"])].ToString();
                        status = row[Convert.ToInt16(sheetVariables["Status"])].ToString();

                        // If the movie is marked as done in our DB,
                        // and there is a valid TMDB ID then proceed to check if the movie is in our list.
                        if (!status.Equals("") && status[0].ToString().ToUpper() != "X" && !tmdbIdValue.Equals("") && tmdbIdValue.ToUpper() != "N/A")
                        {
                            do
                            {
                                tmdbResponse = TmdbApi.ListsCheckItemStatus(tmdbIdValue);

                                if (tmdbResponse.item_present != null)
                                {
                                    if (tmdbResponse.item_present.ToString().ToUpper() == "TRUE")
                                    {
                                        movieIsInList = true;
                                    } else
                                    {
                                        movieIsInList = false;
                                    }
                                }
                                else if (tmdbResponse.message != null)
                                {
                                    Type(CleanTitle + " | " + tmdbResponse.message, 0, 0, 1, "Red");
                                }
                                else
                                {
                                    responseIsBroken = true;
                                }
                            } while (responseIsBroken);

                            if (movieIsInList)
                            {
                                tmdbResponse = TmdbApi.ListsRemoveMovie(tmdbIdValue);

                                if (tmdbResponse.status_code == DELETED_SUCCESSFULLY)
                                {
                                    Type(CleanTitle + " | " + tmdbResponse.message, 0, 0, 1, "Green");
                                    intMoviesRemovedCount++;
                                } else if (tmdbResponse.message != null)
                                {
                                    Type(CleanTitle + " | " + tmdbResponse.message, 0, 0, 1, "Red");
                                } else
                                {
                                    Type("Something went wrong", 0, 0, 1, "Red");
                                }
                            } else
                            {
                                intMoviesNotInListCount++;
                            }
                        } else
                        {
                            intMoviesSkippedCount++;
                        }

                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong..." + e.Message, 3, 100, 1, "Red");
                    }

                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("Movies removed: " + intMoviesRemovedCount, 0, 0, 1, "Green");
            Type("Movies skipped: " + intMoviesSkippedCount, 0, 0, 1, "Yellow");
            Type("Movies not in list: " + intMoviesNotInListCount, 0, 0, 1, "Blue");
        } // End RemoveMoviesFromTmdbList()


        protected static string AskForData(string questionToAsk)
        {
            bool invalidResponse = true;
            var response = "";
            do
            {
                Type(questionToAsk, 0, 0, 1, "Cyan");

                response = Console.ReadLine();

                if (!response.Equals(""))
                {
                    Type("Thank you!", 0, 0, 1, "Green");
                    invalidResponse = false;
                } else
                {
                    Type("Invalid response, try again..", 0, 0, 1, "Red");
                }

            } while (invalidResponse);
            
            return response;
        }

        protected static void MoveKidsMovies(IList<IList<Object>> data, Dictionary<string, int> sheetVariables)
        {
            Type("We will now start moving the kids movies around...", 0, 0, 1, "Gray");
            int moviesMovedCount = 0, moviesNotFoundCount = 0, moviesNotMovedCount = 0, moviesSkippedCount = 0;

            var driveLetter = AskForData("First I need to know what hard drive to look at:");

            foreach (var row in data)
            {
                if (row.Count > 20) // If it's an empty row then it should have less than this.
                {
                    try
                    {
                        // driverLetters[] \ baseFolderLocation \ kidsOrMovieFolder \ cleanTitle
                        // P:\Temp Movies\Kids Movies\A Bug's Life (1998) -- Temp Kids Movie.

                        string kidsOrMovieFolder = "", replaceText = "", sourceDirectory = "";
                        var kids = row[Convert.ToInt16(sheetVariables["Kids"])].ToString();
                        var movieLetter = row[Convert.ToInt16(sheetVariables["Movie Letter"])].ToString();
                        var movieDirectory = row[Convert.ToInt16(sheetVariables["Directory"])].ToString();
                        var status = row[Convert.ToInt16(sheetVariables["Status"])].ToString();
                        var fullLocationPath = driveLetter + movieDirectory;

                        if (kids.ToUpper() == "X")
                        {
                            kidsOrMovieFolder = "\\Kids Movies\\";
                            replaceText = "\\" + movieLetter + "\\";
                        }
                        else
                        {
                            kidsOrMovieFolder = "\\" + movieLetter + "\\";
                            replaceText = "\\Kids Movies\\";
                        }

                        if (!status.Equals("") && status[0].ToString().ToUpper() != "X") // If the first letter of status is an 'X' or empty then don't even look for the directory.
                        {
                            if (!Directory.Exists(fullLocationPath))
                            {
                                Type("We did not find: " + fullLocationPath, 0, 0, 1, "Yellow");
                                Type("We will now look for the Directory in the other folder to move it.", 0, 0, 1, "Yellow");

                                sourceDirectory = fullLocationPath.Replace(kidsOrMovieFolder, replaceText);
                                Type("Now searching here: " + sourceDirectory, 0, 0, 1, "Yellow");

                                if (!Directory.Exists(sourceDirectory))
                                {
                                    Type("We did not find the Directory in the other folder either.", 0, 0, 1, "Red");
                                    moviesNotFoundCount++;
                                }
                                else
                                {
                                    MoveDirectory(sourceDirectory, fullLocationPath);
                                    Type("Moved movie to: " + fullLocationPath, 0, 0, 1, "Green");
                                    moviesMovedCount++;
                                }
                            }
                            else moviesNotMovedCount++;
                        }
                        else moviesSkippedCount++;

                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong..." + e.Message, 3, 100, 1, "Red");
                    }

                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("Movies moved: " + moviesMovedCount, 0, 0, 1, "Green");
            Type("Movies skipped due to Status: " + moviesSkippedCount, 0, 0, 1, "Yellow");
            Type("Movies not found: " + moviesNotFoundCount, 0, 0, 1, "Red");
            Type("Movies not needing to move: " + moviesNotMovedCount, 0, 0, 1, "Blue");

        } // End MoveKidsMovies()

        protected static void InputTmdbId(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, int type)
        {
            int intTmdbIdDoneCount = 0, intTmdbIdCorrectedCount = 0, intTmdbIdSkippedCount = 0, intTmdbIdNotFoundCount = 0, intRowNum = 3;

            string text = "", tmdbIdValue = "", ImdbId = "", ImdbTitle = "", tmdbId = "", strCellToPutData = "";
            int tmdbIdColumnNum = 0; // Used to input the returned ID back into the Google Sheet.
            dynamic tmdbResponse;

            foreach (var row in data)
            {
                if (row.Count > 20) // If it's an empty row then it should have less than this.
                {
                    try
                    {
                        bool responseIsBroken = true;
                        tmdbIdValue = row[Convert.ToInt16(sheetVariables["TMDB ID"])].ToString();
                        tmdbIdColumnNum = Convert.ToInt16(sheetVariables["TMDB ID"]);
                        ImdbId = row[Convert.ToInt16(sheetVariables["IMDB ID"])].ToString();
                        ImdbTitle = row[Convert.ToInt16(sheetVariables["IMDB Title"])].ToString();

                        if (type == 1) // Input only missing TMDB IDs.
                        {
                            if (tmdbIdValue.Equals(""))
                            {
                                do
                                {
                                    tmdbResponse = TmdbApi.MoviesGetDetails(ImdbId);

                                    if (tmdbResponse.id != null)
                                    {
                                        tmdbId = tmdbResponse.id.ToString();
                                        responseIsBroken = false;
                                    }
                                    else if (tmdbResponse.status_message != null)
                                    {
                                        Type(ImdbTitle + " | " + tmdbResponse.status_message, 0, 0, 1, "Red");
                                        tmdbId = "";
                                        responseIsBroken = false;
                                    }
                                    else
                                    {
                                        Thread.Sleep(5000);
                                    }
                                } while (responseIsBroken);

                                strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbIdColumnNum) + intRowNum;

                                if (tmdbId != "")
                                {
                                    if (WriteSingleCellToSheet(tmdbId, strCellToPutData))
                                    {
                                        Type("TMDB ID saved for: " + ImdbTitle, 0, 0, 1, "Green");
                                        intTmdbIdDoneCount++;
                                        text += ImdbTitle + " Success, ";
                                    }
                                    else
                                    {
                                        Type("An error occured!", 0, 0, 1, "Red");
                                        text += ImdbTitle + " Failed, ";
                                    }
                                } else
                                {
                                    intTmdbIdNotFoundCount++;
                                }
                            }
                            
                        }
                        else if (type == 2) // Input ALL TMDB IDs including fixing wrong ones.
                        {
                            do
                            {
                                tmdbResponse = TmdbApi.MoviesGetDetails(ImdbId);

                                if (tmdbResponse.id != null)
                                {
                                    tmdbId = tmdbResponse.id.ToString();
                                    responseIsBroken = false;
                                }
                                else if (tmdbResponse.status_message != null)
                                {
                                    Type(ImdbTitle + " | " + tmdbResponse.status_message, 0, 0, 1, "Red");
                                    tmdbId = "";
                                    responseIsBroken = false;
                                }
                                else
                                {
                                    Thread.Sleep(5000);
                                }
                            } while (responseIsBroken);


                            strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbIdColumnNum) + intRowNum;

                            if (tmdbId != "")
                            {
                                if (tmdbIdValue.Equals("")) // If the ID is missing insert it.
                                {
                                    if (WriteSingleCellToSheet(tmdbId, strCellToPutData))
                                    {
                                        Type("TMDB ID saved for: " + ImdbTitle, 0, 0, 1, "Green");
                                        intTmdbIdDoneCount++;
                                        text += ImdbTitle + " Success, ";
                                    }
                                    else
                                    {
                                        Type("An error occured!", 0, 0, 1, "Red");
                                        text += ImdbTitle + " Failed, ";
                                    }

                                }
                                else if (tmdbIdValue != tmdbId) // Or if the new ID doesn't equal the old one overwrite it.
                                {
                                    if (WriteSingleCellToSheet(tmdbId, strCellToPutData))
                                    {
                                        Type("TMDB ID corrected for: " + ImdbTitle, 0, 0, 1, "Blue");
                                        intTmdbIdCorrectedCount++;
                                        text += ImdbTitle + " Success, ";
                                    }
                                    else
                                    {
                                        Type("An error occured!", 0, 0, 1, "Red");
                                        text += ImdbTitle + " Failed, ";
                                    }

                                }
                                else // Else just skip it.
                                {
                                    intTmdbIdSkippedCount++;
                                    text += ImdbTitle + " Skipped, ";
                                }

                            }
                            else
                            {
                                Type("We didn't find a TMDB ID for: " + ImdbTitle, 0, 0, 1, "Yellow");
                            }
                        }

                        intRowNum++;

                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong with " + ImdbTitle + " | " + e.Message, 3, 100, 1, "Red");
                        Type("Here's what I have: " + text, 0, 0, 1, "Blue"); // What was I thinking?!? Remove all this!
                    }

                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("TMDB IDs inserted: " + intTmdbIdDoneCount, 0, 0, 1, "Green");
            Type("TMDB IDs skipped: " + intTmdbIdSkippedCount, 0, 0, 1, "Yellow");
            Type("TMDB IDs corrected: " + intTmdbIdCorrectedCount, 0, 0, 1, "Blue");
            Type("TMDB IDs not available: " + intTmdbIdNotFoundCount, 0, 0, 1, "Red");

        } // End InputTmdbId()

        protected static void MoveMovies(IList<IList<Object>> data, Dictionary<string, int> sheetVariables)
        {
            int intDirectoriesMoviedCount = 0, intDirectoriesSkippedCount = 0;

            string  oldDirectory = "", newDirectory = "";

            foreach (var row in data)
            {
                if (row.Count > 20) // If it's an empty row then it should have less than this.
                {
                    try
                    {
                        oldDirectory = row[Convert.ToInt16(sheetVariables["Old Directory"])].ToString();
                        newDirectory = row[Convert.ToInt16(sheetVariables["Directory"])].ToString();

                        if (Directory.Exists(oldDirectory))
                        {
                            MoveDirectory(oldDirectory, newDirectory);
                            intDirectoriesMoviedCount++;
                        } else
                        {
                            intDirectoriesSkippedCount++;
                        }

                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong moving " + oldDirectory + " to " + newDirectory + " | " + e.Message, 3, 100, 1, "Red");
                    }

                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("Directories moved: " + intDirectoriesMoviedCount, 0, 0, 1, "Green");
            Type("Directories skipped: " + intDirectoriesSkippedCount, 0, 0, 1, "Yellow");
        }

        protected static void MoveDirectory(string source, string destination)
        {
            try
            {
                //if (!Directory.Exists(destination)) Directory.CreateDirectory(destination);
                Directory.Move(source, destination);
                Type("Moved ", 0, 0, 0);
                Type(source, 0, 0, 1, "Blue");
                Type(" to ", 0, 0, 0);
                Type(destination, 0, 0, 1, "Blue");
            }
            catch (Exception exp)
            {
                Console.WriteLine(exp.Message);
            }
        } // End MoveDirectory()

        protected static void CopyFile(string source, string destination)
        {
            try
            {
                File.Copy(source, destination);
                //Type("Copied ", 0, 0, 0);
                //Type(source, 0, 0, 1, "Green");
                //Type(" to ", 0, 0, 0);
                //Type(destination, 0, 0, 1, "Green");
            }
            catch (Exception e)
            {
                DisplayMessage("error", "An error occured while trying to copy the file. | ");
                DisplayMessage("harderror", e.Message);
                throw;
            }
        } // End MoveDirectory()

        /// <summary>
        /// Takes a directory location without the drive letter and then searches for that directory across all drives to find the current location.
        /// NOTE: This won't work if there are multiple hard drives with the same directory location. It will end up just returning the last hard drive found.
        /// </summary>
        /// <param name="directoryLocation">The directory location without the preceding drive letter.</param>
        /// <returns>The drive letter that has that directory.</returns>
        protected static String FindDriveLetter(String directoryLocation)
        {
            string[] driveLetters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            String driveLetter = "";
            foreach (var letter in driveLetters)
            {
                string withDriveLetter = letter + directoryLocation;
                if (Directory.Exists(withDriveLetter))
                {
                    driveLetter = letter;
                }
            }

            return driveLetter;
        } // End FindDirectory()

        //protected static void MarkMoviesAsOld()
        //{
        //    // Declare variables.
        //    UserCredential credential;
        //    Dictionary<string, int> SheetVariables = new Dictionary<string, int>
        //    {
        //        { "Old", -1 },
        //        { "Directory", -1 },
        //        { "Clean Title", -1 }
        //    };
        //    Dictionary<string, int> NotFoundColumns = new Dictionary<string, int>();

        //    GetTitleRowData(ref SheetVariables, MOVIES_TITLE_RANGE);
        //    bool lessThanZero = CheckColumns(ref NotFoundColumns, SheetVariables);

        //    if (lessThanZero)
        //    {
        //        Type("We didn't find a column that we were looking for...", 1, 100, 1, "Red");
        //        foreach (KeyValuePair<string, int> variable in NotFoundColumns)
        //        {
        //            Type("Key: " + variable.Key.ToString() + ", Value: " + variable.Value.ToString(), 1, 100, 1, "Red");

        //        }
        //        Console.WriteLine();
        //    }
        //    else
        //    {
        //        using (var stream =
        //            new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
        //        {
        //            string credPath = "token.json";
        //            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
        //                GoogleClientSecrets.Load(stream).Secrets,
        //                SCOPES,
        //                "user",
        //                CancellationToken.None,
        //                new FileDataStore(credPath, true)).Result;
        //        }

        //        // Create Google Sheets API service.
        //        var service = new SheetsService(new BaseClientService.Initializer()
        //        {
        //            HttpClientInitializer = credential,
        //            ApplicationName = APLICATION_NAME,
        //        });

        //        SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
        //                service.Spreadsheets.Values.Get(SPREADSHEET_ID, MOVIES_DATA_RANGE);

        //        ValueRange dataRowResponse = dataRowRequest.Execute();
        //        IList<IList<Object>> dataValues = dataRowResponse.Values;
        //        if (dataValues != null)
        //        {
        //            foreach (var row in dataValues)
        //            {
        //                if (row.Count > 20) // If it's an empty row then it should have less than this.
        //                {
        //                    tryMoveKidsMovies
        //                    {
        //                        if (row[Convert.ToInt16(SheetVariables["Old"])].ToString() != "") // Check that the movie is marked.
        //                        {
        //                            string OldFileLocation = row[Convert.ToInt16(SheetVariables["Directory"])].ToString() + "\\" + row[Convert.ToInt16(SheetVariables["Clean Title"])].ToString() + ".mkv";
        //                            string NewFileLocation = row[Convert.ToInt16(SheetVariables["Directory"])].ToString() + "\\" + row[Convert.ToInt16(SheetVariables["Clean Title"])].ToString() + "_OLD.mkv";

        //                            if (File.Exists(OldFileLocation))
        //                            {
        //                                File.Move(OldFileLocation, NewFileLocation);

        //                                Type(row[Convert.ToInt16(SheetVariables["Clean Title"])].ToString() + " has been renamed.", 0, 0, 1, "Green");
        //                            }
        //                            else
        //                            {
        //                                Type(row[Convert.ToInt16(SheetVariables["Clean Title"])].ToString() + " was set to be renamed, but doesn't exist.", 0, 0, 1,"Yellow");
        //                            }

        //                        }

        //                    }
        //                    catch (Exception e)
        //                    {
        //                        Type("Something went wrong..." + e.Message, 3, 100, 1, "Red");
        //                    }

        //                }
        //            }

        //        }
        //        else
        //        {
        //            Console.WriteLine("No data found.");
        //        }
        //        Type("It looks like that's the end of it.", 3, 100, 2);
        //    }
        //} // End MarkMoviesAsOld()

        /// <summary>
        /// Simply tells the user we didn't understand their request.
        /// </summary>
        /// <param name="choice">The users input.</param>
        protected static void DidntUnderstand(string choice)
        {
            Type("I'm sorry I didn't quite understand " + choice + ".", 14, 100, 1);
            Type("Please make sure your choice matches an option exactly from the menu.", 14, 100, 2);
        }

        /// <summary>
        /// Steps through the given data and determines which types of NFO Files to write then sends them to be written.
        /// </summary>
        /// <param name="data">The movie data to be stepped through.</param>
        /// <param name="sheetVariables">The dictionary that holds the column data.</param>
        /// <param name="type">The type of NFO file to write: 1 = ALL movies, 2 = Only selected movies, 3 = Only missing NFO Files.</param>
        /// <param name="isYouTubeFile">For the YouTube filenames we need to trim the title so we don't run into the character limit issue.</param>
        protected static void CreateNfoFiles(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, int type, bool isYouTubeFile = false)
        {
            int nfoFileNotFoundCount = 0, nfoFileOverwrittenCount = 0, nfoFileCreatedCount = 0;

            foreach (var row in data)
            {
                if (row.Count > 20)
                {
                    var directoryFound = false;
                    var cleanTitle = row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString();
                    var movieDirectory = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString();
                    var nfoBody = row[Convert.ToInt16(sheetVariables[NFO_BODY])].ToString();
                    var status = "";
                    var quickCreate = "";
                    var quickCreateInt = 0;

                    // If we are creating NFO files for YouTube videos then we need to trim the titles,
                    // also, we don't need to worry about checking for status.
                    if (isYouTubeFile)
                    {
                        if (cleanTitle.Length > 20)
                        {
                            cleanTitle = cleanTitle.Substring(0, 20).Trim();
                        }
                    } else
                    {
                        status = row[Convert.ToInt16(sheetVariables[STATUS])].ToString();
                    }

                    try
                    {

                        if (sheetVariables.ContainsKey("Quick Create") && row.Count > Convert.ToInt16(sheetVariables["Quick Create"]))
                        {
                            quickCreate = row[Convert.ToInt16(sheetVariables[QUICK_CREATE])].ToString();
                            quickCreateInt = Convert.ToInt16(sheetVariables[QUICK_CREATE]);
                        }


                        if (isYouTubeFile || (!status.Equals("") && status[0].ToString().ToUpper() != "X"))
                        {
                            // Let's go ahead and look for the hard drive letter now.
                            var hardDriveLetter = FindDriveLetter(movieDirectory);

                            if (hardDriveLetter != "")
                            {
                                // Now that we found the hard drive letter let's create the full path variable to create the directory.
                                var pathWithDriveLetter = hardDriveLetter + movieDirectory;
                                Directory.CreateDirectory(pathWithDriveLetter);

                                directoryFound = true;
                                string fileLocation = pathWithDriveLetter + "\\" + cleanTitle + ".nfo";

                                if (type == 1) // All movies, overwrite old NFO files AND put in new ones, but only if the folder exists (I don't want folders with only NFO files sitting in them).
                                {

                                    if (File.Exists(fileLocation))
                                    {
                                        File.Delete(fileLocation);
                                        nfoFileOverwrittenCount++;
                                        Type("NFO overwritten at: " + fileLocation, 0, 0, 1, "Blue");

                                    }
                                    else
                                    {
                                        nfoFileCreatedCount++;
                                        Type("NFO created at: " + fileLocation, 0, 0, 1, "Green");
                                    }
                                    WriteNfoFile(fileLocation, nfoBody);

                                }
                                else if (type == 2) // Only selected movies marked with an x.
                                {
                                    if (row.Count > quickCreateInt && quickCreate.ToUpper() == "X")
                                    {
                                        WriteNfoFile(fileLocation, nfoBody);

                                        if (File.Exists(fileLocation))
                                        {
                                            nfoFileOverwrittenCount++;
                                            Type("NFO overwritten at: " + fileLocation, 0, 0, 1, "Blue");

                                        }
                                        else
                                        {
                                            nfoFileCreatedCount++;
                                            Type("NFO created at: " + fileLocation, 0, 0, 1, "Green");
                                        }

                                    }

                                }
                                else if (type == 3) // Only the movies that are missing NFO files.
                                {
                                    if (!File.Exists(fileLocation))
                                    {
                                        WriteNfoFile(fileLocation, nfoBody);
                                        nfoFileCreatedCount++;
                                        Type("NFO created at: " + fileLocation, 0, 0, 1, "Green");
                                    }
                                }

                            }
                            else
                            {
                                directoryFound = true; // However, it will still try to spit out that it couldn't find the directory, so just set it to true.
                                Type("We did not find the hard drive for: " + movieDirectory, 0, 0, 1, "Red");
                                nfoFileNotFoundCount++;
                            }

                            if (!directoryFound)
                            {
                                Type("We did not find the directory for: " + movieDirectory, 0, 0, 1, "Red");
                                nfoFileNotFoundCount++;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong when looking for: " + movieDirectory + " | " + e.Message, 3, 100, 1, "Red");
                        throw;
                    }

                }
            }

            // Print out results.
            Type("It looks like that's the end of it.", 3, 100, 2);
            Type("NFO Files not found: ", 0, 0, 0); Type(nfoFileNotFoundCount.ToString(), 0, 0, 1, "Red");
            Type("NFO Files overwritten: ", 0, 0, 0); Type(nfoFileOverwrittenCount.ToString(), 0, 0, 1, "Blue");
            Type("NFO Files created: ", 0, 0, 0); Type(nfoFileCreatedCount.ToString(), 0, 0, 2, "Green");
            
        } // End CreateNfoFiles()

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sType"></param>
        /// <param name="onlyQuick"></param>
        //protected static void CheckForMovie(string sType)
        //{
        //    // Declare variables.
        //    UserCredential credential;
        //    string[] driveLetters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
        //    int directoriesFoundCount = 0, directoriesNotFoundCount = 0, intRowNum = STARTING_ROW_NUMBER;
        //    Dictionary<string, int> SheetVariables = new Dictionary<string, int>
        //    {
        //        { "Clean Title", -1 },
        //        { "Movie Letter", -1 },
        //        { "Ownership", -1 },
        //        { "Status", -1 }
        //    };
        //    Dictionary<string, int> NotFoundColumns = new Dictionary<string, int>();
        //    string sheetTitleRange = "", sheetDataRange = "", baseFolderLocation = "";
        //    bool lessThanZero = false;

        //    if (sType.ToUpper() == "MAIN")
        //    {
        //        sheetTitleRange = MOVIES_TITLE_RANGE;
        //        sheetDataRange = MOVIES_DATA_RANGE;
        //        baseFolderLocation = ":\\Movies\\";

        //    }
        //    else if (sType.ToUpper() == "TEMP")
        //    {
        //        sheetTitleRange = TEMP_MOVIES_TITLE_RANGE;
        //        sheetDataRange = TEMP_MOVIES_DATA_RANGE;
        //        baseFolderLocation = ":\\Temp Movies\\";

        //    }

        //    GetTitleRowData(ref SheetVariables, sheetTitleRange);
        //    lessThanZero = CheckColumns(ref NotFoundColumns, SheetVariables);

        //    if (lessThanZero)
        //    {
        //        Type("We didn't find a column that we were looking for...", 1, 100, 1, "Red");
        //        foreach (KeyValuePair<string, int> variable in NotFoundColumns)
        //        {
        //            Type("Key: " + variable.Key.ToString() + ", Value: " + variable.Value.ToString(), 1, 100, 1, "Red");
        //        }
        //        Console.WriteLine();
        //    }
        //    else
        //    {
        //        using (var stream =
        //            new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
        //        {
        //            string credPath = "token.json";
        //            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
        //                GoogleClientSecrets.Load(stream).Secrets,
        //                SCOPES,
        //                "user",
        //                CancellationToken.None,
        //                new FileDataStore(credPath, true)).Result;
        //        }

        //        // Create Google Sheets API service.
        //        var service = new SheetsService(new BaseClientService.Initializer()
        //        {
        //            HttpClientInitializer = credential,
        //            ApplicationName = APLICATION_NAME,
        //        });

        //        SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
        //                service.Spreadsheets.Values.Get(SPREADSHEET_ID, sheetDataRange);

        //        ValueRange dataRowResponse = dataRowRequest.Execute();
        //        IList<IList<Object>> dataValues = dataRowResponse.Values;
        //        if (dataValues != null)
        //        {
        //            foreach (var row in dataValues)
        //            {
        //                if (row.Count > 20) // If it's an empty row then it should have less than this.
        //                {
        //                    //Type("Row Count: " + row.Count.ToString() + ", Quick Create Column: " + Convert.ToInt16(SheetVariables["Quick Create"]), 0, 0, 1, "DarkGray");
        //                    try
        //                    {
        //                        string DirectoryLocation = baseFolderLocation + row[Convert.ToInt16(SheetVariables["Movie Letter"])].ToString() + "\\" + row[Convert.ToInt16(SheetVariables["Clean Title"])].ToString();
        //                        var directoryFound = false;
        //                        var ownership = row[Convert.ToInt16(SheetVariables["Ownership"])].ToString();
        //                        string strCellToSaveData = "Movies!" + ColumnNumToLetter(SheetVariables["Status"]) + intRowNum;

        //                        if ((sType.ToUpper() == "MAIN" && ownership.ToUpper() == "O") || (sType.ToUpper() == "TEMP" && ownership.ToUpper() == "T"))
        //                        {
        //                            foreach (var letter in driveLetters)
        //                            {
        //                                string withDriveLetter = letter + DirectoryLocation;
        //                                if (Directory.Exists(withDriveLetter))
        //                                {
        //                                    directoryFound = true;
        //                                    directoriesFoundCount++;
        //                                    WriteSingleCellToSheet("D", strCellToSaveData);
        //                                }
        //                            }

        //                            if (!directoryFound)
        //                            {
        //                                Type("We did not find the directory for: " + DirectoryLocation, 0, 0, 1, "Red");
        //                                directoriesNotFoundCount++;
        //                                WriteSingleCellToSheet("X", strCellToSaveData);
        //                            }
        //                            else
        //                            {
        //                                Type("We found the directory for: " + DirectoryLocation, 0, 0, 1, "Green");
        //                                NFO_FILE_NOT_FOUND_COUNT++;
        //                            }
        //                        }
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        Type("Something went wrong..." + e.Message, 3, 100, 1, "Red");
        //                    }

        //                }
        //                intRowNum++;
        //            }

        //        }
        //        else
        //        {
        //            Console.WriteLine("No data found.");
        //        }

        //        // Print out results.
        //        Type("It looks like that's the end of it.", 3, 100, 2);
        //        Type("Directories found: ", 0, 0, 0); Type(directoriesFoundCount.ToString(), 0, 0, 1, "Green");
        //        Type("Directories not found: ", 0, 0, 0); Type(directoriesNotFoundCount.ToString(), 0, 0, 1, "Red");

        //    }
        //} // End CheckForMovie()

        public static IList<IList<Object>> CallGetData(Dictionary<string, int> sheetVariables, string titleRowDataRange, string mainDataRange)
        {
            Type("Gathering sheet data... ", 10, 0, 0, "Yellow");
            // Get the title row data.
            IList<IList<Object>> titleData = GetData(titleRowDataRange);
            IList<IList<Object>> movieData = new List<IList<object>> { };

            // Update the sheetVariables Dictionary variable with the correct column number.
            sheetVariables = UpdateSheetVariables(sheetVariables, titleData);

            // Check the dictionary to verify that every column was found.
            // If a column wasn't found then one of the values will still be -1.
            Dictionary<string, int> NotFoundColumns = CheckColumns(sheetVariables);

            // If NotFoundColumns has data then we didn't find a column and need to show the user the missing column(s).
            // Then stop the program.
            if (NotFoundColumns.Count > 0)
            {
                Type("ERROR", 0, 0, 1, "Red");
                MissingColumn(NotFoundColumns);
            }
            else // Else it found the columns and can continue.
            {
                // Now that we have the Title Row Data, let's get the actual movie data to return.
                movieData = GetData(mainDataRange);
                Type("DONE", 100, 0, 1, "Green");
            }

            return movieData;
        }

        /// <summary>
        /// Grabs the data from the Google Sheet.
        /// Used for both the title row data, and the main data.
        /// </summary>
        /// <param name="sheetDataRange">The range in the sheet to pull data from.</param>
        /// <returns>The data from the selected range.</returns>
        public static IList<IList<Object>> GetData(string sheetDataRange)
        {
            try
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
                        service.Spreadsheets.Values.Get(SPREADSHEET_ID, sheetDataRange);

                ValueRange rowresponse = titleRowRequest.Execute();
                IList<IList<Object>> data = rowresponse.Values;
                return data;
            }
            catch (Exception ex)
            {
                Type("An error has occured: " + ex.Message, 0, 0, 1, "Red");
                throw;
            }

        } // End GetData()

        protected static Dictionary<string, int> CheckColumns(Dictionary<string, int> SheetVariables)
        {
            Dictionary<string, int> NotFoundColumns = new Dictionary<string, int>();

            foreach (KeyValuePair<string, int> variable in SheetVariables)
            {
                if (variable.Value < 0)
                {
                    NotFoundColumns.Add(variable.Key, variable.Value);
                }
            }
            return NotFoundColumns;
        } // End CheckColumns()

        /// <summary>
        /// Grabs the list of movies from the Google Sheet. Sends each IMDB ID to theMovieDB.org API to get the movie data.
        /// Inserts missing movie data into the Google Sheet (TMDB Rating, Plot, TMDB ID).
        /// </summary>
        /// <param name="data">The movie data to run through.</param>
        /// <param name="sheetVariables">The column names to look at.</param>
        protected static void InputMovieData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, Boolean overwiteData = false)
        {
            int intTmdbIdDoneCount = 0,
                intTmdbRatingDoneCount = 0,
                intPlotDoneCount = 0,
                intTmdbIdNotFoundCount = 0,
                intTmdbRatingNotFoundCount = 0,
                intPlotNotFoundCount = 0;

            string tmdbIdValue = "", // Our current TMDB ID value from the Google Sheet.
                rowNum = "", // Holds the row number we are on.
                tmdbRatingValue = "", // Our current TMDB Rating value from the Google Sheet.
                plotValue = "", // Our current Plot value from the Google Sheet.
                imdbId = "", // The IMDB ID from the Google Sheet to call the TMDB API with.
                imdbTitle = "", // The Title of the movie we are looking at from the Google Sheet to print out.
                tmdbId = "", // Holds the TMDB ID returned from our API call.
                tmdbRating = "", // Holds the TMDB Rating returned from our API call.
                tmdbPlot = "", // Holds the Plot returned from our API call.
                strCellToPutData = ""; // The string of the location to write the data to.

            int tmdbIdColumnNum = 0, // Used to input the returned ID back into the Google Sheet.
                tmdbRatingColumnNum = 0, // Used to input the returned rating into the Google Sheet.
                plotColumnNum = 0, // Used to input the returned overview into the Google Sheet.
                quickCreateColumnNum = 0; // Used to mark the movies that Plot gets updated.

            dynamic tmdbResponse; // The API call response.

            foreach (var row in data)
            {
                if (row.Count > 20) // If it's an empty row then it should have less than this.
                {
                    try
                    {
                        bool MovieFoundAtTmdb = true;
                        rowNum = row[Convert.ToInt16(sheetVariables["RowNum"])].ToString();
                        tmdbIdValue = row[Convert.ToInt16(sheetVariables["TMDB ID"])].ToString();
                        tmdbIdColumnNum = Convert.ToInt16(sheetVariables["TMDB ID"]);
                        tmdbRatingValue = row[Convert.ToInt16(sheetVariables["TMDB Rating"])].ToString();
                        tmdbRatingColumnNum = Convert.ToInt16(sheetVariables["TMDB Rating"]);
                        plotValue = row[Convert.ToInt16(sheetVariables["Plot"])].ToString();
                        plotColumnNum = Convert.ToInt16(sheetVariables["Plot"]);
                        imdbId = row[Convert.ToInt16(sheetVariables["IMDB ID"])].ToString();
                        imdbTitle = row[Convert.ToInt16(sheetVariables["IMDB Title"])].ToString();
                        if (sheetVariables.ContainsKey("Quick Create")) quickCreateColumnNum = Convert.ToInt16(sheetVariables["Quick Create"]);

                        if (tmdbIdValue.Equals("") || tmdbRatingValue.Equals("") || plotValue.Equals("") || overwiteData)
                        {
                            tmdbResponse = TmdbApi.MoviesGetDetails(imdbId);
                            try
                            {
                                if (tmdbResponse.movie_results[0].id.Value.ToString() != "")
                                {
                                    tmdbId = tmdbResponse.movie_results[0].id.Value.ToString();
                                    tmdbRating = tmdbResponse.movie_results[0].vote_average.ToString();
                                    tmdbPlot = tmdbResponse.movie_results[0].overview.ToString();

                                    if (!tmdbRating.Contains(".")) tmdbRating += ".0";
                                }
                                else
                                {
                                    DisplayMessage("error", "No record adfasdfa found at TMDB for: ", 0);
                                    DisplayMessage("warning", imdbTitle);

                                    MovieFoundAtTmdb = false;
                                }
                            }
                            catch (Exception ex)
                            {
                                DisplayMessage("error", "No record found at TMDB for: ", 0);
                                DisplayMessage("warning", imdbTitle);

                                MovieFoundAtTmdb = false;
                            }

                            if (MovieFoundAtTmdb && (tmdbIdValue.Equals("") || (overwiteData && !tmdbIdValue.Equals(tmdbId))))
                            {
                                if (tmdbId != "")
                                {
                                    strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbIdColumnNum) + rowNum;

                                    if (WriteSingleCellToSheet(tmdbId, strCellToPutData))
                                    {
                                        if (tmdbIdValue.Equals("")) Type("TMDB ID saved for: " + imdbTitle, 0, 0, 1, "Green");
                                        else
                                        {
                                            DisplayMessage("success", "TMDB ID updated from '", 0);
                                            DisplayMessage("info", tmdbIdValue, 0);
                                            DisplayMessage("success", "' to '", 0);
                                            DisplayMessage("info", tmdbId, 0);
                                            DisplayMessage("success", "' for: ", 0);
                                            DisplayMessage("warning", imdbTitle);
                                        }
                                        intTmdbIdDoneCount++;
                                    }
                                    else
                                    {
                                        Type("An error occured writing the ID for: " + imdbTitle, 0, 0, 1, "Red");
                                    }
                                }
                                else
                                {
                                    intTmdbIdNotFoundCount++;
                                }
                            }

                            if (MovieFoundAtTmdb && (tmdbRatingValue.Equals("") || (overwiteData && !tmdbRatingValue.Equals(tmdbRating))))
                            {
                                if (tmdbRating != "")
                                {
                                    strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbRatingColumnNum) + rowNum;

                                    if (WriteSingleCellToSheet(tmdbRating, strCellToPutData))
                                    {
                                        if (tmdbRatingValue.Equals("")) Type("TMDB Rating saved for: " + imdbTitle, 0, 0, 1, "Green");
                                        else
                                        {
                                            DisplayMessage("success", "TMDB Rating updated from '", 0);
                                            DisplayMessage("info", tmdbRatingValue, 0);
                                            DisplayMessage("success", "' to '", 0);
                                            DisplayMessage("info", tmdbRating, 0);
                                            DisplayMessage("success", "' for: ", 0);
                                            DisplayMessage("warning", imdbTitle);
                                        }
                                        intTmdbRatingDoneCount++;
                                    }
                                    else
                                    {
                                        Type("An error occured writing the Rating for: " + imdbTitle, 0, 0, 1, "Red");
                                    }
                                }
                                else
                                {
                                    intTmdbRatingNotFoundCount++;
                                }
                            }

                            if (MovieFoundAtTmdb && (plotValue.Equals("") || (overwiteData && !plotValue.Equals(tmdbPlot))))
                            {
                                if (tmdbPlot != "")
                                {
                                    strCellToPutData = "Movies!" + ColumnNumToLetter(plotColumnNum) + rowNum;

                                    if (WriteSingleCellToSheet(tmdbPlot, strCellToPutData))
                                    {
                                        if (plotValue.Equals("")) Type("Plot saved for: " + imdbTitle, 0, 0, 1, "Green");
                                        else
                                        {
                                            DisplayMessage("success", "Plot updated for: ", 0);
                                            DisplayMessage("warning", imdbTitle);
                                            strCellToPutData = "Movies!" + ColumnNumToLetter(quickCreateColumnNum) + rowNum;
                                            WriteSingleCellToSheet("X", strCellToPutData);
                                        }
                                        intPlotDoneCount++;
                                    }
                                    else
                                    {
                                        Type("An error occured writing the Plot for: " + imdbTitle, 0, 0, 1, "Red");
                                    }
                                }
                                else
                                {
                                    intPlotNotFoundCount++;
                                }
                            }

                        }

                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong while putting in movie data for: " + imdbTitle, 3, 100, 1, "Red");
                        Type(e.Message, 3, 100, 2, "DarkRed");
                    }

                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("TMDB IDs inserted: " + intTmdbIdDoneCount, 0, 0, 1, "Green");
            Type("TMDB Ratings inserted: " + intTmdbRatingDoneCount, 0, 0, 1, "Green");
            Type("Plots inserted: " + intPlotDoneCount, 0, 0, 1, "Green");
            Type("TMDB IDs not available: " + intTmdbIdNotFoundCount, 0, 0, 1, "Red");
            Type("TMDB Ratings not available: " + intTmdbRatingNotFoundCount, 0, 0, 1, "Red");
            Type("Plots not available: " + intPlotNotFoundCount, 0, 0, 1, "Red");

        } // End InputMovieData()

        public static void Countdown(int time)
        {
            do
            {
                DisplayMessage("data", time.ToString(), 0, 0, 1000);
                ClearCurrentConsoleLine();
                time--;
            } while (time > 0);
        }

        /// <summary>
        /// Clears the current console line text.
        /// </summary>
        protected static void ClearCurrentConsoleLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }

        protected static void ConvertVideo(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string presetChoice)
        {
            // Declare variables.
            int intTotalMoviesCount = 0,
                intImagesCount = 0,
                intAlreadyConvertedFilesCount = 0,
                intNoTitleCount = 0,
                intConvertedFilesCount = 0;
            
            foreach (var row in data)
            {
                if (row[1].ToString() != "") // If it's an empty row then this cell should be empty.
                {
                    intTotalMoviesCount++;
                    string i = "",
                            o = "",
                            title = "",
                            additionalCommands = "",
                            chapter = "",
                            directoryLocation = "",
                            showTitle = "",
                            SeasonNum = "",
                            convertPath = "";
                    try
                    {
                        i = row[Convert.ToInt16(sheetVariables["ISO Input"])].ToString();
                        o = row[Convert.ToInt16(sheetVariables["Directory"])].ToString() + "\\" + row[Convert.ToInt16(sheetVariables["Clean Title"])].ToString() + ".mp4";
                        showTitle = row[Convert.ToInt16(sheetVariables["Show"])].ToString();
                        if (row[Convert.ToInt16(sheetVariables["Override Show"])].ToString() != "") showTitle = row[Convert.ToInt16(sheetVariables["Show"])].ToString();
                        SeasonNum = row[Convert.ToInt16(sheetVariables["Season #"])].ToString();
                        string pathRoot = Path.GetPathRoot(i.ToString()),
                            cleanTitle = row[Convert.ToInt16(sheetVariables["Clean Title"])].ToString();
                        convertPath = pathRoot + "These are finished running through HandBrake\\" + showTitle + "\\Season " + SeasonNum + "\\" + cleanTitle + ".mp4";
                        title = row[Convert.ToInt16(sheetVariables["ISO Title #"])].ToString();
                        additionalCommands = " " + row[Convert.ToInt16(sheetVariables["Additional Commands"])].ToString();
                        chapter = row[Convert.ToInt16(sheetVariables["ISO Ch #"])].ToString();
                        directoryLocation = row[Convert.ToInt16(sheetVariables["Directory"])].ToString();

                        if (File.Exists(i))
                        {
                            intImagesCount++;
                            if (File.Exists(o))
                            {
                                DisplayMessage("info", "We found: ", 0);
                                DisplayMessage("default", cleanTitle);
                                intAlreadyConvertedFilesCount++;
                            }
                            else if (File.Exists(convertPath))
                            {
                                DisplayMessage("info", "We found: ", 0);
                                DisplayMessage("default", cleanTitle);
                                intAlreadyConvertedFilesCount++;
                            }
                            else
                            {
                                DisplayMessage("warning", "We didn't find: ", 0);
                                DisplayMessage("default", cleanTitle);

                                if (title != "")
                                {
                                    Directory.CreateDirectory(pathRoot + "These are finished running through HandBrake");
                                    Directory.CreateDirectory(pathRoot + "These are finished running through HandBrake\\" + showTitle);
                                    Directory.CreateDirectory(pathRoot + "These are finished running through HandBrake\\" + showTitle + "\\Season " + SeasonNum);
                                    string strMyConversionString = "HandBrakeCLI -i \"" + i + "\" -o \"" + convertPath + "\" " + presetChoice + " -t " + title + additionalCommands;

                                    //Type("We will use title #" + title, 0, 0, 1, "blue");

                                    if (chapter != "")
                                    {
                                        Type("And we will use Chapter #" + chapter, 0, 0, 1, "blue");
                                        strMyConversionString += " -c " + chapter;
                                    }

                                    //Type("Here is our command: " + strMyConversionString, 0, 0, 1, "blue");

                                    HandBrake(strMyConversionString);
                                    intConvertedFilesCount++;
                                }
                                else
                                {
                                    Type("We don't have a title to go off of.", 0, 0, 1, "gray");
                                    intNoTitleCount++;
                                }
                                Type("-------------------------------------------------------------------", 0, 0, 1);
                            }
                        }
                        else
                        {
                            //Type("We didn't find " + i, 0, 0, 1, "yellow");
                            //Type("We won't be able to convert this one at this time.", 0, 0, 1, "yellow");
                            //Type("-------------------------------------------------------------------", 0, 0, 1);
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong converting the following video: " + title, 3, 100, 1, "Red");
                        Type(e.Message, 3, 100, 2, "DarkRed");
                        break;
                    }
                }
            } // End foreach
            Type("-----SUMMARY-----", 7, 100, 1);
            Type(intTotalMoviesCount + " total episodes in list to convert.", 7, 100, 1);
            Type(intImagesCount + " disc images found.", 7, 100, 1);
            Type(intAlreadyConvertedFilesCount + " episodes already converted and were skipped.", 7, 100, 1);
            Type(intConvertedFilesCount + " episodes converted.", 7, 100, 1);
            Type(intNoTitleCount + " missing titles to convert.", 7, 100, 1);

            Type("It looks like that's the end of it.", 3, 100, 2, "magenta");
        } // End ConvertVideo()

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

        protected static void HandBrake(string command, int count = -1)
        {
            try
            {
                //Type("Here is our command:", 0, 0, 1, "Blue");
                //Type(command.ToString(), 0, 0, 1, "DarkBlue");

                // create the ProcessStartInfo using "cmd" as the program to be run,
                // and "/c " as the parameters.
                // Incidentally, /c tells cmd that we want it to execute the command that follows,
                // and then exit.
                ProcessStartInfo procStartInfo =
                    new ProcessStartInfo("cmd", "/c " + command)
                    {
                        // The following commands are needed to redirect the standard output.
                        // This means that it will be redirected to the Process.StandardOutput StreamReader.
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        // Do not create the black window.
                        CreateNoWindow = true
                    };
                // Now we create a process, assign its ProcessStartInfo and start it
                Process proc = new Process
                {
                    StartInfo = procStartInfo
                };
                DateTime startTime = DateTime.Now;
                Type("Start Time: ", 0, 0, 0, "Blue");
                Type(startTime.ToString("MM/dd/yyyy, h:mm:ss tt"), 0, 0, 1, "Green");
                
                proc.Start();
                // Get the output into a string
                string result = proc.StandardOutput.ReadToEnd();

                // Display the command output.
                //Console.WriteLine(result);

                // Display the time the conversion ended.
                DateTime endTime = DateTime.Now;
                Type("End Time: ", 0, 0, 0, "Blue");
                Type(endTime.ToString("MM/dd/yyyy, h:mm:ss tt"), 0, 0, 1, "Cyan");

                // Display the amount of time that conversion took.
                TimeSpan duration = endTime - startTime;
                Type("Conversion duration: ", 0, 0, 0, "Blue");
                Type(duration.ToString(), 0, 0, 1, "Yellow");

                // Add the duration to display the total running time.
                runningTotalConversionTime += duration;
                Type("Total duration: ", 0, 0, 0, "Blue");
                Type(runningTotalConversionTime.ToString(), 0, 0, 1, "Cyan");

                // Add the duration to display the total session running time.
                sessionDuration += duration;
                Type("Total session duration: ", 0, 0, 0, "Blue");
                Type(sessionDuration.ToString(), 0, 0, 1, "Green");

                // If there are more than one file to convert guesstimate the amount of time remaining.
                if (count > 0)
                {
                    const int DAY = 60 * 24;
                    int daysRemaining = 0,
                        hoursRemaining = 0,
                        minutesRemaining = 0;
                    string ETR = "Roughly ";
                    double durationInMinutes = duration.TotalMinutes;
                    double timeLeftInMinutes = durationInMinutes * count;

                    if (timeLeftInMinutes > DAY)
                    {
                        daysRemaining = (int)timeLeftInMinutes / DAY;
                        ETR += daysRemaining.ToString() + (daysRemaining == 1 ? " day, " : " days, ");
                    }
                    if (timeLeftInMinutes > 60)
                    {
                        hoursRemaining = (daysRemaining > 0 ? ((int)timeLeftInMinutes - (daysRemaining * 1440)) / 60 : (int)timeLeftInMinutes / 60);
                        ETR += hoursRemaining.ToString() + (hoursRemaining == 1 ? " hour, " : " hours, ");
                    }
                        
                    minutesRemaining = (int)timeLeftInMinutes % 60;

                    ETR += minutesRemaining.ToString() + (minutesRemaining == 1 ? " minute remaining" : " minutes remaining");
                    Type("ETR: ", 0, 0, 0, "Blue");
                    Type(ETR, 0, 0, 1, "Red");
                    Type("(Based on the time it took to convert that last one)", 0, 0, 1);
                }

            }
            catch (Exception objException)
            {
                Type("Unable to convert file. | " + objException.Message, 7, 100, 1);
            }
        } // End HandBrake()

        //protected static void GetDataToConvertEpisodes(string itemType, string presetFile)
        //{
        //    UserCredential credential;
        //    Dictionary<string, int> SheetVariables = new Dictionary<string, int>
        //    {
        //        { "Image Location", -1 },
        //        { "Episode Location", -1 },
        //        { "ISO Title #", -1 },
        //        { "Chapter", -1 },
        //        { "Additional Commands", -1 }
        //    };

        //    string titleRange = "", dataRange = "";

        //    if (itemType == "main")
        //    {
        //        titleRange = EPISODES_TITLE_RANGE;
        //        dataRange = EPISODES_DATA_RANGE;
        //    }
        //    else if(itemType == "temp")
        //    {
        //        titleRange = TEMP_EPISODES_TITLE_RANGE;
        //        dataRange = TEMP_EPISODES_DATA_RANGE;
        //    }

        //    Dictionary<string, int> NotFoundColumns = new Dictionary<string, int>();
        //    bool lessThanZero = false;

        //    // Send the variables off to update the column numbers.
        //    GetTitleRowData(ref SheetVariables, titleRange);

        //    // Declare variables.
        //    int intInputFolderColumn = -1,
        //        intOutputFolderColumn = -1,
        //        intIsoTitleNumberColumn = -1,
        //        intChapterNumberColumn = -1,
        //        intTotalEpisodesCount = 0,
        //        intImagesCount = 0,
        //        intAlreadyConvertedFilesCount = 0,
        //        intNoTitleCount = 0,
        //        intConvertedFilesCount = 0;

        //    if (lessThanZero)
        //    {
        //        Type("We didn't find a column that we were looking for...", 1, 100, 1, "Red");
        //        foreach (KeyValuePair<string, int> variable in NotFoundColumns)
        //        {
        //            Type("Key: " + variable.Key.ToString() + ", Value: " + variable.Value.ToString(), 1, 100, 1, "Red");

        //        }
        //        Console.WriteLine();
        //    }
        //    else
        //    {

        //        using (var stream =
        //            new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
        //        {
        //            string credPath = "token.json";
        //            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
        //                GoogleClientSecrets.Load(stream).Secrets,
        //                SCOPES,
        //                "user",
        //                CancellationToken.None,
        //                new FileDataStore(credPath, true)).Result;
        //        }

        //        // Create Google Sheets API service.
        //        var service = new SheetsService(new BaseClientService.Initializer()
        //        {
        //            HttpClientInitializer = credential,
        //            ApplicationName = APLICATION_NAME,
        //        });

        //        SpreadsheetsResource.ValuesResource.GetRequest dataRowRequest =
        //                service.Spreadsheets.Values.Get(SPREADSHEET_ID, dataRange);

        //        ValueRange dataRowResponse = dataRowRequest.Execute();
        //        IList<IList<Object>> dataValues = dataRowResponse.Values;
        //        if (dataValues != null)
        //        {
        //            foreach (var row in dataValues)
        //            {
        //                //Console.WriteLine("Row count is: " + row.Count);
        //                if (row.Count > 19)
        //                {
        //                    intTotalEpisodesCount++;
        //                    try
        //                    {
        //                        string i = row[Convert.ToInt16(SheetVariables["Image Location"])].ToString(),
        //                                o = row[Convert.ToInt16(SheetVariables["Episode Location"])].ToString(),
        //                                title = row[Convert.ToInt16(SheetVariables["ISO Title #"])].ToString(),
        //                                additionalCommands = " " + row[Convert.ToInt16(SheetVariables["Additional Commands"])].ToString(),
        //                                chapter = row[Convert.ToInt16(SheetVariables["Chapter"])].ToString();

        //                        if (File.Exists(i))
        //                        {
        //                            //Type("We found " + i, 0, 0, 1);
        //                            intImagesCount++;
        //                            if (File.Exists(o))
        //                            {
        //                                //Type("We found " + o, 0, 0, 1);
        //                                //Type("We won't have to convert this one.", 0, 0, 1);
        //                                intAlreadyConvertedFilesCount++;
        //                            }
        //                            else
        //                            {
        //                                Type("We found " + i, 0, 0, 1, "green");
        //                                Type("We didn't find " + o, 0, 0, 1, "yellow");

        //                                // Create the directory if needed.
        //                                int lastIndexOf = o.LastIndexOf("\\");
        //                                string fileLocation = o.Substring(0, lastIndexOf);
        //                                Directory.CreateDirectory(fileLocation);
        //                                Type("Directory created at: " + fileLocation, 0, 0, 1, "darkgreen");

        //                                if (title != "")
        //                                {
        //                                    string strMyConversionString = "HandBrakeCLI -i \"" + i + "\" -o \"" + o + "\" " + presetFile + " -t " + title + additionalCommands;

        //                                    Type("We will use title #" + title, 0, 0, 1, "blue");

        //                                    if (chapter != "")
        //                                    {
        //                                        Type("And we will use Chapter #" + chapter, 0, 0, 1, "blue");
        //                                        strMyConversionString += " -c " + chapter;
        //                                    }

        //                                    Type(strMyConversionString, 0, 0, 1);
        //                                    HandBrake(strMyConversionString);
        //                                    intConvertedFilesCount++;
        //                                    Type("-------------------------------------------------------------------", 0, 0, 1);
        //                                }
        //                                else
        //                                {
        //                                    Type("We don't have a title to go off of.", 0, 0, 1, "gray");
        //                                    intNoTitleCount++;
        //                                }
        //                            }
        //                        }
        //                        else
        //                        {
        //                            Type("We didn't find " + i, 0, 0, 1, "yellow");
        //                            Type("We won't be able to convert this one at this time.", 0, 0, 1, "yellow");
        //                        }
        //                        //Type("-------------------------------------------------------------------", 0, 0, 1);

        //                    }
        //                    catch (Exception e)
        //                    {
        //                        Type("Something went wrong... | " + e.Message, 3, 100, 1);
        //                        Type("Row Count: " + row.Count().ToString(), 0, 0, 1);
        //                        break;
        //                    }
        //                }
        //            } // End foreach
        //            Type("-----SUMMARY-----", 7, 100, 1);
        //            Type(intTotalEpisodesCount + " Total Episodes.", 7, 100, 1);
        //            Type(intImagesCount + " Images Found.", 7, 100, 1);
        //            Type(intAlreadyConvertedFilesCount + " Episode Files Found.", 7, 100, 1);
        //            Type(intConvertedFilesCount + " Episodes converted.", 7, 100, 1);
        //            Type(intNoTitleCount + " Missing Titles to convert.", 7, 100, 1);
        //        }
        //        else
        //        {
        //            Console.WriteLine("No data found.");
        //        }
        //        Type("It looks like that's the end of it.", 3, 100, 2, "magenta");
        //    }
        //} // End GetDataToConvertEpisodes()

        protected static void CountFiles()
        {
            var missingDirectories = new List<List<string>>();
            Dictionary<int, string> missingDirectoriesList = new Dictionary<int, string> { };
            bool keepAskingForDirectory = true, keepAskingForList = true;
            do
            {
                ClearDirectories();

                var directory = AskForDirectory();

                //Type("Enter your directory", 7, 100, 1);
                //var directory = Console.ReadLine();
                string[] fileEntries = Directory.GetFiles(directory);
                string[] subdirectoryEntries = Directory.GetDirectories(directory);
                Type("The chosen directory contains " + subdirectoryEntries.Length + " sub folders and " + fileEntries.Length + " files.", 0, 100, 1);
                string directoryPlural = "";
                if (File.Exists(directory))
                {
                    // This path is a file
                    ProcessFile(directory);
                    keepAskingForDirectory = false;
                }
                else if (Directory.Exists(directory))
                {
                    int i = 1;
                    // This path is a directory
                    ProcessDirectory(directory);

                    if (missingNfo.Count() > 0)
                    {
                        directoryPlural = missingNfo.Count() == 1 ? " directory is " : " directories are ";
                        Type(missingNfo.Count().ToString() + directoryPlural + "missing NFO files.", 0, 0, 1, "DarkRed");
                        missingDirectories.Add(missingNfo);
                        missingDirectoriesList.Add(i, missingNfo.Count().ToString() + " Missing NFO");
                        i++;
                    }
                    if (missingJpg.Count() > 0)
                    {
                        directoryPlural = missingJpg.Count() == 1 ? " directory is " : " directories are ";
                        Type(missingJpg.Count().ToString() + directoryPlural + "missing JPG files.", 0, 0, 1, "DarkYellow");
                        missingDirectories.Add(missingJpg);
                        missingDirectoriesList.Add(i, missingJpg.Count().ToString() + " Missing JPG");
                        i++;
                    }
                    if (missingMovie.Count() > 0)
                    {
                        directoryPlural = missingMovie.Count() == 1 ? " directory is " : " directories are ";
                        Type(missingMovie.Count().ToString() + directoryPlural + "missing Movie files.", 0, 0, 1, "Blue");
                        missingDirectories.Add(missingMovie);
                        missingDirectoriesList.Add(i, missingMovie.Count().ToString() + " Missing Movie");
                        i++;
                    }
                    if (missingIso.Count() > 0)
                    {
                        directoryPlural = missingIso.Count() == 1 ? " directory is " : " directories are ";
                        Type(missingIso.Count().ToString() + directoryPlural + "missing ISO files.", 0, 0, 1, "DarkCyan");
                        missingDirectories.Add(missingIso);
                        missingDirectoriesList.Add(i, missingIso.Count().ToString() + " Missing ISO");
                        i++;
                    }
                    if (partFiles.Count() > 0)
                    {
                        directoryPlural = partFiles.Count() == 1 ? " directory has " : " directories have ";
                        Type(partFiles.Count().ToString() + directoryPlural + "a part of a file.", 0, 0, 1, "DarkCyan");
                        missingDirectories.Add(partFiles);
                        missingDirectoriesList.Add(i, partFiles.Count().ToString() + " Part File");
                        i++;
                    }
                    if (emptyDirectory.Count() > 0)
                    {
                        directoryPlural = emptyDirectory.Count() == 1 ? " directory was " : " directories were ";
                        Type(emptyDirectory.Count().ToString() + directoryPlural + "empty.", 0, 0, 1, "Magenta");
                        missingDirectories.Add(emptyDirectory);
                        missingDirectoriesList.Add(i, "Empty Directory");
                        i++;
                    }

                    keepAskingForDirectory = false;
                }
                else
                {
                    Type(directory + " is not a valid file or directory.", 14, 100, 1);
                }
            } while (keepAskingForDirectory);

            if (missingDirectoriesList.Count() > 0)
            {
                do
                {
                    Console.WriteLine();
                    Type("Input the number to list the corresponding info.", 0, 0, 1, "Blue");
                    Type("Press 0 to exit.", 0, 0, 1);

                    foreach (KeyValuePair<int, string> kvp in missingDirectoriesList)
                    {
                        Type("Enter " + kvp.Key + " to view the " + kvp.Value + " list.", 0, 0, 1);
                    }

                    var response = Console.ReadLine();

                    if (response == "0")
                    {
                        keepAskingForList = false;
                        missingDirectories.Clear();
                        missingDirectoriesList.Clear();
                    }
                    else if (missingDirectoriesList.ContainsKey(int.Parse(response)))
                    {
                        Type("The following directories are " + missingDirectoriesList[int.Parse(response)] + " files:", 0, 0, 1);
                        foreach (var item in missingDirectories[int.Parse(response) - 1])
                        {
                            Type(item, 0, 0, 1, "Cyan");
                        }
                        AskForMenu();
                    }
                    else
                    {
                        Type("The list doesn't contain that option, try again.", 0, 0, 1);
                    }
                } while (keepAskingForList);
            }
            else
            {
                Console.WriteLine();
                Type("We didn't find anything out of place", 0, 0, 1, "Green");
            }

        } // End CountFiles()

        protected static void DisplayResults(Dictionary<string, int> results)
        {
            var fontColor = "";
            var i = 0;
            Type("---SUMMARY---", 0, 0, 1, "Magenta");
            foreach (KeyValuePair<string, int> variable in results)
            {
                if (i == 0)
                {
                    fontColor = "Green";
                } else if (i == 1)
                {
                    fontColor = "Yellow";
                }
                 else if (i == 2)
                {
                    fontColor = "Red";
                }
                Type(variable.Key.ToString() + ": " + variable.Value.ToString(), 1, 100, 1, fontColor);
                i++;
            }
            //Type("It looks like that's the end of it.", 0, 0, 1);
        } // End DisplayResults()

        protected static string AskForDirectory(string message = "Enter your directory:")
        {
            bool keepAskingForDirectory = true;
            string directory;
            do
            {
                DisplayMessage("question", message + " (0 to cancel)");
                directory = RemoveCharFromString(Console.ReadLine(), '"');
                if (directory == "0")
                {
                    keepAskingForDirectory = false;
                } 
                else if (File.Exists(directory))
                {
                    Type("No, I need the path to a folder location, not a file.", 0, 0, 1, "Red");
                }
                else if (Directory.Exists(directory))
                {
                    keepAskingForDirectory = false;
                }
            } while (keepAskingForDirectory);

            return directory;
        } // End AskForDirectory()

        protected static void ConvertHandbrakeList(ArrayList videoFiles)
        {
            Type("Now converting " + videoFiles.Count + " files... ", 10, 0, 1, "Yellow");

            // An ArrayList to hold the files that have finished converting so that we can remove the metadata from them.
            ArrayList outputFiles = new ArrayList();

            try
            {
                int count = 1;
                foreach (var myFile in videoFiles)
                {
                    Type("Converting " + count + " of " + videoFiles.Count + " files", 0, 0, 1, "Blue");

                    string fileName = Path.GetFileName(myFile.ToString());
                    string pathRoot = Path.GetPathRoot(myFile.ToString());
                    string i = myFile.ToString(),
                            o = pathRoot + "These are finished running through HandBrake\\" + fileName,
                            presetChoice = "--preset-import-file MP4_RF22f.json -Z \"MP4 RF22f\"";


                    ArrayList inputArrayList = new ArrayList{i};
                    long sizeOfInputFile = SizeOfFiles(inputArrayList);
                    ArrayList outputArrayList = new ArrayList{o};
                    // Since the output file MAY not exist yet we wait to get the size of it.
                    long sizeOfOutputFile = 0;

                    if (!File.Exists(o))
                    {
                        outputFiles.Add(o);

                        string strMyConversionString = "HandBrakeCLI -i \"" + i + "\" -o \"" + o + "\" " + presetChoice;

                        Type("Now converting: " + fileName, 0, 0, 1, "Magenta");

                        HandBrake(strMyConversionString, videoFiles.Count - count);

                        // Now that the output file definitely exists we can grab the size of it.
                        sizeOfOutputFile = SizeOfFiles(outputArrayList);

                        // Display the amount of bytes that conversion saved.
                        DisplaySavings(sizeOfOutputFile, sizeOfInputFile);

                        // Remove the Metadata.
                        RemoveMetadata(outputFiles);

                        // Add a comment to the file.
                        DateTime convertedTime = DateTime.Now;
                        foreach (string file in outputFiles)
                        {
                            AddComment(file, "Recorded in HD, re-encoded with black bars.\nConverted on: " + convertedTime.ToString("MM/dd/yyyy"));
                        }

                        // Now clear the outputFiles arraylist.
                        outputFiles.Clear();

                    } else
                    {
                        // Now that the output file definitely exists we can grab the size of it.
                        sizeOfOutputFile = SizeOfFiles(outputArrayList);

                        // Display the amount of bytes that conversion saved.
                        DisplaySavings(sizeOfOutputFile, sizeOfInputFile);

                        Type(fileName + " already exists at destination. --Skipping to next file.", 0, 0, 1, "Yellow");
                    }
                    // Now delete the input file.
                    File.Delete(i);
                    DisplayMessage("info", "Input file deleted.");

                    count++;
                    DisplayEndOfCurrentProcessLines();
                }

                Type("DONE", 100, 0, 1, "Green");

            }
            catch (Exception e)
            {
                Type("Something happened | " + e.Message, 0, 0, 1, "Red");
                throw;
            }
        } // End ConvertHandbrakeList()

        protected static void DisplaySavings(long oFile, long iFile)
        {
            // Display the amount of bytes that conversion saved.
            long difference = iFile - oFile;
            //Console.WriteLine("iFile: " + iFile.ToString("N"));
            //Console.WriteLine("oFile: " + oFile.ToString("N"));
            //Console.WriteLine("difference: " + difference.ToString("N"));
            if(difference >= 0)
            {
                Type("Conversion savings: ", 0, 0, 0, "Blue");
                Type(FormatSize(difference) + " of " + FormatSize(iFile) + " -" + FormatPercentage(difference, iFile) + "%", 0, 0, 1, "Yellow");
            } else
            {
                Type("Conversion loss: ", 0, 0, 0, "Red");
                Type(FormatSize(difference * -1) + " more than " + FormatSize(iFile) + " +" + FormatPercentage(difference * -1, oFile) + "%", 0, 0, 1, "Yellow");
            }

            // Add the difference to display the total running difference in bytes.
            runningDifference += difference;
            runningFileSize += iFile;
            Type("Total savings: ", 0, 0, 0, "Blue");
            Type(FormatSize(runningDifference) + " of " + FormatSize(runningFileSize) + " " + FormatPercentage(runningDifference, runningFileSize) + "% saved", 0, 0, 1, "Cyan");

            // Add the difference to display the total session difference in bytes.
            runningSessionSavings += difference;
            runningSessionFileSize += iFile;
            Type("Session savings: ", 0, 0, 0, "Blue");
            Type(FormatSize(runningSessionSavings) + " of " + FormatSize(runningSessionFileSize) + " " + FormatPercentage(runningSessionSavings, runningSessionFileSize) + "% saved", 0, 0, 1, "Green");
        }

        protected static string FormatPercentage(long oFile, long iFile)
        {
            return ((decimal.Parse(oFile.ToString()) / decimal.Parse(iFile.ToString())) * 100).ToString("N2");
        }

        public static void DisplayEndOfCurrentProcessLines()
        {
            Type("-------------------------------------------------------------------", 0, 0, 2);
        }

        protected static void AddComment(string myFile, string comment)
        {
            try
            {
                using (TagLib.File videoFile = TagLib.File.Create(myFile))
                {
                    if (videoFile.Tag.Comment == null || videoFile.Tag.Comment == "")
                    {
                        videoFile.Tag.Comment = comment;
                        videoFile.Save();
                        Type("Comment added to: ", 0, 0, 0, "Yellow");
                        Type(Path.GetFileName(myFile.ToString()), 0, 0, 1, "Green");
                    } else
                    {
                        var oldComment = videoFile.Tag.Comment;
                        var newComment = oldComment + "\n" + comment;
                        videoFile.Tag.Comment = newComment;
                        videoFile.Save();
                        DisplayMessage("warning", "Comment Added onto: ", 0);
                        DisplayMessage("success", Path.GetFileName(myFile.ToString()));
                    }
                }

            }
            catch (Exception e)
            {
                Type("Something went wrong | " + e.Message, 0, 0, 1, "Red");
            }
            
        }

        protected static void RemoveMetadata(ArrayList videoFiles)
        {
            Type("Removing Metadata from the video files... ", 10, 0, 0, "Yellow");
            string performersRemovedCount = "Performers Removed Count", titlesRemovedCount = "Titles Removed Count", commentsRemovedCount = "Comments Removed Count";
            Dictionary<string, int> resultVariables = new Dictionary<string, int> { };
            resultVariables.Add(performersRemovedCount, 0);
            resultVariables.Add(titlesRemovedCount, 0);
            resultVariables.Add(commentsRemovedCount, 0);
                
            try
            {
                foreach (var myFile in videoFiles)
                {
                    if (myFile.ToString().ToUpper().Contains(".MP4") || myFile.ToString().ToUpper().Contains(".M4V"))
                    {
                        bool saveFile = false;
                        using (TagLib.File videoFile = TagLib.File.Create(myFile.ToString()))
                        {
                            if (videoFile.Tag.Performers.Length > 0)
                            {
                                videoFile.Tag.Performers = null;
                                resultVariables[performersRemovedCount] += 1;
                                saveFile = true;
                            }
                            if (videoFile.Tag.Title != null)
                            {
                                videoFile.Tag.Title = null;
                                resultVariables[titlesRemovedCount] += 1;
                                saveFile = true;
                            }
                            if (videoFile.Tag.Comment != null)
                            {
                                videoFile.Tag.Comment = null;
                                resultVariables[commentsRemovedCount] += 1;
                                saveFile = true;
                            }

                            if (saveFile)
                            {
                                videoFile.Save();
                            }
                        }
                    }
                            
                }
                Type("DONE", 100, 0, 1, "Green");

                DisplayResults(resultVariables);
            }
            catch (Exception e)
            {
                Type("Unable to remove the metadata on a file | " + e.Message, 0, 0, 1, "Red");
            }

        } // End RemoveMetadata()

        protected static void CopyJpgFiles()
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

                    if (missingNfo.Count() > 0)
                    {
                        Type("The following " + missingNfo.Count().ToString() + " movies are missing NFO files:", 0, 0, 1, "DarkRed");
                        foreach (string item in missingNfo)
                        {
                            Type(item, 0, 0, 1, "Red");
                        }
                        Console.WriteLine();
                    }
                    if (missingJpg.Count() > 0)
                    {
                        Type("The following " + missingJpg.Count().ToString() + " movies are missing JPG files:", 0, 0, 1, "DarkYellow");
                        foreach (string item in missingJpg)
                        {
                            Type(item, 0, 0, 1, "Yellow");
                        }
                        Console.WriteLine();
                    }
                    if (missingMovie.Count() > 0)
                    {
                        Type("The following " + missingMovie.Count().ToString() + " movies are missing MKV files:", 0, 0, 1, "Blue");
                        foreach (string item in missingMovie)
                        {
                            Type(item, 0, 0, 1, "Cyan");
                        }
                        Console.WriteLine();
                    }
                    if (missingIso.Count() > 0)
                    {
                        Type("The following " + missingIso.Count().ToString() + " movies are missing ISO files:", 0, 0, 1, "DarkCyan");
                        foreach (string item in missingIso)
                        {
                            Type(item, 0, 0, 1, "Cyan");
                        }
                        Console.WriteLine();
                    }

                    keepAskingForDirectory = false;
                }
                else
                {
                    Type(directory + " is not a valid file or directory.", 14, 100, 1);
                }
            } while (keepAskingForDirectory);

            Type("It looks like that's it.", 3, 100, 2);
        } // End CopyJpgFiles()

        protected static void CopyMovieFiles(IList<IList<object>> data, Dictionary<string, int> sheetVariables)
        {
            var repeatProcess = false;

            do
            {
                int intFileSkippedCount = 0, intFileAlreadyThereCount = 0, intFileCopiedCount = 0, intFileNotFoundCount = 0;
                string person = "", chosenDestination = "", sourceHardDriveLetter = "";

                DisplayMessage("question", "Who's list of movies are we copyings? (number only please)");
                DisplayMessage("default", "0 - ", 0);
                DisplayMessage("info", "Exit");
                DisplayMessage("default", "1 - ", 0);
                DisplayMessage("info", "Cindy");
                DisplayMessage("default", "2 - ", 0);
                DisplayMessage("info", "Dave");

                person = Console.ReadLine();

                if (person == "0")
                {
                    break;
                } else if (person == "1" || person == "2")
                {
                    Type("What hard drive am I copying to? (Just the hard drive letter)", 0, 0, 1, "Yellow");

                    chosenDestination = Console.ReadLine().ToUpper();

                    if (!HardDriveHasSpace(chosenDestination))
                    {
                        DisplayMessage("error", "We won't be able to copy any more movies to this hard drive because available space is below 10%");
                        break;
                    } else
                    {
                        Console.WriteLine("We will copy to hard drive " + chosenDestination);
                    }

                    Type("What hard drive am I copying from? (Just the hard drive letter)", 0, 0, 1, "Yellow");

                    sourceHardDriveLetter = Console.ReadLine().ToUpper();

                    if (chosenDestination != sourceHardDriveLetter)
                    {
                        Console.WriteLine("We will copy from the " + sourceHardDriveLetter + " drive.");
                    } else
                    {
                        DisplayMessage("error", "I'm sorry the source hard drive can't be the same as the destination hard drive.");
                        repeatProcess = true;
                    }

                } else
                {
                    DisplayMessage("error", "I'm sorry, I don't recognise " + person + " yet. Could you add them to my DB before continuing?");
                    repeatProcess = true;
                }

                if (!repeatProcess)
                {
                    foreach (var row in data)
                    {
                        if (HardDriveHasSpace(chosenDestination))
                        {
                            DisplayMessage("error", "We have stopped copying movies because available hard drive space is below 10%");
                            break;
                        }

                        if (row.Count > 4) // If it's an empty row then it should have less than this.
                        {
                            var cleanTitle = row[Convert.ToInt16(sheetVariables["Clean Title"])].ToString();
                            var movieDirectory = row[Convert.ToInt16(sheetVariables["Directory"])].ToString();
                            var status = row[Convert.ToInt16(sheetVariables["Status"])].ToString();
                            var cindy = row[Convert.ToInt16(sheetVariables["Cindy"])].ToString();
                            var dave = row[Convert.ToInt16(sheetVariables["Dave"])].ToString();
                            var folderLetter = row[Convert.ToInt16(sheetVariables["Movie Letter"])].ToString();
                            var kids = row[Convert.ToInt16(sheetVariables["Kids"])].ToString();

                            var selected = "";

                            if (person == "1")
                            {
                                selected = cindy;
                            }
                            else if (person == "2")
                            {
                                selected = dave;
                            }
                            else
                            {
                                DisplayMessage("error", "I'm sorry, I don't recognise " + person + " yet. Could you add them to my DB before continuing?");
                                repeatProcess = true;
                            }

                            try
                            {
                                // If the first letter of status is an 'x' or is empty, then we don't need to run through this for loop so don't waste the resources.
                                if (!repeatProcess && !status.Equals("") && status[0].ToString().ToUpper() != "X" && selected.ToUpper() == "Y")
                                {
                                    // Since the movie status is valid let's go ahead and check if the movie is already at the destination.

                                    // Create the string that contains the location where the movie should be on the hard drive so we can easily check
                                    // to see if that movie exists.
                                    var fullDestinationPath = "";

                                    // The directory that holds the movie file.
                                    var containingDirectory = "";

                                    if (kids.ToUpper() == "X")
                                    {
                                        containingDirectory = chosenDestination + ":\\Movies\\Kids Movies";
                                    }
                                    else
                                    {
                                        containingDirectory += chosenDestination + ":\\Movies\\" + folderLetter;
                                    }

                                    // Create the holding directory just in case.
                                    Directory.CreateDirectory(containingDirectory);

                                    // Concatenate to the containing directory.
                                    fullDestinationPath = containingDirectory + "\\" + cleanTitle;

                                    // Grab all files that contain the cleanTitle out of the containgDirectory.
                                    string[] fileEntries = Directory.GetFiles(containingDirectory, cleanTitle + ".*", SearchOption.AllDirectories);

                                    if (fileEntries.Length == 0)
                                    {
                                        // Now that we know the destination location doesn't exist we need to check if the source folder exists to
                                        // copy from.
                                        var sourcePathWithDriveLetter = sourceHardDriveLetter + movieDirectory;
                                        var fullSourcePath = sourcePathWithDriveLetter + "\\" + cleanTitle;
                                        if (Directory.Exists(sourcePathWithDriveLetter))
                                        {
                                            // Process the list of files found in the destination.
                                            string[] sourceFileEntries = Directory.GetFiles(sourcePathWithDriveLetter);
                                            if (sourceFileEntries.Length > 0)
                                            {
                                                var movieFoundAtSource = false;
                                                foreach (string fileName in sourceFileEntries)
                                                {
                                                    movieFoundAtSource = false;
                                                    string mp4 = ".mp4",
                                                           mkv = ".mkv",
                                                           m4v = ".m4v",
                                                           avi = ".avi",
                                                           srt = ".en.forced.srt",
                                                           extension = "";
                                                    if (fileName.ToLower().Contains(mp4))
                                                    {
                                                        extension = mp4;
                                                        movieFoundAtSource = true;
                                                    }
                                                    else if (fileName.ToLower().Contains(mkv))
                                                    {
                                                        movieFoundAtSource = true;
                                                        extension = mkv;
                                                    }
                                                    else if (fileName.ToLower().Contains(m4v))
                                                    {
                                                        movieFoundAtSource = true;
                                                        extension = m4v;
                                                    }
                                                    else if (fileName.ToLower().Contains(avi))
                                                    {
                                                        movieFoundAtSource = true;
                                                        extension = avi;
                                                    }
                                                    else if (fileName.ToLower().Contains(srt))
                                                    {
                                                        movieFoundAtSource = true;
                                                        extension = srt;
                                                    }

                                                    if (movieFoundAtSource)
                                                    {
                                                        DisplayMessage("info", "Copying ", 0);
                                                        DisplayMessage("default", cleanTitle + extension + "... ", 0);
                                                        CopyFile(fullSourcePath + extension, fullDestinationPath + extension);
                                                        DisplayMessage("success", "DONE");
                                                        intFileCopiedCount++;
                                                    }
                                                    

                                                }
                                                //if (!movieFoundAtSource)
                                                //{
                                                //    Type("No movie file was found for " + cleanTitle + ".", 0, 0, 1, "Red");
                                                //    intFileNotFoundCount++;
                                                //}
                                            }
                                            else
                                            {
                                                Type("No files found in source directory.", 0, 0, 1, "Magenta");
                                            }

                                        }
                                        else
                                        {
                                            Type(sourcePathWithDriveLetter + " path was not found.", 0, 0, 1, "Magenta");
                                        }
                                    }
                                    else
                                    {
                                        //DisplayMessage("default", cleanTitle, 0);
                                        //DisplayMessage("info", " is already at destination folder.");
                                        intFileAlreadyThereCount++;
                                    }
                                }
                                else
                                {
                                    //Type("Skipped " + cleanTitle, 0, 0, 1, "Yellow");
                                    intFileSkippedCount++;
                                }
                            }
                            catch (Exception e)
                            {
                                Type("Something went wrong when looking for: " + sourceHardDriveLetter + "\\" + movieDirectory + " | " + e.Message, 3, 100, 1, "Red");
                                //throw;
                            }

                        }
                    } // End foreach
                } // End if
                

                Console.WriteLine();

                if (!repeatProcess)
                {
                    Type("It looks like that's the end of it.", 0, 0, 1);
                    Type("Movies copied: " + intFileCopiedCount, 0, 0, 1, "Green");
                    Type("Movies skipped: " + intFileSkippedCount, 0, 0, 1, "Yellow");
                    Type("Source movies not found: " + intFileNotFoundCount, 0, 0, 1, "Red");
                    Type("Movies already at destination: " + intFileAlreadyThereCount, 0, 0, 1, "Blue");
                    DisplayMessage("question", "Remaining Hard Drive Space: " + GetAvailableHardDrivePercent(chosenDestination) + "%");
                }

            } while (repeatProcess);

        } // End CopyMovieFiles()

        public static double GetAvailableHardDrivePercent(string hd)
        {
            try
            {
                DriveInfo di = new DriveInfo(hd);

                if (di.IsReady)
                {
                    double freeSpace = di.AvailableFreeSpace;
                    double totalSpace = di.TotalSize;

                    double availablePercent = Math.Round((freeSpace / totalSpace) * 100, 2);

                    return availablePercent;
                }
                else
                {
                    DisplayMessage("warning", "Can't check available hard drive space -- Hard drive is not ready");
                    throw new InvalidOperationException();
                }

            }
            catch (IOException e)
            {
                DisplayMessage("error", "An error occured while checking the hard drive space | " + e.Message);
                throw;
            }
        } // End GetAvailableHardDrivePercent()

        public static bool HardDriveHasSpace(string hd, int moreSpaceThan = 11)
        {
            try
            {
                DriveInfo di = new DriveInfo(hd);

                if (di.IsReady)
                {
                    double freeSpace = di.AvailableFreeSpace;
                    double totalSpace = di.TotalSize;

                    double availablePercent = Math.Round((freeSpace / totalSpace) * 100, 2);

                    if (availablePercent > moreSpaceThan)
                        return true;
                    else
                        return false;
                }
                else
                {
                    DisplayMessage("warning", "Can't check available hard drive space -- Hard drive is not ready");
                    throw new InvalidOperationException();
                }

            }
            catch (IOException e)
            {
                DisplayMessage("error", "An error occured while checking the hard drive space | " + e.Message);
                throw;
            }
        } // End GetAvailableHardDrivePercent()

        protected static void DeleteMovieFiles(IList<IList<object>> data, Dictionary<string, int> sheetVariables)
        {
            do
            {
                int intFileSkippedCount = 0, intFileAlreadyThereCount = 0, intFileDeletedCount = 0, intFileNotFoundCount = 0;

                DisplayMessage("question", "Who's hard drive are we deleting from? (number only please)");
                DisplayMessage("default", "0 - ", 0);
                DisplayMessage("info", "Exit");
                DisplayMessage("default", "1 - ", 0);
                DisplayMessage("info", "Cindy");
                DisplayMessage("default", "2 - ", 0);
                DisplayMessage("info", "Dave");

                var person = Console.ReadLine();

                if (person == "0")
                {
                    break;
                }
                else
                {
                    Type("What hard drive am I deleting from? (Just the hard drive letter)", 0, 0, 1, "Yellow");

                    var chosenDestination = Console.ReadLine();

                    Console.WriteLine("We will delete from hard drive " + chosenDestination);

                    foreach (var row in data)
                    {
                        if (row.Count > 4) // If it's an empty row then it should have less than this.
                        {
                            var cleanTitle = row[Convert.ToInt16(sheetVariables["Clean Title"])].ToString();
                            var movieDirectory = row[Convert.ToInt16(sheetVariables["Directory"])].ToString();
                            var status = row[Convert.ToInt16(sheetVariables["Status"])].ToString();
                            var cindy = row[Convert.ToInt16(sheetVariables["Cindy"])].ToString();
                            var dave = row[Convert.ToInt16(sheetVariables["Dave"])].ToString();
                            var folderLetter = row[Convert.ToInt16(sheetVariables["Movie Letter"])].ToString();
                            var kids = row[Convert.ToInt16(sheetVariables["Kids"])].ToString();

                            var selected = "";

                            if (person == "1")
                            {
                                selected = cindy;
                            }
                            else if (person == "2")
                            {
                                selected = dave;
                            }
                            else
                            {
                                DisplayMessage("error", "I'm sorry, I don't recognise " + person + " yet. Could you add them to my DB before continuing?");
                                break;
                            }

                            try
                            {
                                // If the first letter of status is an 'x' or is empty, then we don't need to run through this for loop so don't waste the resources.
                                if (!status.Equals("") && status[0].ToString().ToUpper() != "X" && !selected.ToUpper().Equals("Y"))
                                {
                                    // Since the movie status is valid let's go ahead and check if the movie is already at the destination.

                                    // Create the string that contains the location where the movie should be on the hard drive so we can easily check
                                    // to see if that movie exists.
                                    var fullDestinationPathToFileToDelete = "";

                                    // The directory that holds the movie file.
                                    var containingDirectory = "";

                                    if (kids.ToUpper() == "X")
                                    {
                                        containingDirectory = chosenDestination + ":\\Movies\\Kids Movies";
                                    }
                                    else
                                    {
                                        containingDirectory += chosenDestination + ":\\Movies\\" + folderLetter;
                                    }

                                    // Create the holding directory just in case.
                                    Directory.CreateDirectory(containingDirectory);

                                    // Concatenate to the containing directory.
                                    fullDestinationPathToFileToDelete = containingDirectory + "\\" + cleanTitle;

                                    // Loop through the containing directory to see if the movie is already in there.
                                    string[] fileEntries = Directory.GetFiles(containingDirectory, cleanTitle + ".*");
                                    if (fileEntries.Length > 0)
                                    {
                                        foreach (var movie in fileEntries)
                                        {
                                            File.Delete(movie);
                                            intFileDeletedCount++;
                                        }
                                    }
                                }
                                else
                                {
                                    //Type("Skipped " + cleanTitle, 0, 0, 1, "Yellow");
                                    intFileSkippedCount++;
                                }
                            }
                            catch (Exception e)
                            {
                                DisplayMessage("error", "An error occured deleting the video | " + e.Message);
                                throw;
                            }

                        }
                    }

                    Console.WriteLine();
                    Type("It looks like that's the end of it.", 0, 0, 1);
                    Type("Movies deleted: " + intFileDeletedCount, 0, 0, 1, "Green");
                    Type("Movies skipped: " + intFileSkippedCount, 0, 0, 1, "Yellow");
                    Type("Source movies not found: " + intFileNotFoundCount, 0, 0, 1, "Red");
                    Type("Movies already at destination: " + intFileAlreadyThereCount, 0, 0, 1, "Blue");
                }
            } while (false);
             
        } // End DeleteMovieFiles()

        protected static void ClearDirectories()
        {
            missingNfo.Clear();
            missingJpg.Clear();
            missingMovie.Clear();
            missingIso.Clear();
            partFiles.Clear();
        }

        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        public static void ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            if (fileEntries.Length > 0)
            {
                int nfoCount = 0, jpgCount = 0, mp4Count = 0, mkvCount = 0, m4vCount = 0, aviCount = 0, webmCount = 0, unidentifiedCount = 0, isoCount = 0, partCount = 0;
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
                    else if (fileName.ToUpper().Contains(".M4V"))
                        m4vCount++;
                    else if (fileName.ToUpper().Contains(".AVI"))
                        aviCount++;
                    else if (fileName.ToUpper().Contains(".WEBM"))
                        webmCount++;
                    else if (fileName.ToUpper().Contains(".ISO"))
                        isoCount++;
                    else if (fileName.ToUpper().Contains(".PART"))
                        partCount++;
                    else
                    {
                        //Type("Unidentified file: " + fileName, 0, 0, 1);
                        unidentifiedCount++;
                    }

                }

                if (nfoCount == 0)
                {
                    missingNfo.Add(targetDirectory);
                }
                if (jpgCount < 2)
                {
                    missingJpg.Add(targetDirectory);
                }
                if (mp4Count == 0 && mkvCount == 0 && m4vCount == 0 && aviCount == 0 && webmCount == 0)
                {
                    missingMovie.Add(targetDirectory);
                }
                if (partCount > 0)
                {
                    partFiles.Add(targetDirectory);
                }

                //Type(nfoCount + " nfo, " + jpgCount + " jpg, " + mp4Count + " mp4, " + mkvCount + " mkv, " + m4vCount + " m4v, " + isoCount + " iso, " + unidentifiedCount + " unidentified in " + targetDirectory, 0, 0, 1);
            } else if (subdirectoryEntries.Length == 0)
            {
                Directory.Delete(targetDirectory);
                emptyDirectory.Add(targetDirectory);
            }

            // Recurse into subdirectories of this directory.
            foreach (string subdirectory in subdirectoryEntries)
            {
                // Don't go into the bonus features folders.
                if (!subdirectory.Contains("\\Behind The Scenes") &&
                    !subdirectory.Contains("\\Scenes") &&
                    !subdirectory.Contains("\\Deleted Scenes") &&
                    !subdirectory.Contains("\\Shorts") &&
                    !subdirectory.Contains("\\Featurettes") &&
                    !subdirectory.Contains("\\Trailers") &&
                    !subdirectory.Contains("\\Interviews") &&
                    !subdirectory.Contains("\\Broken apart") &&
                    !subdirectory.Contains("\\Other") &&
                    !subdirectory.Contains("\\_Collections") &&
                    !subdirectory.Contains("\\.sync"))
                {
                    ProcessDirectory(subdirectory);
                }
            }
        }
        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        public static void ProcessCopyDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            if (fileEntries.Length > 0)
            {
                int nfoCount = 0, jpgCount = 0, mp4Count = 0, mkvCount = 0, m4vCount = 0, aviCount = 0, unidentifiedCount = 0, isoCount = 0;
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
                    else if (fileName.ToUpper().Contains(".M4V"))
                        m4vCount++;
                    else if (fileName.ToUpper().Contains(".AVI"))
                        aviCount++;
                    else if (fileName.ToUpper().Contains(".ISO"))
                        isoCount++;
                    else
                    {
                        //Type("Unidentified file: " + fileName, 0, 0, 1);
                        unidentifiedCount++;
                    }

                }

                // If none of these types of files are in here then we probably have an empty ISO folder.
                if (nfoCount == 0 && jpgCount == 0 && mp4Count == 0 && mkvCount == 0 && m4vCount == 0 && aviCount == 0 && isoCount == 0)
                {
                    missingIso.Add(targetDirectory);
                }
                // However if isoCount is not equal to 0 then we are in an image folder and I don't want to count missing NFO files and such.
                else if (isoCount != 0)
                {
                    // Don't do anything.
                }
                else
                {
                    if (nfoCount == 0)
                    {
                        missingNfo.Add(targetDirectory);
                    }
                    if (jpgCount < 2)
                    {
                        missingJpg.Add(targetDirectory);
                    }
                    if (mp4Count == 0 && mkvCount == 0 && m4vCount == 0 && aviCount == 0)
                    {
                        missingMovie.Add(targetDirectory);
                    }
                }

                //Type(nfoCount + " nfo, " + jpgCount + " jpg, " + mp4Count + " mp4, " + mkvCount + " mkv, " + m4vCount + " m4v, " + isoCount + " iso, " + unidentifiedCount + " unidentified in " + targetDirectory, 0, 0, 1);
            }
            //else
            //Type(targetDirectory, 0, 0, 1);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
            {
                // Don't go into the bonus features folders.
                if (!subdirectory.Contains("\\Behind The Scenes") &&
                    !subdirectory.Contains("\\Scenes") &&
                    !subdirectory.Contains("\\Deleted Scenes") &&
                    !subdirectory.Contains("\\Shorts") &&
                    !subdirectory.Contains("\\Featurettes") &&
                    !subdirectory.Contains("\\Trailers") &&
                    !subdirectory.Contains("\\Interviews") &&
                    !subdirectory.Contains("\\Broken apart") &&
                    !subdirectory.Contains("\\_Collections") &&
                    !subdirectory.Contains("\\.sync"))
                {
                    ProcessDirectory(subdirectory);
                }
            }
        }

        // Insert logic for processing found files here.
        public static void ProcessFile(string path)
        {
            Console.WriteLine("Processed file '{0}'.", path);
        }

        /// <summary>
        /// Displays the message in the color based on type.
        /// </summary>
        /// <param name="messageType">error = red, harderror = darkred, success = green, warning = yellow, info = blue, data = cyan, question = darkyellow, log = white, default = grey.</param>
        /// <param name="message">The message to display.</param>
        /// <param name="numLines">The number of new lines to print out after the message.</param>
        /// <param name="speed">The speed in MS at which to type the letters (Higher the number the slower).</param>
        /// <param name="pause">The amount of ms to pause before going to the next line.</param>
        public static void DisplayMessage(string messageType, string message, int numLines = 1, int speed = 0, int pause = 0)
        {
            switch (messageType.ToLower())
            {
                case "error":
                    Type(message, speed, pause, numLines, "red");
                    break;
                case "harderror":
                    Type(message, speed, pause, numLines, "darkred");
                    break;
                case "success":
                    Type(message, speed, pause, numLines, "green");
                    break;
                case "warning":
                    Type(message, speed, pause, numLines, "yellow");
                    break;
                case "info":
                    Type(message, speed, pause, numLines, "blue");
                    break;
                case "data":
                    Type(message, speed, pause, numLines, "cyan");
                    break;
                case "question":
                    Type(message, speed, pause, numLines, "darkyellow");
                    break;
                case "log":
                    Type(message, speed, pause, numLines, "white");
                    break;
                case "default":
                    Type(message, speed, pause, numLines);
                    break;
                default:
                    break;
            }
        } // End DisplayMessage()

        /// <summary>
        /// Simply types out the text in a typewriter manner. Then adds the number of new lines.
        /// </summary>
        /// <param name="myString"></param>
        /// <param name="speed"></param>
        /// <param name="timeToPauseBeforeNewLine"></param>
        /// <param name="numberOfNewLines"></param>
        /// <param name="color">Red, Green, Yellow, Blue, Magenta, Gray, Cyan, DarkBlue, DarkCyan, DarkGray, DarkGreen, DarkRed, DarkYellow</param>
        public static void Type(string myString, int speed, int timeToPauseBeforeNewLine, int numberOfNewLines, string color = "gray")
        {
            // First set the text color.
            switch (color.ToLower())
            {
                case "black":
                    Console.ForegroundColor = ConsoleColor.Black;
                    break;
                case "blue":
                    Console.ForegroundColor = ConsoleColor.Blue;
                    break;
                case "cyan":
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    break;
                case "darkblue":
                    Console.ForegroundColor = ConsoleColor.DarkBlue;
                    break;
                case "darkcyan":
                    Console.ForegroundColor = ConsoleColor.DarkCyan;
                    break;
                case "darkgray":
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    break;
                case "darkgreen":
                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                    break;
                case "darkmagenta":
                    Console.ForegroundColor = ConsoleColor.DarkMagenta;
                    break;
                case "darkred":
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    break;
                case "darkyellow":
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    break;
                case "gray":
                    Console.ForegroundColor = ConsoleColor.Gray;
                    break;
                case "green":
                    Console.ForegroundColor = ConsoleColor.Green;
                    break;
                case "magenta":
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    break;
                case "red":
                    Console.ForegroundColor = ConsoleColor.Red;
                    break;
                case "white":
                    Console.ForegroundColor = ConsoleColor.White;
                    break;
                case "yellow":
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;
                default:
                    Console.ForegroundColor = ConsoleColor.Gray;
                    break;
            }

            // Now type out the text.
            for (int i = 0; i < myString.Length; i++)
            {
                Console.Write(myString[i]);
                Thread.Sleep(speed);
            }

            // Reset the color back to normal.
            Console.ResetColor();

            // Pause the desired amount before moving onto the next line.
            Thread.Sleep(timeToPauseBeforeNewLine);

            // Finally print the number of lines.
            while (numberOfNewLines > 0)
            {
                Console.WriteLine();
                numberOfNewLines--;
            }
        } // End Type()

        //private static void SendEmail()
        //{
        //    try
        //    {
        //        MailMessage mail = new MailMessage();
        //        mail.From = new MailAddress("qtip16@gmail.com");

        //        // The important part -- configuring the SMTP client
        //        using (SmtpClient smtp = new SmtpClient())
        //        {
        //            smtp.Port = 25;
        //            //smtp.Port = 587;   // [1] You can try with 465 also, I always used 587 and got success
        //            smtp.DeliveryMethod = SmtpDeliveryMethod.Network; // [2] Added this
        //            //smtp.UseDefaultCredentials = true;
        //            smtp.UseDefaultCredentials = false; // [3] Changed this
        //            smtp.Credentials = new NetworkCredential("qtip16@gmail.com", "Carson#1");  // [4] Added this. Note, first parameter is NOT string.
        //            smtp.EnableSsl = true;
        //            //smtp.EnableSsl = false;
        //            smtp.Host = "smtp.gmail.com";

        //            //recipient address
        //            mail.To.Add(new MailAddress("brandon.birschbach@gmail.com"));

        //            //Formatted mail body
        //            mail.IsBodyHtml = true;
        //            string st = "Test";

        //            mail.Body = st;
        //            smtp.Send(mail);
        //        };
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine("Woops | " + e.Message);
        //        throw;
        //    }
        //}

        //protected static void SendEmails()
        //{
        //    string reportBody = "";
        //    try
        //    {
        //        reportBody += gReportHeader + "<br /><br />";
        //        reportBody += "Click <a href=" + gReportLocation + ">here</a> to view the whole report.";

        //        SmtpClient Smtp_Server = new SmtpClient();
        //        MailMessage e_mail = new MailMessage();

        //        Smtp_Server.UseDefaultCredentials = true;

        //        Smtp_Server.Port = 25;
        //        Smtp_Server.EnableSsl = false;
        //        Smtp_Server.Host = mail.amcllc.net;

        //        e_mail = new MailMessage();
        //        e_mail.From = new MailAddress("Support@rentegi.com");
        //        e_mail.To.Add(steve@rentegi.com);
        //        e_mail.Subject = "AIM Opiniion Survey Send Alert";
        //        e_mail.IsBodyHtml = true;
        //        e_mail.Body = reportBody;

        //        Smtp_Server.Send(e_mail);

        //    } // End try
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error sending email.");
        //        Console.WriteLine(ex.ToString());
        //        Console.WriteLine("Writing results to error log.");
        //        LogReport("Error sending email.<br />" + ex.Message + "<br /><br />" + reportBody);
        //    } // End catch

        //} // End SendEmail()


        ///////////////////////////////////////////////////////////////////////////
        // THE FOLLOWING METHODS HAVE BEEN COPIED OVER TO THE BachFlixNfo.cs FILE FOR NEW CODE.
        // HOWEVER, KEEP THEM COPIED HERE UNTIL I MOVE THE REST OF THE CODE OVER.
        ///////////////////////////////////////////////////////////////////////////



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
                Type("Something went wrong writing to path: " + path + " | " + e.Message, 3, 100, 1, "Red");
                Type(e.Message, 3, 100, 2, "DarkRed");
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

        public static bool WriteSingleCellToSheet(string strDataToSave, string strCellToSaveData)
        {
            var tryAgain = false;
            do
            {
                try
                {
                    // Thread.Sleep(1000); // Sleep for a second so we don't go over the Google allotted requests.
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
                    tryAgain = false;
                }
                catch (Exception e)
                {
                    var m = e.Message;
                    if (m.Contains("Quota exceeded"))
                    {
                        DisplayMessage("error", "Broke the max calls to the Google Sheet.");
                        DisplayMessage("info", "Pausing for 30 seconds and trying again.");
                        Countdown(30);
                        tryAgain = true;
                        DisplayMessage("info", "Let's try again.");
                    } else
                    {
                        DisplayMessage("error", "An error has occurred.");
                        DisplayMessage("harderror", m);
                    }
                }

            } while (tryAgain);
            return true;

        } // End WriteSingleCellToSheet

    } // End class
} // End namespace