using BachFlixNfo.Features;
using BachFlixNfo.Subtitles;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Google.Apis.YouTube.v3;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using TmdbApiCall;

namespace SheetsQuickstart
{
    class Program
    {
        // Emby server settings. API key comes from the local environment.
        private const string EMBY_SERVER_URL_ENV = "EMBY_SERVER_URL";
        private const string EMBY_API_KEY_ENV = "EMBY_API_KEY";
        private const string EMBY_MOVIES_LIBRARY_ID_ENV = "EMBY_MOVIES_LIBRARY_ID";
        private const string DEFAULT_EMBY_SERVER_URL = "http://192.168.0.5:9000";
        private const string DEFAULT_EMBY_MOVIES_LIBRARY_ID = "3";

        // Data ranges for each sheet.
        private const string MOVIES_TITLE_RANGE = "Movies!A2:2";
        private const string MOVIES_DATA_RANGE = "Movies!A3:22002";
        private const string MOVIES_DB_TITLE_RANGE = "Movies!A1:1";
        private const string MOVIES_DB_DATA_RANGE = "Movies!A2:12002";
        private const string TEMP_MOVIES_TITLE_RANGE = "Temp!A2:2";
        private const string TEMP_MOVIES_DATA_RANGE = "Temp!A3:2001";
        private const string YOUTUBE_TITLE_RANGE = "YouTube!A2:2";
        private const string YOUTUBE_DATA_RANGE = "YouTube!A3:4000";
        private const string FITNESS_VIDEO_TITLE_RANGE = "Fitness Videos!A1:1";
        private const string FITNESS_VIDEO_DATA_RANGE = "Fitness Videos!A2:401";
        private const string BONUS_TITLE_RANGE = "Bonus!A1:1";
        private const string BONUS_DATA_RANGE = "Bonus!A2:2036";
        private const string RENAME_EPISODES_TITLE_RANGE = "Rename Episodes!A2:2";
        private const string RENAME_EPISODES_DATA_RANGE = "Rename Episodes!A3:102";
        private const string TEMP_EPISODES_TITLE_RANGE = "Temp Episodes!A1:1";
        private const string TEMP_EPISODES_DATA_RANGE = "Temp Episodes!A2:1000";
        private const string COMBINED_EPISODES_TITLE_RANGE = "Combined Episodes!A2:2";
        private const string COMBINED_EPISODES_DATA_RANGE = "Combined Episodes!A3:3502";
        private const string RECORDED_NAMES_TITLE_RANGE = "Recorded Names!A2:2";
        private const string RECORDED_NAMES_DATA_RANGE = "Recorded Names!A3:1102";
        private const string DB_TITLE_RANGE = "DB!A2:2";
        private const string DB_DATA_RANGE = "DB!A3:2002";
        private const string SEVERAL_COMBINED_EPISODES_TITLE_RANGE = "Several Combined Episodes!A2:2";
        private const string SEVERAL_COMBINED_EPISODES_DATA_RANGE = "Several Combined Episodes!A3:1000";
        private const string AUTOPOULATE_ACTORS_TITLE_RANGE = "Autopopulate Actors!A2:B2";
        private const string AUTOPOULATE_ACTORS_DATA_RANGE = "Autopopulate Actors!A3:B";
        private const string SKIP_ACTORS_ID_TITLE_RANGE = "Autopopulate Actors!D2";
        private const string SKIP_ACTORS_ID_DATA_RANGE = "Autopopulate Actors!D3:D";

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
        private const string ROW_NUM = "RowNum";

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
        static List<string> res240List = new List<string>();
        static List<string> res360List = new List<string>();
        static List<string> res480List = new List<string>();
        static List<string> res576List = new List<string>();
        static List<string> res720List = new List<string>();
        static List<string> res1080List = new List<string>();
        static List<string> res1440List = new List<string>();
        static List<string> res2160List = new List<string>();
        static List<string> resNAList = new List<string>();

        // Variables for the SRT ENG method.
        static List<string> missingEng = new List<string>();
        static int fileEntriesCount = 0,
            srtEntriesCount = 0,
            missingEngCount = 0;
        private static readonly Regex SrtEngFlagRegex = new Regex(@"(^|[._-])eng($|[._-])", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex SrtForcedFlagRegex = new Regex(@"(^|[._-])forced($|[._-])", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex SrtLanguageTokenRegex = new Regex(@"(?<sep>[._-])(?:en(?:[-_]?us)?|som|ger|bre|nor|hat|ht|so)(?=[._-]|$)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // The method in which to input the data into the Google Sheet.
        const int INSERT_MISSING_DATA = 1;
        const int OVERWRITE_DATA = 2;

        // Menu item variables.
        static string exitChoice,
            missingMovieNfoFilesChoice,
            overwriteAllMovieNfoFilesChoice,
            selectedMovieNfoFilesChoice,
            missingCombinedEpisodeNfoFilesChoice,
            missingSeveralCombinedEpisodeNfoFilesChoice,
            missingYoutubeNfoFilesChoice,
            overwriteAllYoutubeNfoFilesChoice,
            selectedYoutubeNfoFilesChoice,
            missingFitnessVideoNfoFilesChoice,
            overwriteAllFitnessVideoNfoFilesChoice,
            selectedFitnessVideoNfoFilesChoice,
            missingTvShowNfoFilesChoice,
            overwriteAllTvShowNfoFilesChoice,
            selectedTvShowNfoFilesChoice,
            convertMoviesChoice,
            convertDirectoryChoice,
            convertMoviesSlowChoice,
            convertBonusFeaturesChoice,
            convertBonusFeaturesSlowChoice,
            convertTvShowsChoice,
            convertTvShowsSlowChoice,
            convertTempTvShowsChoice,
            convertTempTVShowsSlowChoice,
            audioCompatibilityConverterChoice,
            audioCensorRunnerChoice,
            insertMissingMovieDataChoice,
            repeatInsertMissingMovieDataChoice,
            updateMovieDataChoice,
            insertMissingTmdbIdsChoice,
            insertAndOverwriteTmdbIdsChoice,
            addAllActorsCredits,
            clearSelectedRowInMoviesSheet,
            clearSelectedRowInMoviesSheetAndAddToSkipList,
            copyMovieFilesToDestinationChoice,
            deleteMovieFilesAtDestinationChoice,
            removeMetadataChoice,
            createFoldersAndMoveFilesChoice,
            createFoldersAndMoveFilesAndSortChoice,
            trimTitlesInDirectoryChoice,
            bothTrimAndCreateFoldersChoice,
            addSizeOfTvShowDirectories,
            overwriteSizeOfTvShowDirectories,
            addSizeOfMovieDirectories,
            overwriteSizeOfMovieDirectories,
            fetchTvShowPlotsChoice,
            insertMissingCastMembers,
            fixRecordedNamesChoice,
            copyMultipleFilesToOneLocationChoice,
            insertMissingDbDataChoice,
            updateDbSheetChoice,
            insertMissingCombinedEpisodesChoice,
            insertMissingSeveralCombinedEpisodesChoice,
            updateCombinedEpisodesChoice,
            writeToCombinedEpisodesChoice,
            writeToSeveralCombinedEpisodesChoice,
            insertMissingEpisodeDataChoice,
            writeVideoFileNamesToYoutubeSheet,
            moveTvShowEpisodesToFolders,
            moveSameMovieFilesTopLevel,
            deleteMovieFiles,
            deleteJpgFiles,
            addSeasonToFolderName,
            getMovieWatchProviders,
            getMovieWatchProvidersOfSelectedMovies,
            getTvShowWatchProviders,
            moveFolderContentsChoice,
            getVideoResolutionChoice,
            overwriteVideoResolutionChoice,
            getDirectorySizeChoice,
            changeTheSeason,
            resetEpisodeNumbersChoice,
            resetEpisodeNumbersToChosenNumberChoice,
            chosenDirectory,
            searchYoutubeAndDownloadMovieTrailersChoice,
            downloadMovieTrailersChoice,
            changeEpisodesIntoTwoPartsChoice,
            revertEpisodeCombineChoice,
            checkSrtFileNamesChoice,
            checkSelectedSrtFileNamesChoice,
            fileTypeOverwriteChoice,
            fileTypeChoice,
            movieFilenameScannerChoice,
            renameMoviesFromSheetChoice,
            moveReadyVideosChoice,
            srtScoreRunnerChoice,
            fileHealthRunnerChoice,
            plexPopulateRatingKeysChoice,
            plexOverwriteRatingKeysChoice,
            plexSyncMetadataChoice,
            profileRequestProcessorChoice,
            rebuildWebMoviesChoice,
            fullMovieHealthCheck;

        private static SheetsService _cachedSheetsService;

        private static SheetsService GetSheetsService()
        {
            if (_cachedSheetsService != null)
                return _cachedSheetsService;

            UserCredential credential;
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            _cachedSheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });

            return _cachedSheetsService;
        }

        private static string GetConfigValue(string environmentVariableName, string defaultValue = "")
        {
            string value = Environment.GetEnvironmentVariable(environmentVariableName);
            return string.IsNullOrWhiteSpace(value) ? defaultValue : value;
        }

        private static string GetEmbyServerUrl()
        {
            return GetConfigValue(EMBY_SERVER_URL_ENV, DEFAULT_EMBY_SERVER_URL);
        }

        private static string GetEmbyApiKey()
        {
            return GetConfigValue(EMBY_API_KEY_ENV);
        }

        private static string GetEmbyMoviesLibraryId()
        {
            return GetConfigValue(EMBY_MOVIES_LIBRARY_ID_ENV, DEFAULT_EMBY_MOVIES_LIBRARY_ID);
        }
        static Dictionary<string, int> movieSheetVariables = new Dictionary<string, int>
        {
            { ROW_NUM, -1 },
            { "Special Title", -1 },
            { STATUS, -1 },
            { "Ownership", -1 },
            { "Recorded Source", -1 },
            { "StreamFab", -1 },
            { "Possible Record Source", -1 },
            { "Kids", -1 },
            { "Teens", -1 },
            { "Resolution", -1 },
            { "File Type", -1 },
            { "Size", -1 },
            { "Auto Title", -1 },
            { "IMDB Title", -1 },
            { "Sort Title", -1 },
            { "Auto Content Rating", -1 },
            { "Content Rating", -1 },
            { "IMDB URL", -1 },
            { "TMDB Rating", -1 },
            { "Plot", -1 },
            { "Release Date", -1 },
            { "Auto MPAA", -1 },
            { "MPAA", -1 },
            { "YouTube Trailer ID", -1 },
            { "Movie Has Trailer", -1 },
            { "SRT Score", -1 },
            { "Can't Find", -1 },
            { "Runtime", -1 },
            { ISO_TITLE_NUM, -1 },
            { ISO_CH_NUM, -1 },
            { ADDITIONAL_COMMANDS, -1 },
            { "Cast", -1 },
            { "Poster", -1 },
            { "TMDB ID", -1 },
            { "Rating Key", -1 },
            { "Movie Letter", -1 },
            { CLEAN_TITLE, -1 },
            { "IMDB ID", -1 },
            { NFO_BODY, -1 },
            { DIRECTORY, -1 },
            { QUICK_CREATE, -1 }
        };

        static Dictionary<string, int> clearableMovieSheetVariables = new Dictionary<string, int>
        {
            { ROW_NUM, -1 },
            { "Same Name", -1 },
            { "Overflow", -1 },
            { STATUS, -1 },
            { "Ownership", -1 },
            { "Recorded Source", -1 },
            { "StreamFab", -1 },
            { "Playon", -1 },
            { "Removed Splashes", -1 },
            { "Include Subtitles", -1 },
            { "Verify Subtitles Sync", -1 },
            { "Note", -1 },
            { "Possible Record Source", -1 },
            { "Special", -1 },
            { "Kids", -1 },
            { "Teens", -1 },
            { "Selected Resolution", -1 },
            { "Recorded Version", -1 },
            { "Resolution", -1 },
            { "Date Added", -1 },
            { "Size", -1 },
            { "File Type", -1 },
            { "IMDB Title", -1 },
            { "Sort Title", -1 },
            { "Content Rating", -1 },
            { "IMDB URL", -1 },
            { "TMDB Rating", -1 },
            { "Plot", -1 },
            { "Release Date", -1 },
            { "MPAA", -1 },
            { "YouTube Trailer ID", -1 },
            { "Movie Has Trailer", -1 },
            { "SRT Score", -1 },
            { "Can't Find", -1 },
            { "Runtime", -1 },
            { "Cast", -1 },
            { "Poster", -1 },
            { "TMDB ID", -1 },
            { "Rating Key", -1 },
            { QUICK_CREATE, -1 }
        };

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

        private static readonly string YOUTUBE_API_KEY = "AIzaSyBomk4BPUovSEGFGVrJGZIABVzCA1tSNKU";

        // Match: "Movie Title (2010)"
        private static readonly Regex TitleYearRegex =
            new Regex(@"^(?<title>.+?)\s*\((?<year>\d{4})\)\s*$", RegexOptions.Compiled);

        private static readonly Regex DuplicateYearRegex =
            new Regex(@"\((\d{4})\)\s*\(\1\)\s*$", RegexOptions.Compiled);

        private enum MediaKind
        {
            Movie,
            Tv
        }

        private sealed class MoveItem
        {
            public string PcName { get; set; }
            public string ReadyRoot { get; set; }
            public MediaKind Kind { get; set; }
            public string SourcePath { get; set; }
            public string Title { get; set; }
            public string DestinationPath { get; set; }

            public MoveItem()
            {
                PcName = "";
                ReadyRoot = "";
                SourcePath = "";
                Title = "";
                DestinationPath = "";
            }
        }

        private sealed class ScanSummary
        {
            public string PcName { get; set; }
            public string ReadyRoot { get; set; }
            public int MovieCount { get; set; }
            public int TvCount { get; set; }

            public ScanSummary()
            {
                PcName = "";
                ReadyRoot = "";
            }
        }

        private enum MoveKind
        {
            MovieItem,   // a single movie video file (+ sidecars) living in Ready/<Source>/
            MovieFolder, // a folder with a movie video inside
            TvShow       // a folder with subfolders (seasons)
        }

        private sealed class PlannedMove
        {
            public string ReadyRoot { get; set; }
            public string RecordSource { get; set; }

            public MoveKind Kind { get; set; }

            // For folders (TV show / movie folder)
            public string FolderPath { get; set; }
            public string FolderName { get; set; }

            // For MovieItem (video file + sidecars)
            public string MovieVideoFile { get; set; } // full path to video file

            public PlannedMove()
            {
                ReadyRoot = "";
                RecordSource = "";
                FolderPath = "";
                FolderName = "";
                MovieVideoFile = "";
            }

            public string DisplayName
            {
                get
                {
                    if (Kind == MoveKind.MovieItem && !string.IsNullOrWhiteSpace(MovieVideoFile))
                        return Path.GetFileNameWithoutExtension(MovieVideoFile);
                    if (!string.IsNullOrWhiteSpace(FolderName))
                        return FolderName;
                    return "(unknown)";
                }
            }
        }

        #region Disk space (UNC-safe)

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool GetDiskFreeSpaceEx(
            string lpDirectoryName,
            out ulong lpFreeBytesAvailable,
            out ulong lpTotalNumberOfBytes,
            out ulong lpTotalNumberOfFreeBytes);

        private static bool TryGetFreePercent(string path, out double percentFree)
        {
            percentFree = 0;

            try
            {
                if (string.IsNullOrWhiteSpace(path))
                    return false;

                // GetDiskFreeSpaceEx wants a directory. If they pass a file path, normalize it.
                string target = path;

                // If it's a file, use its directory. If it's a directory, use it.
                if (File.Exists(target))
                    target = Path.GetDirectoryName(target);

                if (string.IsNullOrWhiteSpace(target))
                    return false;

                ulong freeAvail, totalBytes, totalFree;
                if (!GetDiskFreeSpaceEx(target, out freeAvail, out totalBytes, out totalFree))
                    return false;

                if (totalBytes == 0)
                    return false;

                percentFree = Math.Round((freeAvail / (double)totalBytes) * 100.0, 2);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Chooses a destination root based on free-space:
        /// - If desiredRoot has >= minPercentFree (or space can't be verified), use desiredRoot.
        /// - If desiredRoot has < minPercentFree, try fallbackRoot.
        /// - If fallbackRoot also has < minPercentFree OR fallbackRoot can't be verified, returns false (can't move).
        /// </summary>
        private static bool ApplyLowSpaceReroute(
            string desiredRoot,
            string desiredLabel,
            string fallbackRoot,
            string recordSource,
            double minPercentFree,
            out string chosenRoot,
            out string chosenLabel,
            out double? desiredPercentFreeOrNull,
            out double? fallbackPercentFreeOrNull,
            out bool reroutedToFallback)
        {
            chosenRoot = desiredRoot;
            chosenLabel = desiredLabel;
            desiredPercentFreeOrNull = null;
            fallbackPercentFreeOrNull = null;
            reroutedToFallback = false;

            // Check desired
            double desiredPct;
            bool desiredOk = TryGetFreePercent(desiredRoot, out desiredPct);
            if (desiredOk)
            {
                desiredPercentFreeOrNull = desiredPct;

                // Good -> use desired
                if (desiredPct >= minPercentFree)
                    return true;

                // Low -> attempt fallback
            }
            else
            {
                // If we can't verify desired, proceed with desired (your previous behavior)
                chosenLabel = desiredLabel + " (space unknown)";
                return true;
            }

            // Desired is low -> check fallback
            double fallbackPct;
            bool fallbackOk = TryGetFreePercent(fallbackRoot, out fallbackPct);
            if (!fallbackOk)
            {
                // User asked: verify fallback isn't full either.
                // If we can't verify fallback, fail safe: don't move.
                return false;
            }

            fallbackPercentFreeOrNull = fallbackPct;

            if (fallbackPct < minPercentFree)
            {
                // Both are low -> can't move
                return false;
            }

            // Fallback is OK -> reroute
            chosenRoot = fallbackRoot;
            chosenLabel = "Other (low space)";
            reroutedToFallback = true;
            return true;
        }

        #endregion

        public static class DriveMapping
        {
            // =============================================================
            // Single source of truth for all QuadPlex/BachFlix network paths
            //
            // Design goal:
            // - You only edit ONE list of range strings when drives change.
            // - Everything else (Start/End chars, share labels, UNC roots, scans)
            //   is generated from those definitions.
            // =============================================================

            // ---- Server
            public const string QUADPLEX = @"\\QUADPLEX";

            // ---- Share / range definitions (EDIT THESE ONLY when you re-bucket drives)
            //
            // IMPORTANT:
            // - TV shares are literally named like: "#-b", "c-d", "e-i", ...
            //   So we derive the share name from the definition by lowercasing it.
            // - Movie shares are literally named like: "Movies #-D", "Movies E-M", ...
            //   So we prefix "Movies " to the definition.
            //
            // Examples:
            //   If you later change "N-R" to "N-O", you update it HERE only.
            private static readonly string[] TvRangeDefinitions =
            {
                "#-B",
                "C-D",
                "E-I",
                "J-M",
                "N-R",
                "S",
                "T-Z"
            };

            private static readonly string[] MovieRangeDefinitions =
            {
                "#-D",
                "E-M",
                "N-Z"
            };

            // ---- Misc shares / folders
            public const string OTHER_SHARE = "Other";

            public const string WORKING_AREA_FOLDER = "Working Area";
            public const string TV_SHOWS_FOLDER = "TV Shows";
            public const string MOVIES_FOLDER = "Movies";
            public const string OTHER_BUS_FOLDER = "The Bus";

            // Helper: build UNC paths without worrying about slashes
            private static string Unc(params string[] parts)
            {
                if (parts == null || parts.Length == 0) return string.Empty;

                var cleaned = new List<string>();
                foreach (var p in parts)
                {
                    if (string.IsNullOrWhiteSpace(p)) continue;
                    cleaned.Add(p.Trim().Trim('\\'));
                }
                if (cleaned.Count == 0) return string.Empty;

                return @"\\" + string.Join(@"\\", cleaned);
            }

            private static List<LetterRange> BuildRanges(string[] definitions, Func<string, string> toShareLabel, string serverUnc, string folderUnderShare)
            {
                return definitions.Select(def =>
                {
                    var trimmed = (def ?? string.Empty).Trim();
                    if (trimmed.Length == 0)
                        throw new ArgumentException("Range definition cannot be empty.");

                    var parts = trimmed.Split('-');
                    char start = parts[0][0];
                    char end = (parts.Length > 1 && parts[1].Length > 0) ? parts[1][0] : parts[0][0];

                    // Share name on the server
                    string shareLabel = toShareLabel(trimmed);

                    // RootPath points at the Working Area for that share
                    string root = Unc(serverUnc, shareLabel, folderUnderShare);

                    return new LetterRange
                    {
                        Start = char.ToUpperInvariant(start),
                        End = char.ToUpperInvariant(end),
                        Label = shareLabel,
                        RootPath = root
                    };
                }).ToList();
            }

            // Generated ranges (do not edit manually)
            public static readonly List<LetterRange> TvRanges =
                BuildRanges(
                    TvRangeDefinitions,
                    def => def.ToLowerInvariant(),              // "#-B" -> "#-b"
                    QUADPLEX,
                    WORKING_AREA_FOLDER
                );

            public static readonly List<LetterRange> MovieRanges =
                BuildRanges(
                    MovieRangeDefinitions,
                    def => $"Movies {def}",                     // "E-M" -> "Movies E-M"
                    QUADPLEX,
                    WORKING_AREA_FOLDER
                );

            // Common roots
            public static string OtherWorkingAreaRoot => Unc(QUADPLEX, OTHER_SHARE, WORKING_AREA_FOLDER);
            public static string OtherMoviesLibraryRoot => Unc(QUADPLEX, OTHER_SHARE, OTHER_BUS_FOLDER, MOVIES_FOLDER);

            public static string GetTvWorkingAreaRoot(char letter, out string label)
            {
                letter = char.ToUpperInvariant(letter);

                // Anything non-letter should go to '#'
                if (letter < 'A' || letter > 'Z') letter = '#';

                foreach (var r in TvRanges)
                {
                    if (r.Contains(letter))
                    {
                        label = r.Label;
                        return r.RootPath;
                    }
                }

                // Fail-safe: last range
                var last = TvRanges[TvRanges.Count - 1];
                label = last.Label;
                return last.RootPath;
            }

            public static string GetMovieWorkingAreaRoot(char letter, out string label)
            {
                letter = char.ToUpperInvariant(letter);
                if (letter < 'A' || letter > 'Z') letter = '#';

                foreach (var r in MovieRanges)
                {
                    if (r.Contains(letter))
                    {
                        label = r.Label;
                        return r.RootPath;
                    }
                }

                var last = MovieRanges[MovieRanges.Count - 1];
                label = last.Label;
                return last.RootPath;
            }

            // Convenience: locations where SRT scans commonly run
            public static string[] GetAllSrtScanLocations()
            {
                var list = new List<string>();

                // TV libraries
                foreach (var r in TvRanges)
                    list.Add(Unc(QUADPLEX, r.Label, TV_SHOWS_FOLDER));

                // Movie libraries
                foreach (var r in MovieRanges)
                    list.Add(Unc(QUADPLEX, r.Label, MOVIES_FOLDER));

                // Other
                list.Add(Unc(QUADPLEX, OTHER_SHARE, OTHER_BUS_FOLDER));

                return list.ToArray();
            }

            public static string[] GetMovieWorkingAreaRoots()
            {
                return MovieRanges
                    .Select(r => r.RootPath)
                    .Concat(new[] { OtherWorkingAreaRoot })
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            }
            public static string[] GetTvWorkingAreaRoots()
            {
                return TvRanges
                    .Select(r => r.RootPath)
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            }
        }

        static void Main(string[] args)
        {
            Type("Welcome to the BachFlix NFO Filer 3000! v1.14.4", 0, 0, 1, "blue");

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
                            keepAskingForChoice = CallSwitch(choice.Trim().ToLower());
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

        [DataContract]
        public class RenameTransactionLog
        {
            [DataMember] public string createdUtc { get; set; }
            [DataMember] public string mode { get; set; }
            [DataMember] public string sourceDirectory { get; set; }
            [DataMember] public bool dryRun { get; set; }
            [DataMember] public List<RenameMove> moves { get; set; } = new List<RenameMove>();
        }

        [DataContract]
        public class RenameMove
        {
            [DataMember] public string from { get; set; }
            [DataMember] public string to { get; set; }
        }

        /// <summary>
        /// Gives the main menu.
        /// </summary>
        private static string[] Menu()
        {
            chosenDirectory = "";

            Type("Please choose from one of the following options..", 0, 0, 1);
            Type("(Or do multiple options by separating them with a comma. i.e. 1,3)", 0, 0, 1);
            exitChoice = "0";
            Type(exitChoice + "- Exit", 0, 0, 1, "darkgray");

            const string NFO_FILE_CREATION_COLOR = "DarkGreen";
            Type("");
            Type("--- NFO File Creation ---", 0, 0, 1, NFO_FILE_CREATION_COLOR);
            missingCombinedEpisodeNfoFilesChoice = "n2";
            Type(missingCombinedEpisodeNfoFilesChoice + "- Create missing NFO files for Combined TV Show episodes.", 0, 0, 1, NFO_FILE_CREATION_COLOR);
            missingSeveralCombinedEpisodeNfoFilesChoice = "n6";
            Type(missingSeveralCombinedEpisodeNfoFilesChoice + "- Create missing NFO files for Several-Combined TV Show episodes.", 0, 0, 1, NFO_FILE_CREATION_COLOR);
            missingYoutubeNfoFilesChoice = "n3";
            Type(missingYoutubeNfoFilesChoice + "- Missing YouTube NFO Files", 0, 0, 1, NFO_FILE_CREATION_COLOR);
            overwriteAllYoutubeNfoFilesChoice = "n3o";
            Type(overwriteAllYoutubeNfoFilesChoice + "- Overwrite ALL YouTube NFO Files", 0, 0, 1, NFO_FILE_CREATION_COLOR);
            selectedYoutubeNfoFilesChoice = "n3s";
            Type(selectedYoutubeNfoFilesChoice + "- Selected YouTube NFO Files", 0, 0, 1, NFO_FILE_CREATION_COLOR);
            missingFitnessVideoNfoFilesChoice = "n4";
            Type(missingFitnessVideoNfoFilesChoice + "- Missing Fitness Video NFO Files", 0, 0, 1, NFO_FILE_CREATION_COLOR);
            overwriteAllFitnessVideoNfoFilesChoice = "n4o";
            Type(overwriteAllFitnessVideoNfoFilesChoice + "- Overwrite ALL Fitness Video NFO Files", 0, 0, 1, NFO_FILE_CREATION_COLOR);
            selectedFitnessVideoNfoFilesChoice = "n4s";
            Type(selectedFitnessVideoNfoFilesChoice + "- Selected Fitness Video NFO Files", 0, 0, 1, NFO_FILE_CREATION_COLOR);

            const string CONVERT_FILES_COLOR = "DarkCyan";
            Type("");
            Type("-- Convert Files ---", 0, 0, 1, CONVERT_FILES_COLOR);
            convertMoviesChoice = "5";
            Type(convertMoviesChoice + "- Movies", 0, 0, 1, CONVERT_FILES_COLOR);
            convertMoviesSlowChoice = "5s";
            Type(convertMoviesSlowChoice + "- Movies (Slow)", 0, 0, 1, CONVERT_FILES_COLOR);
            convertBonusFeaturesChoice = "6";
            Type(convertBonusFeaturesChoice + "- Bonus Features", 0, 0, 1, CONVERT_FILES_COLOR);
            convertBonusFeaturesSlowChoice = "6s";
            Type(convertBonusFeaturesSlowChoice + "- Bonus Features (Slow)", 0, 0, 1, CONVERT_FILES_COLOR);
            convertTvShowsChoice = "7";
            Type(convertTvShowsChoice + "- TV Shows", 0, 0, 1, CONVERT_FILES_COLOR);
            convertTvShowsSlowChoice = "7s";
            Type(convertTvShowsSlowChoice + "- TV Shows (Slow)", 0, 0, 1, CONVERT_FILES_COLOR);
            convertTempTvShowsChoice = "7t";
            Type(convertTempTvShowsChoice + "- Temp TV Shows", 0, 0, 1, CONVERT_FILES_COLOR);
            convertTempTVShowsSlowChoice = "7ts";
            Type(convertTempTVShowsSlowChoice + "- Temp TV Shows (Slow)", 0, 0, 1, CONVERT_FILES_COLOR);
            convertDirectoryChoice = "19";
            Type(convertDirectoryChoice + "- Convert a selected directory.", 0, 0, 1, CONVERT_FILES_COLOR);
            audioCompatibilityConverterChoice = "19a";
            Type(audioCompatibilityConverterChoice + "- Add AC3 5.1 compatibility audio from English EAC3 tracks.", 0, 0, 1, CONVERT_FILES_COLOR);
            audioCensorRunnerChoice = "19c";
            Type(audioCensorRunnerChoice + "- Scan SRT profanity for future clean audio muting.", 0, 0, 1, CONVERT_FILES_COLOR);

            const string TMDB_CALL_COLOR = "Green";
            Type("");
            Type("--- TMDB Call ---", 0, 0, 1, TMDB_CALL_COLOR);
            fetchTvShowPlotsChoice = "25";
            Type(fetchTvShowPlotsChoice + "- Insert TV Show plots into the Combined Episodes sheet.", 0, 0, 1, TMDB_CALL_COLOR);
            insertMissingCastMembers = "37";
            Type(insertMissingCastMembers + "- Insert Missing Cast members into the Google Sheet.", 0, 0, 1, TMDB_CALL_COLOR);

            const string UPDATE_GOOGLE_SHEET_COLOR = "Blue";
            Type("");
            Type("--- Update Google Sheet ---", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            Type("-- Movies Sheet", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            missingMovieNfoFilesChoice = "n1";
            Type(missingMovieNfoFilesChoice + "- Missing Movie NFO Files", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            overwriteAllMovieNfoFilesChoice = "n1o";
            Type(overwriteAllMovieNfoFilesChoice + "- Overwrite ALL Movie NFO Files", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            selectedMovieNfoFilesChoice = "n1s";
            Type(selectedMovieNfoFilesChoice + "- Selected Movie NFO Files", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            insertMissingMovieDataChoice = "10";
            Type(insertMissingMovieDataChoice + "- Insert movie data into the Google Sheet (plot, rating, & TMDB ID).", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            repeatInsertMissingMovieDataChoice = "10r";
            Type(repeatInsertMissingMovieDataChoice + "- Insert movie data into the Google Sheet and repeat if data is still loading.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            updateMovieDataChoice = "10o";
            Type(updateMovieDataChoice + "- Insert and overwrite movie data into the Google Sheet (plot, rating, & TMDB ID).", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            insertMissingTmdbIdsChoice = "11";
            Type(insertMissingTmdbIdsChoice + "- Insert missing TMDB IDs into the Google Sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            insertAndOverwriteTmdbIdsChoice = "11o";
            Type(insertAndOverwriteTmdbIdsChoice + "- Insert and override TMDB IDs in the Google Sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            getVideoResolutionChoice = "45";
            Type(getVideoResolutionChoice + "- Add Video Resolutions to Google Sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            overwriteVideoResolutionChoice = "45o";
            Type(overwriteVideoResolutionChoice + "- Overwrite Video Resolutions in Google Sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            getMovieWatchProviders = "43";
            Type(getMovieWatchProviders + "- Get movie streaming providers.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            getMovieWatchProvidersOfSelectedMovies = "43s";
            Type(getMovieWatchProvidersOfSelectedMovies + "- Get movie streaming providers of selected movies.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            addAllActorsCredits = "48";
            Type(addAllActorsCredits + "- Add all actors credits.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            clearSelectedRowInMoviesSheet = "51";
            Type(clearSelectedRowInMoviesSheet + "- Clear the data in selected rows.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            clearSelectedRowInMoviesSheetAndAddToSkipList = "51b";
            Type(clearSelectedRowInMoviesSheetAndAddToSkipList + "- Clear the data in selected rows & add the movie to the skip list.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            addSizeOfMovieDirectories = "52";
            Type(addSizeOfMovieDirectories + "- Add the size of Movie directories to the Movies Sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            overwriteSizeOfMovieDirectories = "52o";
            Type(overwriteSizeOfMovieDirectories + "- Overwrite the size of Movie directories to the Movies Sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            fileTypeChoice = "54";
            Type(fileTypeChoice + "- Add missing File Types to Google Sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            fileTypeOverwriteChoice = "54o";
            Type(fileTypeOverwriteChoice + "- Overwrite File Types in Google Sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            rebuildWebMoviesChoice = "59w";
            Type(rebuildWebMoviesChoice + "- Rebuild Web Movies from Movies sheet only.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            profileRequestProcessorChoice = "59";
            Type(profileRequestProcessorChoice + "- Process Profile Requests, rebuild Web Movies, sync Plex metadata, and refresh media libraries.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);



            Type("");
            Type("-- DB Sheet", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            missingTvShowNfoFilesChoice = "n7";
            Type(missingTvShowNfoFilesChoice + "- Missing TV Show NFO Files", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            overwriteAllTvShowNfoFilesChoice = "n7o";
            Type(overwriteAllTvShowNfoFilesChoice + "- Overwrite ALL TV Show NFO Files", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            selectedTvShowNfoFilesChoice = "n7s";
            Type(selectedTvShowNfoFilesChoice + "- Selected TV Show NFO Files", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            getTvShowWatchProviders = "43t";
            Type(getTvShowWatchProviders + "- Get TV Show possible record sources.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            addSizeOfTvShowDirectories = "24";
            Type(addSizeOfTvShowDirectories + "- Add the size of the TV Shows directories to the DB Sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            overwriteSizeOfTvShowDirectories = "24o";
            Type(overwriteSizeOfTvShowDirectories + "- Overwrite the size of the TV Shows directories to the DB Sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            insertMissingDbDataChoice = "28";
            Type(insertMissingDbDataChoice + "- Insert missing data into the DB sheet from TVDB.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            updateDbSheetChoice = "29";
            Type(updateDbSheetChoice + "- Update the DB sheet with updated info.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);

            Type("");
            Type("-- Combined Episodes Sheet", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            insertMissingCombinedEpisodesChoice = "30";
            Type(insertMissingCombinedEpisodesChoice + "- Insert missing data into the Combined Episodes sheet from TVDB.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            updateCombinedEpisodesChoice = "31";
            Type(updateCombinedEpisodesChoice + "- Update data in the Combined Episodes sheet from TVDB.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            writeToCombinedEpisodesChoice = "32";
            Type(writeToCombinedEpisodesChoice + "- Write video file names to the Combined Episodes sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);

            Type("");
            Type("-- Several Combined Episodes Sheet", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            insertMissingSeveralCombinedEpisodesChoice = "34";
            Type(insertMissingSeveralCombinedEpisodesChoice + "- Insert missing data into the Several Combined Episodes sheet from TVDB.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            writeToSeveralCombinedEpisodesChoice = "35";
            Type(writeToSeveralCombinedEpisodesChoice + "- Write video file names to the Several Combined Episodes sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);

            Type("");
            Type("-- Rename Episode Sheet", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            insertMissingEpisodeDataChoice = "33";
            Type(insertMissingEpisodeDataChoice + "- Write episode names to the Rename Episode sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);

            Type("");
            Type("-- YouTube Sheet", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);
            writeVideoFileNamesToYoutubeSheet = "36";
            Type(writeVideoFileNamesToYoutubeSheet + "- Write episode names to the YouTube sheet.", 0, 0, 1, UPDATE_GOOGLE_SHEET_COLOR);

            const string MISC_COLOR = "DarkYellow";
            Type("");
            Type("--- Misc. ---", 0, 0, 1, MISC_COLOR);
            Type("8- Count Files", 0, 0, 1, MISC_COLOR);
            Type("9- Remove the UPC numbers from the folder name.", 0, 0, 1, MISC_COLOR);
            Type("12- Move Kids movies.", 0, 0, 1, MISC_COLOR);
            Type("13- Copy JPG files. (Work in progress)", 0, 0, 1, MISC_COLOR);
            copyMovieFilesToDestinationChoice = "14c";
            Type(copyMovieFilesToDestinationChoice + "- Copy Movie files from Google Sheet list to chosen hard drive.", 0, 0, 1, MISC_COLOR);
            deleteMovieFilesAtDestinationChoice = "14d";
            Type(deleteMovieFilesAtDestinationChoice + "- Delete Movie files from Google Sheet list on chosen hard drive.", 0, 0, 1, MISC_COLOR);
            Type("15- Mark Owned Movies as D=Done || X=Not Done.", 0, 0, 1, MISC_COLOR);
            Type("16- Remove movies from TMDB List. (Work in progress)", 0, 0, 1, MISC_COLOR);
            removeMetadataChoice = "18";
            Type(removeMetadataChoice + "- Remove Metadata (including sub-folders).", 0, 0, 1, MISC_COLOR);
            Type("20- Add Comment to file.", 0, 0, 1, MISC_COLOR);
            createFoldersAndMoveFilesChoice = "21";
            Type(createFoldersAndMoveFilesChoice + "- Create directories and move files into them.", 0, 0, 1, MISC_COLOR);
            createFoldersAndMoveFilesAndSortChoice = "21b";
            Type(createFoldersAndMoveFilesAndSortChoice + "- Create directories and move files into them AND sort them into sub folders.", 0, 0, 1, MISC_COLOR);
            trimTitlesInDirectoryChoice = "22";
            Type(trimTitlesInDirectoryChoice + "- Trim titles in chosen directory.", 0, 0, 1, MISC_COLOR);
            bothTrimAndCreateFoldersChoice = "23";
            Type(bothTrimAndCreateFoldersChoice + "- Trim the titles AND create directories then move files into directories.", 0, 0, 1, MISC_COLOR);
            fixRecordedNamesChoice = "26";
            Type(fixRecordedNamesChoice + "- Fix recorded names.", 0, 0, 1, MISC_COLOR);
            copyMultipleFilesToOneLocationChoice = "27";
            Type(copyMultipleFilesToOneLocationChoice + "- Copy multiple chosen files to one location.", 0, 0, 1, MISC_COLOR);
            moveTvShowEpisodesToFolders = "38";
            Type(moveTvShowEpisodesToFolders + "- Sort TV Show Episodes to folders.", 0, 0, 1, MISC_COLOR);
            moveSameMovieFilesTopLevel = "39";
            Type(moveSameMovieFilesTopLevel + "- Move SameMovie files to top level.", 0, 0, 1, MISC_COLOR);
            deleteMovieFiles = "40";
            Type(deleteMovieFiles + "- Delete video files (including sub-folders).", 0, 0, 1, MISC_COLOR);
            deleteJpgFiles = "41";
            Type(deleteJpgFiles + "- Delete JPG files (including sub-folders).", 0, 0, 1, MISC_COLOR);
            addSeasonToFolderName = "42";
            Type(addSeasonToFolderName + "- Add Season to the Season name (i.e. 'S01' = 'Season 01').", 0, 0, 1, MISC_COLOR);
            moveFolderContentsChoice = "44";
            Type(moveFolderContentsChoice + "- Move folder contents (including sub-folders) == Under Construction.", 0, 0, 1, MISC_COLOR);
            getDirectorySizeChoice = "46";
            Type(getDirectorySizeChoice + "- Calculate the size of a folder.", 0, 0, 1, MISC_COLOR);
            changeTheSeason = "47";
            Type(changeTheSeason + "- Change the episodes in a folder to a different season.", 0, 0, 1, MISC_COLOR);
            changeEpisodesIntoTwoPartsChoice = "47b";
            Type(changeEpisodesIntoTwoPartsChoice + "- Change the episodes in a folder to combined episodes. (i.e. S01E01 becomes S01E01-E02)", 0, 0, 1, MISC_COLOR);
            revertEpisodeCombineChoice = "47br";
            Type(revertEpisodeCombineChoice + "- Revert a Combine Episodes run using the JSON log file.", 0, 0, 1, MISC_COLOR);
            resetEpisodeNumbersChoice = "47c";
            Type(resetEpisodeNumbersChoice + "- Reset the episode numbers in a folder to start at 01.", 0, 0, 1, MISC_COLOR);
            resetEpisodeNumbersToChosenNumberChoice = "47d";
            Type(resetEpisodeNumbersToChosenNumberChoice + "- Reset the episode numbers in a folder to a chosen number.", 0, 0, 1, MISC_COLOR);
            searchYoutubeAndDownloadMovieTrailersChoice = "49";
            Type(searchYoutubeAndDownloadMovieTrailersChoice + "- Search for and download movie trailers from YouTube. (This eats up the YouTube API quota very fast, use sparingly)", 0, 0, 1, MISC_COLOR);
            downloadMovieTrailersChoice = "50";
            Type(downloadMovieTrailersChoice + "- Download movies from YouTube using the YouTube IDs in the Google sheet", 0, 0, 1, MISC_COLOR);
            checkSrtFileNamesChoice = "53";
            Type(checkSrtFileNamesChoice + "- Verify that all SRT files have the ENG flag.", 0, 0, 1, MISC_COLOR);
            checkSelectedSrtFileNamesChoice = "53s";
            Type(checkSelectedSrtFileNamesChoice + "- Verify that the SRT files in the selected directory have the ENG flag.", 0, 0, 1, MISC_COLOR);
            renameMoviesFromSheetChoice = "55";
            Type(renameMoviesFromSheetChoice + "- Rename movies and SRT files using the IMDB Title from the Google Sheet.", 0, 0, 1, MISC_COLOR);
            movieFilenameScannerChoice = "55v";
            Type(movieFilenameScannerChoice + "- Verify movie/NFO/SRT/BIF/trailer/trickplay names against the Movies sheet and approve suggested renames.", 0, 0, 1, MISC_COLOR);
            moveReadyVideosChoice = "55m";
            Type(moveReadyVideosChoice + "- Move video files that are in 'Ready' folders to the 'Working Area' folders.", 0, 0, 1, MISC_COLOR);
            srtScoreRunnerChoice = "56";
            Type(srtScoreRunnerChoice + "- Score movie SRT files and add to the Google Sheet.", 0, 0, 1, MISC_COLOR);
            fileHealthRunnerChoice = "57";
            Type(fileHealthRunnerChoice + "- Scan Movies sheet for video File Health and write OK/BAD.", 0, 0, 1, "DarkYellow");
            fullMovieHealthCheck = "57s";
            Type(fullMovieHealthCheck + "- Check full video health for selected file/folder (no sheet update).", 0, 0, 1, "DarkYellow");
            plexPopulateRatingKeysChoice = "58";
            Type(plexPopulateRatingKeysChoice + "- Populate missing Plex ratingKeys into Movies sheet and mark Quick Create (match by IMDB/TMDB).", 0, 0, 1, "DarkYellow");
            plexOverwriteRatingKeysChoice = "58o";
            Type(plexOverwriteRatingKeysChoice + "- Overwrite changed Plex ratingKeys in Movies sheet and mark Quick Create (match by IMDB/TMDB).", 0, 0, 1, "DarkYellow");
            plexSyncMetadataChoice = "60";
            Type(plexSyncMetadataChoice + "- Sync Plex labels (Tags/NFO Body <tag>), sort title, and plot from Movies sheet.", 0, 0, 1, "DarkYellow");

            return Console.ReadLine().Split(',');

        } // End Menu()
        static bool CallSwitch(string choice)
        {
            bool keepAskingForChoice = true;
            try
            {
                Dictionary<string, int> autopopulateActorsSheetVariables = new Dictionary<string, int>
                {
                    { "Name", -1 },
                    { "person_id", -1 }
                };

                Dictionary<string, int> skipMovieIdsSheetVariables = new Dictionary<string, int>
                {
                    { "Skip", -1 }
                };

                Dictionary<string, int> dbSheetVariables = new Dictionary<string, int>
                {
                    { "Clean Name", -1 },
                    { "Clean Name with Year", -1 },
                    { "Content Rating", -1 },
                    { "Continuing?", -1 },
                    { "Found Locations", -1 },
                    { "Hard Drive Letter", -1 },
                    { NFO_BODY, -1 },
                    { QUICK_CREATE, -1 },
                    { ROW_NUM, -1 },
                    { "Season Count", -1 },
                    { "Series Name", -1 },
                    { "Size", -1 },
                    { "TVDB ID", -1 },
                    { "TVDB Slug", -1 },
                    { "Year", -1 }
                };

                Dictionary<string, int> tvShowStreamingProviderSheetVariables = new Dictionary<string, int>
                {
                    { ROW_NUM, -1 },
                    { "Clean Name with Year", -1 },
                    { "Series Name", -1 },
                    { "TVDB ID", -1 },
                    { "StreamFab", -1 },
                    { "Possible Record Source", -1 },
                    { "Resolution", -1 },
                    { "Format", -1 }
                };

                Dictionary<string, int> combinedEpisodeSheetVariables = new Dictionary<string, int>
                {
                    { "Combined Episode Name", -1 },
                    { "Episode 1 No.", -1 },
                    { "Episode 1 Plot", -1 },
                    { "Episode 1 Season", -1 },
                    { "Episode 1 Title", -1 },
                    { "Episode 2 No.", -1 },
                    { "Episode 2 Plot", -1 },
                    { "Episode 2 Season", -1 },
                    { "Episode 2 Title", -1 },
                    { "Lock Plot 1", -1 },
                    { "Lock Plot 2", -1 },
                    { "New Episode Name", -1 },
                    { NFO_BODY, -1 },
                    { ROW_NUM, -1 },
                    { "Show Title", -1 },
                    { "TMDB ID", -1 },
                    { "TVDB ID", -1 }
                };

                Dictionary<string, int> severalCombinedSheetVariables = new Dictionary<string, int>
                {
                    { "Combined Episode Name", -1 },
                    { "Episode 1 No.", -1 },
                    { "Episode 1 Plot", -1 },
                    { "Episode 1 Season", -1 },
                    { "Episode 1 Title", -1 },
                    { "Episode 2 No.", -1 },
                    { "Episode 2 Plot", -1 },
                    { "Episode 2 Season", -1 },
                    { "Episode 2 Title", -1 },
                    { "Episode 3 No.", -1 },
                    { "Episode 3 Plot", -1 },
                    { "Episode 3 Season", -1 },
                    { "Episode 3 Title", -1 },
                    { "Episode 4 No.", -1 },
                    { "Episode 4 Plot", -1 },
                    { "Episode 4 Season", -1 },
                    { "Episode 4 Title", -1 },
                    { "Episode 5 No.", -1 },
                    { "Episode 5 Plot", -1 },
                    { "Episode 5 Season", -1 },
                    { "Episode 5 Title", -1 },
                    { "Episode 6 No.", -1 },
                    { "Episode 6 Plot", -1 },
                    { "Episode 6 Season", -1 },
                    { "Episode 6 Title", -1 },
                    { "Episode 7 No.", -1 },
                    { "Episode 7 Plot", -1 },
                    { "Episode 7 Season", -1 },
                    { "Episode 7 Title", -1 },
                    { "Episode 8 No.", -1 },
                    { "Episode 8 Plot", -1 },
                    { "Episode 8 Season", -1 },
                    { "Episode 8 Title", -1 },
                    { "Episode 9 No.", -1 },
                    { "Episode 9 Plot", -1 },
                    { "Episode 9 Season", -1 },
                    { "Episode 9 Title", -1 },
                    { "Episode 10 No.", -1 },
                    { "Episode 10 Plot", -1 },
                    { "Episode 10 Season", -1 },
                    { "Episode 10 Title", -1 },
                    { "New Episode Name", -1 },
                    { NFO_BODY, -1 },
                    { ROW_NUM, -1 },
                    { "Series Name", -1 },
                    { "TMDB ID", -1 },
                    { "TVDB ID", -1 }
                };

                Dictionary<string, int> youtubeSheetVariables = new Dictionary<string, int>
                {
                    { CLEAN_TITLE, -1 },
                    { DIRECTORY, -1 },
                    { NFO_BODY, -1 },
                    { QUICK_CREATE, -1 },
                    { ROW_NUM, -1 },
                    { STATUS, -1 }
                };

                Dictionary<string, int> fitnessSheetVariables = new Dictionary<string, int>
                {
                    { "Program", -1 },
                    { "Subfolder", -1 },
                    { "Name", -1 },
                    { "Title", -1 },
                    { NFO_BODY, -1 }
                };

                Dictionary<string, int> recordedNamesSheetVariables = new Dictionary<string, int>
                {
                    { "Recorded Name", -1 },
                    { "Actual Name", -1 }
                };

                Dictionary<string, int> renameEpisodesSheetVariables = new Dictionary<string, int>
                {
                    { "Original Name", -1 },
                    { ROW_NUM, -1 }
                };

                if (choice.Equals(exitChoice))
                {
                    Type("Thank you, have a nice day! \\(^.^)/", 0, 0, 1);
                    keepAskingForChoice = false;

                }
                else if (choice.Equals(missingMovieNfoFilesChoice)) // NFO files for New Movies - does not overwrite any, just puts in missing NFO files.
                {
                    Type("Insert missing NFO Files. Let's go!", 0, 0, 1);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    DisplayMessage("info", "Searching for missing NFO files...");
                    CreateNfoFiles(movieData, movieSheetVariables, 3);

                }
                else if (choice.Equals(overwriteAllMovieNfoFilesChoice)) // NFO files for All Movies - overwrite old NFO files AND put in new ones.
                {
                    Type("Insert missing AND overwrite NFO Files. Let's go!", 0, 0, 1);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    DisplayMessage("info", "Overwriting ALL NFO files...");
                    CreateNfoFiles(movieData, movieSheetVariables, 1);
                }
                else if (choice.Equals(selectedMovieNfoFilesChoice)) // NFO files for Selected Movies - overwrite or put in new ones. if they are selected.
                {
                    Type("Insert selected NFO Files. Let's go!", 0, 0, 1);

                    RunSelectedMovieNfoFiles();
                }
                else if (choice.Equals(missingTvShowNfoFilesChoice))
                {
                    Type("Insert missing NFO Files for TV Shows. Let's go!", 0, 0, 1);

                    IList<IList<Object>> tvShowData = CallGetData(dbSheetVariables, DB_TITLE_RANGE, DB_DATA_RANGE);

                    DisplayMessage("info", "Searching for missing NFO files...");
                    CreateTvShowNfoFiles(tvShowData, dbSheetVariables, 3);
                }
                else if (choice.Equals(overwriteAllTvShowNfoFilesChoice))
                {
                    Type("Insert missing NFO Files for TV Shows. Let's go!", 0, 0, 1);

                    IList<IList<Object>> tvShowData = CallGetData(dbSheetVariables, DB_TITLE_RANGE, DB_DATA_RANGE);

                    DisplayMessage("info", "Searching for missing NFO files...");
                    CreateTvShowNfoFiles(tvShowData, dbSheetVariables, 1);
                }
                else if (choice.Equals(selectedTvShowNfoFilesChoice))
                {
                    Type("Insert selected NFO Files for TV Shows. Let's go!", 0, 0, 1);

                    IList<IList<Object>> tvShowData = CallGetData(dbSheetVariables, DB_TITLE_RANGE, DB_DATA_RANGE);

                    DisplayMessage("info", "Adding/Overwriting selected TV Show NFO files...");
                    CreateTvShowNfoFiles(tvShowData, dbSheetVariables, 2);
                }
                else if (choice.Equals(missingCombinedEpisodeNfoFilesChoice)) // Create missing NFO files for TV Show episodes.
                {
                    Type("Create missing NFO files for Combined TV Show episodes. Let's go!", 0, 0, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        IList<IList<Object>> movieData = CallGetData(combinedEpisodeSheetVariables, COMBINED_EPISODES_TITLE_RANGE, COMBINED_EPISODES_DATA_RANGE);

                        CreateMissingCombinedEpisodeNfoFiles(movieData, combinedEpisodeSheetVariables, directory);
                    }
                }
                else if (choice.Equals(missingSeveralCombinedEpisodeNfoFilesChoice)) // Create missing NFO files for TV Show episodes.
                {
                    Type("Create missing NFO files for Several-Combined TV Show episodes. Let's go!", 0, 0, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        IList<IList<Object>> movieData = CallGetData(severalCombinedSheetVariables, SEVERAL_COMBINED_EPISODES_TITLE_RANGE, SEVERAL_COMBINED_EPISODES_DATA_RANGE);

                        CreateMissingSeveralCombinedEpisodeNfoFiles(movieData, severalCombinedSheetVariables, directory);
                    }
                }
                else if (choice.Equals(missingYoutubeNfoFilesChoice)) // NFO files for New videos - does not overwrite any, just puts in missing NFO files.
                {
                    Type("Create missing YouTube NFO Files. Let's go!", 0, 0, 1);

                    IList<IList<Object>> videoData = CallGetData(youtubeSheetVariables, YOUTUBE_TITLE_RANGE, YOUTUBE_DATA_RANGE);

                    CreateNfoFiles(videoData, youtubeSheetVariables, 3, true);
                }
                else if (choice.Equals(overwriteAllYoutubeNfoFilesChoice)) // NFO files for All videos - overwrite old NFO files AND put in new ones.
                {
                    Type("Overwrite ALL YouTube NFO Files. Let's go!", 0, 0, 1);

                    IList<IList<Object>> movieData = CallGetData(youtubeSheetVariables, YOUTUBE_TITLE_RANGE, YOUTUBE_DATA_RANGE);

                    CreateNfoFiles(movieData, youtubeSheetVariables, 1, true);
                }
                else if (choice.Equals(selectedYoutubeNfoFilesChoice)) // NFO files for Selected videos - overwrite or put in new ones. if they are selected.
                {
                    Type("Create/Overwrite selected YouTube NFO Files. Let's go!", 0, 0, 1);

                    IList<IList<Object>> movieData = CallGetData(youtubeSheetVariables, YOUTUBE_TITLE_RANGE, YOUTUBE_DATA_RANGE);

                    CreateNfoFiles(movieData, youtubeSheetVariables, 2, true);
                }
                else if (choice.Equals(missingFitnessVideoNfoFilesChoice)) // NFO files for New videos - does not overwrite any, just puts in missing NFO files.
                {
                    Type("This method is still in the works, please try another one.", 0, 0, 1, "Yellow");
                    //Type("Create missing Fitness Video NFO Files. Let's go!", 0, 0, 1, "Blue");

                    //IList<IList<Object>> videoData = CallGetData(fitnessSheetVariables, FITNESS_VIDEO_TITLE_RANGE, FITNESS_VIDEO_DATA_RANGE);

                    //BachFlixNfo.MissingFitnessVideoNfoFiles(videoData, fitnessSheetVariables);
                }
                else if (choice.Equals(overwriteAllFitnessVideoNfoFilesChoice)) // NFO files for All videos - overwrite old NFO files AND put in new ones.
                {
                    Type("Overwrite ALL YouTube NFO Files. Let's go!", 0, 0, 1, "Blue");

                    IList<IList<Object>> videoData = CallGetData(fitnessSheetVariables, FITNESS_VIDEO_TITLE_RANGE, FITNESS_VIDEO_DATA_RANGE);

                    BachFlixNfo.BachFlixNfo.OverwriteFitnessVideoNfoFiles(videoData, fitnessSheetVariables);
                }
                else if (choice.Equals(convertMoviesChoice))
                {
                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    ConvertVideo(movieData, movieSheetVariables, "--preset-import-file MP4_RF22f.json -Z \"MP4 RF22f\"");
                }
                else if (choice.Equals(convertMoviesSlowChoice))
                {
                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    ConvertVideo(movieData, movieSheetVariables, "--preset-import-file MP4_RF22s.json -Z \"MP4 RF22s\"");
                }
                else if (choice.Equals(convertDirectoryChoice))
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

                            fileSize = FormatSize(fileSizeBytes, true);


                            string plural = videoFiles.Count == 1 ? " file " : " files ";

                            DisplayMessage("info", "The size of the " + videoFiles.Count + plural + "is: ", 0);
                            DisplayMessage("data", fileSize);

                            // Send those video files off to be converted.
                            ConvertHandbrakeList(videoFiles);

                            ResetGlobals();

                            // Now move video files over to the conversion folder.
                            MoveVideoFilesToHoldFolder(directory);

                            // Check the Priority folder for more videos.
                            fileEntries = Directory.GetFiles(directory + @"\Priority");

                            // Filter out the files that aren't video files.
                            ArrayList videoFilesToMove = GrabMovieFiles(fileEntries);
                            if (videoFilesToMove.Count > 0)
                            {
                                DisplayMessage("info", "Grabbing the next video from the Priority folder");
                                int i = 0;
                                foreach (var moveFile in videoFilesToMove)
                                {
                                    if (i < 1)
                                    {
                                        MoveDirectory(moveFile.ToString(), Path.GetFullPath(Path.Combine(moveFile.ToString(), @"..\..\" + Path.GetFileName(moveFile.ToString()))));
                                    }
                                    i++;
                                }
                            }
                            else
                            {
                                // Check the Hold folder for more videos.
                                fileEntries = Directory.GetFiles(directory + @"\Hold");

                                // Filter out the files that aren't video files.
                                videoFilesToMove = GrabMovieFiles(fileEntries);
                                if (videoFilesToMove.Count > 0)
                                {
                                    DisplayMessage("info", "Grabbing the next video from the Hold folder");
                                    int i = 0;
                                    foreach (var moveFile in videoFilesToMove)
                                    {
                                        if (i < 1)
                                        {
                                            MoveDirectory(moveFile.ToString(), Path.GetFullPath(Path.Combine(moveFile.ToString(), @"..\..\" + Path.GetFileName(moveFile.ToString()))));
                                        }
                                        i++;
                                    }
                                }
                            }

                            // Check for more files.
                            DisplayMessage("warning", "Checking for more files... ");
                            fileEntries = Directory.GetFiles(directory);
                            ArrayList videoFilesCheck = GrabMovieFiles(fileEntries);
                            if (videoFilesCheck.Count == 0) keepLooping = false;
                            else DisplayMessage("info", "More files found. Restarting conversion...");
                        } while (keepLooping);
                    }

                }
                else if (choice.Equals(moveTvShowEpisodesToFolders))
                {
                    DisplayMessage("info", "Filter TV Show Episodes to season folders. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        // Grab all files in the directory.
                        DisplayMessage("warning", "Grabbing all files... ", 0);
                        string[] fileEntries = Directory.GetFiles(directory);
                        DisplayMessage("success", "DONE");

                        if (fileEntries.Length > 0)
                        {
                            foreach (var fileToMove in fileEntries)
                            {
                                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileToMove.ToString());
                                string fileName = Path.GetFileName(fileToMove.ToString());
                                string fileDirectory = Path.GetDirectoryName(fileToMove.ToString());
                                string[] splitTitle = fileName.Split(new[] { " - " }, StringSplitOptions.None);

                                if (splitTitle.Length > 1)
                                {
                                    try
                                    {
                                        string season = "\\Season " + splitTitle[1].Substring(1, splitTitle[1].ToUpper().IndexOf('E') - 1);
                                        Directory.CreateDirectory(fileDirectory + season);
                                        MoveDirectory(fileToMove.ToString(), fileDirectory + season + "\\" + fileName);
                                    }
                                    catch (Exception err)
                                    {
                                        DisplayMessage("error", err.Message);
                                    }
                                }
                            }
                        }
                    }

                }
                else if (choice.Equals(moveSameMovieFilesTopLevel))
                {
                    DisplayMessage("info", "Move SameMovie file to top level. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0" && Directory.Exists(directory))
                    {
                        RecurseSameMovieFolder(directory, directory);
                    }

                }
                else if (choice.Equals(deleteMovieFiles))
                {
                    DisplayMessage("info", "Delete video files. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0" && Directory.Exists(directory))
                    {
                        DeleteMoviesInFolder(directory);
                    }

                }
                else if (choice.Equals(deleteJpgFiles))
                {
                    DisplayMessage("info", "Delete JPG files. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0" && Directory.Exists(directory))
                    {
                        DeleteJpgsInFolder(directory);
                    }

                }
                else if (choice.Equals(addSeasonToFolderName))
                {
                    DisplayMessage("info", "Add Season to folder name. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0" && Directory.Exists(directory))
                    {
                        string[] subdirectoryEntries = Directory.GetDirectories(directory);
                        foreach (string subdirectory in subdirectoryEntries)
                        {
                            string directoryName = Path.GetDirectoryName(subdirectory);
                            string fileName = Path.GetFileName(subdirectory);
                            if (!fileName.Contains("Season "))
                            {
                                string newName = fileName.Replace("S", "Season ");
                                string newDirectory = Path.Combine(directoryName, newName);
                                Directory.Move(subdirectory, newDirectory);
                            }
                        }
                    }

                }
                else if (choice.Equals(getMovieWatchProviders))
                {
                    DisplayMessage("info", "Fill in the movie streaming providers. Let's go!");

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    FillInStreamingProviders(movieData, movieSheetVariables);
                }
                else if (choice.Equals(getMovieWatchProvidersOfSelectedMovies))
                {
                    DisplayMessage("info", "Fill in the movie streaming providers of selected movies. Let's go!");

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    FillInStreamingProviders(movieData, movieSheetVariables, true);
                }
                else if (choice.Equals(getTvShowWatchProviders))
                {
                    DisplayMessage("info", "Fill in the TV Show possible record sources. Let's go!");

                    IList<IList<Object>> tvShowData = CallGetData(tvShowStreamingProviderSheetVariables, DB_TITLE_RANGE, DB_DATA_RANGE);
                    FillInTvShowStreamingProviders(tvShowData, tvShowStreamingProviderSheetVariables);
                }
                else if (choice.Equals(addAllActorsCredits))
                {
                    DisplayMessage("info", "Fill in all the actors credits. Let's go!");

                    IList<IList<Object>> autoPopulateActorsData = CallGetData(autopopulateActorsSheetVariables, AUTOPOULATE_ACTORS_TITLE_RANGE, AUTOPOULATE_ACTORS_DATA_RANGE, "Gathering list of actors... ");
                    IList<IList<Object>> skipMovieIdsData = CallGetData(skipMovieIdsSheetVariables, SKIP_ACTORS_ID_TITLE_RANGE, SKIP_ACTORS_ID_DATA_RANGE, "Gathering list of movie IDs to skip... ");

                    if (autoPopulateActorsData != null)
                    {
                        string actorPlural = autoPopulateActorsData.Count == 1 ? " actor " : " actors ";
                        int count = 1;
                        IList<IList<Object>> movieData = new List<IList<object>> { };
                        DisplayMessage("info", autoPopulateActorsData.Count + actorPlural + "found. Now stepping through each actor to get data.");
                        bool moviesAddedToSheet = true;

                        foreach (var actor in autoPopulateActorsData)
                        {
                            Type("");
                            DisplayMessage("info", "----------------------------------------------------");

                            if (moviesAddedToSheet)
                            {
                                movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                            }

                            string message = $"{count} of {autoPopulateActorsData.Count} - {actor[0].ToString()} - ";

                            DisplayMessage("warning", "Gathering movie credits for ", 0);
                            DisplayMessage("success", actor[0].ToString(), 0);
                            DisplayMessage("warning", "... ", 0);
                            dynamic actorMovieCredits = TmdbApi.ActorsGetMovieCredits(actor[1].ToString());
                            DisplayMessage("success", "DONE");

                            moviesAddedToSheet = FillInActorMovieCredits(movieData, movieSheetVariables, actorMovieCredits.cast, ref skipMovieIdsData, message);
                            count++;
                        }
                    }
                    else
                    {
                        DisplayMessage("warning", "No Actors found, Add names and person_id to the Autopopulate Actors sheet.");
                    }
                }
                else if (choice.Equals(clearSelectedRowInMoviesSheet))
                {
                    DisplayMessage("info", "Clear the data from selected rows. Let's go!");

                    IList<IList<Object>> movieData = CallGetData(clearableMovieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    ClearSelectedRowData(movieData, clearableMovieSheetVariables);
                }
                else if (choice.Equals(clearSelectedRowInMoviesSheetAndAddToSkipList))
                {
                    DisplayMessage("info", "Clear the data from selected rows, and add the selected videos to the skip list. Let's go!");

                    IList<IList<Object>> movieData = CallGetData(clearableMovieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    IList<IList<Object>> skipMovieIdsData = CallGetData(skipMovieIdsSheetVariables, SKIP_ACTORS_ID_TITLE_RANGE, SKIP_ACTORS_ID_DATA_RANGE, "Gathering list of movie IDs to skip... ");
                    ClearSelectedRowData(movieData, clearableMovieSheetVariables, skipMovieIdsData);
                }
                else if (choice.Equals(moveFolderContentsChoice))
                {
                    DisplayMessage("info", "Move contents of a folder. Let's go!");
                    var srcDirectory = AskForDirectory("Give me the path to the folder of content you want to move");

                    if (srcDirectory != "0")
                    {
                        var destDirectory = AskForDirectory("Now give me the folder of where you want it moved to");

                        if (destDirectory != "0")
                        {
                            MoveFolderContent(srcDirectory, destDirectory);
                        }
                    }
                }
                else if (choice.Equals(getDirectorySizeChoice))
                {
                    DisplayMessage("info", "Calculate the size of a folder. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        long folderSize = CalculateFolderSize(directory);
                        Type("Folder size is: ", 0, 0, 0, "Blue");
                        Type(FormatSize(folderSize, true), 0, 0, 1, "Green");
                    }
                }
                else if (choice.Equals(changeTheSeason))
                {
                    DisplayMessage("info", "This method takes a directory path full of TV Show episodes and a number to change the season number to", 2);

                    DisplayMessage("question", "Enter 1 or y to continue, or 0 or n to go back to the menu.");
                    string wantToContinue = RemoveCharFromString(Console.ReadLine(), '"');
                    if (wantToContinue == "1" || wantToContinue.ToUpper() == "Y")
                    {
                        var directory = AskForDirectory("Give me the directory full of TV show episodes.");
                        if (directory != "0")
                        {
                            // Grab all files in the directory.
                            Type("Grabbing all files... ", 0, 0, 0, "Yellow");
                            string[] fileEntries = Directory.GetFiles(directory);
                            Type("DONE", 0, 0, 2, "Green");

                            DisplayMessage("question", "Now, what season would you like these changed to? (Number only including the leading zero i.e. 01)");
                            string newSeasonNumber = RemoveCharFromString(Console.ReadLine(), '"');

                            // Use a regular expression to find and replace the season number
                            string pattern = @"S(\d{2})E(\d{2})";
                            string replacement = $"S{newSeasonNumber:D2}E$2";

                            foreach (var originalFile in fileEntries)
                            {
                                string newFile = Regex.Replace(originalFile, pattern, replacement);
                                File.Move(originalFile, newFile);
                            }
                            Type("DONE", 0, 0, 1, "Green");
                        }
                    }
                }
                else if (choice.Equals(changeEpisodesIntoTwoPartsChoice))
                {
                    DisplayMessage("info", "Smart normalize/combine episodes in selected folder, let's go!", 2);

                    var directory = AskForDirectory("Give me the directory full of TV show episodes.");
                    if (directory != "0")
                    {

                        // --- Dry run prompt
                        DisplayMessage("question", "Dry run first? (y/n)  (Dry run prints + logs but does NOT rename files)", 1);
                        string dryInput = Console.ReadLine();
                        bool dryRun = !string.IsNullOrWhiteSpace(dryInput) &&
                                      dryInput.Trim().Equals("y", StringComparison.OrdinalIgnoreCase);

                        // Grab all MKVs in the directory (top-level only).
                        Type("Grabbing all MKV files... ", 0, 0, 0, "Yellow");
                        string[] fileEntries = Directory.GetFiles(directory, "*.mkv", SearchOption.TopDirectoryOnly);
                        Type("DONE", 0, 0, 2, "Green");

                        if (fileEntries.Length == 0)
                        {
                            Type("No videos found! Are they the right format?", 0, 0, 2, "Red");
                        }
                        else
                        {
                            // Match: S04E03  or  S04E03-E04
                            Regex epRegex = new Regex(
                                @"S(?<season>\d{2})E(?<ep>\d{2})(?:-E(?<ep2>\d{2}))?",
                                RegexOptions.IgnoreCase | RegexOptions.Compiled);

                            // Build an ordered list based on the MKVs (driver files)
                            var parsed = new List<(string MkvPath, int Season, int Ep, bool IsCombined)>();

                            foreach (var mkv in fileEntries)
                            {
                                string name = Path.GetFileNameWithoutExtension(mkv);
                                Match m = epRegex.Match(name);

                                if (!m.Success)
                                    continue;

                                int season = int.Parse(m.Groups["season"].Value);
                                int ep = int.Parse(m.Groups["ep"].Value);

                                // Combined heuristic:
                                // - already has -E## in the episode token
                                // - OR underscore in filename with spaces around it
                                bool isCombined = m.Groups["ep2"].Success || name.Contains(" _ ");

                                parsed.Add((mkv, season, ep, isCombined));
                            }

                            if (parsed.Count == 0)
                            {
                                Type("No MKVs matched SxxEyy naming pattern.", 0, 0, 2, "Red");
                            }
                            else
                            {
                                parsed = parsed
                                    .OrderBy(x => x.Season)
                                    .ThenBy(x => x.Ep)
                                    .ToList();

                                // Starting episode = first episode found in folder
                                int counter = parsed.Min(x => x.Ep);

                                DisplayMessage("info", $"Starting episode counter at E{counter:D2}.", 1);
                                DisplayMessage("info", dryRun ? "DRY RUN enabled (no files will be moved)." : "LIVE RUN enabled (files WILL be renamed).", 2);

                                // Build the log
                                var log = new RenameTransactionLog
                                {
                                    createdUtc = DateTime.UtcNow.ToString("o"),
                                    mode = "SmartCombine47b",
                                    sourceDirectory = directory,
                                    dryRun = dryRun
                                };

                                foreach (var entry in parsed)
                                {
                                    string mkvPath = entry.MkvPath;
                                    string baseName = Path.GetFileNameWithoutExtension(mkvPath);

                                    // Compute replacement token
                                    string newToken = entry.IsCombined
                                        ? $"S{entry.Season:D2}E{counter:D2}-E{counter + 1:D2}"
                                        : $"S{entry.Season:D2}E{counter:D2}";

                                    // Rename all files that share the same base name (mkv + srt + nfo + etc.)
                                    string[] sidecars = Directory.GetFiles(directory, baseName + ".*", SearchOption.TopDirectoryOnly);

                                    foreach (var file in sidecars)
                                    {
                                        string fileNameOnly = Path.GetFileName(file);

                                        if (!epRegex.IsMatch(fileNameOnly))
                                            continue;

                                        string newFileNameOnly = epRegex.Replace(fileNameOnly, newToken, 1);
                                        if (newFileNameOnly.Equals(fileNameOnly, StringComparison.OrdinalIgnoreCase))
                                            continue;

                                        string newFullPath = Path.Combine(directory, newFileNameOnly);

                                        // Safety: do not overwrite
                                        if (File.Exists(newFullPath))
                                        {
                                            DisplayMessage("error", $"Target already exists, skipping: {newFileNameOnly}", 1);
                                            continue;
                                        }

                                        // Log it
                                        log.moves.Add(new RenameMove { from = file, to = newFullPath });

                                        // Print it
                                        Console.WriteLine(entry.IsCombined ? "[COMBINE]" : "[SINGLE]");
                                        Console.WriteLine("Renamed:");
                                        Console.WriteLine(Path.GetFileName(file));
                                        Console.WriteLine(newFileNameOnly);
                                        Console.WriteLine();

                                        // Move it (only if live)
                                        if (!dryRun)
                                        {
                                            File.Move(file, newFullPath);
                                        }
                                    }

                                    // advance the counter
                                    counter += entry.IsCombined ? 2 : 1;
                                }

                                // Write the JSON log (even for dry run)
                                string logDir = @"C:\Users\Brandon\Desktop\BachFlixNfo Logs\Combine Episodes";
                                Directory.CreateDirectory(logDir);

                                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                                ExtractShowAndSeasonFromPath(directory, out string showName, out string seasonLabel);

                                string modeTag = dryRun ? "DRYRUN" : "LIVE";
                                string logFileName = $"{showName} - {seasonLabel} - {modeTag} - {stamp}.json";

                                string logPath = Path.Combine(logDir, logFileName);
                                WriteJsonLog(logPath, log);

                                DisplayMessage("success", $"Combine log written: {logPath}", 2);
                                Type("DONE", 0, 0, 1, "Green");
                            }
                        }
                    }
                }
                else if (choice == revertEpisodeCombineChoice)
                {
                    RevertEpisodeCombineFromLog();
                }
                else if (choice.Equals(resetEpisodeNumbersChoice))
                {
                    DisplayMessage("info", "Reset episode numbers in chosen folder, let's go!", 2);

                    UpdateEpisodeNumbersInChosenDirectory();
                }
                else if (choice.Equals(checkSrtFileNamesChoice))
                {
                    DisplayMessage("info", "Verify that all SRT files have the ENG flag, let's go!", 2);

                    // Centralized in DriveMapping
                    string[] srtLocations = DriveMapping.GetAllSrtScanLocations();

                    verifySrtFileNames(srtLocations);

                    displayMissingEngResults();
                }
                else if (choice.Equals(checkSelectedSrtFileNamesChoice))
                {
                    DisplayMessage("info", "Verify that the SRT files in the selected directory have the ENG flag, let's go!", 2);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        string[] srtLocation = { directory };

                        verifySrtFileNames(srtLocation);
                    }

                    displayMissingEngResults();
                }
                else if (choice.Equals(resetEpisodeNumbersToChosenNumberChoice))
                {
                    DisplayMessage("info", "Reset episode numbers in chosen folder to selected number, let's go!", 2);

                    DisplayMessage("question", "Now, what episode number would you like to reset to? (Number only including the leading zero i.e. 01)");
                    string newEpisodeNumber = RemoveCharFromString(Console.ReadLine(), '"');

                    UpdateEpisodeNumbersInChosenDirectory(int.Parse(newEpisodeNumber));
                }
                else if (choice.Equals(searchYoutubeAndDownloadMovieTrailersChoice))
                {
                    DisplayMessage("info", "Search for and download movie trailers from YouTube. Let's go!", 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    SearchAndDownloadMovieTrailers(movieData, movieSheetVariables);
                }
                else if (choice.Equals(downloadMovieTrailersChoice))
                {
                    DisplayMessage("info", "Get the YouTube IDs from Google sheets and use yt-dlp to download them. Let's go!", 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    DownloadMovieTrailers(movieData, movieSheetVariables);
                }
                else if (choice.Equals(getVideoResolutionChoice))
                {
                    DisplayMessage("info", "Add Video Resolutions to the Google Sheet. Let's go!");

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    FillInVideoResolution(movieData, movieSheetVariables, false);
                }
                else if (choice.Equals(overwriteVideoResolutionChoice))
                {
                    DisplayMessage("info", "Overwrite Video Resolutions in Google Sheet. Let's go!");

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    FillInVideoResolution(movieData, movieSheetVariables, true);
                }
                else if (choice.Equals(addSizeOfTvShowDirectories))
                {
                    Type("Add the size of TV Show directories to the DB sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(dbSheetVariables, DB_TITLE_RANGE, DB_DATA_RANGE);

                    InsertTvShowDirectorySizesIntoDbSheet(movieData, dbSheetVariables, false);
                }
                else if (choice.Equals(overwriteSizeOfTvShowDirectories))
                {
                    Type("Overwrite the size of TV Show directories to the DB sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(dbSheetVariables, DB_TITLE_RANGE, DB_DATA_RANGE);

                    InsertTvShowDirectorySizesIntoDbSheet(movieData, dbSheetVariables, true);
                }
                else if (choice.Equals(addSizeOfMovieDirectories))
                {
                    Type("Add the size of Movie directories to the Movies Sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    InsertMovieDirectorySizesIntoMoviesSheet(movieData, movieSheetVariables, false);
                }
                else if (choice.Equals(overwriteSizeOfMovieDirectories))
                {
                    Type("Overwrite the size of Movie directories to the Movies Sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    InsertMovieDirectorySizesIntoMoviesSheet(movieData, movieSheetVariables, true);
                }
                else if (choice.Equals(fileTypeChoice))
                {
                    Type("Add missing File Types to the Google Sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    FillInFileTypes(movieData, movieSheetVariables, false);
                }
                else if (choice.Equals(fileTypeOverwriteChoice))
                {
                    Type("Add missing File Types to the Google Sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    FillInFileTypes(movieData, movieSheetVariables, true);
                }
                else if (choice.Equals(renameMoviesFromSheetChoice))
                {
                    DisplayMessage("info", "Rename movies and SRT files using the IMDB Title from the Google Sheet. Let's go!", 2);

                    var directory = AskForDirectory("Enter the directory that contains the movies to rename:");
                    if (directory != "0")
                    {
                        // Grab all files and filter to just video files (movies) using your existing helper.
                        string[] fileEntries = Directory.GetFiles(directory, "*.*", SearchOption.AllDirectories);
                        ArrayList movieFiles = GrabMovieFiles(fileEntries, showMessage: false);

                        if (movieFiles.Count == 0)
                        {
                            DisplayMessage("warning", "No video files found in the selected directory.");
                        }
                        else
                        {
                            IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                            RenameMoviesAndSrtsFromSheet(movieData, movieSheetVariables, movieFiles, directory);
                        }
                    }
                }
                else if (choice.Equals(movieFilenameScannerChoice))
                {
                    Console.WriteLine();
                    DisplayMessage("info", "Verifying movie, sidecar, trailer, and trickplay names against the Movies sheet. Let's go!", 2);

                    SheetsService sheetsService = CreateSheetsService();
                    MovieFilenameSheetScanner.RunInteractive(
                        sheetsService,
                        SPREADSHEET_ID,
                        "Movies",
                        (t, msg, ind) => DisplayMessage(t, msg, ind));
                }
                else if (choice.Equals(moveReadyVideosChoice))
                {
                    DisplayMessage("info", "55m - Scan Ready folders, rename movies from sheet (ask early), then move Movies/TV to Working Areas. Let's go!", 2);

                    Run55m_RenameMoviesFirst_ThenMove();
                }
                else if (choice.Equals(srtScoreRunnerChoice))
                {
                    Console.WriteLine();
                    DisplayMessage("info", "Run subtitle scoring. Let's go!", 2);

                    // --- duplicate auth code for now (same pattern as WriteSingleCellToSheet / BulkWriteToSheet) ---
                    UserCredential credential;

                    using (var stream =
                        new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                    {
                        string credPath = "token.json";
                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.FromStream(stream).Secrets,
                            SCOPES,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;
                    }

                    SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = "Google-SheetsSample/0.1",  // same as your other helpers
                    });
                    // --- end local SheetsService setup ---

                    SubtitleScoreRunner.Run(sheetsService, SPREADSHEET_ID, "Movies");
                }
                else if (choice.Equals(fileHealthRunnerChoice))
                {
                    Console.WriteLine();
                    DisplayMessage("info", "Scanning Movies sheet for video File Health and updating 'File Health' column. Let's go!", 2);

                    // You already have a helper that builds a SheetsService:
                    SheetsService sheetsService = CreateSheetsService();

                    VideoHealthSheetRunner.Run(sheetsService, SPREADSHEET_ID, "Movies");
                }
                else if (choice.Equals(plexPopulateRatingKeysChoice) || choice.Equals(plexOverwriteRatingKeysChoice))
                {
                    Console.WriteLine();
                    bool overwriteRatingKeys = choice.Equals(plexOverwriteRatingKeysChoice);
                    DisplayMessage(
                        "info",
                        overwriteRatingKeys
                            ? "Overwriting Plex ratingKeys in Movies sheet (match by IMDB/TMDB). Let's go!"
                            : "Populating missing Plex ratingKeys into Movies sheet (match by IMDB/TMDB). Let's go!",
                        2);

                    // Sheets
                    SheetsService sheetsService = CreateSheetsService();

                    string plexBase = GetPlexBaseUrlFromVarsSheet();
                    if (string.IsNullOrWhiteSpace(plexBase))
                    {
                        Type("Plex base URL (Enter for http://192.168.0.5:32400): ", 0, 0, 0);
                        plexBase = Console.ReadLine();
                        if (string.IsNullOrWhiteSpace(plexBase)) plexBase = "http://192.168.0.5:32400";
                    }

                    string plexToken = GetPlexTokenFromVarsSheet();
                    if (string.IsNullOrWhiteSpace(plexToken))
                    {
                        Type("Plex token (paste X-Plex-Token): ", 0, 0, 0);
                        plexToken = Console.ReadLine();
                    }

                    Type("Movies library section id (Enter for 1): ", 0, 0, 0);
                    string secStr = Console.ReadLine();
                    int sectionId = 1;
                    if (!string.IsNullOrWhiteSpace(secStr)) int.TryParse(secStr, out sectionId);

                    PlexRatingKeySheetPopulator.Summary ratingKeySummary = PlexRatingKeySheetPopulator.Run(
                        sheetsService,
                        SPREADSHEET_ID,
                        MOVIES_TITLE_RANGE,
                        MOVIES_DATA_RANGE,
                        (t, msg, ind) => DisplayMessage(t, msg, ind),
                        new PlexRatingKeySheetPopulator.Options
                        {
                            PlexBaseUrl = plexBase,
                            PlexToken = plexToken,
                            MoviesSectionId = sectionId,
                            TmdbIdColumnLetterFallback = "CF",
                            PlexRatingKeyColumnLetterFallback = "CG",
                            ImdbIdColumnLetterFallback = "CN",
                            QuickCreateHeaderName = QUICK_CREATE,
                            QuickCreateValue = "X",
                            OverwriteExistingRatingKeys = overwriteRatingKeys,
                            MarkQuickCreateForWrittenRatingKeys = true
                        });

                    if (ratingKeySummary.QuickCreateRowNumbersMarked.Count > 0)
                    {
                        DisplayMessage("info", "Updating Plex metadata for rows marked Quick Create by ratingKey sync...", 2);
                        PlexMetadataSheetUpdater.Summary metadataSummary = RunPlexMetadataSheetSync(sheetsService, ratingKeySummary.QuickCreateRowNumbersMarked);

                        if (metadataSummary.RowsFailed == 0)
                        {
                            ClearMovieQuickCreateMarks(sheetsService, ratingKeySummary.QuickCreateRowNumbersMarked);
                        }
                        else
                        {
                            DisplayMessage("warning", "Leaving Quick Create marks in place because one or more targeted Plex metadata updates failed.", 2);
                        }
                    }
                    else
                    {
                        DisplayMessage("info", "No ratingKey rows were marked Quick Create, so targeted Plex metadata sync was skipped.", 2);
                    }
                }
                else if (choice.Equals(plexSyncMetadataChoice))
                {
                    Console.WriteLine();
                    DisplayMessage("info", "60 - Syncing Plex labels (Tags/NFO Body <tag>), sort title, and plot from Movies sheet. Let's go!", 2);

                    SheetsService sheetsService = CreateSheetsService();
                    RunPlexMetadataSheetSync(sheetsService, null);
                }
                else if (choice.Equals(rebuildWebMoviesChoice))
                {
                    Console.WriteLine();
                    DisplayMessage("info", "Rebuilding Web Movies from Movies sheet only. Let's go!", 2);

                    SheetsService sheetsService = CreateSheetsService();

                    ProfileRequestProcessor.RebuildWebMoviesOnly(
                        sheetsService,
                        SPREADSHEET_ID,
                        (t, msg, ind) => DisplayMessage(t, msg, ind),
                        new ProfileRequestProcessor.Options
                        {
                            MoviesSheetName = "Movies",
                            MoviesHeaderRow = 2,
                            MoviesDataStartRow = 3,
                            WebMoviesSheetName = "Web Movies",
                            WebMoviesHeaderRow = 1,
                            WebMoviesDataStartRow = 2,
                            WebMoviesStatusValue = "n"
                        });
                }
                else if (choice.Equals(profileRequestProcessorChoice))
                {
                    Console.WriteLine();
                    DisplayMessage("info", "Processing Profile Requests, rebuilding Web Movies, and refreshing media libraries. Let's go!", 2);

                    SheetsService sheetsService = CreateSheetsService();

                    var options = new ProfileRequestProcessor.Options
                    {
                        MoviesSheetName = "Movies",
                        MoviesHeaderRow = 2,
                        MoviesDataStartRow = 3,
                        RequestsSheetName = "Profile Requests",
                        RequestsHeaderRow = 1,
                        RequestsDataStartRow = 2,
                        ArchiveSheetName = "Profile Requests Archive",
                        WebMoviesSheetName = "Web Movies",
                        SyncQuickCreateMovieMetadata = rowNumbers =>
                        {
                            DisplayMessage("info", "Updating Plex metadata for Quick Create Movies rows before selected NFO recreation...", 2);
                            RunPlexMetadataSheetSync(sheetsService, rowNumbers);
                        },
                        RecreateSelectedMovieNfoFiles = RunSelectedMovieNfoFiles
                    };

                    options.MediaServers.Add(new ProfileRequestProcessor.MediaServerOptions
                    {
                        Name = "Emby",
                        BaseUrl = GetEmbyServerUrl(),
                        ApiKey = GetEmbyApiKey(),
                        LibraryItemId = GetEmbyMoviesLibraryId(),
                        ApiPathPrefix = "emby"
                    });

                    string jellyfinUrl = Environment.GetEnvironmentVariable("JELLYFIN_SERVER_URL");
                    string jellyfinApiKey = Environment.GetEnvironmentVariable("JELLYFIN_API_KEY");
                    string jellyfinMoviesLibraryId = Environment.GetEnvironmentVariable("JELLYFIN_MOVIES_LIBRARY_ID");
                    if (!string.IsNullOrWhiteSpace(jellyfinUrl) &&
                        !string.IsNullOrWhiteSpace(jellyfinApiKey) &&
                        !string.IsNullOrWhiteSpace(jellyfinMoviesLibraryId))
                    {
                        options.MediaServers.Add(new ProfileRequestProcessor.MediaServerOptions
                        {
                            Name = "Jellyfin",
                            BaseUrl = jellyfinUrl,
                            ApiKey = jellyfinApiKey,
                            LibraryItemId = jellyfinMoviesLibraryId,
                            ApiPathPrefix = ""
                        });
                    }

                    ProfileRequestProcessor.Run(
                        sheetsService,
                        SPREADSHEET_ID,
                        (t, msg, ind) => DisplayMessage(t, msg, ind),
                        options);
                }
                else if (choice.Equals(fullMovieHealthCheck))
                {
                    Console.WriteLine();
                    DisplayMessage("info", "Check full video health (single file or entire folder). Let's go!", 2);

                    VideoHealthCheck.FullMovieHealthCheck();
                }
                else if (choice.Equals(audioCompatibilityConverterChoice))
                {
                    Console.WriteLine();
                    DisplayMessage("info", "FFmpeg audio compatibility converter. Let's go!", 2);

                    AudioCompatibilityConverter.RunInteractive(
                        (t, msg, ind) => DisplayMessage(t, msg, ind),
                        DriveMapping.GetAllSrtScanLocations());
                }
                else if (choice.Equals(audioCensorRunnerChoice))
                {
                    Console.WriteLine();
                    DisplayMessage("info", "Audio censor subtitle scan. Let's go!", 2);

                    AudioCensorRunner.RunInteractive(
                        (t, msg, ind) => DisplayMessage(t, msg, ind));
                }
                else if (choice.Equals(fetchTvShowPlotsChoice)) // Fetch TV Show episode plots from TVDB.
                {
                    Type("Gather the TV Show episode plots from TVDB. Let's go!", 0, 0, 2);

                    IList<IList<Object>> videoData = CallGetData(combinedEpisodeSheetVariables, COMBINED_EPISODES_TITLE_RANGE, COMBINED_EPISODES_DATA_RANGE);

                    BachFlixNfo.BachFlixNfo.InputTvShowPlots(videoData, combinedEpisodeSheetVariables);

                }
                else if (choice.Equals(fixRecordedNamesChoice)) // Fix recorded names.
                {
                    DisplayMessage("info", "Let's go!", 0, 0, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        IList<IList<Object>> videoData = CallGetData(recordedNamesSheetVariables, RECORDED_NAMES_TITLE_RANGE, RECORDED_NAMES_DATA_RANGE);

                        BachFlixNfo.BachFlixNfo.FixRecordedNames(videoData, recordedNamesSheetVariables, directory);
                    }

                }
                else if (choice.Equals(copyMultipleFilesToOneLocationChoice)) // Copy an array of files to one location.
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
                        }
                    }

                }
                else if (choice.Equals("20")) // Add a comment to a directory.
                {
                    Type("Add a comment to a directory. Let's go!", 0, 0, 1);
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
                else if (choice.Equals(createFoldersAndMoveFilesChoice)) // Create directories that match the names of files in a directory, then move those files into their respective directories.
                {
                    Type("Create directories then move files into directories. Let's go!", 0, 0, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        CreateFoldersAndMoveFiles(directory);
                    }
                }
                else if (choice.Equals(createFoldersAndMoveFilesAndSortChoice))
                {
                    Type("Create directories then move files into directories AND sort them. Let's go!", 0, 0, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        CreateFoldersAndMoveFiles(directory, true);
                    }
                }
                else if (choice.Equals(trimTitlesInDirectoryChoice))
                {
                    Type("Trim the titles in a chosen directory. Let's go!", 0, 0, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        TrimTitlesInDirectory(directory);
                    }
                }
                else if (choice.Equals(bothTrimAndCreateFoldersChoice))
                {
                    Type("Trim the titles AND create directories then move files into directories. Let's go!", 0, 0, 1);
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        TrimTitlesInDirectory(directory);
                        CreateFoldersAndMoveFiles(directory);
                    }
                }
                else if (choice.Equals(insertMissingMovieDataChoice))
                {
                    Type("Insert missing movie data into the Google Sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    DisplayMessage("warning", "Looking through sheet data for missing data... ");
                    InputMovieData(movieData, movieSheetVariables);
                    DisplayMessage("success", "Done", 2);
                    DisplayMessage("warning", "Now looking through sheet data for missing cast... ");
                    InputMovieCredits(movieData, movieSheetVariables);
                    DisplayMessage("success", "Done", 2);
                    DisplayMessage("warning", "Now looking through sheet data for auto populated data... ");
                    CopyAutoPopulatedData(movieData, movieSheetVariables);
                    DisplayMessage("success", "Done", 2);
                    DisplayMessage("warning", "Now looking through sheet data for missing runtimes... ");
                    movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    InputMovieRuntimes(movieData, movieSheetVariables);
                    DisplayMessage("success", "Done", 2);
                }
                else if (choice.Equals(repeatInsertMissingMovieDataChoice))
                {
                    bool repeatDataCall = true;
                    Type("Repeat insert missing movie data into the Google Sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    DisplayMessage("warning", "Looking through sheet data for missing data... ");
                    InputMovieData(movieData, movieSheetVariables);
                    DisplayMessage("success", "Done", 2);
                    DisplayMessage("warning", "Now looking through sheet data for missing cast... ");
                    InputMovieCredits(movieData, movieSheetVariables);
                    DisplayMessage("success", "Done", 2);
                    DisplayMessage("warning", "Now looking through sheet data for auto populated data... ");
                    do
                    {
                        if (CopyAutoPopulatedData(movieData, movieSheetVariables))
                        {
                            DisplayMessage("warning", "There is still data loading.");
                            DisplayMessage("info", "We will try again in 5 minutes.");
                            Countdown(300);
                            movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                            DisplayMessage("info", "Let's try again.");
                        }
                        else
                        {
                            repeatDataCall = false;
                        }
                    } while (repeatDataCall);
                    DisplayMessage("warning", "Now looking through sheet data for missing runtimes... ");
                    movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                    InputMovieRuntimes(movieData, movieSheetVariables);
                    DisplayMessage("success", "Done", 2);

                    DisplayMessage("success", "Done", 2);
                }
                else if (choice.Equals(insertMissingCastMembers))
                {
                    Type("Insert missing movie cast into the Google Sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    DisplayMessage("warning", "Looking through sheet data for missing cast... ");
                    InputMovieCredits(movieData, movieSheetVariables);
                    DisplayMessage("success", "Done");
                }
                else if (choice.Equals(updateMovieDataChoice))
                {
                    Type("Insert missing and update movie data into the Google Sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    DisplayMessage("warning", "Looking through sheet data to update data... ");
                    InputMovieData(movieData, movieSheetVariables, true);
                    DisplayMessage("success", "Done");
                }
                else if (choice.Equals(insertMissingTmdbIdsChoice)) // Insert TMDB IDs into the Google Sheet.
                {
                    Type("Insert missing TMDB IDs into the Google Sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    InputTmdbId(movieData, movieSheetVariables, 1);
                }
                else if (choice.Equals(insertAndOverwriteTmdbIdsChoice)) // Insert TMDB IDs into the Google Sheet.
                {
                    Type("Insert missing AND overwrite TMDB IDs into the Google Sheet. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    InputTmdbId(movieData, movieSheetVariables, 2);
                }
                else if (choice.Equals("12")) // Move kids movies.
                {
                    Type("Move kids movies. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    MoveKidsMovies(movieData, movieSheetVariables);
                }
                else if (choice.Equals(copyMovieFilesToDestinationChoice)) // Copy movies.
                {
                    Type("Copy movies. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    CopyMovieFiles(movieData, movieSheetVariables);
                }
                else if (choice.Equals(deleteMovieFilesAtDestinationChoice)) // Delete movies from moms hard drive.
                {
                    Type("Delete movies. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    DeleteMovieFiles(movieData, movieSheetVariables);
                }
                else if (choice.Equals("16")) // Remove movies from TMDB List.
                {
                    Type("Remove movies from TMDB List. Let's go!", 0, 0, 2);

                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    RemoveMoviesFromTmdbList(movieData, movieSheetVariables);
                }
                else if (choice.Equals("44"))
                {
                    Type("Getting authorization.", 0, 0, 1, "Blue");
                    dynamic tmdbResponse = TmdbApi.AuthenticationCreateRequestToken();

                    var requestToken = tmdbResponse.request_token.ToString();

                    tmdbResponse = TmdbApi.AuthenticationSendRequestToken(requestToken);

                    //Type(tmdbResponse.request_token.ToString(), 0, 0, 1);
                    Type("Authorization received.", 0, 0, 1, "Green");
                }
                else if (choice.Equals(removeMetadataChoice))
                {
                    Type("Remove Metadata. Let's go!", 0, 0, 1);
                    var directory = AskForDirectory();

                    if (directory != "0" && Directory.Exists(directory))
                    {
                        RemoveMetadataInFolder(directory);
                    }
                }
                else if (choice.Equals(insertMissingDbDataChoice))
                {
                    DisplayMessage("info", "Insert missing data into the DB sheet from TVDB. Let's go!");

                    var jwtToken = TvdbApiCall.TvdbApi.GetTvdbJwtKey();

                    IList<IList<Object>> dbData = CallGetData(dbSheetVariables, DB_TITLE_RANGE, DB_DATA_RANGE);

                    InsertMissingDbData(dbData, dbSheetVariables, jwtToken);
                }
                else if (choice.Equals(updateDbSheetChoice))
                {
                    DisplayMessage("info", "Update DB sheet info. Let's go!");

                    var jwtToken = TvdbApiCall.TvdbApi.GetTvdbJwtKey();

                    IList<IList<Object>> dbData = CallGetData(dbSheetVariables, DB_TITLE_RANGE, DB_DATA_RANGE);

                    UpdateDbData(dbData, dbSheetVariables, jwtToken);
                }
                else if (choice.Equals(insertMissingCombinedEpisodesChoice))
                {
                    DisplayMessage("info", "Insert missing data into the Combined Episodes sheet from TVDB. Let's go!");

                    var jwtToken = TvdbApiCall.TvdbApi.GetTvdbJwtKey();

                    IList<IList<Object>> dbData = CallGetData(combinedEpisodeSheetVariables, COMBINED_EPISODES_TITLE_RANGE, COMBINED_EPISODES_DATA_RANGE);

                    InsertMissingCombinedEpisodeData(dbData, combinedEpisodeSheetVariables, jwtToken);
                }
                else if (choice.Equals(insertMissingSeveralCombinedEpisodesChoice))
                {
                    DisplayMessage("info", "Insert missing data into the Several Combined Episodes sheet from TVDB. Let's go!");

                    var jwtToken = TvdbApiCall.TvdbApi.GetTvdbJwtKey();

                    IList<IList<Object>> dbData = CallGetData(severalCombinedSheetVariables, SEVERAL_COMBINED_EPISODES_TITLE_RANGE, SEVERAL_COMBINED_EPISODES_DATA_RANGE);

                    InsertMissingSeveralCombinedEpisodeData(dbData, severalCombinedSheetVariables, jwtToken);
                }
                else if (choice.Equals(insertMissingEpisodeDataChoice))
                {
                    Type("This method is dead...");
                    // DisplayMessage("info", "Insert missing data into the Episodes sheet from TVDB. Let's go!");

                    // IList<IList<Object>> dbData = CallGetData(renameEpisodesSheetVariables, RENAME_EPISODES_TITLE_RANGE, RENAME_EPISODES_DATA_RANGE);

                    // InsertEpisodesIntoRenameEpisodesSheet(dbData, renameEpisodesSheetVariables);
                }
                else if (choice.Equals(updateCombinedEpisodesChoice))
                {
                    DisplayMessage("info", "Update data in the Combined Episodes sheet from TVDB. Let's go!");

                    var jwtToken = TvdbApiCall.TvdbApi.GetTvdbJwtKey();

                    IList<IList<Object>> dbData = CallGetData(combinedEpisodeSheetVariables, COMBINED_EPISODES_TITLE_RANGE, COMBINED_EPISODES_DATA_RANGE);

                    UpdateCombinedEpisodeData(dbData, combinedEpisodeSheetVariables, jwtToken);
                }
                else if (choice.Equals(writeToCombinedEpisodesChoice))
                {

                    DisplayMessage("info", "Add video file names to the Combined Episodes sheet. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        string[] fileEntries = Directory.GetFiles(directory); // Gather ALL files from the directory.

                        ArrayList videoFiles = GrabMovieFiles(fileEntries); // Now filter out anything that isn't a video file.

                        IList<IList<Object>> sheetData = CallGetData(combinedEpisodeSheetVariables, COMBINED_EPISODES_TITLE_RANGE, COMBINED_EPISODES_DATA_RANGE);

                        WriteToSheetColumn(videoFiles, sheetData, "Combined Episodes", Convert.ToInt16(combinedEpisodeSheetVariables[ROW_NUM]), Convert.ToInt16(combinedEpisodeSheetVariables["Combined Episode Name"]));
                    }
                }
                else if (choice.Equals(writeToSeveralCombinedEpisodesChoice))
                {

                    DisplayMessage("info", "Add video file names to the Several-Combined Episodes sheet. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        string[] fileEntries = Directory.GetFiles(directory); // Gather ALL files from the directory.

                        ArrayList videoFiles = GrabMovieFiles(fileEntries); // Now filter out anything that isn't a video file.

                        IList<IList<Object>> sheetData = CallGetData(severalCombinedSheetVariables, SEVERAL_COMBINED_EPISODES_TITLE_RANGE, SEVERAL_COMBINED_EPISODES_DATA_RANGE);

                        WriteToSheetColumn(videoFiles, sheetData, "Several Combined Episodes", Convert.ToInt16(severalCombinedSheetVariables[ROW_NUM]), Convert.ToInt16(severalCombinedSheetVariables["Combined Episode Name"]));
                    }
                }
                else if (choice.Equals(writeVideoFileNamesToYoutubeSheet))
                {
                    DisplayMessage("info", "Add video file names to the YouTube sheet. Let's go!");
                    var directory = AskForDirectory();

                    if (directory != "0")
                    {
                        string[] fileEntries = Directory.GetFiles(directory); // Gather ALL files from the directory.

                        ArrayList videoFiles = GrabMovieFiles(fileEntries); // Now filter out anything that isn't a video file.

                        IList<IList<Object>> sheetData = CallGetData(youtubeSheetVariables, YOUTUBE_TITLE_RANGE, YOUTUBE_DATA_RANGE);

                        WriteToSheetColumn(videoFiles, sheetData, "YouTube", Convert.ToInt16(youtubeSheetVariables[ROW_NUM]), Convert.ToInt16(youtubeSheetVariables[CLEAN_TITLE]));
                    }
                }

                switch (choice.Trim())
                {
                    case "8": // Count files.
                        DisplayMessage("info", "Count the files. Let's go!");
                        CountFiles();
                        break;
                    case "13": // Copy JPG files.
                        CopyJpgFiles();
                        break;
                    //case "14": // Copy movie files.
                    //    CopyMovieFiles();
                    //    break;
                    //case "15": // Mark Owned Movies as D=Done || X=Not Done.
                    //    Type("Mark Owned Movies as D=Done || X=Not Done.", 0, 0, 1);
                    //    CheckForMovie("Main");
                    //    break;
                    //case "21": // testing rewriting to the same console line.
                    //    Type("1", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("2", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("3", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("4", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("5", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("6", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("7", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("8", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("9", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("10", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("11", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("12", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("13", 0, 0, 0);
                    //    Console.SetCursorPosition(0, Console.CursorTop);
                    //    Type("14", 0, 0, 0);
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

        private static bool IsSrtFile(string fileName)
        {
            return string.Equals(Path.GetExtension(fileName), ".srt", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasEngSrtFlag(string fileName)
        {
            string srtName = Path.GetFileNameWithoutExtension(fileName);
            return !string.IsNullOrWhiteSpace(srtName) && SrtEngFlagRegex.IsMatch(srtName);
        }

        private static bool IsForcedSrtFile(string fileName)
        {
            string srtName = Path.GetFileNameWithoutExtension(fileName);
            return !string.IsNullOrWhiteSpace(srtName) && SrtForcedFlagRegex.IsMatch(srtName);
        }

        private static string BuildEngSrtRenameTarget(string fileName)
        {
            if (!IsSrtFile(fileName) || HasEngSrtFlag(fileName))
                return null;

            string directory = Path.GetDirectoryName(fileName);
            string srtName = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);

            if (string.IsNullOrWhiteSpace(srtName))
                return null;

            string newSrtName = SrtLanguageTokenRegex.Replace(srtName, "${sep}eng", 1);

            if (string.Equals(srtName, newSrtName, StringComparison.OrdinalIgnoreCase) && IsForcedSrtFile(fileName))
            {
                newSrtName = SrtForcedFlagRegex.Replace(
                    srtName,
                    match => ".eng" + match.Groups[1].Value + "forced" + match.Groups[2].Value,
                    1);
            }

            if (string.Equals(srtName, newSrtName, StringComparison.OrdinalIgnoreCase))
                newSrtName = srtName + ".eng";

            return Path.Combine(directory ?? "", newSrtName + extension);
        }

        public static void displayMissingEngResults()
        {
            Console.WriteLine();
            DisplayMessage("info", "That's the end of it.");
            DisplayMessage("data", "Of the: ", 0);
            DisplayMessage("log", fileEntriesCount.ToString(), 0);
            DisplayMessage("data", " file entries.");
            DisplayMessage("log", srtEntriesCount.ToString(), 0);
            DisplayMessage("data", " of them were SRT files");
            DisplayMessage("log", missingEng.Count().ToString(), 0);
            DisplayMessage("data", " of the SRT files were missing the ENG flag.", 2);

            if (missingEng.Count() > 0)
            {
                DisplayMessage("log", "Press 1 to view the list, press 0 to exit");
                var keepAsking = true;

                do
                {
                    var response = Console.ReadLine();

                    if (response == "0")
                    {
                        keepAsking = false;
                    }
                    else if (response == "1")
                    {
                        foreach (var missingEngDirectory in missingEng)
                        {
                            DisplayMessage("warning", missingEngDirectory);
                        }

                        Console.WriteLine();
                        DisplayMessage("info", "Would you like to fix these names? 1=Yes or 2=No");

                        var fixResponse = Console.ReadLine();

                        if (fixResponse == "1")
                        {
                            foreach (var fixDirectory in missingEng)
                            {
                                if (fixDirectory.ToUpper().Contains("-TRAILER"))
                                {
                                    File.Delete(fixDirectory);
                                    DisplayMessage("warning", "We detected a trailer subtitle file and deleted it:");
                                    DisplayMessage("info", fixDirectory.ToString(), 2);
                                    continue;
                                }
                                var newName = BuildEngSrtRenameTarget(fixDirectory);

                                if (string.IsNullOrWhiteSpace(newName) ||
                                    string.Equals(fixDirectory, newName, StringComparison.OrdinalIgnoreCase))
                                {
                                    DisplayMessage("warning", "Skipping SRT because it does not need an ENG rename:");
                                    DisplayMessage("info", fixDirectory.ToString(), 2);
                                    continue;
                                }

                                if (File.Exists(newName))
                                {
                                    DisplayMessage("error", "There appears to already be a file named:");
                                    DisplayMessage("harderror", newName);
                                    DisplayMessage("error", "Please navigate to the containing folder and decide which file you want to keep.", 2);
                                }
                                else
                                {
                                    File.Move(fixDirectory, newName);
                                    Type(fixDirectory, 0, 0, 0);
                                    Type(" renamed to:", 0, 0, 1, "Green");
                                    Type(newName, 0, 0, 2);
                                }
                            }
                        }
                        keepAsking = false;
                    }
                    else
                    {
                        Type("That wasn't an option, try again.", 0, 0, 1);
                    }
                } while (keepAsking);
            }
            // Clear the variables
            missingEng.Clear();
            fileEntriesCount = 0;
            srtEntriesCount = 0;
            missingEngCount = 0;
        }

        private static void WriteJsonLog(string path, RenameTransactionLog log)
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(dir))
                Directory.CreateDirectory(dir);

            var serializer = new DataContractJsonSerializer(typeof(RenameTransactionLog));
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                serializer.WriteObject(fs, log);
            }
        }

        private static RenameTransactionLog ReadJsonLog(string path)
        {
            var serializer = new DataContractJsonSerializer(typeof(RenameTransactionLog));
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                return (RenameTransactionLog)serializer.ReadObject(fs);
            }
        }

        // Drag-drop paths often come in wrapped in quotes
        private static string CleanPath(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return input;
            input = input.Trim();
            if (input.StartsWith("\"") && input.EndsWith("\"") && input.Length >= 2)
                input = input.Substring(1, input.Length - 2);
            return input;
        }

        private static void NormalizeSeasonEpisodesSmart_WithLog(string directory, bool dryRun = true)
        {
            if (!Directory.Exists(directory))
            {
                Console.WriteLine($"[ERR] Directory not found: {directory}");
                return;
            }

            var epRegex = new Regex(
                @"S(?<season>\d{2})E(?<ep>\d{2})(?:-E(?<ep2>\d{2}))?",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

            var mkvs = Directory.GetFiles(directory, "*.mkv", SearchOption.TopDirectoryOnly);

            var entries = new List<(string MkvPath, int Season, int Ep, bool IsCombined)>();

            foreach (var mkv in mkvs)
            {
                var name = Path.GetFileNameWithoutExtension(mkv);
                var m = epRegex.Match(name);
                if (!m.Success) continue;

                int season = int.Parse(m.Groups["season"].Value);
                int ep = int.Parse(m.Groups["ep"].Value);

                // Combined heuristic: underscore OR already has -E##
                bool isCombined = name.Contains("_") || m.Groups["ep2"].Success;

                entries.Add((mkv, season, ep, isCombined));
            }

            if (entries.Count == 0)
            {
                Console.WriteLine("[WARN] No MKVs with SxxEyy pattern found.");
                return;
            }

            entries = entries
                .OrderBy(e => e.Season)
                .ThenBy(e => e.Ep)
                .ToList();

            int counter = entries.Min(e => e.Ep);

            Console.WriteLine($"[INFO] Starting episode counter at: E{counter:D2}");
            Console.WriteLine(dryRun ? "[INFO] DRY RUN (no files will be moved)\n" : "[INFO] LIVE RUN\n");

            // --- Create the log object now (even dry run can log "planned moves" if you want)
            var log = new RenameTransactionLog
            {
                createdUtc = DateTime.UtcNow.ToString("o"),
                mode = "NormalizeSeasonEpisodesSmart",
                sourceDirectory = directory,
                dryRun = dryRun
            };

            foreach (var entry in entries)
            {
                string mkvPath = entry.MkvPath;
                string baseName = Path.GetFileNameWithoutExtension(mkvPath);

                string newToken = entry.IsCombined
                    ? $"S{entry.Season:D2}E{counter:D2}-E{counter + 1:D2}"
                    : $"S{entry.Season:D2}E{counter:D2}";

                // Rename all sidecars: base.* (mkv, srt, nfo, etc)
                string[] sidecars = Directory.GetFiles(directory, baseName + ".*", SearchOption.TopDirectoryOnly);

                foreach (var path in sidecars)
                {
                    string fileNameOnly = Path.GetFileName(path);

                    if (!epRegex.IsMatch(fileNameOnly))
                        continue;

                    string newFileNameOnly = epRegex.Replace(fileNameOnly, newToken, 1);
                    if (newFileNameOnly.Equals(fileNameOnly, StringComparison.OrdinalIgnoreCase))
                        continue;

                    string newFullPath = Path.Combine(directory, newFileNameOnly);

                    Console.WriteLine(entry.IsCombined ? "[COMBINE]" : "[SINGLE]");
                    Console.WriteLine($"  {fileNameOnly}");
                    Console.WriteLine($"  -> {newFileNameOnly}\n");

                    // Record the move in the log (even in dry run; you can change this if you only want real moves)
                    log.moves.Add(new RenameMove { from = path, to = newFullPath });

                    if (!dryRun)
                    {
                        if (File.Exists(newFullPath))
                        {
                            Console.WriteLine($"[ERR] Target already exists, skipping: {newFullPath}\n");
                            continue;
                        }

                        File.Move(path, newFullPath);
                    }
                }

                counter += entry.IsCombined ? 2 : 1;
            }

            // --- Write log to your folder
            string logDir = @"C:\Users\Brandon\Desktop\BachFlixNfo Logs\Combine Episodes";
            Directory.CreateDirectory(logDir);

            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            ExtractShowAndSeasonFromPath(directory, out string showName, out string seasonLabel);

            string modeTag = dryRun ? "DRYRUN" : "LIVE";
            string logFileName = $"{showName} - {seasonLabel} - {modeTag} - {stamp}.json";

            string logPath = Path.Combine(logDir, logFileName);
            WriteJsonLog(logPath, log);

            Console.WriteLine($"[OK] Log written: {logPath}");
        }

        private static void ExtractShowAndSeasonFromPath(
            string directory,
            out string showName,
            out string seasonLabel
        )
        {
            showName = "Unknown Show";
            seasonLabel = "Season ??";

            if (string.IsNullOrWhiteSpace(directory))
                return;

            try
            {
                var dirInfo = new DirectoryInfo(directory);

                // Last folder should be "Season 01"
                if (dirInfo.Name.StartsWith("Season", StringComparison.OrdinalIgnoreCase))
                {
                    seasonLabel = dirInfo.Name;

                    // Parent folder should be the show name
                    if (dirInfo.Parent != null)
                    {
                        showName = dirInfo.Parent.Name;
                    }
                }
            }
            catch
            {
                // Fail silently and keep defaults
            }

            showName = SanitizeForFileName(showName);
            seasonLabel = SanitizeForFileName(seasonLabel);
        }


        private static string SanitizeForFileName(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "Unknown";

            foreach (char c in Path.GetInvalidFileNameChars())
                input = input.Replace(c, ' ');

            // collapse double spaces
            input = Regex.Replace(input, @"\s+", " ").Trim();

            return input;
        }

        private static void RevertEpisodeCombineFromLog()
        {
            Console.WriteLine("Drag/drop the JSON log file to revert, then press Enter:");
            string input = Console.ReadLine();
            string logPath = CleanPath(input);

            if (string.IsNullOrWhiteSpace(logPath) || !File.Exists(logPath))
            {
                Console.WriteLine("[ERR] Log file not found.");
                return;
            }

            RenameTransactionLog log;
            try
            {
                log = ReadJsonLog(logPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("[ERR] Failed to read log JSON: " + ex.Message);
                return;
            }

            if (log?.moves == null || log.moves.Count == 0)
            {
                Console.WriteLine("[WARN] Log contains no moves.");
                return;
            }

            Console.WriteLine($"[INFO] Reverting {log.moves.Count} moves from log.");
            Console.WriteLine($"[INFO] Source directory (for reference): {log.sourceDirectory}");
            Console.WriteLine();

            // Revert in reverse order to reduce collisions
            foreach (var move in log.moves.AsEnumerable().Reverse())
            {
                string from = move.from; // original
                string to = move.to;     // renamed

                // Revert means: move "to" back to "from"
                if (!File.Exists(to))
                {
                    Console.WriteLine($"[SKIP] Missing (already reverted?): {to}");
                    continue;
                }

                // Avoid overwriting if something is already there
                if (File.Exists(from))
                {
                    Console.WriteLine($"[ERR] Cannot revert, target exists: {from}");
                    continue;
                }

                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(from));
                    File.Move(to, from);

                    Console.WriteLine("[OK] Reverted:");
                    Console.WriteLine("  " + Path.GetFileName(to));
                    Console.WriteLine("  -> " + Path.GetFileName(from));
                    Console.WriteLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ERR] Failed to revert {to} -> {from}: {ex.Message}");
                }
            }

            Console.WriteLine("[DONE] Revert attempt complete.");
        }

        public static void CreateMissingCombinedEpisodeNfoFiles(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string directory)
        {
            try
            {
                string[] fileEntries = Directory.GetFiles(directory);
                var fileCount = fileEntries.Length;
                string plural = fileCount == 1 ? " file " : " files ";
                DisplayMessage("warning", "Searching directory for Combined Episodes... ", 0);
                foreach (var row in data)
                {
                    var tvdbId = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                    var oldName = row[Convert.ToInt16(sheetVariables["Combined Episode Name"])].ToString();
                    var newName = row[Convert.ToInt16(sheetVariables["New Episode Name"])].ToString();
                    var nfoBody = row[Convert.ToInt16(sheetVariables[NFO_BODY])].ToString();

                    if (!tvdbId.Equals("")) // If the ID is empty then just skip the rest.
                    {
                        var sourceFiles = Directory.GetFiles(directory, oldName + ".*");
                        if (sourceFiles.Length > 0)
                        {
                            // Loop through all sourceFiles to rename. (i.e. srt, mp4, and jpg files)
                            for (int i = 0; i < sourceFiles.Length; i++)
                            {
                                string srtVariables = "";
                                if (Path.GetExtension(sourceFiles[i]) == ".srt")
                                {
                                    srtVariables += ".eng";
                                    if (sourceFiles[i].ToLower().Contains(".forced."))
                                    {
                                        srtVariables += ".forced";
                                    }
                                }
                                var destinationFile = Path.Combine(directory, newName + srtVariables) + Path.GetExtension(sourceFiles[i]);
                                File.Move(sourceFiles[i], destinationFile);
                            }
                            // Write the NFO file once.
                            var nfoFile = Path.Combine(directory, newName) + ".nfo";
                            WriteNfoFile(nfoFile, nfoBody);
                        }
                    }
                }
                DisplayMessage("success", "DONE");
            }
            catch (Exception e)
            {
                DisplayMessage("error", "An error occured in CreateMissingCombinedEpisodeNfoFiles method:");
                DisplayMessage("harderror", e.Message);
            }
        } // End CreateMissingCombinedEpisodeNfoFiles()

        public static void CreateMissingSeveralCombinedEpisodeNfoFiles(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string directory)
        {
            try
            {
                string[] fileEntries = Directory.GetFiles(directory);
                var fileCount = fileEntries.Length;
                string plural = fileCount == 1 ? " file " : " files ";
                var count = 1;
                DisplayMessage("warning", "Searching directory for Several-Combined Episodes...");
                DisplayMessage("info", "Found " + fileCount + plural);
                foreach (var row in data)
                {
                    var tvdbId = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                    var oldName = row[Convert.ToInt16(sheetVariables["Combined Episode Name"])].ToString();
                    var newName = row[Convert.ToInt16(sheetVariables["New Full Episode Name"])].ToString();
                    var nfoBody = row[Convert.ToInt16(sheetVariables[NFO_BODY])].ToString();

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
                DisplayMessage("error", "An error occured in CreateMissingSeveralCombinedEpisodeNfoFiles method:");
                DisplayMessage("harderror", e.Message);
            }
        } // End CreateMissingSeveralCombinedEpisodeNfoFiles()

        public static void CreateMissingEpisodeNfoFiles(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string directory)
        {
            try
            {
                string[] fileEntries = Directory.GetFiles(directory);
                var fileCount = fileEntries.Length;
                string plural = fileCount == 1 ? " file " : " files ";
                var count = 1;
                DisplayMessage("warning", "Searching directory for Episodes missing info files...");
                DisplayMessage("info", "Found " + fileCount + " " + plural);
                foreach (var row in data)
                {
                    var tvdbId = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                    var episodeName = row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString();
                    var nfoBody = row[Convert.ToInt16(sheetVariables[NFO_BODY])].ToString();

                    if (!tvdbId.Equals("")) // If the ID is empty then just skip the rest.
                    {
                        var sourceFile = Directory.GetFiles(directory, episodeName + ".*");
                        if (sourceFile.Length > 0)
                        {
                            var destinationFile = Path.Combine(directory, episodeName) + Path.GetExtension(sourceFile[0]);
                            var nfoFile = Path.Combine(directory, episodeName) + ".nfo";
                            File.Move(sourceFile[0], destinationFile);
                            WriteNfoFile(nfoFile, nfoBody);
                            DisplayMessage("default", "(" + count + " of " + fileCount + ")" + " File renamed and NFO file created for: ", 0);
                            DisplayMessage("success", episodeName);
                            count++;
                        }
                    }
                }
                DisplayMessage("success", "DONE");
            }
            catch (Exception e)
            {
                DisplayMessage("error", "An error occured in CreateMissingEpisodeNfoFiles method:");
                DisplayMessage("harderror", e.Message);
            }
        } // End CreateMissingEpisodeNfoFiles()

        private static void InsertMissingDbData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string token)
        {
            DisplayMessage("info", "Now inserting info into DB sheet.");

            int intSeriesNamesInsertedCount = 0,
                intSeasonCountsInsertedCount = 0,
                intContinuingsInsertedCount = 0,
                intContentRatingInsertedCount = 0,
                intYearValueInsertedCount = 0,
                intTvdbIdsInsertedCount = 0;

            string tvdbIdValue = "", // Our current TVDB ID value from the Google Sheet.
                seriesNameValue = "", // Our current TVDB Series Name from the Google Sheet.
                seasonCountValue = "", // Our current Season Count value from the Google Sheet.
                continuingValue = "", // Our current Continuing? value from the Google Sheet.
                currentYearValue = "", // The current year value in the Google Sheet.
                contentRatingValue = "",
                tvdbSlugValue = "", // Our current TVDB Slug value from the Google Sheet.
                rowNum = "", // Holds the row number we are on.
                strCellToPutData = ""; // The string of the location to write the data to.

            int tvdbIdColumnNum = 0, // Used to input the returned ID back into the Google Sheet.
                seriesNameColumnNum = 0, // Used to input the returned rating into the Google Sheet.
                seasonCountColumnNum = 0, // Used to input the returned rating into the Google Sheet.
                currentYearColumnNum = 0,
                contentRatingColumnNum = 0,
                continuingColumnNum = 0; // Used to input the returned overview into the Google Sheet.

            foreach (var row in data)
            {
                if (row.Count > 14)
                {
                    rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                    tvdbSlugValue = row[Convert.ToInt16(sheetVariables["TVDB Slug"])].ToString();
                    tvdbIdValue = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                    tvdbIdColumnNum = Convert.ToInt16(sheetVariables["TVDB ID"]);
                    seriesNameValue = row[Convert.ToInt16(sheetVariables["Series Name"])].ToString();
                    seriesNameColumnNum = Convert.ToInt16(sheetVariables["Series Name"]);
                    seasonCountValue = row[Convert.ToInt16(sheetVariables["Season Count"])].ToString();
                    seasonCountColumnNum = Convert.ToInt16(sheetVariables["Season Count"]);
                    contentRatingValue = row[Convert.ToInt16(sheetVariables["Content Rating"])].ToString();
                    contentRatingColumnNum = Convert.ToInt16(sheetVariables["Content Rating"]);
                    continuingValue = row[Convert.ToInt16(sheetVariables["Continuing?"])].ToString();
                    continuingColumnNum = Convert.ToInt16(sheetVariables["Continuing?"]);
                    currentYearValue = row[Convert.ToInt16(sheetVariables["Year"])].ToString();
                    currentYearColumnNum = Convert.ToInt16(sheetVariables["Year"]);

                    if (!tvdbSlugValue.Equals("")) // If there is no slug then the row is considered empty and should be skipped.
                    {
                        // First check to see if the id is empty and populate it if it is.
                        if (tvdbIdValue.Equals(""))
                        {
                            try
                            {
                                tvdbIdValue = TvdbApiCall.TvdbApi.GetSeriesIdAsync(ref token, tvdbSlugValue);
                                strCellToPutData = "DB!" + ColumnNumToLetter(tvdbIdColumnNum) + rowNum;

                                WriteSingleCellToSheet(tvdbIdValue, strCellToPutData);

                                DisplayMessage("default", "ID saved for ", 0, 10);
                                DisplayMessage("success", tvdbSlugValue, 1, 5, 10);
                                intTvdbIdsInsertedCount++;
                            }
                            catch (Exception err)
                            {
                                Type("Unable to find '" + tvdbSlugValue + "' on TVDB", 0, 0, 2, "Red");
                                continue;
                            }
                        }
                        if (seriesNameValue.Equals("") || seasonCountValue.Equals("") || continuingValue.Equals("") || contentRatingValue.Equals("") || currentYearValue.Equals(""))
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
                            if (currentYearValue.Equals(""))
                            {
                                string airedDate = response.data.firstAired;

                                int year = DateTime.Parse(airedDate).Year;

                                strCellToPutData = "DB!" + ColumnNumToLetter(currentYearColumnNum) + rowNum;

                                WriteSingleCellToSheet(year.ToString(), strCellToPutData);
                                DisplayMessage("default", "Year Value saved for ", 0);
                                DisplayMessage("success", seriesNameValue);
                                intYearValueInsertedCount++;
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
                            if (contentRatingValue.Equals(""))
                            {
                                contentRatingValue = response.data.rating.ToString();
                                if (!contentRatingValue.Equals(""))
                                {
                                    strCellToPutData = "DB!" + ColumnNumToLetter(contentRatingColumnNum) + rowNum;

                                    WriteSingleCellToSheet(contentRatingValue, strCellToPutData);
                                    DisplayMessage("default", "Content Rating saved for ", 0);
                                    DisplayMessage("success", seriesNameValue);
                                    intContentRatingInsertedCount++;
                                }
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
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("New TVDB IDs inserted: " + intTvdbIdsInsertedCount, 0, 0, 1, "Green");
            Type("New Series Names inserted: " + intSeriesNamesInsertedCount, 0, 0, 1, "Green");
            Type("New Season Counts inserted: " + intSeasonCountsInsertedCount, 0, 0, 1, "Green");
            Type("New Continuings inserted: " + intContinuingsInsertedCount, 0, 0, 1, "Green");
            Type("New Content Ratings inserted: " + intContentRatingInsertedCount, 0, 0, 1, "Green");
            Type("New Year Values inserted: " + intYearValueInsertedCount, 0, 0, 1, "Green");

        } // End InsertMissingDbData()

        private static void UpdateDbData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string token)
        {
            int intSeasonCountsUpdatedCount = 0,
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

            int tvdbIdColumnNum = 0,
                seasonCountColumnNum = 0,
                continuingColumnNum = 0;

            foreach (var row in data)
            {
                try
                {
                    rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                    tvdbSlugValue = row[Convert.ToInt16(sheetVariables["TVDB Slug"])].ToString();
                    tvdbIdValue = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                    tvdbIdColumnNum = Convert.ToInt16(sheetVariables["TVDB ID"]);
                    seriesNameValue = row[Convert.ToInt16(sheetVariables["Series Name"])].ToString();
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
                        }
                    }
                }
                catch (Exception e)
                {
                    Type(e.Message, 0, 0, 2, "DarkRed");
                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.");
            Type("New TVDB IDs inserted: " + intTvdbIdsInsertedCount, 0, 0, 1, "Green");
            Type("New Season Counts updated: " + intSeasonCountsUpdatedCount, 0, 0, 1, "Green");
            Type("New Continuings updated: " + intContinuingsUpdatedCount, 0, 0, 1, "Green");

        } // End UpdateDbData()

        private static void UpdateEpisodeNumbersInChosenDirectory(int episodeNumber = 1)
        {
            var directory = AskForDirectory("Give me the directory full of TV show episodes.");
            if (directory != "0")
            {
                // Grab all files in the directory.
                Type("Grabbing all MKV files... ", 0, 0, 0, "Yellow");
                string[] fileEntries = Directory.GetFiles(directory, "*.mkv");
                Type("DONE", 0, 0, 2, "Green");

                if (fileEntries.Count() == 0)
                {
                    Type("No videos found! Are they the right format?", 0, 0, 2, "Yellow");
                }

                string pattern = @"S(\d{2})E(\d{2,3})";

                foreach (var originalFile in fileEntries)
                {
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalFile);

                    string searchString = $"*{fileNameWithoutExtension}*";

                    string[] filteredEntries = Directory.GetFiles(directory, searchString);

                    foreach (var filteredFile in filteredEntries)
                    {
                        // Use regex to check if the file name matches the expected pattern
                        Match match = Regex.Match(filteredFile, pattern);

                        if (match.Success)
                        {
                            string season = match.Groups[1].Value; // Preserve original season

                            string replacement = $"S{season}E{episodeNumber:D2}";

                            // Combine the folder path with the new file name
                            string newFile = Regex.Replace(filteredFile, pattern, replacement);

                            // Rename the file
                            File.Move(filteredFile, newFile);
                            Console.WriteLine($"Renamed:");
                            Console.WriteLine(Path.GetFileName(filteredFile));
                            Console.WriteLine(Path.GetFileName(newFile));
                            Console.WriteLine();
                        }
                    }
                    // Increment the starting episode by 2 for the next file
                    episodeNumber++;
                }

                Type("DONE", 0, 0, 1, "Green");
            }
        }

        private static void InsertTvShowDirectorySizesIntoDbSheet(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, bool overwriteData)
        {
            string message = overwriteData ? "Overwriting TV Show directory sizes..." : "Inserting missing TV Show directory sizes...";

            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            DisplayMessage("warning", message);
            string[] parentFolders = { "TV Shows", @"The Bus\TV Shows" };
            string[] networkFolders = { "#-B", "C-D", "E-I", "J-M", "N-R", "S", "T-Z", "Other" };
            string cleanName = "";

            foreach (var row in data)
            {
                try
                {
                    if (row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString() != "") // If there is no id then the row is considered empty and should be skipped.
                    {
                        cleanName = row[Convert.ToInt16(sheetVariables["Clean Name"])].ToString();

                        string rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString(),
                            cleanNameWithYear = row[Convert.ToInt16(sheetVariables["Clean Name with Year"])].ToString(),
                            hardDriveLetter = row[Convert.ToInt16(sheetVariables["Hard Drive Letter"])].ToString(),
                            size = row[Convert.ToInt16(sheetVariables["Size"])].ToString(),
                            foundLocations = row[Convert.ToInt16(sheetVariables["Found Locations"])].ToString();

                        int sizeColumnNum = Convert.ToInt16(sheetVariables["Size"]),
                            foundLocationsColumnNum = Convert.ToInt16(sheetVariables["Found Locations"]);

                        long folderSize = 0;

                        List<string> foundLocationsList = new List<string>();

                        if (size.Equals("") || overwriteData)
                        {
                            bool directoryFound = false;
                            foreach (var parentFolder in parentFolders)
                            {
                                string pathWithoutYear = Path.Combine(hardDriveLetter, parentFolder, cleanName);
                                string pathWithYear = Path.Combine(hardDriveLetter, parentFolder, cleanNameWithYear);

                                if (Directory.Exists(pathWithoutYear))
                                {
                                    directoryFound = true;
                                    foundLocationsList.Add(pathWithoutYear);
                                    folderSize += CalculateFolderSize(pathWithoutYear);
                                }

                                if (Directory.Exists(pathWithYear))
                                {
                                    directoryFound = true;
                                    foundLocationsList.Add(pathWithYear);
                                    folderSize += CalculateFolderSize(pathWithYear);
                                }
                            }

                            foreach (var networkFolder in networkFolders)
                            {
                                foreach (var parentFolder in parentFolders)
                                {
                                    string pathWithoutYear = Path.Combine("\\\\QUADPLEX", networkFolder, parentFolder, cleanName);
                                    string pathWithYear = Path.Combine("\\\\QUADPLEX", networkFolder, parentFolder, cleanNameWithYear);

                                    if (Directory.Exists(pathWithoutYear))
                                    {
                                        directoryFound = true;
                                        foundLocationsList.Add(pathWithoutYear);
                                        folderSize += CalculateFolderSize(pathWithoutYear);
                                    }

                                    if (Directory.Exists(pathWithYear))
                                    {
                                        directoryFound = true;
                                        foundLocationsList.Add(pathWithYear);
                                        folderSize += CalculateFolderSize(pathWithYear);
                                    }

                                }
                            }

                            if (directoryFound)
                            {
                                string formattedSize = ConvertBytesToGBytes(folderSize);

                                if (formattedSize != size)
                                {
                                    try
                                    {
                                        string strCellToPutData = "DB!" + ColumnNumToLetter(sizeColumnNum) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { formattedSize } }
                                        });

                                        if (overwriteData)
                                        {
                                            Type("Successfully overwrote the previous size from: ", 0, 0, 0, "Green");
                                            Type(size, 0, 0, 0, "Yellow");
                                            Type(" to ", 0, 0, 0, "Green");
                                            Type(formattedSize, 0, 0, 0, "Blue");
                                            Type(" at: ", 0, 0, 0, "Green");
                                        }
                                        else
                                        {
                                            Type(formattedSize, 0, 0, 0, "Blue");
                                            Type(" Successfully saved at: ", 0, 0, 0, "Green");
                                        }
                                        Type(strCellToPutData, 0, 0, 0, "Blue");
                                        Type(" For: ", 0, 0, 0, "Green");
                                        Type(cleanNameWithYear, 0, 0, 1, "Blue");
                                    }
                                    catch (Exception ex)
                                    {
                                        DisplayMessage("error", "An error occured saving the folder size for: ", 0);
                                        DisplayMessage("warning", cleanNameWithYear);
                                        DisplayMessage("harderror", ex.Message);
                                    }

                                }
                                else
                                {
                                    DisplayMessage("info", "Size of TV Show directory is same in Google sheet for: ", 0);
                                    DisplayMessage("success", cleanNameWithYear, 0);
                                    DisplayMessage("info", " | Nothing to update");
                                }

                                if (foundLocationsList[0].ToString() != foundLocations)
                                {
                                    string strCellToPutData = "DB!" + ColumnNumToLetter(foundLocationsColumnNum) + rowNum;

                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { string.Join(",", foundLocationsList) } }
                                    });
                                }
                            }
                            else
                            {
                                if (!size.Equals(""))
                                {
                                    string strCellToPutData = "DB!" + ColumnNumToLetter(sizeColumnNum) + rowNum;
                                    WriteSingleCellToSheet("", strCellToPutData);

                                    strCellToPutData = "DB!" + ColumnNumToLetter(foundLocationsColumnNum) + rowNum;
                                    WriteSingleCellToSheet("", strCellToPutData);
                                }

                                Type("Directory not found for: ", 0, 0, 0, "Red");
                                Type(cleanNameWithYear, 0, 0, 1, "Yellow");
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Type("Error while getting directory size for:", 0, 0, 1, "Red");
                    Type(cleanName, 0, 0, 1, "Yellow");
                    Type(e.Message, 0, 0, 2, "DarkRed");
                }
            }
            var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
            DisplayMessage("info", "Looks like that's it.", 2);
        }

        private static void InsertMovieDirectorySizesIntoMoviesSheet(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, bool overwriteData)
        {
            string message = overwriteData ? "Overwriting Movie directory sizes..." : "Inserting missing Movie directory sizes...";

            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            DisplayMessage("warning", message);
            string imdbTitle = "",
                rowNum = "",
                status = "",
                size = "",
                directory = "";

            foreach (var row in data)
            {
                try
                {
                    if (row.Count > 6) // If there is no url then the row is considered empty and should be skipped.
                    {
                        imdbTitle = row[Convert.ToInt16(sheetVariables["IMDB Title"])].ToString();
                        rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                        status = row[Convert.ToInt16(sheetVariables["Status"])].ToString();
                        size = row[Convert.ToInt16(sheetVariables["Size"])].ToString();
                        directory = row[Convert.ToInt16(sheetVariables["Directory"])].ToString();

                        int sizeColumnNum = Convert.ToInt16(sheetVariables["Size"]);

                        long folderSize = 0;

                        if (status.Equals("n") && (size.Equals("") || overwriteData))
                        {
                            if (Directory.Exists(directory))
                            {
                                folderSize += CalculateFolderSize(directory);

                                string formattedSize = ConvertBytesToGBytes(folderSize);

                                if (formattedSize != size)
                                {
                                    try
                                    {
                                        string strCellToPutData = "Movies!" + ColumnNumToLetter(sizeColumnNum) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { formattedSize } }
                                        });

                                        if (overwriteData)
                                        {
                                            Type("Successfully overwrote the previous size from: ", 0, 0, 0, "Green");
                                            Type(size, 0, 0, 0, "Yellow");
                                            Type(" to ", 0, 0, 0, "Green");
                                            Type(formattedSize, 0, 0, 0, "Blue");
                                            Type(" at: ", 0, 0, 0, "Green");
                                        }
                                        else
                                        {
                                            Type(formattedSize, 0, 0, 0, "Blue");
                                            Type(" Successfully saved at: ", 0, 0, 0, "Green");
                                        }
                                        Type(strCellToPutData, 0, 0, 0, "Blue");
                                        Type(" For: ", 0, 0, 0, "Green");
                                        Type(imdbTitle, 0, 0, 1, "Blue");
                                    }
                                    catch (Exception ex)
                                    {
                                        DisplayMessage("error", "An error occured saving the folder size for: ", 0);
                                        DisplayMessage("warning", imdbTitle);
                                        DisplayMessage("harderror", ex.Message);
                                    }



                                }
                                else
                                {
                                    DisplayMessage("info", "Size of Movie directory is same in Google sheet for: ", 0);
                                    DisplayMessage("success", imdbTitle, 0);
                                    DisplayMessage("info", " | Nothing to update");
                                }
                            }
                            else
                            {
                                Type("Directory not found for: ", 0, 0, 0, "Red");
                                Type(imdbTitle, 0, 0, 1, "Yellow");
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Type("Error while getting directory size for:", 0, 0, 1, "Red");
                    Type(imdbTitle, 0, 0, 1, "Yellow");
                    Type(e.Message, 0, 0, 2, "DarkRed");
                }
            }
            var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
            DisplayMessage("info", "Looks like that's it.", 2);
        }

        protected static void FillInFileTypes(IList<IList<object>> data, Dictionary<string, int> sheetVariables, bool overwrite)
        {
            string message = overwrite
                ? "Overwriting File Types in Movies sheet..."
                : "Adding missing File Types to Movies sheet...";

            DisplayMessage("warning", message);

            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>()
            };

            int rowNumCol = sheetVariables[ROW_NUM];
            int statusCol = sheetVariables["Status"];
            int directoryCol = sheetVariables["Directory"];
            int fileTypeCol = sheetVariables["File Type"];
            int imdbTitleCol = sheetVariables["IMDB Title"];

            foreach (var row in data)
            {
                try
                {
                    // Make sure row is long enough to contain File Type column
                    if (row.Count <= fileTypeCol)
                        continue;

                    string rowNum = row[rowNumCol].ToString();
                    string status = row[statusCol].ToString();
                    string directory = row[directoryCol].ToString();
                    string currentType = row[fileTypeCol].ToString();
                    string imdbTitle = row[imdbTitleCol].ToString();

                    // Only touch movies with Status == "n"
                    if (!status.Equals("n", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Directory must exist in sheet
                    if (string.IsNullOrWhiteSpace(directory))
                        continue;

                    // Non-overwrite mode → only fill blanks
                    if (!overwrite && !string.IsNullOrWhiteSpace(currentType))
                        continue;

                    // Let helper decide: extension / "0"
                    string newType = DetectFileTypeFromDirectory(directory);

                    if (string.IsNullOrEmpty(newType))
                        newType = "0"; // ultra-safe fallback

                    // Skip if nothing changed
                    if (currentType == newType)
                    {
                        DisplayMessage("info", $"File Type '{currentType}' already correct for: {imdbTitle}");
                        continue;
                    }

                    string cellRange =
                        "Movies!" + ColumnNumToLetter(fileTypeCol) + rowNum;

                    batchRequest.Data.Add(new ValueRange
                    {
                        Range = cellRange,
                        MajorDimension = "ROWS",
                        Values = new List<IList<object>> { new List<object> { newType } }
                    });

                    // Logging
                    DisplayMessage("success", "File Type ", 0);
                    DisplayMessage("info", newType, 0);

                    if (overwrite && !string.IsNullOrWhiteSpace(currentType))
                    {
                        DisplayMessage("success", " overwritten from: ", 0);
                        DisplayMessage("info", currentType, 0);
                        DisplayMessage("success", " for: ", 0);
                    }
                    else
                    {
                        DisplayMessage("success", " saved for: ", 0);
                    }

                    DisplayMessage("info", imdbTitle, 0);
                    DisplayMessage("log", " at- ", 0);
                    DisplayMessage("info", cellRange);
                }
                catch (Exception e)
                {
                    DisplayMessage("error", "An error occured saving the File Type for this row.");
                    DisplayMessage("harderror", e.Message);
                }
            }

            if (batchRequest.Data.Count > 0)
            {
                BulkWriteToSheet(batchRequest);
                DisplayMessage("info", "Looks like that's it.", 2);
            }
            else
            {
                DisplayMessage("warning", "No File Types needed updating.", 2);
            }
        }

        private static string DetectFileTypeFromDirectory(string directory)
        {
            try
            {
                if (!Directory.Exists(directory))
                {
                    DisplayMessage("warning", $"Directory does not exist: {directory}");
                    return "0";
                }

                // Grab all files and then filter using your existing helper
                string[] fileEntries = Directory.GetFiles(directory);
                ArrayList videoFiles = GrabMovieFiles(fileEntries, false);

                var nonTrailerFiles = videoFiles
                    .Cast<object>()
                    .Select(v => v.ToString())
                    .Where(path => !Path.GetFileName(path)
                                       .ToLower()
                                       .Contains("-trailer"))
                    .ToList();

                if (nonTrailerFiles.Count == 0)
                {
                    DisplayMessage("warning", $"No movie file found in: {directory}");
                    return "0";
                }

                if (nonTrailerFiles.Count > 1)
                {
                    DisplayMessage("warning", $"Multiple movie files found in: {directory}");
                    return "0";
                }

                // Exactly one
                string ext = Path.GetExtension(nonTrailerFiles[0])
                                    .ToLower()
                                    .TrimStart('.');
                return ext;
            }
            catch (Exception e)
            {
                DisplayMessage("error", $"Error scanning directory: {directory}");
                DisplayMessage("harderror", e.Message);
                return "0";
            }
        }

        /// <summary>
        /// Uses the Clean Title column from the Movies Google Sheet to rename movie files
        /// and their corresponding SRT files in a chosen directory.
        /// Matching logic:
        ///   1) Try exact match on Clean Title (Title (Year))
        ///   2) If no match, strip the year and match on title only
        ///   3) If multiple title-only matches, ask user which one
        ///   4) If still no match, ask user to paste a Clean Title from the sheet
        /// </summary>
        static void RenameMoviesAndSrtsFromSheet(
            IList<IList<object>> movieData,
            Dictionary<string, int> movieSheetVariables,
            ArrayList movieFiles,
            string rootDirectory)
        {
            var renameLog = new List<string>();

            if (!movieSheetVariables.ContainsKey(CLEAN_TITLE) ||
                movieSheetVariables[CLEAN_TITLE] < 0)
            {
                DisplayMessage("error", "Clean Title column not found in Movies sheet.");
                return;
            }

            int cleanTitleIndex = movieSheetVariables[CLEAN_TITLE];
            int imdbTitleIndex = movieSheetVariables.ContainsKey("IMDB Title")
                ? movieSheetVariables["IMDB Title"]
                : -1;

            // ===== Build AllMovies list from the sheet =====
            AllMovies = movieData
                .Where(row => row.Count > cleanTitleIndex)
                .Select(row => new SheetMovie
                {
                    CleanTitle = row[cleanTitleIndex]?.ToString()?.Trim() ?? ""
                })
                .Where(m => !string.IsNullOrEmpty(m.CleanTitle))
                .ToList();

            if (AllMovies.Count == 0)
            {
                DisplayMessage("error", "No Clean Titles found in Movies sheet.");
                return;
            }

            // Also keep a simple list of Clean Titles for the manual fallback prompt
            var allCleanTitles = AllMovies
                .Select(m => m.CleanTitle)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            // ===== Process each movie file =====
            foreach (string moviePath in movieFiles)
            {
                if (string.IsNullOrWhiteSpace(moviePath) || !File.Exists(moviePath)) continue;

                string originalMoviePath = moviePath;
                string originalFileName = Path.GetFileName(originalMoviePath);
                string originalBase = Path.GetFileNameWithoutExtension(originalMoviePath);
                originalBase = FixDuplicateYear(originalBase);

                string folder = Path.GetDirectoryName(originalMoviePath) ?? rootDirectory;

                bool isTrailer = IsTrailerFileName(originalBase);

                // For matching ONLY, strip "trailer"/"teaser" so we can find the real title
                string baseForMatching = isTrailer ? StripTrailerTokens(originalBase) : originalBase;

                // Normalize filename for matching (handle underscores, extra spaces, etc.)
                string normalizedBaseForMatch = NormalizeTitleForMatch(baseForMatching);

                DisplayMessage("data", "");
                DisplayMessage("data", isTrailer
                    ? $"Processing trailer: {normalizedBaseForMatch}"
                    : $"Processing movie: {normalizedBaseForMatch}");

                // ===== 1) Try automatic matching using Clean Title logic =====
                SheetMovie matchedMovie = FindBestMatchFromSheet(normalizedBaseForMatch);

                // ===== 2) If still no match, ask user to paste a Clean Title =====
                if (matchedMovie == null)
                {
                    string manualCleanTitle = AskUserForImdbTitle(normalizedBaseForMatch, allCleanTitles);
                    if (manualCleanTitle == null)
                    {
                        renameLog.Add($"SKIP: {normalizedBaseForMatch} (no Clean Title selected)");
                        DisplayMessage("warning", "Skipping this movie; no title selected.");
                        continue;
                    }

                    matchedMovie = AllMovies
                        .FirstOrDefault(m =>
                            string.Equals(m.CleanTitle,
                                          manualCleanTitle,
                                          StringComparison.OrdinalIgnoreCase));

                    if (matchedMovie == null)
                    {
                        renameLog.Add($"SKIP: {normalizedBaseForMatch} (manual Clean Title not found in sheet)");
                        DisplayMessage("warning", "Clean Title not found in sheet, skipping.");
                        continue;
                    }
                }

                // ===== 3) Use Clean Title (Title (Year)) as final base name =====
                string finalBaseTitle = matchedMovie.CleanTitle;

                // IMPORTANT: trailers keep the movie Clean Title, but add the Plex suffix
                if (isTrailer)
                    finalBaseTitle = finalBaseTitle + "-trailer";

                string sanitizedBase = SanitizeFileName(finalBaseTitle);

                string currentMoviePath = originalMoviePath;
                string targetMoviePath = Path.Combine(folder, sanitizedBase + Path.GetExtension(originalMoviePath));

                string currentName = Path.GetFileName(currentMoviePath);
                string targetName = Path.GetFileName(targetMoviePath);

                bool movieRenamedOrAlreadyCorrect = false;

                // ===== 4) Rename movie file if needed =====
                if (string.Equals(currentMoviePath, targetMoviePath, StringComparison.OrdinalIgnoreCase))
                {
                    // Same file path ignoring case, so this may be a capitalization-only fix
                    if (!string.Equals(currentName, targetName, StringComparison.Ordinal))
                    {
                        string tempPath = Path.Combine(
                            folder,
                            Guid.NewGuid().ToString() + Path.GetExtension(originalMoviePath));

                        File.Move(currentMoviePath, tempPath);
                        File.Move(tempPath, targetMoviePath);

                        renameLog.Add($"RENAMED MOVIE (case fix): {currentName} -> {targetName}");
                        DisplayMessage("success", $"Fixed movie capitalization: {currentName} -> {targetName}");
                    }
                    else
                    {
                        DisplayMessage("info", $"Movie already correct: {currentName}");
                    }

                    movieRenamedOrAlreadyCorrect = true;
                }
                else if (!File.Exists(targetMoviePath))
                {
                    File.Move(currentMoviePath, targetMoviePath);

                    renameLog.Add($"RENAMED MOVIE: {currentName} -> {targetName}");
                    DisplayMessage("success", $"Renamed movie: {currentName} -> {targetName}");

                    movieRenamedOrAlreadyCorrect = true;
                }
                else
                {
                    renameLog.Add($"SKIP MOVIE: target already exists: {targetName}");
                    DisplayMessage("warning", $"Movie target already exists, skipping rename: {targetName}");
                }

                // ===== 5) Rename associated SRT files only if movie is now correct/aligned =====
                if (movieRenamedOrAlreadyCorrect)
                {
                    RenameSrtsForMovie(originalBase, sanitizedBase, folder, renameLog);
                }
                else
                {
                    renameLog.Add($"SKIP SRT RENAME: {originalBase} -> {sanitizedBase} (movie rename skipped due to conflicting target)");
                    DisplayMessage("warning", "Skipping SRT rename because movie rename was skipped due to an existing conflicting target.");
                }
            }

            // ===== 6) Write log to disk =====
            if (renameLog.Count > 0)
            {
                string path = BachFlixLog.WriteBachFlixLog(
                    renameLog,
                    category: "File Renamer",
                    fileNamePrefix: "FileRenamer",
                    out string err
                );

                if (path != null)
                    DisplayMessage("success", $"Rename log saved to: {path}");
                else
                    DisplayMessage("error", $"Failed to write rename log: {err}");
            }
            else
            {
                DisplayMessage("info", "No changes were made; no log file created.");
            }
        }

        private static string NormalizeUserPath(string rawInput)
        {
            if (string.IsNullOrWhiteSpace(rawInput))
                return string.Empty;

            // Trim whitespace first
            string s = rawInput.Trim();

            // Keep peeling off matching outer quotes (single or double)
            bool changed = true;
            while (changed && s.Length > 1)
            {
                changed = false;

                if (s.StartsWith("\"") && s.EndsWith("\""))
                {
                    s = s.Substring(1, s.Length - 2).Trim();
                    changed = true;
                }
                else if (s.StartsWith("'") && s.EndsWith("'"))
                {
                    s = s.Substring(1, s.Length - 2).Trim();
                    changed = true;
                }
            }

            return s;
        }

        private static string GetTitleWithoutYear(string cleanTitle)
        {
            if (string.IsNullOrWhiteSpace(cleanTitle)) return string.Empty;

            var match = TitleYearRegex.Match(cleanTitle.Trim());
            if (match.Success)
            {
                return match.Groups["title"].Value.Trim();
            }

            // If it doesn't follow "Title (YYYY)" pattern, just return trimmed
            return cleanTitle.Trim();
        }

        /// <summary>
        /// Fixes cases like "Title (2017) (2017)" -> "Title (2017)".
        /// Only touches exactly repeated year-at-the-end patterns.
        /// </summary>
        private static string FixDuplicateYear(string title)
        {
            if (string.IsNullOrWhiteSpace(title)) return title;

            return DuplicateYearRegex.Replace(title, "($1)");
        }

        class SheetMovie
        {
            public string CleanTitle { get; set; }
        }

        static List<SheetMovie> AllMovies;


        private static SheetMovie FindBestMatchFromSheet(string fileBasedCleanTitle)
        {
            if (string.IsNullOrWhiteSpace(fileBasedCleanTitle))
                return null;

            // Fix weird cases like "Title (2017) (2017)"
            fileBasedCleanTitle = FixDuplicateYear(fileBasedCleanTitle);

            // Normalize the incoming filename-based title
            var candidateFull = NormalizeForMatching(fileBasedCleanTitle);

            // ===== 1. Exact match on Clean Title (title + year) =====
            var exactMatch = AllMovies
                .FirstOrDefault(m =>
                    NormalizeForMatching(m.CleanTitle) == candidateFull);

            if (exactMatch != null)
            {
                // Found perfect match, we're done
                return exactMatch;
            }

            // ===== 2. Fallback: match by title only (ignore year) =====
            var candidateTitleOnly = NormalizeForMatching(GetTitleWithoutYear(fileBasedCleanTitle));

            var titleMatches = AllMovies
                .Where(m =>
                    NormalizeForMatching(GetTitleWithoutYear(m.CleanTitle)) == candidateTitleOnly)
                .ToList();

            if (titleMatches.Count == 0)
            {
                // No match even by title-only
                return null;
            }

            if (titleMatches.Count == 1)
            {
                // Exactly one title-only match: assume it’s right
                Console.WriteLine($"[INFO] No exact year match for \"{fileBasedCleanTitle}\",");
                Console.WriteLine($"       but found single title match: \"{titleMatches[0].CleanTitle}\"");
                return titleMatches[0];
            }

            // ===== 3. Multiple title-only matches: ask user =====
            Console.WriteLine();
            Console.WriteLine($"[INPUT] Multiple matches found for \"{candidateTitleOnly}\":");
            for (int i = 0; i < titleMatches.Count; i++)
            {
                Console.WriteLine($"  {i + 1}) {titleMatches[i].CleanTitle}");
            }
            Console.WriteLine("  0) None of the above / skip");

            while (true)
            {
                Console.Write("Select the correct movie # (or 0 to skip): ");
                var input = Console.ReadLine();

                if (int.TryParse(input, out int choice))
                {
                    if (choice == 0)
                        return null;

                    if (choice >= 1 && choice <= titleMatches.Count)
                        return titleMatches[choice - 1];
                }

                Console.WriteLine("Invalid choice, please try again.");
            }
        }

        private static string NormalizeForMatching(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
                return string.Empty;

            // Replace underscores (StreamFab)
            title = title.Replace('_', ' ');

            // Remove punctuation except letters/numbers/spaces
            title = Regex.Replace(title, @"[^\w\s]", "");

            // Collapse multiple spaces
            title = Regex.Replace(title, @"\s+", " ").Trim();

            return title.ToLowerInvariant();
        }

        private static string NormalizeTitleForMatch(string title)
        {
            if (string.IsNullOrWhiteSpace(title)) return string.Empty;

            // 1) Replace StreamFab underscores with spaces
            title = title.Replace('_', ' ');

            // 2) Collapse multiple spaces
            title = System.Text.RegularExpressions.Regex.Replace(title, @"\s+", " ");

            // 3) Trim whitespace
            return title.Trim();
        }

        /// <summary>
        /// Prompt user when automatic match failed.
        /// User pastes the IMDB Title exactly, or types 'skip'.
        /// </summary>
        static string AskUserForImdbTitle(string fileName, List<string> imdbTitles)
        {
            while (true)
            {
                DisplayMessage("warning", $"I can't find the match for \"{fileName}\", what do you suggest?");
                Type("Paste the IMDB Title exactly as it appears in your Google Sheet, or type 'skip' to skip this file.", 0, 0, 1, "Yellow");
                Type("> ", 0, 0, 0);
                string input = Console.ReadLine()?.Trim();

                if (string.IsNullOrEmpty(input))
                {
                    DisplayMessage("warning", "Empty input; please paste a title or type 'skip'.");
                    continue;
                }

                if (string.Equals(input, "skip", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                // First try exact match
                var exact = imdbTitles
                    .FirstOrDefault(t => string.Equals(t, input, StringComparison.OrdinalIgnoreCase));

                if (exact != null)
                {
                    DisplayMessage("success", $"Using IMDB Title: {exact}");
                    return exact;
                }

                // Then try "contains" to help the user find typos
                var candidates = imdbTitles
                    .Where(t => t.IndexOf(input, StringComparison.OrdinalIgnoreCase) >= 0)
                    .Take(10)
                    .ToList();

                if (candidates.Count == 0)
                {
                    DisplayMessage("warning", "No matching titles found in the sheet. Try again or type 'skip'.");
                    continue;
                }

                if (candidates.Count == 1)
                {
                    DisplayMessage("success", $"Found one candidate: {candidates[0]}");
                    return candidates[0];
                }

                // Let user choose from multiple
                Type("Multiple matches found:", 0, 0, 1, "Yellow");
                for (int i = 0; i < candidates.Count; i++)
                {
                    Type($"  [{i + 1}] {candidates[i]}", 0, 0, 1);
                }

                Type("Choose a number (or 0 to try again): ", 0, 0, 0);
                string choice = Console.ReadLine();

                if (int.TryParse(choice, out int idx) &&
                    idx >= 1 &&
                    idx <= candidates.Count)
                {
                    string chosen = candidates[idx - 1];
                    DisplayMessage("success", $"Using IMDB Title: {chosen}");
                    return chosen;
                }

                DisplayMessage("warning", "Invalid choice. Let's try again.");
            }
        }

        /// <summary>
        /// For a given movie, find all SRT files in the folder that start with the original base name
        /// and rename them to the normalized pattern based on detected type:
        ///  - cc / closed captions -> .eng.cc.srt
        ///  - forced              -> .eng.forced.srt
        ///  - SDH / HOH           -> .eng.sdh.srt
        ///  - otherwise           -> .eng.srt
        /// </summary>
        static void RenameSrtsForMovie(string originalBase, string finalBase, string folder, List<string> renameLog)
        {
            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder)) return;

            var srtFiles = Directory.GetFiles(folder, "*.srt", SearchOption.TopDirectoryOnly)
                                    .Where(path =>
                                    {
                                        string srtBase = Path.GetFileNameWithoutExtension(path);
                                        // Handle names like "Frankenstein (2025).en.closedcaptions.srt"
                                        return srtBase.StartsWith(originalBase, StringComparison.OrdinalIgnoreCase) ||
                                               srtBase.StartsWith(finalBase, StringComparison.OrdinalIgnoreCase);
                                    })
                                    .ToList();

            if (!srtFiles.Any())
            {
                return;
            }

            foreach (var srtPath in srtFiles)
            {
                string oldName = Path.GetFileName(srtPath);
                string lower = oldName.ToLowerInvariant();

                SrtKind kind = DetectSrtKind(lower);

                string newSrtName = BuildNewSrtName(finalBase, kind);
                string newSrtPath = Path.Combine(folder, newSrtName);

                // Already correct?
                if (string.Equals(oldName, newSrtName, StringComparison.OrdinalIgnoreCase))
                {
                    renameLog.Add($"SRT: {oldName} (already correct)");
                    continue;
                }

                if (File.Exists(newSrtPath))
                {
                    DisplayMessage("warning", $"SRT target already exists, skipping rename: {newSrtName}");
                    renameLog.Add($"SKIP-SRT: {oldName} -> {newSrtName} (target exists)");
                    continue;
                }

                File.Move(srtPath, newSrtPath);
                DisplayMessage("success", $"Renamed SRT: {oldName} -> {newSrtName}", 1);
                renameLog.Add($"SRT: {oldName} -> {newSrtName}");
            }
        }

        enum SrtKind
        {
            Normal,
            Sdh,
            Cc,
            Forced
        }

        /// <summary>
        /// Detect if an SRT filename is SDH, CC, or forced based on keywords.
        /// - Example mappings:
        ///   Frankenstein (2025).en.closedcaptions.srt          -> CC
        ///   Frankenstein (2025).en.forced.srt                 -> Forced
        ///   Frankenstein (2025).en.STRIPPED_SDH.subtitles.srt -> SDH
        /// </summary>
        static SrtKind DetectSrtKind(string lowerFileName)
        {
            if (string.IsNullOrEmpty(lowerFileName)) return SrtKind.Normal;

            if (lowerFileName.Contains("forced"))
                return SrtKind.Forced;

            if (lowerFileName.Contains("sdh") ||
                lowerFileName.Contains("hearing") ||
                lowerFileName.Contains("hard.of.hearing") ||
                lowerFileName.Contains("hoh"))
                return SrtKind.Sdh;

            if (lowerFileName.Contains(".cc.") ||
                lowerFileName.EndsWith(".cc.srt") ||
                lowerFileName.Contains("closedcaption") ||
                lowerFileName.Contains("closed_captions") ||
                lowerFileName.Contains("closedcaptions") ||
                lowerFileName.Contains("closed.captions"))
                return SrtKind.Cc;

            return SrtKind.Normal;
        }

        /// <summary>
        /// Build final SRT filename using the IMDB Title base and detected kind.
        /// </summary>
        static string BuildNewSrtName(string baseTitle, SrtKind kind)
        {
            var sb = new StringBuilder();
            sb.Append(baseTitle);
            sb.Append(".eng");

            switch (kind)
            {
                case SrtKind.Forced:
                    sb.Append(".forced");
                    break;
                case SrtKind.Sdh:
                    sb.Append(".sdh");
                    break;
                case SrtKind.Cc:
                    sb.Append(".cc");
                    break;
                case SrtKind.Normal:
                default:
                    // no extra suffix
                    break;
            }

            sb.Append(".srt");
            return sb.ToString();
        }

        /// <summary>
        /// Simple Windows filename sanitizer: replaces invalid characters with ''.
        /// </summary>
        static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return name;

            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);

            foreach (var ch in name)
            {
                if (!invalid.Contains(ch))
                    sb.Append(ch);
            }

            return sb.ToString().Trim();
        }

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
                        rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
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

                            plot1Data = response.data[0].overview.ToString().Trim();

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

                            plot2Data = response.data[0].overview.ToString().Trim();

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
                    Type(e.Message, 0, 0, 2, "DarkRed");
                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("New Plots inserted for episode one: " + intPlot1InsertedCount, 0, 0, 1, "Green");
            Type("New Plots inserted for episode two: " + intPlot2InsertedCount, 0, 0, 1, "Green");
            Type("Plots skipped due to no plot for episode one: " + intPlot1EmptyCount, 0, 0, 1, "Yellow");
            Type("Plots skipped due to no plot for episode two: " + intPlot2EmptyCount, 0, 0, 1, "Yellow");

        } // End InsertMissingCombinedEpisodeData()

        private static void InsertMissingSeveralCombinedEpisodeData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string token)
        {
            int plot1InsertedCount = 0,
                plot2InsertedCount = 0,
                plot3InsertedCount = 0,
                plot4InsertedCount = 0,
                plot5InsertedCount = 0,
                plot6InsertedCount = 0,
                plot7InsertedCount = 0,
                plot8InsertedCount = 0,
                plot9InsertedCount = 0,
                plot10InsertedCount = 0,
                plot1EmptyCount = 0,
                plot2EmptyCount = 0,
                plot3EmptyCount = 0,
                plot4EmptyCount = 0,
                plot5EmptyCount = 0,
                plot6EmptyCount = 0,
                plot7EmptyCount = 0,
                plot8EmptyCount = 0,
                plot9EmptyCount = 0,
                plot10EmptyCount = 0,
                title1InsertedCount = 0,
                title2InsertedCount = 0,
                title3InsertedCount = 0,
                title4InsertedCount = 0,
                title5InsertedCount = 0,
                title6InsertedCount = 0,
                title7InsertedCount = 0,
                title8InsertedCount = 0,
                title9InsertedCount = 0,
                title10InsertedCount = 0,
                title1EmptyCount = 0,
                title2EmptyCount = 0,
                title3EmptyCount = 0,
                title4EmptyCount = 0,
                title5EmptyCount = 0,
                title6EmptyCount = 0,
                title7EmptyCount = 0,
                title8EmptyCount = 0,
                title9EmptyCount = 0,
                title10EmptyCount = 0;

            foreach (var row in data)
            {
                try
                {
                    if (row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString() != "")
                    {
                        string rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                        string tvdbIdValue = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                        string strSeriesName = row[Convert.ToInt16(sheetVariables["Series Name"])].ToString();
                        string episode1Season = row[Convert.ToInt16(sheetVariables["Episode 1 Season"])].ToString();
                        string episode1No = row[Convert.ToInt16(sheetVariables["Episode 1 No."])].ToString();
                        string episode1Plot = row[Convert.ToInt16(sheetVariables["Episode 1 Plot"])].ToString();
                        string episode1Title = row[Convert.ToInt16(sheetVariables["Episode 1 Title"])].ToString();
                        string episode2Season = row[Convert.ToInt16(sheetVariables["Episode 2 Season"])].ToString();
                        string episode2No = row[Convert.ToInt16(sheetVariables["Episode 2 No."])].ToString();
                        string episode2Plot = row[Convert.ToInt16(sheetVariables["Episode 2 Plot"])].ToString();
                        string episode2Title = row[Convert.ToInt16(sheetVariables["Episode 2 Title"])].ToString();
                        string episode3Season = row[Convert.ToInt16(sheetVariables["Episode 3 Season"])].ToString();
                        string episode3No = row[Convert.ToInt16(sheetVariables["Episode 3 No."])].ToString();
                        string episode3Plot = row[Convert.ToInt16(sheetVariables["Episode 3 Plot"])].ToString();
                        string episode3Title = row[Convert.ToInt16(sheetVariables["Episode 3 Title"])].ToString();
                        string episode4Season = row[Convert.ToInt16(sheetVariables["Episode 4 Season"])].ToString();
                        string episode4No = row[Convert.ToInt16(sheetVariables["Episode 4 No."])].ToString();
                        string episode4Plot = row[Convert.ToInt16(sheetVariables["Episode 4 Plot"])].ToString();
                        string episode4Title = row[Convert.ToInt16(sheetVariables["Episode 4 Title"])].ToString();
                        string episode5Season = row[Convert.ToInt16(sheetVariables["Episode 5 Season"])].ToString();
                        string episode5No = row[Convert.ToInt16(sheetVariables["Episode 5 No."])].ToString();
                        string episode5Plot = row[Convert.ToInt16(sheetVariables["Episode 5 Plot"])].ToString();
                        string episode5Title = row[Convert.ToInt16(sheetVariables["Episode 5 Title"])].ToString();
                        string episode6Season = row[Convert.ToInt16(sheetVariables["Episode 6 Season"])].ToString();
                        string episode6No = row[Convert.ToInt16(sheetVariables["Episode 6 No."])].ToString();
                        string episode6Plot = row[Convert.ToInt16(sheetVariables["Episode 6 Plot"])].ToString();
                        string episode6Title = row[Convert.ToInt16(sheetVariables["Episode 6 Title"])].ToString();
                        string episode7Season = row[Convert.ToInt16(sheetVariables["Episode 7 Season"])].ToString();
                        string episode7No = row[Convert.ToInt16(sheetVariables["Episode 7 No."])].ToString();
                        string episode7Plot = row[Convert.ToInt16(sheetVariables["Episode 7 Plot"])].ToString();
                        string episode7Title = row[Convert.ToInt16(sheetVariables["Episode 7 Title"])].ToString();
                        string episode8Season = row[Convert.ToInt16(sheetVariables["Episode 8 Season"])].ToString();
                        string episode8No = row[Convert.ToInt16(sheetVariables["Episode 8 No."])].ToString();
                        string episode8Plot = row[Convert.ToInt16(sheetVariables["Episode 8 Plot"])].ToString();
                        string episode8Title = row[Convert.ToInt16(sheetVariables["Episode 8 Title"])].ToString();
                        string episode9Season = row[Convert.ToInt16(sheetVariables["Episode 9 Season"])].ToString();
                        string episode9No = row[Convert.ToInt16(sheetVariables["Episode 9 No."])].ToString();
                        string episode9Plot = row[Convert.ToInt16(sheetVariables["Episode 9 Plot"])].ToString();
                        string episode9Title = row[Convert.ToInt16(sheetVariables["Episode 9 Title"])].ToString();
                        string episode10Season = row[Convert.ToInt16(sheetVariables["Episode 10 Season"])].ToString();
                        string episode10No = row[Convert.ToInt16(sheetVariables["Episode 10 No."])].ToString();
                        string episode10Plot = row[Convert.ToInt16(sheetVariables["Episode 10 Plot"])].ToString();
                        string episode10Title = row[Convert.ToInt16(sheetVariables["Episode 10 Title"])].ToString();
                        string strCellToPutData = "";

                        int episode1PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 1 Plot"]);
                        int episode1TitleColumnNum = Convert.ToInt16(sheetVariables["Episode 1 Title"]);
                        int episode2PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 2 Plot"]);
                        int episode2TitleColumnNum = Convert.ToInt16(sheetVariables["Episode 2 Title"]);
                        int episode3PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 3 Plot"]);
                        int episode3TitleColumnNum = Convert.ToInt16(sheetVariables["Episode 3 Title"]);
                        int episode4PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 4 Plot"]);
                        int episode4TitleColumnNum = Convert.ToInt16(sheetVariables["Episode 4 Title"]);
                        int episode5PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 5 Plot"]);
                        int episode5TitleColumnNum = Convert.ToInt16(sheetVariables["Episode 5 Title"]);
                        int episode6PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 6 Plot"]);
                        int episode6TitleColumnNum = Convert.ToInt16(sheetVariables["Episode 6 Title"]);
                        int episode7PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 7 Plot"]);
                        int episode7TitleColumnNum = Convert.ToInt16(sheetVariables["Episode 7 Title"]);
                        int episode8PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 8 Plot"]);
                        int episode8TitleColumnNum = Convert.ToInt16(sheetVariables["Episode 8 Title"]);
                        int episode9PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 9 Plot"]);
                        int episode9TitleColumnNum = Convert.ToInt16(sheetVariables["Episode 9 Title"]);
                        int episode10PlotColumnNum = Convert.ToInt16(sheetVariables["Episode 10 Plot"]);
                        int episode10TitleColumnNum = Convert.ToInt16(sheetVariables["Episode 10 Title"]);

                        if ((!episode1Season.Equals("") && !episode1No.Equals("")) && (episode1Plot.Equals("") || episode1Title.Equals("")))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode1Season, episode1No);

                            string plot = response.data[0].overview.ToString().Trim(),
                                name = response.data[0].episodeName.ToString().Trim();

                            if (episode1Plot.Equals(""))
                            {
                                if (plot.Equals(""))
                                {
                                    DisplayMessage("default", "No Plot available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode1Season + "E" + episode1No);

                                    plot1EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode1PlotColumnNum) + rowNum;

                                    WriteSingleCellToSheet(plot, strCellToPutData);
                                    DisplayMessage("default", "Plot saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode1Season + "E" + episode1No + " - at " + strCellToPutData);

                                    plot1InsertedCount++;
                                }
                            }

                            if (episode1Title.Equals(""))
                            {
                                if (name.Equals(""))
                                {
                                    DisplayMessage("default", "No Episode Title available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode1Season + "E" + episode1No);

                                    title1EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode1TitleColumnNum) + rowNum;
                                    name = CleanAmpersands(name);
                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("default", "Episode Title saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode1Season + "E" + episode1No + " - at " + strCellToPutData);

                                    title1InsertedCount++;
                                }
                            }
                        }
                        if ((!episode2Season.Equals("") && !episode2No.Equals("")) && (episode2Plot.Equals("") || episode2Title.Equals("")))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode2Season, episode2No);

                            string plot = response.data[0].overview.ToString().Trim(),
                                name = response.data[0].episodeName.ToString().Trim();

                            if (episode2Plot.Equals(""))
                            {
                                if (plot.Equals(""))
                                {
                                    DisplayMessage("default", "No Plot available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode2Season + "E" + episode2No);

                                    plot2EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode2PlotColumnNum) + rowNum;

                                    WriteSingleCellToSheet(plot, strCellToPutData);
                                    DisplayMessage("default", "Plot saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode2Season + "E" + episode2No + " - at " + strCellToPutData);

                                    plot2InsertedCount++;
                                }
                            }

                            if (episode2Title.Equals(""))
                            {
                                if (name.Equals(""))
                                {
                                    DisplayMessage("default", "No Episode Title available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode2Season + "E" + episode2No);

                                    title2EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode2TitleColumnNum) + rowNum;
                                    name = CleanAmpersands(name);
                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("default", "Episode Title saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode2Season + "E" + episode2No + " - at " + strCellToPutData);

                                    title2InsertedCount++;
                                }
                            }
                        }
                        if ((!episode3Season.Equals("") && !episode3No.Equals("")) && (episode3Plot.Equals("") || episode3Title.Equals("")))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode3Season, episode3No);

                            string plot = response.data[0].overview.ToString().Trim(),
                                name = response.data[0].episodeName.ToString().Trim();

                            if (episode3Plot.Equals(""))
                            {
                                if (plot.Equals(""))
                                {
                                    DisplayMessage("default", "No Plot available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode3Season + "E" + episode3No);

                                    plot3EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode3PlotColumnNum) + rowNum;

                                    WriteSingleCellToSheet(plot, strCellToPutData);
                                    DisplayMessage("default", "Plot saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode3Season + "E" + episode3No + " - at " + strCellToPutData);

                                    plot3InsertedCount++;
                                }
                            }

                            if (episode3Title.Equals(""))
                            {
                                if (name.Equals(""))
                                {
                                    DisplayMessage("default", "No Episode Title available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode3Season + "E" + episode3No);

                                    title3EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode3TitleColumnNum) + rowNum;
                                    name = CleanAmpersands(name);
                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("default", "Episode Title saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode3Season + "E" + episode3No + " - at " + strCellToPutData);

                                    title3InsertedCount++;
                                }
                            }
                        }
                        if ((!episode4Season.Equals("") && !episode4No.Equals("")) && (episode4Plot.Equals("") || episode4Title.Equals("")))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode4Season, episode4No);

                            string plot = response.data[0].overview.ToString().Trim(),
                                name = response.data[0].episodeName.ToString().Trim();

                            if (episode4Plot.Equals(""))
                            {
                                if (plot.Equals(""))
                                {
                                    DisplayMessage("default", "No Plot available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode4Season + "E" + episode4No);

                                    plot4EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode4PlotColumnNum) + rowNum;

                                    WriteSingleCellToSheet(plot, strCellToPutData);
                                    DisplayMessage("default", "Plot saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode4Season + "E" + episode4No + " - at " + strCellToPutData);

                                    plot4InsertedCount++;
                                }
                            }

                            if (episode4Title.Equals(""))
                            {
                                if (name.Equals(""))
                                {
                                    DisplayMessage("default", "No Episode Title available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode4Season + "E" + episode4No);

                                    title4EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode4TitleColumnNum) + rowNum;
                                    name = CleanAmpersands(name);
                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("default", "Episode Title saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode4Season + "E" + episode4No + " - at " + strCellToPutData);

                                    title4InsertedCount++;
                                }
                            }
                        }
                        if ((!episode5Season.Equals("") && !episode5No.Equals("")) && (episode5Plot.Equals("") || episode5Title.Equals("")))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode5Season, episode5No);

                            string plot = response.data[0].overview.ToString().Trim(),
                                name = response.data[0].episodeName.ToString().Trim();

                            if (episode5Plot.Equals(""))
                            {
                                if (plot.Equals(""))
                                {
                                    DisplayMessage("default", "No Plot available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode5Season + "E" + episode5No);

                                    plot5EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode5PlotColumnNum) + rowNum;

                                    WriteSingleCellToSheet(plot, strCellToPutData);
                                    DisplayMessage("default", "Plot saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode5Season + "E" + episode5No + " - at " + strCellToPutData);

                                    plot5InsertedCount++;
                                }
                            }

                            if (episode5Title.Equals(""))
                            {
                                if (name.Equals(""))
                                {
                                    DisplayMessage("default", "No Episode Title available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode5Season + "E" + episode5No);

                                    title5EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode5TitleColumnNum) + rowNum;
                                    name = CleanAmpersands(name);
                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("default", "Episode Title saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode5Season + "E" + episode5No + " - at " + strCellToPutData);

                                    title5InsertedCount++;
                                }
                            }
                        }
                        if ((!episode6Season.Equals("") && !episode6No.Equals("")) && (episode6Plot.Equals("") || episode6Title.Equals("")))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode6Season, episode6No);

                            string plot = response.data[0].overview.ToString().Trim(),
                                name = response.data[0].episodeName.ToString().Trim();

                            if (episode6Plot.Equals(""))
                            {
                                if (plot.Equals(""))
                                {
                                    DisplayMessage("default", "No Plot available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode6Season + "E" + episode6No);

                                    plot6EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode6PlotColumnNum) + rowNum;

                                    WriteSingleCellToSheet(plot, strCellToPutData);
                                    DisplayMessage("default", "Plot saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode6Season + "E" + episode6No + " - at " + strCellToPutData);

                                    plot6InsertedCount++;
                                }
                            }

                            if (episode6Title.Equals(""))
                            {
                                if (name.Equals(""))
                                {
                                    DisplayMessage("default", "No Episode Title available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode6Season + "E" + episode6No);

                                    title6EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode6TitleColumnNum) + rowNum;
                                    name = CleanAmpersands(name);
                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("default", "Episode Title saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode6Season + "E" + episode6No + " - at " + strCellToPutData);

                                    title6InsertedCount++;
                                }
                            }
                        }
                        if ((!episode7Season.Equals("") && !episode7No.Equals("")) && (episode7Plot.Equals("") || episode7Title.Equals("")))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode7Season, episode7No);

                            string plot = response.data[0].overview.ToString().Trim(),
                                name = response.data[0].episodeName.ToString().Trim();

                            if (episode7Plot.Equals(""))
                            {
                                if (plot.Equals(""))
                                {
                                    DisplayMessage("default", "No Plot available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode7Season + "E" + episode7No);

                                    plot7EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode7PlotColumnNum) + rowNum;

                                    WriteSingleCellToSheet(plot, strCellToPutData);
                                    DisplayMessage("default", "Plot saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode7Season + "E" + episode7No + " - at " + strCellToPutData);

                                    plot7InsertedCount++;
                                }
                            }

                            if (episode7Title.Equals(""))
                            {
                                if (name.Equals(""))
                                {
                                    DisplayMessage("default", "No Episode Title available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode7Season + "E" + episode7No);

                                    title7EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode7TitleColumnNum) + rowNum;
                                    name = CleanAmpersands(name);
                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("default", "Episode Title saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode7Season + "E" + episode7No + " - at " + strCellToPutData);

                                    title7InsertedCount++;
                                }
                            }
                        }
                        if ((!episode8Season.Equals("") && !episode8No.Equals("")) && (episode8Plot.Equals("") || episode8Title.Equals("")))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode8Season, episode8No);

                            string plot = response.data[0].overview.ToString().Trim(),
                                name = response.data[0].episodeName.ToString().Trim();

                            if (episode8Plot.Equals(""))
                            {
                                if (plot.Equals(""))
                                {
                                    DisplayMessage("default", "No Plot available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode8Season + "E" + episode8No);

                                    plot8EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode8PlotColumnNum) + rowNum;

                                    WriteSingleCellToSheet(plot, strCellToPutData);
                                    DisplayMessage("default", "Plot saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode8Season + "E" + episode8No + " - at " + strCellToPutData);

                                    plot8InsertedCount++;
                                }
                            }

                            if (episode8Title.Equals(""))
                            {
                                if (name.Equals(""))
                                {
                                    DisplayMessage("default", "No Episode Title available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode8Season + "E" + episode8No);

                                    title8EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode8TitleColumnNum) + rowNum;
                                    name = CleanAmpersands(name);
                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("default", "Episode Title saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode8Season + "E" + episode8No + " - at " + strCellToPutData);

                                    title8InsertedCount++;
                                }
                            }
                        }
                        if ((!episode9Season.Equals("") && !episode9No.Equals("")) && (episode9Plot.Equals("") || episode9Title.Equals("")))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode9Season, episode9No);

                            string plot = response.data[0].overview.ToString().Trim(),
                                name = response.data[0].episodeName.ToString().Trim();

                            if (episode9Plot.Equals(""))
                            {
                                if (plot.Equals(""))
                                {
                                    DisplayMessage("default", "No Plot available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode9Season + "E" + episode9No);

                                    plot9EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode9PlotColumnNum) + rowNum;

                                    WriteSingleCellToSheet(plot, strCellToPutData);
                                    DisplayMessage("default", "Plot saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode9Season + "E" + episode9No + " - at " + strCellToPutData);

                                    plot9InsertedCount++;
                                }
                            }

                            if (episode9Title.Equals(""))
                            {
                                if (name.Equals(""))
                                {
                                    DisplayMessage("default", "No Episode Title available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode9Season + "E" + episode9No);

                                    title9EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode9TitleColumnNum) + rowNum;
                                    name = CleanAmpersands(name);
                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("default", "Episode Title saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode9Season + "E" + episode9No + " - at " + strCellToPutData);

                                    title9InsertedCount++;
                                }
                            }
                        }
                        if ((!episode10Season.Equals("") && !episode10No.Equals("")) && (episode10Plot.Equals("") || episode10Title.Equals("")))
                        {
                            var response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode10Season, episode10No);

                            string plot = response.data[0].overview.ToString().Trim(),
                                name = response.data[0].episodeName.ToString().Trim();

                            if (episode10Plot.Equals(""))
                            {
                                if (plot.Equals(""))
                                {
                                    DisplayMessage("default", "No Plot available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode10Season + "E" + episode10No);

                                    plot10EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode10PlotColumnNum) + rowNum;

                                    WriteSingleCellToSheet(plot, strCellToPutData);
                                    DisplayMessage("default", "Plot saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode10Season + "E" + episode10No + " - at " + strCellToPutData);

                                    plot10InsertedCount++;
                                }
                            }

                            if (episode10Title.Equals(""))
                            {
                                if (name.Equals(""))
                                {
                                    DisplayMessage("default", "No Episode Title available for ", 0);
                                    DisplayMessage("warning", strSeriesName + " - S" + episode10Season + "E" + episode10No);

                                    title10EmptyCount++;
                                }
                                else
                                {
                                    strCellToPutData = "Several Combined Episodes!" + ColumnNumToLetter(episode10TitleColumnNum) + rowNum;
                                    name = CleanAmpersands(name);
                                    WriteSingleCellToSheet(name, strCellToPutData);
                                    DisplayMessage("default", "Episode Title saved for ", 0);
                                    DisplayMessage("success", strSeriesName + " - S" + episode10Season + "E" + episode10No + " - at " + strCellToPutData);

                                    title10InsertedCount++;
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Type(e.Message, 0, 0, 2, "DarkRed");
                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            if (plot1InsertedCount > 0) Type("New Plots inserted for episode 1: " + plot1InsertedCount, 0, 0, 1, "Green");
            if (plot2InsertedCount > 0) Type("New Plots inserted for episode 2: " + plot2InsertedCount, 0, 0, 1, "Green");
            if (plot3InsertedCount > 0) Type("New Plots inserted for episode 3: " + plot3InsertedCount, 0, 0, 1, "Green");
            if (plot4InsertedCount > 0) Type("New Plots inserted for episode 4: " + plot4InsertedCount, 0, 0, 1, "Green");
            if (plot5InsertedCount > 0) Type("New Plots inserted for episode 5: " + plot5InsertedCount, 0, 0, 1, "Green");
            if (plot6InsertedCount > 0) Type("New Plots inserted for episode 6: " + plot6InsertedCount, 0, 0, 1, "Green");
            if (plot7InsertedCount > 0) Type("New Plots inserted for episode 7: " + plot7InsertedCount, 0, 0, 1, "Green");
            if (plot8InsertedCount > 0) Type("New Plots inserted for episode 8: " + plot8InsertedCount, 0, 0, 1, "Green");
            if (plot9InsertedCount > 0) Type("New Plots inserted for episode 9: " + plot9InsertedCount, 0, 0, 1, "Green");
            if (plot10InsertedCount > 0) Type("New Plots inserted for episode 10: " + plot10InsertedCount, 0, 0, 1, "Green");
            if (plot1EmptyCount > 0) Type("Plots skipped due to no plot for episode 1: " + plot1EmptyCount, 0, 0, 1, "Yellow");
            if (plot2EmptyCount > 0) Type("Plots skipped due to no plot for episode 2: " + plot2EmptyCount, 0, 0, 1, "Yellow");
            if (plot3EmptyCount > 0) Type("Plots skipped due to no plot for episode 3: " + plot3EmptyCount, 0, 0, 1, "Yellow");
            if (plot4EmptyCount > 0) Type("Plots skipped due to no plot for episode 4: " + plot4EmptyCount, 0, 0, 1, "Yellow");
            if (plot5EmptyCount > 0) Type("Plots skipped due to no plot for episode 5: " + plot5EmptyCount, 0, 0, 1, "Yellow");
            if (plot6EmptyCount > 0) Type("Plots skipped due to no plot for episode 6: " + plot6EmptyCount, 0, 0, 1, "Yellow");
            if (plot7EmptyCount > 0) Type("Plots skipped due to no plot for episode 7: " + plot7EmptyCount, 0, 0, 1, "Yellow");
            if (plot8EmptyCount > 0) Type("Plots skipped due to no plot for episode 8: " + plot8EmptyCount, 0, 0, 1, "Yellow");
            if (plot9EmptyCount > 0) Type("Plots skipped due to no plot for episode 9: " + plot9EmptyCount, 0, 0, 1, "Yellow");
            if (plot10EmptyCount > 0) Type("Plots skipped due to no plot for episode 10: " + plot10EmptyCount, 0, 0, 1, "Yellow");
            if (title1InsertedCount > 0) Type("New Titles inserted for episode 1: " + title1InsertedCount, 0, 0, 1, "Green");
            if (title2InsertedCount > 0) Type("New Titles inserted for episode 2: " + title2InsertedCount, 0, 0, 1, "Green");
            if (title3InsertedCount > 0) Type("New Titles inserted for episode 3: " + title3InsertedCount, 0, 0, 1, "Green");
            if (title4InsertedCount > 0) Type("New Titles inserted for episode 4: " + title4InsertedCount, 0, 0, 1, "Green");
            if (title5InsertedCount > 0) Type("New Titles inserted for episode 5: " + title5InsertedCount, 0, 0, 1, "Green");
            if (title6InsertedCount > 0) Type("New Titles inserted for episode 6: " + title6InsertedCount, 0, 0, 1, "Green");
            if (title7InsertedCount > 0) Type("New Titles inserted for episode 7: " + title7InsertedCount, 0, 0, 1, "Green");
            if (title8InsertedCount > 0) Type("New Titles inserted for episode 8: " + title8InsertedCount, 0, 0, 1, "Green");
            if (title9InsertedCount > 0) Type("New Titles inserted for episode 9: " + title9InsertedCount, 0, 0, 1, "Green");
            if (title10InsertedCount > 0) Type("New Titles inserted for episode 10: " + title10InsertedCount, 0, 0, 1, "Green");
            if (title1EmptyCount > 0) Type("Titles skipped due to no title for episode 1: " + title1EmptyCount, 0, 0, 1, "Yellow");
            if (title2EmptyCount > 0) Type("Titles skipped due to no title for episode 2: " + title2EmptyCount, 0, 0, 1, "Yellow");
            if (title3EmptyCount > 0) Type("Titles skipped due to no title for episode 3: " + title3EmptyCount, 0, 0, 1, "Yellow");
            if (title4EmptyCount > 0) Type("Titles skipped due to no title for episode 4: " + title4EmptyCount, 0, 0, 1, "Yellow");
            if (title5EmptyCount > 0) Type("Titles skipped due to no title for episode 5: " + title5EmptyCount, 0, 0, 1, "Yellow");
            if (title6EmptyCount > 0) Type("Titles skipped due to no title for episode 6: " + title6EmptyCount, 0, 0, 1, "Yellow");
            if (title7EmptyCount > 0) Type("Titles skipped due to no title for episode 7: " + title7EmptyCount, 0, 0, 1, "Yellow");
            if (title8EmptyCount > 0) Type("Titles skipped due to no title for episode 8: " + title8EmptyCount, 0, 0, 1, "Yellow");
            if (title9EmptyCount > 0) Type("Titles skipped due to no title for episode 9: " + title9EmptyCount, 0, 0, 1, "Yellow");
            if (title10EmptyCount > 0) Type("Titles skipped due to no title for episode 10: " + title10EmptyCount, 0, 0, 1, "Yellow");

        } // End InsertMissingSeveralCombinedEpisodeData()

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
                        rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
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
                            }
                            else
                            {
                                string plot = response.data[0].overview.ToString().Trim(),
                                    name = response.data[0].episodeName.ToString().Trim();

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
                    Type(e.Message, 0, 0, 2, "DarkRed");
                }
            }
            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            DisplayMessage("success", "New Names inserted: ", 0);
            DisplayMessage("default", intNamesInsertedCount.ToString());
            DisplayMessage("success", "New Plots inserted: ", 0);
            DisplayMessage("default", intPlotsInsertedCount.ToString());
            DisplayMessage("warning", "Plots skipped due to no plot for the episode: ", 0);
            DisplayMessage("default", intPlotEmptyCount.ToString());

        } // End InsertMissingEpisodeData()

        private static void InsertEpisodesIntoRenameEpisodesSheet(IList<IList<Object>> data, Dictionary<string, int> sheetVariables)
        {
            foreach (var row in data)
            {
                try
                {
                    //string rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                    //string tvdbIdValue = row[Convert.ToInt16(sheetVariables["TVDB ID"])].ToString();
                    //plot1Data = row[Convert.ToInt16(sheetVariables["Episode 1 Plot"])].ToString();
                    //plot1ColumnNum = Convert.ToInt16(sheetVariables["Episode 1 Plot"]);
                    //plot2Data = row[Convert.ToInt16(sheetVariables["Episode 2 Plot"])].ToString();
                    //plot2ColumnNum = Convert.ToInt16(sheetVariables["Episode 2 Plot"]);
                    //episode1SeasonNum = row[Convert.ToInt16(sheetVariables["Episode 1 Season"])].ToString();
                    //episode1Num = row[Convert.ToInt16(sheetVariables["Episode 1 No."])].ToString();
                    //episode2SeasonNum = row[Convert.ToInt16(sheetVariables["Episode 2 Season"])].ToString();
                    //episode2Num = row[Convert.ToInt16(sheetVariables["Episode 2 No."])].ToString();
                    //showTitle = row[Convert.ToInt16(sheetVariables["Show Title"])].ToString();
                    //lockPLot1 = row[Convert.ToInt16(sheetVariables["Lock Plot 1"])].ToString();
                    //lockPlot2 = row[Convert.ToInt16(sheetVariables["Lock Plot 2"])].ToString();

                    //if (!tvdbIdValue.Equals("")) // If there is no id then the row is considered empty and should be skipped.
                    //{
                    //    var episode1Response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode1SeasonNum, episode1Num);
                    //    plot1Call = episode1Response.data[0].overview.ToString().Trim();

                    //    if (plot1Call.Equals(""))
                    //    {
                    //        DisplayMessage("default", "No plot available for ", 0);
                    //        DisplayMessage("warning", showTitle + " - S" + episode1SeasonNum + "E" + episode1Num);

                    //        intPlot1EmptyCount++;
                    //    }
                    //    else if (lockPLot1.ToUpper().Equals("X"))
                    //    {
                    //        DisplayMessage("default", "Plot locked for ", 0);
                    //        DisplayMessage("info", showTitle + " - S" + episode1SeasonNum + "E" + episode1Num);

                    //        intPlot1LockedCount++;
                    //    }
                    //    else if (plot1Call != plot1Data)
                    //    {
                    //        strCellToPutData = "Combined Episodes!" + ColumnNumToLetter(plot1ColumnNum) + rowNum;

                    //        WriteSingleCellToSheet(plot1Call, strCellToPutData);
                    //        DisplayMessage("default", "Plot updated for ", 0);
                    //        DisplayMessage("success", showTitle + " - S" + episode1SeasonNum + "E" + episode1Num + " - at " + strCellToPutData);

                    //        intPlot1InsertedCount++;
                    //    }

                    //    var episode2Response = TvdbApiCall.TvdbApi.GetTvEpisodeDetails(ref token, tvdbIdValue, episode2SeasonNum, episode2Num);
                    //    plot2Call = episode2Response.data[0].overview.ToString().Trim();

                    //    if (plot2Call.Equals(""))
                    //    {
                    //        DisplayMessage("default", "No plot available for ", 0);
                    //        DisplayMessage("warning", showTitle + " - S" + episode2SeasonNum + "E" + episode2Num);

                    //        intPlot2EmptyCount++;
                    //    }
                    //    else if (lockPlot2.ToUpper().Equals("X"))
                    //    {
                    //        DisplayMessage("default", "Plot locked for ", 0);
                    //        DisplayMessage("info", showTitle + " - S" + episode2SeasonNum + "E" + episode2Num);

                    //        intPlot2LockedCount++;
                    //    }
                    //    else if (plot2Call != plot2Data)
                    //    {
                    //        strCellToPutData = "Combined Episodes!" + ColumnNumToLetter(plot2ColumnNum) + rowNum;

                    //        WriteSingleCellToSheet(plot2Call, strCellToPutData);
                    //        DisplayMessage("default", "Plot saved for ", 0);
                    //        DisplayMessage("success", showTitle + " - S" + episode2SeasonNum + "E" + episode2Num + " - at " + strCellToPutData);

                    //        intPlot2InsertedCount++;
                    //    }
                    //}
                }
                catch (Exception e)
                {
                    Type(e.Message, 0, 0, 2, "DarkRed");
                }
            }
        }

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
                    rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
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
                        plot1Call = episode1Response.data[0].overview.ToString().Trim();

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
                        plot2Call = episode2Response.data[0].overview.ToString().Trim();

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
                    Type(e.Message, 0, 0, 2, "DarkRed");
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

        private static void WriteToSheetColumn(ArrayList videoFilesList, IList<IList<Object>> sheetData, string sheetName, int dataRowNum, int dataColumnNum)
        {
            var i = 0;
            DisplayMessage("info", "Adding " + videoFilesList.Count + " names to sheet '" + sheetName + "'");
            foreach (var row in sheetData)
            {
                string intRowNum = row[dataRowNum].ToString(),
                    columnToWriteTo = row[dataColumnNum].ToString();

                if (columnToWriteTo.Equals("") && i < videoFilesList.Count)
                {
                    string fileName = Path.GetFileNameWithoutExtension(videoFilesList[i].ToString());
                    string strCellToPutData = sheetName + "!" + ColumnNumToLetter(dataColumnNum) + int.Parse(intRowNum);
                    WriteSingleCellToSheet(fileName, strCellToPutData);
                    i++;
                    DisplayMessage("default", i + " of " + videoFilesList.Count, 0);
                    DisplayMessage("success", " - " + fileName, 0);
                    DisplayMessage("default", " - saved to row ", 0);
                    DisplayMessage("info", intRowNum);
                }
            }
        }

        private static void CreateFoldersAndMoveFiles(string directory, bool sort = false)
        {
            string[] fileEntries = Directory.GetFiles(directory);

            foreach (string sourceFile in fileEntries)
            {
                string extension = Path.GetExtension(sourceFile).ToLowerInvariant();
                string fileName = Path.GetFileName(sourceFile);
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);

                string baseName = fileNameWithoutExtension; // fallback for movie files
                string newFileName = fileName;

                if (extension == ".srt")
                {
                    // Match pattern: base + optional language + optional flags
                    Regex pattern = new Regex(
                        @"^(?<base>.*?)([_\.](eng|en))?([_\.\-](forced))?([_\.\-](sdh))?$",
                        RegexOptions.IgnoreCase
                    );

                    Match match = pattern.Match(fileNameWithoutExtension);

                    if (match.Success)
                    {
                        baseName = match.Groups["base"].Value.Trim();

                        bool hasForced = match.Groups[4].Success;
                        bool hasSdh = match.Groups[6].Success;

                        // Build normalized name: base.eng[.forced][.sdh].srt
                        newFileName = baseName + ".eng";
                        if (hasForced) newFileName += ".forced";
                        if (hasSdh) newFileName += ".sdh";
                        newFileName += ".srt";
                    }
                }
                else
                {
                    // Non-subtitle file: use base name for folder
                    baseName = Path.GetFileNameWithoutExtension(fileName);
                }

                // Build target movie folder (in this step, no letter folder yet)
                string movieFolder = Path.Combine(
                    Path.GetDirectoryName(sourceFile),
                    baseName
                );

                try
                {
                    Directory.CreateDirectory(movieFolder);
                }
                catch (Exception e)
                {
                    Type($"Something went wrong creating folder: {movieFolder}", 0, 0, 1, "Red");
                    Type(e.Message, 0, 0, 2, "DarkRed");
                    continue; // skip moving if folder failed
                }

                string destinationFile = Path.Combine(movieFolder, newFileName);

                try
                {
                    File.Move(sourceFile, destinationFile);
                    Type($"Moved: {newFileName}", 0, 0, 1, "Green");
                }
                catch (Exception e)
                {
                    Type($"Something went wrong moving file: {newFileName}", 0, 0, 1, "Red");
                    Type(e.Message, 0, 0, 2, "DarkRed");
                }
            }

            if (sort)
            {
                SortFoldersIntoSubFolders(directory);
            }
        }

        private static void SortFoldersIntoSubFolders(string directory)
        {
            string[] folderEntries = Directory.GetDirectories(directory);

            foreach (string sourcefolder in folderEntries)
            {
                string directoryName = Path.GetFileName(sourcefolder);

                // Don't sort the folder if it is named Kids Movies or if the folder is only one character long already.
                if (directoryName != "Kids Movies" && directoryName.Length > 1)
                {
                    // If the first four characters equals "THE " then we want to grab the fifth character.
                    string firstChar = directoryName.Substring(0, 4).ToUpper() == "THE " ? directoryName[4].ToString() : directoryName[0].ToString();

                    // If the firstChar is not a number 0-9 then set the subfolder to the firstChar.
                    string subFolder = !new Regex(@"^\d$").IsMatch(firstChar) ? firstChar : "#";

                    try
                    {
                        Directory.CreateDirectory(Path.Combine(directory, subFolder));
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong while creating the subfolder for: " + directoryName, 0, 0, 1, "Red");
                        Type(e.Message, 0, 0, 2, "DarkRed");
                    }

                    string targetDirectory = Path.Combine(directory, subFolder, directoryName);

                    try
                    {
                        Directory.Move(sourcefolder, targetDirectory);
                        Type("Moved into subfolder: ", 0, 0, 0, "Green");
                        Type(directoryName, 0, 0, 1);
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong while moving the folder: " + directoryName, 0, 0, 1, "Red");
                        Type(e.Message, 0, 0, 2, "DarkRed");
                    }
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
                            File.Move(f.FullName, Path.Combine(f.DirectoryName, f.Name.Substring(0, 30).Trim()) + f.Extension);
                            Type(f.Name, 0, 0, 0);
                            Type(" Trimmed", 0, 0, 1, "Green");
                        }
                        else
                        {
                            Type(f.Name, 0, 0, 0);
                            Type(" NOT Trimmed", 0, 0, 1, "Blue");
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong while trimming the file: " + f.Name, 0, 0, 1, "Red");
                        Type(e.Message, 0, 0, 2, "DarkRed");
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

        private static ArrayList GrabMovieFiles(string[] files, bool showMessage = true)
        {
            if (showMessage) Type("Grabbing just the video files... ", 0, 0, 0, "Yellow");

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
                if (showMessage) Type("DONE", 0, 0, 1, "Green");
                return videoFiles;
            }
            catch (Exception e)
            {
                Type("An error occured!", 0, 0, 1, "Red");
                Type(e.Message, 0, 0, 1, "DarkRed");
                throw;
            }

        } // End GrabMovieFiles()

        private static ArrayList GrabJpgFiles(string[] files)
        {
            Type("Grabbing just the JPG files... ", 0, 0, 0, "Yellow");
            ArrayList jpgFiles = new ArrayList();

            try
            {
                foreach (string file in files)
                {
                    if (file.ToUpper().Contains(".JPG"))
                    {
                        jpgFiles.Add(file);
                    }
                }
                Type("DONE", 0, 0, 1, "Green");
                return jpgFiles;
            }
            catch (Exception e)
            {
                Type("An error occured!", 0, 0, 1, "Red");
                Type(e.Message, 0, 0, 1, "DarkRed");
                throw;
            }

        } // End GrabJpgFiles()

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

        protected static long CalculateFolderSize(string directory)
        {
            long folderSize = 0;
            try
            {
                // Checks if the path is valid.
                if (!Directory.Exists(directory))
                    return folderSize; // If it doesn't exist, simply return 0.
                else
                {
                    try
                    {
                        // Calculate the size of each file in the directory and add it to the running total.
                        foreach (string file in Directory.GetFiles(directory))
                        {
                            if (File.Exists(file))
                            {
                                FileInfo finfo = new FileInfo(file);
                                folderSize += finfo.Length;
                            }
                        }

                        // Now recurse through each sub-directory.
                        foreach (string dir in Directory.GetDirectories(directory))
                            folderSize += CalculateFolderSize(dir);
                    }
                    catch (NotSupportedException e)
                    {
                        Console.WriteLine("Unable to calculate folder size: {0}", e.Message);
                    }
                }
            }
            catch (UnauthorizedAccessException e)
            {
                Console.WriteLine("Unable to calculate folder size: {0}", e.Message);
            }
            return folderSize;
        }

        static readonly string[] suffixes =
            { "Bytes", "KB", "MB", "GB", "TB", "PB" };

        /// <summary>
        /// Returns a long number as a
        /// </summary>
        /// <param name="byes">The size in bytes to be converted to MB | GB etc.</param>
        /// <param name="returnSuffix">A bool to return the number with MB/GB count.</param>
        /// <returns>String with MB/GB.</returns>
        public static string FormatSize(Int64 bytes, bool returnSuffix)
        {
            int counter = 0;
            decimal number = (decimal)bytes;
            while (Math.Round(number / 1024) >= 1)
            {
                number = number / 1024;
                counter++;
            }
            if (returnSuffix) return string.Format("{0:n1}{1}", number, suffixes[counter]);
            return string.Format("{0:n2}", number);
        }

        public static string ConvertBytesToGBytes(Int64 bytes)
        {
            double gigabytes = (double)bytes / (1024 * 1024 * 1024);
            return string.Format("{0:n2}", gigabytes);
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
            Type("We didn't find a column that we were looking for...", 0, 0, 1, "Red");
            foreach (KeyValuePair<string, int> variable in NotFoundColumns)
            {
                Type("Missing Column: '" + variable.Key.ToString() + "'", 0, 0, 1, "DarkRed");
            }
            Type("It's likely that the column we are looking for has changed names.", 0, 0, 2, "Red");
            Type("Press ENTER to exit the program.", 0, 0, 1, "DarkRed");
            Console.ReadLine();
            Environment.Exit(0);
        }

        static void AskForMenu()
        {
            Console.WriteLine();
            Type("Press any key to return to the menu...", 0, 0, 1, "Magenta");
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
                        CleanTitle = row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString();
                        status = row[Convert.ToInt16(sheetVariables[STATUS])].ToString();

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
                                    }
                                    else
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
                                }
                                else if (tmdbResponse.message != null)
                                {
                                    Type(CleanTitle + " | " + tmdbResponse.message, 0, 0, 1, "Red");
                                }
                                else
                                {
                                    Type("Something went wrong", 0, 0, 1, "Red");
                                }
                            }
                            else
                            {
                                intMoviesNotInListCount++;
                            }
                        }
                        else
                        {
                            intMoviesSkippedCount++;
                        }

                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong..." + e.Message, 0, 0, 1, "Red");
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
                }
                else
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

            foreach (var row in data)
            {
                if (row.Count > 20) // If it's an empty row then it should have less than this.
                {
                    try
                    {
                        string textToReplace = "", sourceDirectory = "";
                        var kids = row[Convert.ToInt16(sheetVariables["Kids"])].ToString();
                        var ownership = row[Convert.ToInt16(sheetVariables["Ownership"])].ToString();
                        var movieLetter = row[Convert.ToInt16(sheetVariables["Movie Letter"])].ToString();
                        var movieDirectory = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString();
                        var status = row[Convert.ToInt16(sheetVariables[STATUS])].ToString();
                        var directoryFound = false;
                        string[] potentialMovieFolderLocations = { "\\Melanie's Movies\\" + movieLetter + "\\", "\\Melanie's Kids Movies\\", "\\Movies\\" + movieLetter + "\\", "\\Kids Movies\\" };

                        if (!status.Equals("") && status[0].ToString().ToUpper() != "X") // If the first letter of status is an 'X' or empty then don't even look for the directory.
                        {
                            if (!Directory.Exists(movieDirectory))
                            {
                                // We need to figure out where the movie directory is pointing to so we know what we need to replace in the Directory string.
                                if (kids.ToUpper() == "X" && ownership.ToUpper() == "M")
                                {
                                    textToReplace = "\\Melanie's Kids Movies\\";
                                }
                                else if (kids.ToUpper() == "X")
                                {
                                    textToReplace = "\\Kids Movies\\";
                                }
                                else if (ownership.ToUpper() == "M")
                                {
                                    textToReplace = "\\Melanie's Movies\\" + movieLetter + "\\";
                                }
                                else
                                {
                                    textToReplace = "\\Movies\\" + movieLetter + "\\";
                                }

                                Type("We did not find: " + movieDirectory, 0, 0, 1, "Yellow");
                                Type("We will now look in the other directories to move it.", 0, 0, 1, "Yellow");

                                foreach (var location in potentialMovieFolderLocations)
                                {
                                    if (location != textToReplace)
                                    {
                                        sourceDirectory = movieDirectory.Replace(textToReplace, location);
                                        if (Directory.Exists(sourceDirectory))
                                        {
                                            directoryFound = true;
                                            DisplayMessage("info", "We found the movie here: ", 0);
                                            DisplayMessage("success", sourceDirectory);
                                            DisplayMessage("data", "We will now go move it");
                                            break;
                                        }
                                    }
                                }

                                if (directoryFound)
                                {
                                    MoveDirectory(sourceDirectory, movieDirectory);
                                    moviesMovedCount++;
                                }
                                else
                                {
                                    Type("We did not find the Directory in the other folders either.", 0, 0, 1, "Red");
                                    moviesNotFoundCount++;
                                }
                            }
                            else moviesNotMovedCount++;
                        }
                        else moviesSkippedCount++;

                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong..." + e.Message, 0, 0, 1, "Red");
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
            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            string message = type == 1 ? "We will now insert missing TMDB IDs" : "We will now insert AND overwrite TMDB IDs";
            Type(message + "...", 0, 0, 1, "Gray");
            int intTmdbIdDoneCount = 0, intTmdbIdCorrectedCount = 0, intTmdbIdSkippedCount = 0, intTmdbIdNotFoundCount = 0, intRowNum = 3;

            string tmdbIdValue = "", ImdbId = "", ImdbTitle = "", tmdbId = "", strCellToPutData = "";
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
                                    dynamic tmdbR = tmdbResponse.movie_results[0];

                                    if (tmdbR.id != null)
                                    {
                                        tmdbId = tmdbR.id.ToString();
                                        responseIsBroken = false;
                                    }
                                    else if (tmdbR.status_message != null)
                                    {
                                        Type(ImdbTitle + " | " + tmdbR.status_message, 0, 0, 1, "Red");
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
                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { tmdbId } }
                                    });

                                    Type("TMDB ID saved for: " + ImdbTitle, 0, 0, 1, "Green");
                                    intTmdbIdDoneCount++;
                                }
                                else
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

                                if (tmdbResponse != null && tmdbResponse is JObject tmdbObject)
                                {
                                    var movieResults = tmdbResponse["movie_results"] as JArray;

                                    if (movieResults != null && movieResults.Count > 0)
                                    {
                                        var movieResult = movieResults[0];

                                        if (movieResult["id"] != null)
                                        {
                                            tmdbId = movieResult["id"].ToString();
                                            responseIsBroken = false;
                                        }
                                        else if (movieResult["status_message"] != null)
                                        {
                                            Type($"{ImdbTitle} | {movieResult["status_message"]}", 0, 0, 1, "Red");
                                            tmdbId = "";
                                            responseIsBroken = false;
                                        }
                                        else
                                        {
                                            Thread.Sleep(5000); // Fallback logic
                                        }
                                    }
                                    else
                                    {
                                        tmdbId = "";
                                        responseIsBroken = false;
                                    }
                                }
                                else
                                {
                                    tmdbId = "";
                                    responseIsBroken = false;
                                }
                            } while (responseIsBroken);


                            strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbIdColumnNum) + intRowNum;

                            if (tmdbId != "")
                            {
                                if (tmdbIdValue.Equals("")) // If the ID is missing insert it.
                                {
                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { tmdbId } }
                                    });

                                    Type("TMDB ID saved for: ", 0, 0, 0, "Gray");
                                    Type(ImdbTitle, 0, 0, 1, "Green");
                                    intTmdbIdDoneCount++;

                                }
                                else if (tmdbIdValue != tmdbId) // Or if the new ID doesn't equal the old one overwrite it.
                                {
                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { tmdbId } }
                                    });

                                    Type("TMDB ID corrected for: ", 0, 0, 0, "Gray");
                                    Type(ImdbTitle, 0, 0, 0, "Blue");
                                    Type(" From: ", 0, 0, 0, "Gray");
                                    Type(tmdbIdValue, 0, 0, 0, "Blue");
                                    Type(" To: ", 0, 0, 0, "Gray");
                                    Type(tmdbId, 0, 0, 1, "Blue");
                                    intTmdbIdCorrectedCount++;
                                }
                                else // Else just skip it.
                                {
                                    Type("TMDB ID is correct for: ", 0, 0, 0, "Gray");
                                    Type(ImdbTitle, 0, 0, 1, "DarkGray");
                                    intTmdbIdSkippedCount++;
                                }
                            }
                            else
                            {
                                Type("No record found at TMDB for: ", 0, 0, 0, "Gray");
                                Type(ImdbTitle, 0, 0, 1, "Yellow");
                                intTmdbIdNotFoundCount++;
                            }
                        }
                        intRowNum++;
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong with " + ImdbTitle + " | " + e.Message, 0, 0, 1, "Red");
                    }

                }
            }
            var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
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

            string oldDirectory = "", newDirectory = "";

            foreach (var row in data)
            {
                if (row.Count > 20) // If it's an empty row then it should have less than this.
                {
                    try
                    {
                        oldDirectory = row[Convert.ToInt16(sheetVariables["Old Directory"])].ToString();
                        newDirectory = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString();

                        if (Directory.Exists(oldDirectory))
                        {
                            MoveDirectory(oldDirectory, newDirectory);
                            intDirectoriesMoviedCount++;
                        }
                        else
                        {
                            intDirectoriesSkippedCount++;
                        }

                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong moving " + oldDirectory + " to " + newDirectory + " | " + e.Message, 0, 0, 1, "Red");
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

        // =============================================================

        // =============================================================
        // 55m (Pre-step) - Scan Ready folders and rename movies early
        // =============================================================
        private static void Run55m_RenameMoviesFirst_ThenMove()
        {
            // 1) Scan Ready folders for Movies + TV shows
            DisplayMessage("info", "Step 1/4: Scanning Ready folders for Movies and TV shows...", 1);

            ArrayList movieFiles;
            int tvShowCount;
            int movieFolderCount;
            int movieLooseFileCount;

            ScanReadyFoldersForMoviesAndTv(out movieFiles, out tvShowCount, out movieFolderCount, out movieLooseFileCount);

            int totalMovies = movieFiles?.Count ?? 0;

            DisplayMessage("success", "Scan complete.", 1);
            DisplayMessage("info", "Movies found: ", 0); DisplayMessage("data", totalMovies.ToString());
            DisplayMessage("info", "  - movie files: ", 0); DisplayMessage("data", movieLooseFileCount.ToString());
            DisplayMessage("info", "  - movie folders: ", 0); DisplayMessage("data", movieFolderCount.ToString());
            DisplayMessage("info", "TV shows found: ", 0); DisplayMessage("data", tvShowCount.ToString(), 2);

            // 2) If there are movies, rename them FIRST (so prompts happen early)
            if (totalMovies > 0)
            {
                DisplayMessage("info", "Step 2/4: Renaming movies from Google Sheet (this may ask for help early)...", 2);

                try
                {
                    IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

                    // rootDirectory is only used as a fallback display/path; moviePaths already include full paths
                    string rootDirectory = "(Ready folders)";
                    RenameMoviesAndSrtsFromSheet(movieData, movieSheetVariables, movieFiles, rootDirectory);

                    DisplayMessage("success", "Movie rename step complete.", 2);
                }
                catch (Exception ex)
                {
                    DisplayMessage("error", "Movie rename step failed: ", 0);
                    DisplayMessage("data", ex.Message, 2);

                    // We still continue to moving so you're not dead in the water.
                }
            }
            else
            {
                DisplayMessage("warning", "Step 2/4: No movies found. Skipping movie rename step.", 2);
            }

            // 3) Move Movies and/or TV Shows to Working Areas (existing 55m logic)
            DisplayMessage("info", "Step 3/4: Moving Ready items to Working Areas...", 2);
            MoveReadyVideosToWorkingAreas();

            ProcessMoviesFromWorkingAreasThenCommit();
        }

        /// <summary>
        /// Scans all Ready roots (Raph/Leo/Mikey/Donny) and returns:
        /// - movieFiles: ALL movie video file paths (includes files inside movie folders)
        /// - tvShowCount: # of TV show folders detected
        /// - movieFolderCount: # of movie folders detected
        /// - movieLooseFileCount: # of movie files found directly under Ready/<Source>/
        /// </summary>
        private static void ScanReadyFoldersForMoviesAndTv(
            out ArrayList movieFiles,
            out int tvShowCount,
            out int movieFolderCount,
            out int movieLooseFileCount)
        {
            movieFiles = new ArrayList();
            tvShowCount = 0;
            movieFolderCount = 0;
            movieLooseFileCount = 0;

            string[] readyRoots =
            {
                @"\\RAPH\RaphOutput\Ready",
                @"\\LEO\LeoOutput\Ready",
                @"\\MIKEY\MikeyOutput\Ready",
                @"\\DONNY\DonnyOutput\Ready",
                @"C:\Users\Brandon\Downloads\checked"
            };

            foreach (var readyRoot in readyRoots)
            {
                try
                {
                    if (!Directory.Exists(readyRoot))
                    {
                        DisplayMessage("warning", "Ready folder not found: ", 0);
                        DisplayMessage("data", readyRoot);
                        continue;
                    }

                    var sourceFolders = Directory.GetDirectories(readyRoot);

                    foreach (var sourceFolder in sourceFolders)
                    {
                        // 1) Movie files directly under Ready/<Source>/
                        try
                        {
                            string[] topFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
                            var vids = GrabMovieFiles(topFiles, showMessage: false);

                            foreach (string v in vids)
                            {
                                if (!string.IsNullOrWhiteSpace(v) && File.Exists(v))
                                {
                                    movieFiles.Add(v);
                                    movieLooseFileCount++;
                                }
                            }
                        }
                        catch
                        {
                            // ignore per-source file scan errors (MoveReadyVideosToWorkingAreas will log details later)
                        }

                        // 2) Top-level folders under Ready/<Source>/
                        try
                        {
                            var topLevelFolders = Directory.GetDirectories(sourceFolder);

                            foreach (var folder in topLevelFolders)
                            {
                                var subfolders = Directory.GetDirectories(folder, "*", SearchOption.TopDirectoryOnly);

                                // TV show if it has subfolders (seasons)
                                if (subfolders.Length > 0)
                                {
                                    tvShowCount++;
                                    continue;
                                }

                                // Otherwise, treat as movie folder if it has video files
                                string[] files = Directory.GetFiles(folder, "*.*", SearchOption.TopDirectoryOnly);
                                var vids = GrabMovieFiles(files, showMessage: false);

                                if (vids == null || vids.Count == 0)
                                    continue;

                                movieFolderCount++;

                                foreach (string v in vids)
                                {
                                    if (!string.IsNullOrWhiteSpace(v) && File.Exists(v))
                                    {
                                        movieFiles.Add(v);
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // ignore; MoveReadyVideosToWorkingAreas will log later
                        }
                    }
                }
                catch
                {
                    // ignore; MoveReadyVideosToWorkingAreas will log later
                }
            }

            // De-dupe paths (same file could be discovered twice in edge cases)
            if (movieFiles.Count > 1)
            {
                var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var deduped = new ArrayList();

                foreach (string p in movieFiles)
                {
                    if (string.IsNullOrWhiteSpace(p)) continue;
                    if (unique.Add(p))
                        deduped.Add(p);
                }

                movieFiles = deduped;
            }
        }

        // 55m - Move Ready folders from recording PCs to Working Areas
        // =============================================================
        private static void MoveReadyVideosToWorkingAreas()
        {
            // Recording PCs (source)
            string[] readyRoots =
            {
                @"\\RAPH\RaphOutput\Ready",
                @"\\LEO\LeoOutput\Ready",
                @"\\MIKEY\MikeyOutput\Ready",
                @"\\DONNY\DonnyOutput\Ready",
                @"C:\Users\Brandon\Downloads\checked"
            };

            // TV/Movie Working Areas (dest)
            // Now driven from DriveMapping so you only update paths in ONE place.

            int movedMovieFiles = 0;
            int movedMovieFolderCount = 0;
            int movedShowCount = 0;
            int skippedCount = 0;
            int errorCount = 0;

            double MIN_FREE_PERCENT = 10.0;
            string otherFallbackRoot = DriveMapping.OtherWorkingAreaRoot;

            // -------------------------------------------------------
            // Local helpers (C# 7.3 safe)
            // -------------------------------------------------------
            string GetPcNameFromReadyRoot(string readyRoot)
            {
                if (string.IsNullOrWhiteSpace(readyRoot)) return "Unknown";
                var s = readyRoot.Trim();
                if (s.StartsWith(@"\\")) s = s.Substring(2);
                int idx = s.IndexOf('\\');
                if (idx <= 0) return readyRoot;
                return s.Substring(0, idx);
            }

            void PrintScanning(string readyRoot)
            {
                DisplayMessage("info", "Scanning Ready root: ", 0);
                DisplayMessage("data", readyRoot);
            }

            void PrintFoundSummary(string pcName, int movies, int tv)
            {
                DisplayMessage("info", "Found: ", 0);
                DisplayMessage("data", movies + " movies", 0);
                DisplayMessage("info", " and ", 0);
                DisplayMessage("data", tv + " TV shows", 0);
                DisplayMessage("info", " ready to move from ", 0);
                DisplayMessage("data", pcName);
            }

            void PrintTotalSummary(int movies, int tv, int items)
            {
                DisplayMessage("info", "TOTAL: ", 0);
                DisplayMessage("data", movies + " movies", 0);
                DisplayMessage("info", " and ", 0);
                DisplayMessage("data", tv + " TV shows", 0);
                DisplayMessage("info", " (", 0);
                DisplayMessage("data", items.ToString(), 0);
                DisplayMessage("info", " items)");
            }

            void PrintTvMoving(int index, int total, string name)
            {
                DisplayMessage("info", "[TV] Moving ", 0);
                DisplayMessage("data", index + " of " + total, 0);
                DisplayMessage("info", ": ", 0);
                DisplayMessage("data", name);
            }

            void PrintTvLine(string label, string value)
            {
                DisplayMessage("info", "[TV] " + label + ": ", 0);
                DisplayMessage("data", value);
            }

            void PrintMovieMoving(int index, int total, string name)
            {
                DisplayMessage("info", "[MOVIE] Moving ", 0);
                DisplayMessage("data", index + " of " + total, 0);
                DisplayMessage("info", ": ", 0);
                DisplayMessage("data", name);
            }

            void PrintMovieLine(string label, string value)
            {
                DisplayMessage("info", "[MOVIE] " + label + ": ", 0);
                DisplayMessage("data", value);
            }

            string ResolveMovieRoot(char letter, out string label)
            {
                return DriveMapping.GetMovieWorkingAreaRoot(letter, out label);
            }

            string ResolveTvRoot(char letter, out string label)
            {
                return DriveMapping.GetTvWorkingAreaRoot(letter, out label);
            }

            // -------------------------------------------------------
            // PASS 1: BUILD A PLAN
            // Requires:
            //   - enum MoveKind { MovieItem, MovieFolder, TvShow }
            //   - class PlannedMove { ReadyRoot, RecordSource, Kind, FolderPath, FolderName, MovieVideoFile, DisplayName }
            // -------------------------------------------------------
            var plan = new List<PlannedMove>();

            foreach (var readyRoot in readyRoots)
            {
                int foundMoviesThisPc = 0;
                int foundTvThisPc = 0;

                string pcName = GetPcNameFromReadyRoot(readyRoot);

                try
                {
                    PrintScanning(readyRoot);

                    if (!Directory.Exists(readyRoot))
                    {
                        DisplayMessage("warning", "Ready folder not found: ", 0);
                        DisplayMessage("data", readyRoot);
                        continue;
                    }

                    var sourceFolders = Directory.GetDirectories(readyRoot);

                    foreach (var sourceFolder in sourceFolders)
                    {
                        string recordSource = Path.GetFileName(sourceFolder);
                        if (string.IsNullOrWhiteSpace(recordSource))
                        {
                            skippedCount++;
                            continue;
                        }

                        // 1) MOVIES: files directly inside Ready/<Source>/
                        try
                        {
                            string[] topFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
                            var movieVideoFiles = GrabMovieFiles(topFiles, showMessage: false);

                            foreach (string movieVideoFile in movieVideoFiles)
                            {
                                if (string.IsNullOrWhiteSpace(movieVideoFile))
                                {
                                    skippedCount++;
                                    continue;
                                }

                                string bn = Path.GetFileNameWithoutExtension(movieVideoFile);
                                if (IsTrailerFileName(bn))
                                {
                                    // don't let trailers become their own "movie"
                                    continue;
                                }

                                plan.Add(new PlannedMove
                                {
                                    ReadyRoot = readyRoot,
                                    RecordSource = recordSource,
                                    Kind = MoveKind.MovieItem,
                                    MovieVideoFile = movieVideoFile
                                });

                                foundMoviesThisPc++;
                            }
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            DisplayMessage("error", "Error scanning movie files in: ", 0);
                            DisplayMessage("data", sourceFolder, 0);
                            DisplayMessage("error", " | ", 0);
                            DisplayMessage("data", ex.Message);
                        }

                        // 2) FOLDERS inside Ready/<Source>/
                        try
                        {
                            var topLevelFolders = Directory.GetDirectories(sourceFolder);

                            foreach (var folder in topLevelFolders)
                            {
                                string folderName = Path.GetFileName(folder);
                                if (string.IsNullOrWhiteSpace(folderName))
                                {
                                    skippedCount++;
                                    continue;
                                }

                                var subfolders = Directory.GetDirectories(folder, "*", SearchOption.TopDirectoryOnly);

                                // TV show if any subfolders
                                if (subfolders.Length > 0)
                                {
                                    plan.Add(new PlannedMove
                                    {
                                        ReadyRoot = readyRoot,
                                        RecordSource = recordSource,
                                        Kind = MoveKind.TvShow,
                                        FolderPath = folder,
                                        FolderName = folderName
                                    });

                                    foundTvThisPc++;
                                    continue;
                                }

                                // Otherwise: maybe a movie folder (has a movie file at top)
                                string[] files = Directory.GetFiles(folder, "*.*", SearchOption.TopDirectoryOnly);
                                var vids = GrabMovieFiles(files, showMessage: false);

                                if (vids == null || vids.Count == 0)
                                {
                                    skippedCount++;
                                    continue;
                                }

                                plan.Add(new PlannedMove
                                {
                                    ReadyRoot = readyRoot,
                                    RecordSource = recordSource,
                                    Kind = MoveKind.MovieFolder,
                                    FolderPath = folder,
                                    FolderName = folderName
                                });

                                foundMoviesThisPc++;
                            }
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            DisplayMessage("error", "Error scanning folders in: ", 0);
                            DisplayMessage("data", sourceFolder, 0);
                            DisplayMessage("error", " | ", 0);
                            DisplayMessage("data", ex.Message);
                        }
                    }

                    PrintFoundSummary(pcName, foundMoviesThisPc, foundTvThisPc);
                }
                catch (Exception ex)
                {
                    errorCount++;
                    DisplayMessage("error", "Unhandled error scanning Ready root: ", 0);
                    DisplayMessage("data", readyRoot, 0);
                    DisplayMessage("error", " | ", 0);
                    DisplayMessage("data", ex.Message);
                }
            }

            int totalMovies = plan.Count(p => p.Kind == MoveKind.MovieItem || p.Kind == MoveKind.MovieFolder);
            int totalTv = plan.Count(p => p.Kind == MoveKind.TvShow);
            int totalItems = plan.Count;

            PrintTotalSummary(totalMovies, totalTv, totalItems);
            Type(new string('-', 60), 0, 0, 1, "Blue");

            if (totalItems == 0)
            {
                DisplayMessage("success", "55m complete (nothing to move).");

                DisplayMessage("info", "Movies moved (files): ", 0);
                DisplayMessage("data", "0");

                DisplayMessage("info", "Movies moved (folders): ", 0);
                DisplayMessage("data", "0");

                DisplayMessage("info", "TV shows moved (folders): ", 0);
                DisplayMessage("data", "0");

                DisplayMessage("warning", "Skipped: ", 0);
                DisplayMessage("data", skippedCount.ToString());

                DisplayMessage("error", "Errors: ", 0);
                DisplayMessage("data", errorCount.ToString());

                return;
            }

            // -------------------------------------------------------
            // PASS 2: EXECUTE PLAN
            // -------------------------------------------------------
            int movieIndex = 0;
            int tvIndex = 0;

            foreach (var item in plan)
            {
                try
                {
                    if (item.Kind == MoveKind.TvShow)
                    {
                        tvIndex++;
                        PrintTvMoving(tvIndex, totalTv, item.DisplayName);

                        string folder = item.FolderPath;
                        string folderName = item.FolderName;
                        string recordSource = item.RecordSource;

                        PrintTvLine("Detected show", folderName);

                        char sortLetter = GetSortLetter(folderName);

                        string tvLabel;
                        string tvWorkingRoot = ResolveTvRoot(sortLetter, out tvLabel);

                        string chosenTvRoot, chosenTvLabel;
                        double? tvPct, fallbackPct;
                        bool rerouted;
                        bool canMove = ApplyLowSpaceReroute(
                            tvWorkingRoot,
                            tvLabel,
                            otherFallbackRoot,
                            recordSource,
                            MIN_FREE_PERCENT,
                            out chosenTvRoot,
                            out chosenTvLabel,
                            out tvPct,
                            out fallbackPct,
                            out rerouted);

                        if (!canMove)
                        {
                            skippedCount++;
                            DisplayMessage("error", "[TV] Not moved (low space). ", 0);
                            DisplayMessage("data", folderName, 0);
                            DisplayMessage("error", " | ", 0);
                            DisplayMessage("data", "Target=" + tvLabel + " (" + (tvPct.HasValue ? tvPct.Value.ToString("0.##") + "%" : "unknown") + "), " +
                                                   "Fallback=Other (" + (fallbackPct.HasValue ? fallbackPct.Value.ToString("0.##") + "%" : "unknown") + ")");
                            continue;
                        }

                        if (tvPct.HasValue && tvPct.Value < MIN_FREE_PERCENT && rerouted)
                        {
                            DisplayMessage("warning", "[TV] Low free space detected on target (", 0);
                            DisplayMessage("data", tvPct.Value.ToString("0.##") + "%", 0);
                            DisplayMessage("warning", "). Rerouting to Other (", 0);
                            DisplayMessage("data", fallbackPct.Value.ToString("0.##") + "%", 0);
                            DisplayMessage("warning", ").");
                        }
                        else if (!tvPct.HasValue)
                        {
                            DisplayMessage("warning", "[TV] Could not verify free space for destination share. Proceeding anyway.");
                        }

                        string destShowPath = Path.Combine(chosenTvRoot, recordSource, folderName);

                        // For TV shows, DO NOT create (2), (3), etc.
                        // Reuse the existing show folder if it already exists.
                        EnsureDirectory(destShowPath);

                        PrintTvLine("Target drive", chosenTvLabel);
                        PrintTvLine("Source", folder);
                        PrintTvLine("Destination", destShowPath);

                        var subfolders = Directory.GetDirectories(folder, "*", SearchOption.TopDirectoryOnly);
                        var seasonNames = subfolders
                            .Select(Path.GetFileName)
                            .Where(n => !string.IsNullOrWhiteSpace(n))
                            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                            .ToList();

                        if (seasonNames.Count > 0)
                        {
                            foreach (var s in seasonNames)
                            {
                                string seasonDest = Path.Combine(destShowPath, s);

                                if (Directory.Exists(seasonDest))
                                {
                                    DisplayMessage("warning", "[TV] Season folder already exists, will merge into ", 0);
                                    DisplayMessage("data", "'" + seasonDest + "'");
                                }
                                else
                                {
                                    DisplayMessage("info", "[TV] Will create season folder ", 0);
                                    DisplayMessage("data", "'" + s + "'", 0);
                                    DisplayMessage("info", " in ", 0);
                                    DisplayMessage("data", "'" + destShowPath + "'");
                                }
                            }
                        }

                        MoveDirectoryRobust(folder, destShowPath);
                        movedShowCount++;
                    }
                    else if (item.Kind == MoveKind.MovieFolder)
                    {
                        movieIndex++;
                        PrintMovieMoving(movieIndex, totalMovies, item.DisplayName);

                        string folder = item.FolderPath;
                        string folderName = item.FolderName;
                        string recordSource = item.RecordSource;

                        char sortLetter = GetSortLetter(folderName);

                        string movieLabel;
                        string movieWorkingRoot = ResolveMovieRoot(sortLetter, out movieLabel);

                        string chosenMovieRoot, chosenMovieLabel;
                        double? moviePct, fallbackPct;
                        bool rerouted;
                        bool canMove = ApplyLowSpaceReroute(
                            movieWorkingRoot,
                            movieLabel,
                            otherFallbackRoot,
                            recordSource,
                            MIN_FREE_PERCENT,
                            out chosenMovieRoot,
                            out chosenMovieLabel,
                            out moviePct,
                            out fallbackPct,
                            out rerouted);

                        if (!canMove)
                        {
                            skippedCount++;
                            DisplayMessage("error", "[MOVIE] Folder not moved (low space). ", 0);
                            DisplayMessage("data", folderName, 0);
                            DisplayMessage("error", " | ", 0);
                            DisplayMessage("data", "Target=" + movieLabel + " (" + (moviePct.HasValue ? moviePct.Value.ToString("0.##") + "%" : "unknown") + "), " +
                                                   "Fallback=Other (" + (fallbackPct.HasValue ? fallbackPct.Value.ToString("0.##") + "%" : "unknown") + ")");
                            continue;
                        }

                        if (moviePct.HasValue && moviePct.Value < MIN_FREE_PERCENT && rerouted)
                        {
                            DisplayMessage("warning", "[MOVIE] Low free space detected on target (", 0);
                            DisplayMessage("data", moviePct.Value.ToString("0.##") + "%", 0);
                            DisplayMessage("warning", "). Rerouting to Other (", 0);
                            DisplayMessage("data", fallbackPct.Value.ToString("0.##") + "%", 0);
                            DisplayMessage("warning", ").");
                        }
                        else if (!moviePct.HasValue)
                        {
                            DisplayMessage("warning", "[MOVIE] Could not verify free space for destination share. Proceeding anyway.");
                        }

                        string destMovieFolder = Path.Combine(chosenMovieRoot, recordSource, folderName);

                        destMovieFolder = GetUniqueDirectoryPath(destMovieFolder);

                        string destParent = Path.GetDirectoryName(destMovieFolder);
                        if (!string.IsNullOrWhiteSpace(destParent))
                            EnsureDirectory(destParent);

                        PrintMovieLine("Target drive", chosenMovieLabel);
                        PrintMovieLine("Source", folder);
                        PrintMovieLine("Destination", destMovieFolder);

                        MoveDirectoryRobust(folder, destMovieFolder);
                        movedMovieFolderCount++;
                    }
                    else // MovieItem (video file + sidecars)
                    {
                        movieIndex++;
                        string movieVideoFile = item.MovieVideoFile;
                        string recordSource = item.RecordSource;

                        string title = item.DisplayName;
                        PrintMovieMoving(movieIndex, totalMovies, title);

                        string sourceFolder = Path.GetDirectoryName(movieVideoFile);
                        if (string.IsNullOrWhiteSpace(sourceFolder) || !Directory.Exists(sourceFolder))
                        {
                            skippedCount++;
                            DisplayMessage("warning", "Skipped (missing source folder): ", 0);
                            DisplayMessage("data", movieVideoFile);
                            continue;
                        }

                        char sortLetter = GetSortLetter(title);

                        string movieLabel;
                        string movieWorkingRoot = ResolveMovieRoot(sortLetter, out movieLabel);

                        string chosenMovieRoot, chosenMovieLabel;
                        double? moviePct, fallbackPct;
                        bool rerouted;
                        bool canMove = ApplyLowSpaceReroute(
                            movieWorkingRoot,
                            movieLabel,
                            otherFallbackRoot,
                            recordSource,
                            MIN_FREE_PERCENT,
                            out chosenMovieRoot,
                            out chosenMovieLabel,
                            out moviePct,
                            out fallbackPct,
                            out rerouted);

                        if (!canMove)
                        {
                            skippedCount++;
                            DisplayMessage("error", "[MOVIE] File not moved (low space). ", 0);
                            DisplayMessage("data", title, 0);
                            DisplayMessage("error", " | ", 0);
                            DisplayMessage("data", "Target=" + movieLabel + " (" + (moviePct.HasValue ? moviePct.Value.ToString("0.##") + "%" : "unknown") + "), " +
                                                   "Fallback=Other (" + (fallbackPct.HasValue ? fallbackPct.Value.ToString("0.##") + "%" : "unknown") + ")");
                            continue;
                        }

                        if (moviePct.HasValue && moviePct.Value < MIN_FREE_PERCENT && rerouted)
                        {
                            DisplayMessage("warning", "[MOVIE] Low free space detected on target (", 0);
                            DisplayMessage("data", moviePct.Value.ToString("0.##") + "%", 0);
                            DisplayMessage("warning", "). Rerouting to Other (", 0);
                            DisplayMessage("data", fallbackPct.Value.ToString("0.##") + "%", 0);
                            DisplayMessage("warning", ").");
                        }
                        else if (!moviePct.HasValue)
                        {
                            DisplayMessage("warning", "[MOVIE] Could not verify free space for destination share. Proceeding anyway.");
                        }

                        // Create: <WorkingRoot>\<Source>\<Title>\
                        string destMovieFolder = Path.Combine(chosenMovieRoot, recordSource, title);
                        destMovieFolder = GetUniqueDirectoryPath(destMovieFolder);
                        EnsureDirectory(destMovieFolder);

                        PrintMovieLine("Destination folder", destMovieFolder);

                        string baseName = Path.GetFileNameWithoutExtension(movieVideoFile);

                        // Grab video + any sidecars that share the same base name (subs, nfo, etc.)
                        var sidecarsList = new List<string>();

                        // Main file + normal sidecars (subs/nfo/etc.)
                        sidecarsList.AddRange(Directory.GetFiles(sourceFolder, baseName + ".*", SearchOption.TopDirectoryOnly));

                        // Trailer variations that won't match baseName + ".*"
                        sidecarsList.AddRange(Directory.GetFiles(sourceFolder, baseName + " - trailer*.*", SearchOption.TopDirectoryOnly));
                        sidecarsList.AddRange(Directory.GetFiles(sourceFolder, baseName + "-trailer*.*", SearchOption.TopDirectoryOnly));
                        sidecarsList.AddRange(Directory.GetFiles(sourceFolder, baseName + " - teaser*.*", SearchOption.TopDirectoryOnly));
                        sidecarsList.AddRange(Directory.GetFiles(sourceFolder, baseName + "-teaser*.*", SearchOption.TopDirectoryOnly));
                        sidecarsList.AddRange(Directory.GetFiles(sourceFolder, baseName + " - preview*.*", SearchOption.TopDirectoryOnly));
                        sidecarsList.AddRange(Directory.GetFiles(sourceFolder, baseName + "-preview*.*", SearchOption.TopDirectoryOnly));

                        string[] sidecars = sidecarsList
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToArray();

                        foreach (var f in sidecars)
                        {
                            try
                            {
                                string fileName = Path.GetFileName(f);

                                string suffix;

                                if (fileName.Length > baseName.Length)
                                {
                                    suffix = fileName.Substring(baseName.Length);
                                }
                                else
                                {
                                    suffix = Path.GetExtension(fileName);
                                }

                                // Normalize trailer/teaser/preview naming to Plex format
                                string lowerSuffix = suffix.ToLowerInvariant();

                                if (lowerSuffix.Contains("trailer"))
                                {
                                    suffix = "-trailer" + Path.GetExtension(fileName);
                                }
                                else if (lowerSuffix.Contains("teaser"))
                                {
                                    suffix = "-teaser" + Path.GetExtension(fileName);
                                }
                                else if (lowerSuffix.Contains("preview"))
                                {
                                    suffix = "-preview" + Path.GetExtension(fileName);
                                }

                                // Force the moved filename to start with the folder title
                                string destFileName = title + suffix;

                                DisplayMessage("info", "[MOVIE] Moving file: ", 0);
                                DisplayMessage("data", fileName, 0);
                                DisplayMessage("info", " -> ", 0);
                                DisplayMessage("data", destFileName);

                                string destFile = GetUniqueFilePath(Path.Combine(destMovieFolder, destFileName));
                                MoveFileRobust(f, destFile);
                                movedMovieFiles++;
                            }
                            catch (Exception ex)
                            {
                                errorCount++;
                                DisplayMessage("error", "Failed moving file: ", 0);
                                DisplayMessage("data", f, 0);
                                DisplayMessage("error", " | ", 0);
                                DisplayMessage("data", ex.Message);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    errorCount++;
                    DisplayMessage("error", "Failed moving item: ", 0);
                    DisplayMessage("data", item.DisplayName, 0);
                    DisplayMessage("error", " | ", 0);
                    DisplayMessage("data", ex.Message);
                }
            }

            // -------------------------------------------------------
            // Final summary
            // -------------------------------------------------------
            DisplayMessage("success", "55m complete.");

            DisplayMessage("info", "Movies moved (files): ", 0);
            DisplayMessage("data", movedMovieFiles.ToString());

            DisplayMessage("info", "Movies moved (folders): ", 0);
            DisplayMessage("data", movedMovieFolderCount.ToString());

            DisplayMessage("info", "TV shows moved (folders): ", 0);
            DisplayMessage("data", movedShowCount.ToString());

            DisplayMessage("warning", "Skipped: ", 0);
            DisplayMessage("data", skippedCount.ToString());

            DisplayMessage("error", "Errors: ", 0);
            DisplayMessage("data", errorCount.ToString());
        }

        // =============================================================
        // 55m (Part 2) - Process staged movies in Working Areas,
        // then commit to Movies library + run ONE Emby scan
        // =============================================================
        private static void ProcessMoviesFromWorkingAreasThenCommit()
        {
            // Movie Working Areas (sources for this phase)
            string[] movieWorkingRoots = DriveMapping.GetMovieWorkingAreaRoots();

            int processed = 0;
            int skipped = 0;
            int errors = 0;

            var committedDestinations = new List<string>();

            // Collect staged folders first (so we can do “metadata pass” and “sheet+commit pass”)
            var stagedMovieFolders = new List<(string MovieFolder, string RecordSource)>();

            foreach (var workingRoot in movieWorkingRoots)
            {
                if (!Directory.Exists(workingRoot))
                    continue;

                // Working Area\<RecordSource>\MovieFolder...
                var recordSourceFolders = Directory.GetDirectories(workingRoot);

                foreach (var recordSourceFolder in recordSourceFolders)
                {
                    string recordSource = Path.GetFileName(recordSourceFolder);

                    try
                    {
                        // Movies are now ALWAYS:
                        // Working Area\<Source>\<Movie Title>\ (folder containing mkv/mp4 + sidecars)
                        var movieFolders = Directory.GetDirectories(recordSourceFolder, "*", SearchOption.TopDirectoryOnly);

                        foreach (var movieFolder in movieFolders)
                        {
                            // Sanity: only treat it as a staged movie folder if it contains a video file
                            bool hasVideo = Directory.GetFiles(movieFolder, "*.*", SearchOption.TopDirectoryOnly)
                                .Any(f =>
                                    f.EndsWith(".mp4", StringComparison.OrdinalIgnoreCase) ||
                                    f.EndsWith(".mkv", StringComparison.OrdinalIgnoreCase) ||
                                    f.EndsWith(".avi", StringComparison.OrdinalIgnoreCase) ||
                                    f.EndsWith(".webm", StringComparison.OrdinalIgnoreCase));

                            if (!hasVideo)
                                continue;

                            stagedMovieFolders.Add((movieFolder, recordSource));
                        }
                    }
                    catch (Exception ex)
                    {
                        errors++;
                        DisplayMessage("error", "Error enumerating staged movies in: ", 0);
                        DisplayMessage("data", recordSourceFolder, 0);
                        DisplayMessage("error", " | ", 0);
                        DisplayMessage("data", ex.Message);
                    }
                }
            }

            // De-dupe (same folder can be found twice if it had both top file + folder logic)
            stagedMovieFolders = stagedMovieFolders
                .GroupBy(x => x.MovieFolder, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();

            if (stagedMovieFolders.Count == 0)
            {
                DisplayMessage("warning", "No staged movie folders found in Working Areas.");
                return;
            }

            DisplayMessage("info",$"--- Found {stagedMovieFolders.Count} staged movie folder(s): {DateTime.Now} ---");
            foreach (var f in stagedMovieFolders)
                DisplayMessage("info", $"{f.RecordSource} | {f.MovieFolder}",2);

            // PASS 1: Remove metadata
            DisplayMessage("info", $"--- Removing metadata from {stagedMovieFolders.Count} folder(s): {DateTime.Now} ---");
            foreach (var item in stagedMovieFolders)
            {
                try
                {
                    RemoveMetadataInFolder(item.MovieFolder);
                }
                catch (Exception ex)
                {
                    errors++;
                    DisplayMessage("harderror", $"ERROR RemoveMetadataInFolder: {item.MovieFolder} | {ex.Message}");
                }
            }
            DisplayMessage("info", $"--- Removing metadata finished: {DateTime.Now} ---", 2);

            // PASS 2: Update Google Sheet + commit move into library
            DisplayMessage("info", $"--- Now going to update Google Sheet + commit staged movies: {DateTime.Now} ---");

            MoviesSheetBatchContext sheetContext = null;

            try
            {
                sheetContext = BuildMoviesSheetBatchContext();
            }
            catch (Exception ex)
            {
                errors++;
                DisplayMessage("harderror", $"ERROR BuildMoviesSheetBatchContext: {ex.Message}");
                return;
            }

            foreach (var item in stagedMovieFolders)
            {
                try
                {
                    bool ok = ProcessSingleStagedMovieFolderAndCommit(
                        sheetContext: sheetContext,
                        stagedMovieFolder: item.MovieFolder,
                        recordSource: item.RecordSource,
                        committedDestinations: committedDestinations
                    );

                    if (ok) processed++;
                    else skipped++;
                }
                catch (Exception ex)
                {
                    errors++;
                    DisplayMessage("harderror", $"ERROR ProcessSingleStagedMovieFolderAndCommit: {item.MovieFolder} | {ex.Message}");
                }
            }

            try
            {
                FlushMoviesSheetBatchUpdates(sheetContext);
            }
            catch (Exception ex)
            {
                errors++;
                DisplayMessage("harderror", $"ERROR FlushMoviesSheetBatchUpdates: {ex.Message}");
            }

            DisplayMessage("info", $"--- Updating Google Sheet + commit finished: {DateTime.Now} ---", 2);

            // =============================================================
            // Pre-Emby: Update Google Sheet + NFO/trailers/scoring tasks
            // =============================================================
            try
            {
                DisplayMessage("info", $"--- Pre-Emby tasks starting: {DateTime.Now} ---");

                // Update sheet resolution choice
                // getVideoResolutionChoice();
                IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);
                FillInVideoResolution(movieData, movieSheetVariables, false);

                // Add folder sizes to sheet
                //addSizeOfMovieDirectories();
                InsertMovieDirectorySizesIntoMoviesSheet(movieData, movieSheetVariables, false);

                // Set / update file type choice
                //fileTypeChoice();
                FillInFileTypes(movieData, movieSheetVariables, false);

                // Download missing trailers
                //downloadMovieTrailersChoice();
                DownloadMovieTrailers(movieData, movieSheetVariables);

                // Run SRT scoring
                //srtScoreRunnerChoice();
                // --- duplicate auth code for now (same pattern as WriteSingleCellToSheet / BulkWriteToSheet) ---
                UserCredential credential;

                using (var stream =
                    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        SCOPES,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                }

                SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Google-SheetsSample/0.1",  // same as your other helpers
                });
                // --- end local SheetsService setup ---
                SubtitleScoreRunner.Run(sheetsService, SPREADSHEET_ID, "Movies");

                // Find & create missing movie NFO files
                //missingMovieNfoFilesChoice();
                CreateNfoFiles(movieData, movieSheetVariables, 3);

                DisplayMessage("success", $"--- Pre-Emby tasks finished: {DateTime.Now} ---");
            }
            catch (Exception ex)
            {
                errors++;
                DisplayMessage("harderror", $"ERROR Pre-Emby tasks: {ex.Message}");
            }

            // =============================================================
            // ONE Emby refresh at the end
            // =============================================================
            try
            {
                DisplayMessage("info", $"--- Triggering Emby library refresh: {DateTime.Now} ---");
                RefreshEmbyLibrary();   // existing method
                DisplayMessage("success", $"--- Emby library refresh step finished: {DateTime.Now} ---");
            }
            catch (Exception ex)
            {
                errors++;
                DisplayMessage("harderror", $"ERROR RefreshEmbyLibrary: {ex.Message}");
            }

            // =============================================================
            // Post-Emby: File health runner
            // =============================================================
            try
            {
                DisplayMessage("info", $"--- Post-Emby file health starting: {DateTime.Now} ---");
                //fileHealthRunnerChoice();
                SheetsService sheetsService = CreateSheetsService();

                VideoHealthSheetRunner.Run(sheetsService, SPREADSHEET_ID, "Movies");
                DisplayMessage("success", $"--- Post-Emby file health finished: {DateTime.Now} ---");
            }
            catch (Exception ex)
            {
                errors++;
                DisplayMessage("harderror", $"ERROR fileHealthRunnerChoice: {ex.Message}");
            }


            DisplayMessage("success", "Staged movies processed: " + processed);
            DisplayMessage("warning", "Skipped: " + skipped);
            DisplayMessage("error", "Errors: " + errors);
        }

        private static bool IsTrailerFileName(string baseName)
        {
            if (string.IsNullOrWhiteSpace(baseName)) return false;

            string lower = baseName.ToLowerInvariant();

            // filename markers
            return lower.Contains("trailer")
                || lower.Contains("teaser")
                || lower.Contains("preview");
        }

        private static string StripTrailerTokens(string baseName)
        {
            if (string.IsNullOrWhiteSpace(baseName)) return baseName;

            string s = baseName;

            s = Regex.Replace(
                s,
                @"\s*-\s*(trailer|teaser|preview)(?:[\s._-]*[A-Za-z0-9]+)?\s*$",
                "",
                RegexOptions.IgnoreCase);

            s = Regex.Replace(
                s,
                @"\s+(official\s+)?(trailer|teaser|preview)(?:[\s._-]*[A-Za-z0-9]+)?\s*$",
                "",
                RegexOptions.IgnoreCase);

            // Remove common patterns (best-effort)
            string[] tokens =
            {
                " - trailer", "- trailer", " trailer",
                " - teaser", "- teaser", " teaser",
                " official trailer", "official trailer",
                " official teaser", "official teaser",
                " - preview", "- preview", " preview"
            };

            string lower = s.ToLowerInvariant();

            foreach (var t in tokens)
            {
                int idx = lower.IndexOf(t, StringComparison.Ordinal);
                if (idx >= 0)
                {
                    s = s.Remove(idx, t.Length);
                    lower = s.ToLowerInvariant();
                }
            }

            return s.Trim().Trim('-', '_', ' ', '.');
        }

        private static string BuildTargetVideoName(string baseName, string extension, bool isTrailer)
        {
            // baseName: "500 Days of Summer (2009)"
            // trailer: "500 Days of Summer (2009)-trailer.mkv"
            return isTrailer
                ? $"{baseName}-trailer{extension}"
                : $"{baseName}{extension}";
        }

        private static string BuildTargetSrtName(string baseName, string lang, bool isTrailer)
        {
            // "500 Days of Summer (2009).eng.srt" or "...-trailer.eng.srt"
            string trailerPart = isTrailer ? "-trailer" : "";
            if (string.IsNullOrWhiteSpace(lang))
                return $"{baseName}{trailerPart}.srt";

            return $"{baseName}{trailerPart}.{lang}.srt";
        }

        private static void QuarantineFileAndSidecars(string filePath, string reason)
        {
            try
            {
                var parentDir = Path.GetDirectoryName(filePath) ?? "";
                var folderName = "Bad - " + SanitizeFolderName(reason);
                var destDir = Path.Combine(parentDir, folderName);
                Directory.CreateDirectory(destDir);

                // Base name without extension (used to find sidecars)
                var baseName = Path.GetFileNameWithoutExtension(filePath);

                // Collect candidates to move:
                //  1) the original bad file
                //  2) any .srt files whose filename starts with the same base (case-insensitive)
                var toMove = new List<string> { filePath };

                var sidecars = Directory.GetFiles(parentDir, "*.srt", SearchOption.TopDirectoryOnly)
                    .Where(p => Path.GetFileName(p).StartsWith(baseName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                toMove.AddRange(sidecars);

                foreach (var src in toMove.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    try
                    {
                        var fileName = Path.GetFileName(src);
                        var destPath = Path.Combine(destDir, fileName);

                        // If collision, append timestamp
                        if (File.Exists(destPath))
                        {
                            var n = Path.GetFileNameWithoutExtension(fileName);
                            var ext = Path.GetExtension(fileName);
                            destPath = Path.Combine(destDir, $"{n}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");
                        }

                        File.Move(src, destPath);
                        DisplayMessage("info", $"[QUARANTINED] {src} -> {destPath}");
                    }
                    catch (Exception exFile)
                    {
                        DisplayMessage("info", $"[QUARANTINE-FAIL] Could not move {src}: {exFile.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayMessage("info", $"[QUARANTINE-FAIL] Could not prepare quarantine folder for {filePath}: {ex.Message}");
            }
        }

        private static string SanitizeFolderName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Unknown Error";
            // Remove/replace characters invalid for folder names
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (var ch in name)
                sb.Append(invalid.Contains(ch) ? '_' : ch);

            // Trim and collapse spaces
            var cleaned = sb.ToString().Trim();
            // Keep it reasonable in length
            if (cleaned.Length > 80) cleaned = cleaned.Substring(0, 80).Trim();
            return cleaned.Length == 0 ? "Unknown Error" : cleaned;
        }

        private static void RefreshEmbyLibrary()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);
                    string embyApiKey = GetEmbyApiKey();
                    string embyMoviesLibraryId = GetEmbyMoviesLibraryId();
                    var baseUrl = GetEmbyServerUrl()?.TrimEnd('/') ?? "";
                    if (string.IsNullOrWhiteSpace(baseUrl) ||
                        string.IsNullOrWhiteSpace(embyApiKey) ||
                        string.IsNullOrWhiteSpace(embyMoviesLibraryId))
                    {
                        DisplayMessage("info", "Emby refresh skipped: EMBY_SERVER_URL, EMBY_API_KEY, or EMBY_MOVIES_LIBRARY_ID not configured.");
                        return;
                    }

                    // Refresh just the Movies library (ID = 3)
                    var url =
                        $"{baseUrl}/emby/Items/{embyMoviesLibraryId}/Refresh" +
                        $"?Recursive=true" +
                        $"&ImageRefreshMode=FullRefresh" +
                        $"&MetadataRefreshMode=Default" +
                        $"&ReplaceAllMetadata=false" +
                        $"&ReplaceAllImages=false" +
                        $"&api_key={embyApiKey}";

                    DisplayMessage("info", $"Sending Emby Movies library refresh request for library {embyMoviesLibraryId}.");

                    var response = client.PostAsync(url, null).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        DisplayMessage("info", "✅ Emby Movies library refresh triggered successfully.");
                    }
                    else
                    {
                        DisplayMessage("error", $"❌ Emby refresh failed. Status: {(int)response.StatusCode} {response.ReasonPhrase}");
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayMessage("harderror", $"❌ Error triggering Emby Movies library refresh: {ex.Message}");
            }
        }

        private static bool ProcessSingleStagedMovieFolderAndCommit(
            MoviesSheetBatchContext sheetContext,
            string stagedMovieFolder,
            string recordSource,
            List<string> committedDestinations,
            string moviesLibraryRoot_1 = @"\\Quadplex\#-d\Movies",
            string moviesLibraryRoot_2 = @"\\Quadplex\e-m\Movies",
            string moviesLibraryRoot_3 = @"\\Quadplex\n-z\Movies")
        {
            try
            {
                if (!Directory.Exists(stagedMovieFolder))
                    return false;

                // Find the primary video file in the staged folder (largest non-trailer)
                var videoFiles = Directory.GetFiles(stagedMovieFolder, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(f =>
                        f.EndsWith(".mp4", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".mkv", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".avi", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".m4v", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".webm", StringComparison.OrdinalIgnoreCase))
                    .Where(f => f.IndexOf("-trailer", StringComparison.OrdinalIgnoreCase) < 0)
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(fi => fi.Length)
                    .ToList();

                if (videoFiles.Count == 0)
                    return false;

                string cleanTitle = Path.GetFileNameWithoutExtension(videoFiles[0].FullName);

                return UpdateMoviesSheetRowAndCommitMove(
                    sheetContext,
                    cleanTitle,
                    recordSource,
                    stagedMovieFolder,
                    moviesLibraryRoot_1,
                    moviesLibraryRoot_2,
                    moviesLibraryRoot_3,
                    committedDestinations
                );
            }
            catch (Exception ex)
            {
                DisplayMessage("error", "Failed staged movie FOLDER: ", 0);
                DisplayMessage("data", stagedMovieFolder, 0);
                DisplayMessage("error", " | ", 0);
                DisplayMessage("data", ex.Message);
                return false;
            }
        }

        private static bool UpdateMoviesSheetRowAndCommitMove(
            MoviesSheetBatchContext sheetContext,
            string cleanTitle,
            string recordedSource,
            string stagedMovieFolder,
            string moviesLibraryRoot_1,
            string moviesLibraryRoot_2,
            string moviesLibraryRoot_3,
            List<string> committedDestinations)
        {
            try
            {
                var data = sheetContext.DataRows;
                var colIndex = sheetContext.ColIndex;

                int cleanTitleIdx = colIndex[CLEAN_TITLE];

                // Find matching row by Clean Title
                int matchingRowIndex = -1;
                for (int i = 0; i < data.Count; i++)
                {
                    var row = data[i];
                    if (cleanTitleIdx < row.Count)
                    {
                        var cell = row[cleanTitleIdx]?.ToString()?.Trim();
                        if (!string.IsNullOrWhiteSpace(cell) &&
                            string.Equals(cell, cleanTitle, StringComparison.OrdinalIgnoreCase))
                        {
                            matchingRowIndex = i;
                            break;
                        }
                    }
                }

                if (matchingRowIndex == -1)
                {
                    DisplayMessage("warning", $"No matching Movies row found for Clean Title: {cleanTitle}");
                    return false;
                }

                var rowToUpdate = data[matchingRowIndex];

                // Sheet row number: MOVIES_DATA_RANGE starts at row 3
                int sheetRowNumber = matchingRowIndex + STARTING_ROW_NUMBER;

                bool stagedFromOther = IsPathUnderRoot(stagedMovieFolder, DriveMapping.OtherWorkingAreaRoot);

                // Pull Directory from sheet (preferred unless the movie was rerouted to Other)
                string existingDirectoryPath = "";
                int directoryIdx = -1;

                if (colIndex.TryGetValue(DIRECTORY, out directoryIdx) && directoryIdx < rowToUpdate.Count)
                {
                    existingDirectoryPath = rowToUpdate[directoryIdx]?.ToString()?.Trim() ?? "";
                }

                string directoryPath;
                if (stagedFromOther)
                {
                    directoryPath = Path.Combine(DriveMapping.OtherMoviesLibraryRoot, cleanTitle);
                }
                else
                {
                    directoryPath = existingDirectoryPath;

                    // Fallback: compute destination if directory missing
                    if (string.IsNullOrWhiteSpace(directoryPath))
                    {
                        string finalDestRoot = PickMoviesLibraryRootFromTitle(cleanTitle, moviesLibraryRoot_1, moviesLibraryRoot_2, moviesLibraryRoot_3);
                        directoryPath = Path.Combine(finalDestRoot, cleanTitle);
                    }
                }

                string overflowReplacementSourceFolder = "";
                if (stagedFromOther)
                {
                    overflowReplacementSourceFolder = FindOriginalMovieFolderForOverflowReplacement(
                        existingDirectoryPath,
                        cleanTitle,
                        directoryPath,
                        moviesLibraryRoot_1,
                        moviesLibraryRoot_2,
                        moviesLibraryRoot_3);

                    if (!string.IsNullOrWhiteSpace(overflowReplacementSourceFolder))
                    {
                        DisplayMessage("info", "[MOVIE] Overflow replacement source found: ", 0);
                        DisplayMessage("data", overflowReplacementSourceFolder);
                    }
                }

                // Subtitles check
                bool hasSubtitles = Directory.GetFiles(stagedMovieFolder, cleanTitle + "*.srt", SearchOption.TopDirectoryOnly).Any();

                // Delete existing destination contents (replace behavior)
                DeleteExistingMovieFolderContents(directoryPath);

                void QueueCellUpdate(string columnName, string value, string noteRecordSource = null)
                {
                    if (!colIndex.TryGetValue(columnName, out int colIdx))
                        return;

                    // Special handling for Note: preserve if it already mentions the recorded source
                    if (string.Equals(columnName, "Note", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(noteRecordSource))
                    {
                        string existingNote = rowToUpdate.Count > colIdx ? rowToUpdate[colIdx]?.ToString() ?? "" : "";
                        if (!string.IsNullOrWhiteSpace(existingNote) &&
                            existingNote.IndexOf(noteRecordSource, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return;
                        }
                    }

                    sheetContext.PendingUpdates.Add(new ValueRange
                    {
                        Range = $"Movies!{ColumnNumToLetter(colIdx)}{sheetRowNumber}",
                        MajorDimension = "ROWS",
                        Values = new List<IList<object>>
                {
                    new List<object> { value }
                }
                    });

                    // Keep in-memory row data updated too, so future logic sees latest values
                    while (rowToUpdate.Count <= colIdx)
                    {
                        rowToUpdate.Add("");
                    }

                    rowToUpdate[colIdx] = value;
                }

                // Only touch columns that exist in YOUR sheet
                QueueCellUpdate("Recorded Source", recordedSource);
                QueueCellUpdate("Status", "n");
                QueueCellUpdate("Ownership", "R");
                QueueCellUpdate("StreamFab", recordedSource.Equals("goojara", StringComparison.OrdinalIgnoreCase) || recordedSource.Equals("onoflix", StringComparison.OrdinalIgnoreCase) ? "" : "Y");
                QueueCellUpdate("Playon", "");
                QueueCellUpdate("Removed Splashes", "");
                QueueCellUpdate("Include Subtitles", hasSubtitles ? "Y" : "N");
                QueueCellUpdate("Note", "", recordedSource);
                QueueCellUpdate("Date Added", DateTime.Now.ToString("M-d-yyyy"));
                QueueCellUpdate("Possible Record Source", "");
                QueueCellUpdate("Resolution", "");
                QueueCellUpdate("Size", "");
                QueueCellUpdate("File Type", "");
                QueueCellUpdate("Movie Has Trailer", "");
                QueueCellUpdate("SRT Score", "");
                QueueCellUpdate("File Health", "");
                QueueCellUpdate("Rating Key", "");

                // If Directory was missing/blank, write it back. If we rerouted to Other,
                // force the sheet to the real final location on The Bus.
                if ((stagedFromOther || string.IsNullOrWhiteSpace(existingDirectoryPath)) && colIndex.ContainsKey(DIRECTORY))
                {
                    QueueCellUpdate(DIRECTORY, directoryPath);
                }

                // Commit move: move staged folder into directoryPath
                CommitMoveFolderContents(stagedMovieFolder, directoryPath);

                MoveArtworkAndDeleteOriginalMovieFolder(
                    overflowReplacementSourceFolder,
                    directoryPath,
                    cleanTitle);

                // Track for one Emby refresh later
                committedDestinations?.Add(directoryPath);

                DisplayMessage("success", $"Committed: {cleanTitle}", 0);
                DisplayMessage("default", " -> ", 0);
                DisplayMessage("data", directoryPath, 2);

                return true;
            }
            catch (Exception ex)
            {
                DisplayMessage("error", "UpdateMoviesSheetRowAndCommitMove failed: ", 0);
                DisplayMessage("data", ex.Message);
                return false;
            }
        }

        private static void FlushMoviesSheetBatchUpdates(MoviesSheetBatchContext sheetContext)
        {
            if (sheetContext == null || !sheetContext.PendingUpdates.Any())
            {
                DisplayMessage("info", "No Movies sheet updates to flush.", 2);
                return;
            }

            var batchUpdateRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = sheetContext.PendingUpdates
            };

            sheetContext.Service.Spreadsheets.Values
                .BatchUpdate(batchUpdateRequest, SPREADSHEET_ID)
                .Execute();

            DisplayMessage("success", $"Flushed {sheetContext.PendingUpdates.Count} Movies sheet updates in one batch.", 2);

            sheetContext.PendingUpdates.Clear();
        }

        private static string PickMoviesLibraryRootFromTitle(string cleanTitle, string root1, string root2, string root3)
        {
            // Uses your existing GetSortLetter() logic that already exists in Program.cs
            char letter = GetSortLetter(cleanTitle);

            if (letter == '#' || (letter >= 'A' && letter <= 'D'))
                return root1;

            if (letter >= 'E' && letter <= 'M')
                return root2;

            return root3; // N-Z
        }

        private static bool IsPathUnderRoot(string path, string root)
        {
            string normalizedPath = NormalizePathForComparison(path);
            string normalizedRoot = NormalizePathForComparison(root);

            if (string.IsNullOrWhiteSpace(normalizedPath) || string.IsNullOrWhiteSpace(normalizedRoot))
                return false;

            return normalizedPath.Equals(normalizedRoot, StringComparison.OrdinalIgnoreCase) ||
                   normalizedPath.StartsWith(normalizedRoot + "\\", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizePathForComparison(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return "";

            string value = path.Trim().Replace('/', '\\');
            bool isUnc = value.StartsWith(@"\\");
            string prefix = isUnc ? @"\\" : "";
            string remainder = isUnc ? value.Substring(2) : value;
            string[] parts = remainder.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

            return prefix + string.Join(@"\", parts);
        }

        private static string FindOriginalMovieFolderForOverflowReplacement(
            string existingDirectoryPath,
            string cleanTitle,
            string finalDestFolder,
            string moviesLibraryRoot_1,
            string moviesLibraryRoot_2,
            string moviesLibraryRoot_3)
        {
            var candidates = new List<string>();
            AddReplacementCandidate(candidates, existingDirectoryPath);

            string primaryRoot = PickMoviesLibraryRootFromTitle(cleanTitle, moviesLibraryRoot_1, moviesLibraryRoot_2, moviesLibraryRoot_3);
            AddReplacementCandidate(candidates, Path.Combine(primaryRoot, cleanTitle));

            foreach (var candidate in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(candidate) || !Directory.Exists(candidate))
                        continue;

                    if (IsPathUnderRoot(candidate, finalDestFolder) ||
                        IsPathUnderRoot(finalDestFolder, candidate) ||
                        IsLikelyOverflowMoviesPath(candidate))
                    {
                        continue;
                    }

                    if (LooksLikeSameMovieFolder(candidate, cleanTitle))
                        return candidate;
                }
                catch
                {
                    // Keep scanning other candidates if one path is unavailable.
                }
            }

            return "";
        }

        private static void AddReplacementCandidate(List<string> candidates, string path)
        {
            string normalized = TrimTrailingDirectorySeparators(path);
            if (string.IsNullOrWhiteSpace(normalized))
                return;

            if (!candidates.Any(p => string.Equals(p, normalized, StringComparison.OrdinalIgnoreCase)))
                candidates.Add(normalized);
        }

        private static string TrimTrailingDirectorySeparators(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return "";

            string value = path.Trim().Replace('/', '\\');
            while (value.Length > 3 && value.EndsWith("\\", StringComparison.Ordinal))
                value = value.Substring(0, value.Length - 1);

            return value;
        }

        private static bool IsLikelyOverflowMoviesPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            if (IsPathUnderRoot(path, DriveMapping.OtherMoviesLibraryRoot))
                return true;

            string normalized = NormalizePathForComparison(path) + "\\";
            return normalized.IndexOf("\\The Bus\\Movies\\", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   normalized.IndexOf("\\Other\\The Bus\\Movies\\", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool LooksLikeSameMovieFolder(string folderPath, string cleanTitle)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || string.IsNullOrWhiteSpace(cleanTitle) || !Directory.Exists(folderPath))
                return false;

            string target = NormalizeMovieNameForComparison(cleanTitle);
            string folderName = Path.GetFileName(TrimTrailingDirectorySeparators(folderPath));

            if (string.Equals(NormalizeMovieNameForComparison(folderName), target, StringComparison.OrdinalIgnoreCase))
                return true;

            var videoFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                .Where(IsMovieVideoFile)
                .Where(f => !IsTrailerFileName(Path.GetFileNameWithoutExtension(f)))
                .ToList();

            if (videoFiles.Count > 3)
                return false;

            return videoFiles.Any(f =>
                string.Equals(
                    NormalizeMovieNameForComparison(Path.GetFileNameWithoutExtension(f)),
                    target,
                    StringComparison.OrdinalIgnoreCase));
        }

        private static string NormalizeMovieNameForComparison(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "";

            string value = DuplicateYearRegex.Replace(name.Trim(), "($1)");
            value = Regex.Replace(value, @"\s+", " ");
            return value.Trim().TrimEnd('.');
        }

        private static bool HasPrimaryMovieVideo(string folderPath, string cleanTitle)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                return false;

            var videoFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                .Where(IsMovieVideoFile)
                .Where(f => !IsTrailerFileName(Path.GetFileNameWithoutExtension(f)))
                .ToList();

            string target = NormalizeMovieNameForComparison(cleanTitle);
            return videoFiles.Any(f =>
                       string.Equals(
                           NormalizeMovieNameForComparison(Path.GetFileNameWithoutExtension(f)),
                           target,
                           StringComparison.OrdinalIgnoreCase)) ||
                   videoFiles.Count > 0;
        }

        private static bool IsMovieVideoFile(string path)
        {
            string ext = Path.GetExtension(path);
            return ext.Equals(".mp4", StringComparison.OrdinalIgnoreCase) ||
                   ext.Equals(".mkv", StringComparison.OrdinalIgnoreCase) ||
                   ext.Equals(".avi", StringComparison.OrdinalIgnoreCase) ||
                   ext.Equals(".m4v", StringComparison.OrdinalIgnoreCase) ||
                   ext.Equals(".mov", StringComparison.OrdinalIgnoreCase) ||
                   ext.Equals(".webm", StringComparison.OrdinalIgnoreCase);
        }

        private static void MoveArtworkAndDeleteOriginalMovieFolder(
            string originalMovieFolder,
            string finalDestFolder,
            string cleanTitle)
        {
            if (string.IsNullOrWhiteSpace(originalMovieFolder))
                return;

            try
            {
                if (!Directory.Exists(originalMovieFolder))
                    return;

                if (IsLikelyOverflowMoviesPath(originalMovieFolder))
                    return;

                if (!Directory.Exists(finalDestFolder))
                {
                    DisplayMessage("warning", "[MOVIE] Original folder cleanup skipped; final destination is missing: ", 0);
                    DisplayMessage("data", finalDestFolder);
                    return;
                }

                if (!LooksLikeSameMovieFolder(originalMovieFolder, cleanTitle))
                {
                    DisplayMessage("warning", "[MOVIE] Original folder cleanup skipped; source does not look like the same movie: ", 0);
                    DisplayMessage("data", originalMovieFolder);
                    return;
                }

                if (!HasPrimaryMovieVideo(finalDestFolder, cleanTitle))
                {
                    DisplayMessage("warning", "[MOVIE] Original folder cleanup skipped; replacement video was not found at: ", 0);
                    DisplayMessage("data", finalDestFolder);
                    return;
                }

                int movedAssets = MovePreservedReplacementAssets(originalMovieFolder, finalDestFolder);
                DisplayMessage("info", "[MOVIE] Preserved artwork/trailer files moved from original folder: ", 0);
                DisplayMessage("data", movedAssets.ToString());

                DeleteOriginalMovieFolder(originalMovieFolder);
            }
            catch (Exception ex)
            {
                DisplayMessage("warning", "[MOVIE] Original folder cleanup failed: ", 0);
                DisplayMessage("data", originalMovieFolder, 0);
                DisplayMessage("warning", " | ", 0);
                DisplayMessage("data", ex.Message);
            }
        }

        private static int MovePreservedReplacementAssets(string originalMovieFolder, string finalDestFolder)
        {
            int moved = 0;

            foreach (var sourceFile in Directory.GetFiles(originalMovieFolder, "*.*", SearchOption.AllDirectories)
                         .Where(IsPreservedReplacementAsset))
            {
                try
                {
                    string relativePath = GetRelativePathUnderRoot(originalMovieFolder, sourceFile);
                    string destFile = Path.Combine(finalDestFolder, relativePath);

                    if (File.Exists(destFile))
                        continue;

                    EnsureDirectory(Path.GetDirectoryName(destFile));
                    MoveFileAllowingCrossVolume(sourceFile, destFile);
                    moved++;
                }
                catch (Exception ex)
                {
                    DisplayMessage("warning", "[MOVIE] Could not preserve asset: ", 0);
                    DisplayMessage("data", sourceFile, 0);
                    DisplayMessage("warning", " | ", 0);
                    DisplayMessage("data", ex.Message);
                }
            }

            return moved;
        }

        private static bool IsPreservedReplacementAsset(string path)
        {
            string ext = Path.GetExtension(path);
            if (ext.Equals(".jpg", StringComparison.OrdinalIgnoreCase) ||
                ext.Equals(".jpeg", StringComparison.OrdinalIgnoreCase) ||
                ext.Equals(".png", StringComparison.OrdinalIgnoreCase) ||
                ext.Equals(".webp", StringComparison.OrdinalIgnoreCase) ||
                ext.Equals(".bmp", StringComparison.OrdinalIgnoreCase) ||
                ext.Equals(".gif", StringComparison.OrdinalIgnoreCase) ||
                ext.Equals(".tbn", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return IsTrailerFileName(Path.GetFileNameWithoutExtension(path));
        }

        private static string GetRelativePathUnderRoot(string rootPath, string childPath)
        {
            try
            {
                string root = Path.GetFullPath(rootPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                string child = Path.GetFullPath(childPath);
                string prefix = root + Path.DirectorySeparatorChar;

                if (child.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    return child.Substring(prefix.Length);
            }
            catch
            {
            }

            return Path.GetFileName(childPath);
        }

        private static void MoveFileAllowingCrossVolume(string sourceFile, string destFile)
        {
            try
            {
                try { File.SetAttributes(sourceFile, FileAttributes.Normal); } catch { }
                File.Move(sourceFile, destFile);
            }
            catch
            {
                File.Copy(sourceFile, destFile);
                File.SetAttributes(destFile, FileAttributes.Normal);
                File.Delete(sourceFile);
            }
        }

        private static bool DeleteOriginalMovieFolder(string originalMovieFolder)
        {
            try
            {
                foreach (var file in Directory.GetFiles(originalMovieFolder, "*", SearchOption.AllDirectories))
                {
                    try { File.SetAttributes(file, FileAttributes.Normal); } catch { }
                }

                foreach (var dir in Directory.GetDirectories(originalMovieFolder, "*", SearchOption.AllDirectories)
                             .OrderByDescending(d => d.Length))
                {
                    try { File.SetAttributes(dir, FileAttributes.Normal); } catch { }
                }

                try { File.SetAttributes(originalMovieFolder, FileAttributes.Normal); } catch { }
                Directory.Delete(originalMovieFolder, recursive: true);

                if (!Directory.Exists(originalMovieFolder))
                {
                    DisplayMessage("success", "[MOVIE] Deleted original movie folder after overflow replacement: ", 0);
                    DisplayMessage("data", originalMovieFolder);
                    return true;
                }
            }
            catch (Exception ex)
            {
                DisplayMessage("warning", "[MOVIE] Could not delete original movie folder: ", 0);
                DisplayMessage("data", originalMovieFolder, 0);
                DisplayMessage("warning", " | ", 0);
                DisplayMessage("data", ex.Message);
            }

            return false;
        }
        private static void DeleteExistingMovieFolderContents(string directoryPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(directoryPath))
                    return;

                if (!Directory.Exists(directoryPath))
                    return;

                // Delete files
                foreach (var file in Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories))
                {
                    string ext = Path.GetExtension(file).ToLower();
                    string fileName = Path.GetFileName(file).ToLower();

                    // Skip trailer files
                    if (fileName.Contains("-trailer"))
                    {
                        DisplayMessage("data", $"Preserved trailer: {Path.GetFileName(file)}");
                        continue;
                    }

                    // Delete only video, subtitle and nfo types
                    if (ext == ".mp4" || ext == ".mkv" || ext == ".avi" || ext == ".mov" || ext == ".m4v" ||
                        ext == ".srt" || ext == ".sub" || ext == ".nfo" || ext == ".bif")
                    {
                        File.Delete(file);
                        DisplayMessage("data", $"Deleted old file: {Path.GetFileName(file)}");
                    }
                }

                // Delete subdirectories (deepest first)
                foreach (var dir in Directory.GetDirectories(directoryPath, "*", SearchOption.AllDirectories)
                             .OrderByDescending(d => d.Length))
                {
                    try
                    {
                        if (Directory.Exists(dir) && !Directory.EnumerateFileSystemEntries(dir).Any())
                            Directory.Delete(dir, recursive: false);
                    }
                    catch { /* keep going */ }
                }
            }
            catch
            {
                // swallow; we don’t want to abort the run on cleanup issues
            }
        }

        private static void CommitMoveFolderContents(string stagedFolder, string finalDestFolder)
        {
            if (string.IsNullOrWhiteSpace(stagedFolder) || string.IsNullOrWhiteSpace(finalDestFolder))
                return;

            Directory.CreateDirectory(finalDestFolder);

            // Move files (top level + subfolders)
            foreach (var dirPath in Directory.GetDirectories(stagedFolder, "*", SearchOption.TopDirectoryOnly))
            {
                var name = Path.GetFileName(dirPath.TrimEnd(Path.DirectorySeparatorChar));
                var destDir = Path.Combine(finalDestFolder, name);

                if (Directory.Exists(destDir))
                {
                    // merge if exists
                    CommitMoveFolderContents(dirPath, destDir);
                    TryDeleteIfEmpty(dirPath);
                }
                else
                {
                    Directory.Move(dirPath, destDir);
                }
            }

            foreach (var filePath in Directory.GetFiles(stagedFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                var name = Path.GetFileName(filePath);
                var destFile = Path.Combine(finalDestFolder, name);

                if (File.Exists(destFile))
                {
                    try
                    {
                        File.SetAttributes(destFile, FileAttributes.Normal);
                        File.Delete(destFile);
                    }
                    catch { /* ignore */ }
                }

                File.Move(filePath, destFile);
            }

            TryDeleteIfEmpty(stagedFolder);
        }

        private static void TryDeleteIfEmpty(string folder)
        {
            try
            {
                if (Directory.Exists(folder) && !Directory.EnumerateFileSystemEntries(folder).Any())
                    Directory.Delete(folder, recursive: false);
            }
            catch { }
        }

        static void LogTv(string message)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(message);
            Console.ResetColor();
        }

        private static char GetSortLetter(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
                return '#';

            string t = title.Trim();

            // Ignore "The "
            if (t.StartsWith("The ", StringComparison.OrdinalIgnoreCase))
                t = t.Substring(4).TrimStart();

            // If the first meaningful char is a digit/symbol, treat as '#'
            // We only use a LETTER A-Z for sorting, otherwise '#'
            foreach (char ch in t)
            {
                if (char.IsWhiteSpace(ch))
                    continue;

                if (char.IsLetter(ch))
                    return char.ToUpperInvariant(ch);

                // first non-space is NOT a letter -> numbers/symbols go to '#'
                return '#';
            }

            return '#';
        }

        private static string GetTvWorkingRoot(char letter, string r1, string r2, string r3, string r4, string r5, string r6)
        {
            // '#-b' contains non-letters and A-B.
            if (letter == '#') return r1;
            if (letter <= 'B') return r1;
            if (letter >= 'C' && letter <= 'F') return r2;
            if (letter >= 'G' && letter <= 'L') return r3;
            if (letter >= 'M' && letter <= 'P') return r4;
            if (letter >= 'Q' && letter <= 'S') return r5;
            return r6; // T-Z
        }

        private static string GetMovieWorkingRoot(char letter, string r1, string r2, string r3)
        {
            // '#-D' contains non-letters and A-D.
            if (letter == '#') return r1;
            if (letter <= 'D') return r1;
            if (letter >= 'E' && letter <= 'M') return r2;
            return r3; // N-Z
        }

        public class LetterRange
        {
            public char Start { get; set; }
            public char End { get; set; }
            public string Label { get; set; }
            public string RootPath { get; set; }

            public bool Contains(char c)
            {
                return c >= Start && c <= End;
            }
        }

        private static void EnsureDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
        }

        private static string GetUniqueDirectoryPath(string desiredPath)
        {
            if (!Directory.Exists(desiredPath)) return desiredPath;

            int i = 2;
            while (true)
            {
                string candidate = desiredPath + " (" + i + ")";
                if (!Directory.Exists(candidate)) return candidate;
                i++;
            }
        }

        private static string GetUniqueFilePath(string desiredPath)
        {
            if (!File.Exists(desiredPath)) return desiredPath;

            string dir = Path.GetDirectoryName(desiredPath);
            string name = Path.GetFileNameWithoutExtension(desiredPath);
            string ext = Path.GetExtension(desiredPath);

            int i = 2;
            while (true)
            {
                string candidate = Path.Combine(dir, name + " (" + i + ")" + ext);
                if (!File.Exists(candidate)) return candidate;
                i++;
            }
        }

        private static void MoveFileRobust(string sourceFile, string destFile)
        {
            // Try fast move; if that fails (cross-volume / UNC), fall back to copy + delete.
            try
            {
                EnsureDirectory(Path.GetDirectoryName(destFile));
                File.Move(sourceFile, destFile);
                Type("Moved ", 0, 0, 0);
                Type(sourceFile, 0, 0, 0, "Blue");
                Type(" -> ", 0, 0, 0);
                Type(destFile, 0, 0, 1, "Blue");
            }
            catch
            {
                EnsureDirectory(Path.GetDirectoryName(destFile));
                File.Copy(sourceFile, destFile);
                File.SetAttributes(destFile, FileAttributes.Normal);
                File.Delete(sourceFile);
                Type("Copied+Deleted ", 0, 0, 0);
                Type(sourceFile, 0, 0, 0, "Blue");
                Type(" -> ", 0, 0, 0);
                Type(destFile, 0, 0, 1, "Blue");
            }
        }

        private static void MoveDirectoryRobust(string sourceDir, string destDir)
        {
            // Try fast move; if that fails (cross-volume / UNC), fall back to copy+delete.
            try
            {
                EnsureDirectory(Path.GetDirectoryName(destDir));
                Directory.Move(sourceDir, destDir);
                Type("Moved ", 0, 0, 0);
                Type(sourceDir, 0, 0, 0, "Blue");
                Type(" -> ", 0, 0, 0);
                Type(destDir, 0, 0, 1, "Blue");
            }
            catch
            {
                CopyDirectoryRecursive(sourceDir, destDir);
                Directory.Delete(sourceDir, recursive: true);
                Type("Copied+Deleted ", 0, 0, 0);
                Type(sourceDir, 0, 0, 0, "Blue");
                Type(" -> ", 0, 0, 0);
                Type(destDir, 0, 0, 1, "Blue");
            }
        }

        private static void CopyDirectoryRecursive(string sourceDir, string destDir)
        {
            EnsureDirectory(destDir);

            foreach (var file in Directory.GetFiles(sourceDir))
            {
                string destFile = GetUniqueFilePath(Path.Combine(destDir, Path.GetFileName(file)));
                File.Copy(file, destFile);
                File.SetAttributes(destFile, FileAttributes.Normal);
            }

            foreach (var dir in Directory.GetDirectories(sourceDir))
            {
                string childName = Path.GetFileName(dir);
                CopyDirectoryRecursive(dir, Path.Combine(destDir, childName));
            }
        }

        protected static string CleanAmpersands(string dirtyString)
        {
            string newString = "";
            for (int i = 0; i < dirtyString.Length; i++)
            {
                if (dirtyString[i] != '&')
                {
                    newString += dirtyString[i];
                }
                else
                {
                    newString += "and";
                }
            }
            return newString;
        }
        protected static string CleanString(string dirtyString)
        {
            string newString = "";
            for (int i = 0; i < dirtyString.Length; i++)
            {
                if (dirtyString[i] != '\\' && dirtyString[i] != '/' && dirtyString[i] != ':' && dirtyString[i] != '*' && dirtyString[i] != '?' && dirtyString[i] != '"' && dirtyString[i] != '<' && dirtyString[i] != '>' && dirtyString[i] != '|')
                {
                    newString += dirtyString[i];
                }
                else
                {
                    newString += "";
                }
            }
            return newString;
        }

        protected static void CopyFile(string source, string destination)
        {
            try
            {
                File.Copy(source, destination);
                File.SetAttributes(destination, FileAttributes.Normal);
                //Type("Copied ", 0, 0, 0);
                //Type(source, 0, 0, 1, "Green");
                //Type(" to ", 0, 0, 0);
                //Type(destination, 0, 0, 1, "Green");
            }
            catch (Exception e)
            {
                DisplayMessage("error", "An error occured while trying to copy the file. | ");
                DisplayMessage("harderror", e.Message);
            }
        } // End MoveDirectory()

        /// <summary>
        /// Takes a directory location without the drive letter and then searches for that directory across all drives to find the current location.
        /// </summary>
        /// <param name="directoryLocation">The directory location without the preceding drive letter.</param>
        /// <returns>An ArrayList that contains any hard drive letter that contains that directory.</returns>
        protected static ArrayList FindDriveLetters(String directoryLocation)
        {
            string[] driveLetters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            ArrayList foundDriveLetters = new ArrayList();
            foreach (var letter in driveLetters)
            {
                string withDriveLetter = letter + ":" + directoryLocation;
                if (Directory.Exists(withDriveLetter))
                {
                    foundDriveLetters.Add(letter);
                }
            }

            return foundDriveLetters;
        } // End FindDirectory()

        //protected static void MarkMoviesAsOld()
        //{
        //    // Declare variables.
        //    UserCredential credential;
        //    Dictionary<string, int> SheetVariables = new Dictionary<string, int>
        //    {
        //        { "Old", -1 },
        //        { DIRECTORY, -1 },
        //        { CLEAN_TITLE, -1 }
        //    };
        //    Dictionary<string, int> NotFoundColumns = new Dictionary<string, int>();

        //    GetTitleRowData(ref SheetVariables, MOVIES_TITLE_RANGE);
        //    bool lessThanZero = CheckColumns(ref NotFoundColumns, SheetVariables);

        //    if (lessThanZero)
        //    {
        //        Type("We didn't find a column that we were looking for...", 0, 0, 1, "Red");
        //        foreach (KeyValuePair<string, int> variable in NotFoundColumns)
        //        {
        //            Type("Key: " + variable.Key.ToString() + ", Value: " + variable.Value.ToString(), 0, 0, 1, "Red");

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
        //                            string OldFileLocation = row[Convert.ToInt16(SheetVariables[DIRECTORY])].ToString() + "\\" + row[Convert.ToInt16(SheetVariables[CLEAN_TITLE])].ToString() + ".mp4";
        //                            string NewFileLocation = row[Convert.ToInt16(SheetVariables[DIRECTORY])].ToString() + "\\" + row[Convert.ToInt16(SheetVariables[CLEAN_TITLE])].ToString() + "_OLD.mp4";

        //                            if (File.Exists(OldFileLocation))
        //                            {
        //                                File.Move(OldFileLocation, NewFileLocation);

        //                                Type(row[Convert.ToInt16(SheetVariables[CLEAN_TITLE])].ToString() + " has been renamed.", 0, 0, 1, "Green");
        //                            }
        //                            else
        //                            {
        //                                Type(row[Convert.ToInt16(SheetVariables[CLEAN_TITLE])].ToString() + " was set to be renamed, but doesn't exist.", 0, 0, 1,"Yellow");
        //                            }

        //                        }

        //                    }
        //                    catch (Exception e)
        //                    {
        //                        Type("Something went wrong..." + e.Message, 0, 0, 1, "Red");
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

        private static void RunSelectedMovieNfoFiles()
        {
            IList<IList<Object>> movieData = CallGetData(movieSheetVariables, MOVIES_TITLE_RANGE, MOVIES_DATA_RANGE);

            DisplayMessage("info", "Adding/Overwriting selected NFO files...");
            CreateNfoFiles(movieData, movieSheetVariables, 2);
        }

        private static PlexMetadataSheetUpdater.Summary RunPlexMetadataSheetSync(SheetsService sheetsService, IList<int> movieRowNumbersToUpdate)
        {
            string plexBase = GetPlexBaseUrlFromVarsSheet();
            if (string.IsNullOrWhiteSpace(plexBase))
            {
                Type("Plex base URL (Enter for http://192.168.0.5:32400): ", 0, 0, 0);
                plexBase = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(plexBase)) plexBase = "http://192.168.0.5:32400";
            }

            string plexToken = GetPlexTokenFromVarsSheet();
            if (string.IsNullOrWhiteSpace(plexToken))
            {
                Type("Plex token (paste X-Plex-Token): ", 0, 0, 0);
                plexToken = Console.ReadLine();
            }

            if (string.IsNullOrWhiteSpace(plexToken))
            {
                DisplayMessage("error", "Plex metadata sync skipped because no Plex token was found.");
                return new PlexMetadataSheetUpdater.Summary();
            }

            int sectionId = GetPlexMoviesSectionIdFromVarsSheet();

            return PlexMetadataSheetUpdater.Run(
                sheetsService,
                SPREADSHEET_ID,
                (t, msg, ind) => DisplayMessage(t, msg, ind),
                new PlexMetadataSheetUpdater.Options
                {
                    PlexBaseUrl = plexBase,
                    PlexToken = plexToken,
                    MoviesSectionId = sectionId,
                    MoviesSheetName = "Movies",
                    MoviesHeaderRow = 2,
                    MoviesDataStartRow = 3,
                    StatusFilterValue = "n",
                    IncludeNfoBodyTagsAsLabels = true,
                    MovieRowNumbersToUpdate = movieRowNumbersToUpdate == null
                        ? new List<int>()
                        : movieRowNumbersToUpdate.ToList()
                });
        }

        private static void ClearMovieQuickCreateMarks(SheetsService sheetsService, IList<int> movieRowNumbers)
        {
            if (sheetsService == null || movieRowNumbers == null || movieRowNumbers.Count == 0)
                return;

            var rowsToClear = movieRowNumbers
                .Where(r => r > 0)
                .Distinct()
                .OrderBy(r => r)
                .ToList();

            if (rowsToClear.Count == 0)
                return;

            try
            {
                var headerResponse = sheetsService.Spreadsheets.Values.Get(SPREADSHEET_ID, MOVIES_TITLE_RANGE).Execute();
                var headers = headerResponse.Values != null && headerResponse.Values.Count > 0
                    ? headerResponse.Values[0]
                    : null;

                if (headers == null)
                {
                    DisplayMessage("error", "Could not clear Quick Create marks because the Movies header row was not found.");
                    return;
                }

                int quickCreateColumnIndex = -1;
                for (int i = 0; i < headers.Count; i++)
                {
                    string header = (headers[i] ?? "").ToString().Trim();
                    if (string.Equals(header, QUICK_CREATE, StringComparison.OrdinalIgnoreCase))
                    {
                        quickCreateColumnIndex = i;
                        break;
                    }
                }

                if (quickCreateColumnIndex < 0)
                {
                    DisplayMessage("error", "Could not clear Quick Create marks because the Movies 'Quick Create' column was not found.");
                    return;
                }

                string quickCreateColumn = ColumnNumToLetter(quickCreateColumnIndex);
                var updates = new List<ValueRange>();

                foreach (int rowNumber in rowsToClear)
                {
                    updates.Add(new ValueRange
                    {
                        Range = "Movies!" + quickCreateColumn + rowNumber,
                        MajorDimension = "ROWS",
                        Values = new List<IList<object>> { new List<object> { "" } }
                    });
                }

                var batchRequest = new BatchUpdateValuesRequest
                {
                    ValueInputOption = "USER_ENTERED",
                    Data = updates
                };

                sheetsService.Spreadsheets.Values.BatchUpdate(batchRequest, SPREADSHEET_ID).Execute();

                DisplayMessage("success", "Cleared Quick Create marks for ", 0);
                DisplayMessage("data", rowsToClear.Count.ToString(), 0);
                DisplayMessage("success", " ratingKey metadata row(s).", 2);
            }
            catch (Exception ex)
            {
                DisplayMessage("error", "Failed to clear Quick Create marks after ratingKey metadata sync: " + ex.Message, 2);
            }
        }

        private static string GetPlexBaseUrlFromVarsSheet()
        {
            try
            {
                DisplayMessage("warning", "Grabbing Plex base URL from VARS sheet... ", 0);

                string baseUrl = GetVarsValueByName("Plex Base URL", "PLEX_BASE_URL", "PlexBaseUrl", "Plex URL", "PLEX_URL");
                if (!string.IsNullOrWhiteSpace(baseUrl))
                {
                    DisplayMessage("success", "DONE");
                    return baseUrl;
                }

                DisplayMessage("warning", "not found.");
            }
            catch (Exception ex)
            {
                DisplayMessage("warning", "could not read Plex base URL from VARS: " + ex.Message);
            }

            return "";
        }

        private static string GetPlexTokenFromVarsSheet()
        {
            try
            {
                DisplayMessage("warning", "Grabbing Plex token from VARS sheet... ", 0);

                string token = GetVarsValueByName("Plex Token", "PLEX_TOKEN", "X-Plex-Token", "PlexToken");
                if (string.IsNullOrWhiteSpace(token))
                    token = GetFirstCellValue("VARS!A4");

                if (!string.IsNullOrWhiteSpace(token))
                {
                    DisplayMessage("success", "DONE");
                    return token;
                }

                DisplayMessage("warning", "not found.");
            }
            catch (Exception ex)
            {
                DisplayMessage("warning", "could not read Plex token from VARS: " + ex.Message);
            }

            return "";
        }

        private static int GetPlexMoviesSectionIdFromVarsSheet()
        {
            try
            {
                DisplayMessage("warning", "Grabbing Plex Movies section ID from VARS sheet... ", 0);

                string sectionIdText = GetVarsValueByName(
                    "Plex Movies Section ID",
                    "PLEX_MOVIES_SECTION_ID",
                    "Plex Movie Section ID",
                    "PLEX_MOVIE_SECTION_ID",
                    "Movies Section ID",
                    "Plex Section ID");

                int sectionId;
                if (!string.IsNullOrWhiteSpace(sectionIdText) &&
                    int.TryParse(sectionIdText.Trim(), out sectionId) &&
                    sectionId > 0)
                {
                    DisplayMessage("success", "DONE");
                    return sectionId;
                }

                DisplayMessage("warning", "not found. Using 1.");
            }
            catch (Exception ex)
            {
                DisplayMessage("warning", "could not read Plex Movies section ID from VARS: " + ex.Message);
            }

            return 1;
        }

        private static string GetVarsValueByName(params string[] names)
        {
            IList<IList<object>> varsData = GetData("VARS!A:ZZ");
            if (varsData == null || names == null || names.Length == 0)
                return "";

            var normalizedNames = names
                .Select(NormalizeVarsKey)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();

            if (varsData.Count >= 3)
            {
                var labelRow = varsData[1];
                var valueRow = varsData[2];

                for (int i = 0; i < labelRow.Count; i++)
                {
                    string key = (labelRow[i] ?? "").ToString().Trim();
                    if (!normalizedNames.Contains(NormalizeVarsKey(key)))
                        continue;

                    if (valueRow != null && valueRow.Count > i)
                        return (valueRow[i] ?? "").ToString().Trim();

                    return "";
                }
            }

            foreach (var row in varsData)
            {
                if (row == null || row.Count == 0)
                    continue;

                string key = (row[0] ?? "").ToString().Trim();
                if (!normalizedNames.Contains(NormalizeVarsKey(key)))
                    continue;

                if (row.Count > 1)
                    return (row[1] ?? "").ToString().Trim();
            }

            return "";
        }

        private static string GetFirstCellValue(string range)
        {
            IList<IList<object>> rows = GetData(range);
            if (rows == null || rows.Count == 0 || rows[0] == null || rows[0].Count == 0)
                return "";

            return (rows[0][0] ?? "").ToString().Trim();
        }

        private static string NormalizeVarsKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            return new string(value.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
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
            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            int nfoFileNotFoundCount = 0, nfoFileOverwrittenCount = 0, nfoFileCreatedCount = 0;
            bool directoryFound = false;
            var specialTitle = "";
            var movieDirectory = "";
            var nfoBody = "";
            var rowNum = "";
            var status = "";
            var quickCreate = "";
            var movieTitle = "";
            var quickCreateInt = 0;
            var cantFindInt = 0;

            foreach (var row in data)
            {
                if (row.Count > 25)
                {
                    try
                    {
                        directoryFound = false;
                        movieTitle = row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString();
                        specialTitle = row[Convert.ToInt16(sheetVariables["Special Title"])].ToString();
                        movieDirectory = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString();
                        nfoBody = row[Convert.ToInt16(sheetVariables[NFO_BODY])].ToString();
                        rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                        cantFindInt = Convert.ToInt16(sheetVariables["Can't Find"]);
                        status = "";
                        quickCreate = "";
                        quickCreateInt = 0;

                        if (!specialTitle.Equals(""))
                        {
                            movieTitle = specialTitle;
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong trying to create an NFO file from row num: " + row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString() + " | " + e.Message, 0, 0, 1, "Red");
                    }

                    // If we are creating NFO files for YouTube videos then we need to trim the titles,
                    // also, we don't need to worry about checking for status.
                    if (isYouTubeFile)
                    {
                        if (movieTitle.Length > 30)
                        {
                            movieTitle = movieTitle.Substring(0, 30).Trim();
                        }
                    }
                    else
                    {
                        status = row[Convert.ToInt16(sheetVariables[STATUS])].ToString();
                    }

                    try
                    {

                        if (sheetVariables.ContainsKey(QUICK_CREATE) && row.Count > Convert.ToInt16(sheetVariables[QUICK_CREATE]))
                        {
                            quickCreate = row[Convert.ToInt16(sheetVariables[QUICK_CREATE])].ToString();
                            quickCreateInt = Convert.ToInt16(sheetVariables[QUICK_CREATE]);
                        }


                        if (isYouTubeFile || (!status.Equals("") && status[0].ToString().ToUpper() != "X"))
                        {
                            if (Directory.Exists(movieDirectory))
                            {
                                directoryFound = true;

                                string fileLocation = movieDirectory + "\\" + movieTitle + ".nfo";

                                if (type == 1) // All movies, overwrite old NFO files AND put in new ones, but only if the folder exists (I don't want folders with only NFO files sitting in them).
                                {
                                    if (File.Exists(fileLocation))
                                    {
                                        File.Delete(fileLocation);
                                        nfoFileOverwrittenCount++;
                                        Type("NFO overwritten at: ", 0, 0, 0, "Green");
                                        Type(fileLocation, 0, 0, 1, "Blue");

                                    }
                                    else
                                    {
                                        nfoFileCreatedCount++;
                                        Type("NFO created at: ", 0, 0, 0, "Green");
                                        Type(fileLocation, 0, 0, 1, "Blue");
                                    }
                                    WriteNfoFile(fileLocation, nfoBody);

                                    var strCellToPutData = "Movies!" + ColumnNumToLetter(cantFindInt) + rowNum;
                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { "" } }
                                    });

                                }
                                else if (type == 2) // Only selected movies marked with an x.
                                {
                                    if (row.Count > quickCreateInt && quickCreate.ToUpper() == "X")
                                    {
                                        WriteNfoFile(fileLocation, nfoBody);
                                        var strCellToPutData = "Movies!" + ColumnNumToLetter(quickCreateInt) + rowNum;

                                        if (File.Exists(fileLocation))
                                        {
                                            nfoFileOverwrittenCount++;
                                            Type("NFO overwritten at: ", 0, 0, 0, "Green");
                                            Type(fileLocation, 0, 0, 1, "Blue");

                                        }
                                        else
                                        {
                                            nfoFileCreatedCount++;
                                            Type("NFO created at: ", 0, 0, 0, "Green");
                                            Type(fileLocation, 0, 0, 1, "Blue");
                                        }

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { "" } }
                                        });

                                    }

                                }
                                else if (type == 3) // Only the movies that are missing NFO files.
                                {
                                    if (!File.Exists(fileLocation))
                                    {
                                        WriteNfoFile(fileLocation, nfoBody);
                                        nfoFileCreatedCount++;
                                        Type("NFO created at: ", 0, 0, 0, "Green");
                                        Type(fileLocation, 0, 0, 1, "Blue");
                                    }
                                }

                            }

                            if (!directoryFound)
                            {
                                Type("We did not find the directory for: ", 0, 0, 0, "Red");
                                Type(movieDirectory, 0, 0, 1, "Yellow");
                                nfoFileNotFoundCount++;

                                var strCellToPutData = "Movies!" + ColumnNumToLetter(cantFindInt) + rowNum;
                                batchRequest.Data.Add(new ValueRange
                                {
                                    Range = strCellToPutData,
                                    MajorDimension = "ROWS",
                                    Values = new List<IList<object>> { new List<object> { "X" } }
                                });
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong when looking for: " + movieDirectory + " | " + e.Message, 0, 0, 1, "Red");
                    }

                }
            }
            BulkWriteToSheet(batchRequest);

            // Print out results.
            Type("It looks like that's the end of it.", 3, 100, 2);
            Type("NFO Files not found: ", 0, 0, 0); Type(nfoFileNotFoundCount.ToString(), 0, 0, 1, "Red");
            Type("NFO Files overwritten: ", 0, 0, 0); Type(nfoFileOverwrittenCount.ToString(), 0, 0, 1, "Blue");
            Type("NFO Files created: ", 0, 0, 0); Type(nfoFileCreatedCount.ToString(), 0, 0, 2, "Green");

        } // End CreateNfoFiles()

        protected static void CreateTvShowNfoFiles(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, int type)
        {
            int nfoFileNotFoundCount = 0, nfoFileOverwrittenCount = 0, nfoFileCreatedCount = 0;
            bool directoryFound = false;
            var cleanTitle = "";
            var tvShowDirectory = "";
            var nfoBody = "";
            var rowNum = "";
            var quickCreate = "";
            var quickCreateInt = 0;

            foreach (var row in data)
            {
                if (row.Count > 25)
                {
                    try
                    {
                        directoryFound = false;
                        cleanTitle = row[Convert.ToInt16(sheetVariables["Clean Name with Year"])].ToString();
                        tvShowDirectory = row[Convert.ToInt16(sheetVariables["Found Locations"])].ToString();
                        nfoBody = row[Convert.ToInt16(sheetVariables[NFO_BODY])].ToString();
                        rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                        quickCreate = "";
                        quickCreateInt = 0;
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong trying to create an NFO file from row num: " + row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString() + " | " + e.Message, 0, 0, 1, "Red");
                    }

                    try
                    {

                        if (sheetVariables.ContainsKey(QUICK_CREATE) && row.Count > Convert.ToInt16(sheetVariables[QUICK_CREATE]))
                        {
                            quickCreate = row[Convert.ToInt16(sheetVariables[QUICK_CREATE])].ToString();
                            quickCreateInt = Convert.ToInt16(sheetVariables[QUICK_CREATE]);
                        }


                        if (!tvShowDirectory.Equals(""))
                        {
                            if (Directory.Exists(tvShowDirectory))
                            {
                                directoryFound = true;

                                string fileLocation = tvShowDirectory + "\\" + "tvshow.nfo";

                                if (type == 1) // All tv shows, overwrite old NFO files AND put in new ones, but only if the folder exists (I don't want folders with only NFO files sitting in them).
                                {
                                    if (File.Exists(fileLocation))
                                    {
                                        File.Delete(fileLocation);
                                        nfoFileOverwrittenCount++;
                                        Type("NFO overwritten at: ", 0, 0, 0, "Green");
                                        Type(fileLocation, 0, 0, 1, "Blue");

                                    }
                                    else
                                    {
                                        nfoFileCreatedCount++;
                                        Type("NFO created at: ", 0, 0, 0, "Green");
                                        Type(fileLocation, 0, 0, 1, "Blue");
                                    }
                                    WriteNfoFile(fileLocation, nfoBody);

                                }
                                else if (type == 2) // Only selected tv shows marked with an x.
                                {
                                    if (row.Count > quickCreateInt && quickCreate.ToUpper() == "X")
                                    {
                                        WriteNfoFile(fileLocation, nfoBody);
                                        var strCellToPutData = "DB!" + ColumnNumToLetter(quickCreateInt) + rowNum;

                                        if (File.Exists(fileLocation))
                                        {
                                            nfoFileOverwrittenCount++;
                                            Type("NFO overwritten at: ", 0, 0, 0, "Green");
                                            Type(fileLocation, 0, 0, 1, "Blue");

                                        }
                                        else
                                        {
                                            nfoFileCreatedCount++;
                                            Type("NFO created at: ", 0, 0, 0, "Green");
                                            Type(fileLocation, 0, 0, 1, "Blue");
                                        }

                                        WriteSingleCellToSheet("", strCellToPutData);

                                    }

                                }
                                else if (type == 3) // Only the tv shows that are missing NFO files.
                                {
                                    if (!File.Exists(fileLocation))
                                    {
                                        WriteNfoFile(fileLocation, nfoBody);
                                        nfoFileCreatedCount++;
                                        Type("NFO created at: ", 0, 0, 0, "Green");
                                        Type(fileLocation, 0, 0, 1, "Blue");
                                    }
                                }

                            }

                            if (!directoryFound)
                            {
                                Type("We did not find the directory for: ", 0, 0, 0, "Red");
                                Type(tvShowDirectory, 0, 0, 1, "Yellow");
                                nfoFileNotFoundCount++;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong when looking for: " + tvShowDirectory + " | " + e.Message, 0, 0, 1, "Red");
                    }

                }
            }

            // Print out results.
            Type("It looks like that's the end of it.", 3, 100, 2);
            Type("NFO Files not found: ", 0, 0, 0); Type(nfoFileNotFoundCount.ToString(), 0, 0, 1, "Red");
            Type("NFO Files overwritten: ", 0, 0, 0); Type(nfoFileOverwrittenCount.ToString(), 0, 0, 1, "Blue");
            Type("NFO Files created: ", 0, 0, 0); Type(nfoFileCreatedCount.ToString(), 0, 0, 2, "Green");

        }

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
        //        { CLEAN_TITLE, -1 },
        //        { "Movie Letter", -1 },
        //        { "Ownership", -1 },
        //        { STATUS, -1 }
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
        //        Type("We didn't find a column that we were looking for...", 0, 0, 1, "Red");
        //        foreach (KeyValuePair<string, int> variable in NotFoundColumns)
        //        {
        //            Type("Key: " + variable.Key.ToString() + ", Value: " + variable.Value.ToString(), 0, 0, 1, "Red");
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
        //                    //Type("Row Count: " + row.Count.ToString() + ", Quick Create Column: " + Convert.ToInt16(SheetVariables[QUICK_CREATE]), 0, 0, 1, "DarkGray");
        //                    try
        //                    {
        //                        string DirectoryLocation = baseFolderLocation + row[Convert.ToInt16(SheetVariables["Movie Letter"])].ToString() + "\\" + row[Convert.ToInt16(SheetVariables[CLEAN_TITLE])].ToString();
        //                        var directoryFound = false;
        //                        var ownership = row[Convert.ToInt16(SheetVariables["Ownership"])].ToString();
        //                        string strCellToSaveData = "Movies!" + ColumnNumToLetter(SheetVariables[STATUS]) + intRowNum;

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
        //                        Type("Something went wrong..." + e.Message, 0, 0, 1, "Red");
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

        public static IList<IList<Object>> CallGetData(Dictionary<string, int> sheetVariables, string titleRowDataRange, string mainDataRange, string dataMessage = "Gathering sheet data... ")
        {
            Type(dataMessage, 0, 0, 0, "Yellow");
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
                Type("DONE", 0, 0, 1, "Green");
            }

            return movieData;
        }

        /// <summary>
        /// Grabs the data from the Google Sheet.
        /// Used for both the title row data, and the main data.
        /// </summary>
        /// <param name="sheetDataRange">The range in the sheet to pull data from.</param>
        /// <returns>The data from the selected range.</returns>
        public static IList<IList<object>> GetData(string sheetDataRange)
        {
            try
            {
                UserCredential credential;

                using (var stream =
                    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
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
                Type("An error has occured getting data: " + ex.Message, 0, 0, 1, "Red");
                throw;
            }

        } // End GetData()

        private static SheetsService CreateSheetsService()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = APLICATION_NAME,
            });
        }

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

        protected static void InputMovieRuntimes(IList<IList<Object>> data, Dictionary<string, int> sheetVariables)
        {
            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            string strRowNum = "",
                    strRuntime = "",
                    strMovieTitle = "",
                    strTmdbId = "";


            IList<IList<Object>> filteredData = new List<IList<Object>>();

            // First loop through the data to filter out the rows we won't be working with.
            foreach (var row in data)
            {
                if (row.Count > 70)
                {
                    strRuntime = row[Convert.ToInt16(sheetVariables["Runtime"])].ToString();

                    if (strRuntime.Equals(""))
                    {
                        filteredData.Add(row);
                    }
                }
            }

            foreach (var row in filteredData)
            {
                try
                {

                    strRowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                    strRuntime = row[Convert.ToInt16(sheetVariables["Runtime"])].ToString();
                    strMovieTitle = row[Convert.ToInt16(sheetVariables["IMDB Title"])].ToString();
                    strTmdbId = row[Convert.ToInt16(sheetVariables["TMDB ID"])].ToString();

                    if (strRuntime.Equals(""))
                    {
                        DisplayMessage("info", " Getting Runtime for: ", 0);
                        DisplayMessage("data", strMovieTitle);

                        dynamic movieDetails = TmdbApi.MoviesGetDetailsByTmdbId(strTmdbId);

                        if (movieDetails.runtime != null)
                        {
                            string runTime = movieDetails.runtime.ToString("D3");
                            string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Runtime"])) + strRowNum;

                            batchRequest.Data.Add(new ValueRange
                            {
                                Range = strCellToPutData,
                                MajorDimension = "ROWS",
                                Values = new List<IList<object>> { new List<object> { runTime } }
                            });
                        }
                        else if (movieDetails.status_code == 34)
                        {
                            DisplayMessage("warning", " --Runtime was not available");

                            string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Runtime"])) + strRowNum;

                            batchRequest.Data.Add(new ValueRange
                            {
                                Range = strCellToPutData,
                                MajorDimension = "ROWS",
                                Values = new List<IList<object>> { new List<object> { "N/A" } }
                            });
                        }


                    }
                }
                catch (Exception ex)
                {
                    DisplayMessage("error", "Something went wrong getting the Runtime for: ", 0);
                    DisplayMessage("warning", strMovieTitle);
                    DisplayMessage("harderror", ex.Message, 2);
                }
            }
            BulkWriteToSheet(batchRequest);
        }

        /// <summary>
        /// Grabs the list of movies from the Google Sheet. Sends each IMDB ID to theMovieDB.org API to get the movie data.
        /// Inserts missing movie data into the Google Sheet (TMDB Rating, Plot, TMDB ID).
        /// </summary>
        /// <param name="data">The movie data to run through.</param>
        /// <param name="sheetVariables">The column names to look at.</param>
        protected static void InputMovieData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, Boolean overwriteData = false)
        {
            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            int intTmdbIdDoneCount = 0,
                intTmdbRatingDoneCount = 0,
                intPlotDoneCount = 0,
                intReleaseDateDoneCount = 0,
                intTmdbIdNotFoundCount = 0,
                intTmdbRatingNotFoundCount = 0,
                intPlotNotFoundCount = 0,
                intReleaseDateNotFoundCount = 0;

            dynamic tmdbResponse; // The API call response.

            foreach (var row in data)
            {
                if (row.Count > 70) // If it's an empty row then it should have less than this.
                {
                    string rowNum = "";
                    string strCellToPutData = "";

                    // The following variables hold the values from the Google Sheet.
                    string tmdbIdValue = "",
                            tmdbRatingValue = "",
                            plotValue = "",
                            imdbIdValue = "",
                            imdbTitleValue = "",
                            sortTitleValue = "",
                            releaseDateValue = "",
                            youTubeTrailerIdValue = "",
                            posterValue = "";

                    // The following variables will be filled from the TMDB API call.
                    string tmdbId = "",
                            tmdbTitle = "",
                            tmdbSortTitle = "",
                            tmdbRating = "",
                            tmdbReleaseDate = "",
                            tmdbPlot = "",
                            tmdbPosterUrl = "";

                    // The following variables are used to keep track of the column number of each variable to input the data back to the Google Sheet.
                    int tmdbIdColumnNum = 0, // Used to input the returned ID back into the Google Sheet.
                        tmdbRatingColumnNum = 0, // Used to input the returned rating into the Google Sheet.
                        plotColumnNum = 0, // Used to input the returned overview into the Google Sheet.
                        quickCreateColumnNum = 0, // Used to mark the movies that Plot gets updated.
                        imdbTitleColumnNum = 0,
                        sortTitleColumnNum = 0,
                        releaseDateColumnNum = 0,
                        posterColumnNum = 0;

                    try
                    {
                        bool MovieFoundAtTmdb = true;
                        bool dataOverwritten = false;
                        bool logWritten = false;
                        rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                        tmdbIdValue = row[Convert.ToInt16(sheetVariables["TMDB ID"])].ToString();
                        tmdbIdColumnNum = Convert.ToInt16(sheetVariables["TMDB ID"]);
                        tmdbRatingValue = row[Convert.ToInt16(sheetVariables["TMDB Rating"])].ToString();
                        tmdbRatingColumnNum = Convert.ToInt16(sheetVariables["TMDB Rating"]);
                        plotValue = row[Convert.ToInt16(sheetVariables["Plot"])].ToString();
                        plotColumnNum = Convert.ToInt16(sheetVariables["Plot"]);
                        imdbIdValue = row[Convert.ToInt16(sheetVariables["IMDB ID"])].ToString();
                        imdbTitleValue = row[Convert.ToInt16(sheetVariables["IMDB Title"])].ToString();
                        imdbTitleColumnNum = Convert.ToInt16(sheetVariables["IMDB Title"]);
                        sortTitleValue = row[Convert.ToInt16(sheetVariables["Sort Title"])].ToString();
                        sortTitleColumnNum = Convert.ToInt16(sheetVariables["Sort Title"]);
                        releaseDateValue = row[Convert.ToInt16(sheetVariables["Release Date"])].ToString();
                        releaseDateColumnNum = Convert.ToInt16(sheetVariables["Release Date"]);
                        youTubeTrailerIdValue = row[Convert.ToInt16(sheetVariables["YouTube Trailer ID"])].ToString();
                        posterValue = row[Convert.ToInt16(sheetVariables["Poster"])].ToString();
                        posterColumnNum = Convert.ToInt16(sheetVariables["Poster"]);

                        if (sheetVariables.ContainsKey(QUICK_CREATE)) quickCreateColumnNum = Convert.ToInt16(sheetVariables[QUICK_CREATE]);

                        if (imdbTitleValue.Equals("") || sortTitleValue.Equals("") || releaseDateValue.Equals("") || tmdbIdValue.Equals("") || tmdbRatingValue.Equals("") || plotValue.Equals("") || posterValue.Equals("") || overwriteData)
                        {
                            logWritten = true;
                            DisplayMessage("warning", "Making TMDB Call... ", 0);

                            tmdbResponse = TmdbApi.MoviesGetDetails(imdbIdValue);

                            DisplayMessage("success", "DONE");

                            try
                            {
                                if (!tmdbResponse.Equals("") && tmdbResponse?.movie_results[0]?.id?.Value?.ToString() != "")
                                {
                                    dynamic tmdbR = tmdbResponse.movie_results[0];
                                    tmdbId = tmdbR.id.Value.ToString();
                                    double voteAverage = tmdbR.vote_average;
                                    double roundedAverage = Math.Round(voteAverage, 1, MidpointRounding.AwayFromZero);
                                    tmdbRating = roundedAverage.ToString();
                                    tmdbPlot = tmdbR.overview.ToString();
                                    tmdbTitle = tmdbR.title.ToString() + " (" + tmdbR.release_date?.ToString().Substring(0, 4) + ")";
                                    tmdbSortTitle = tmdbTitle.Substring(0, 4) == "The " ? tmdbTitle.Substring(4) : tmdbTitle;
                                    tmdbReleaseDate = tmdbR.release_date.ToString();
                                    string posterPath = tmdbR.poster_path?.ToString();

                                    if (!string.IsNullOrWhiteSpace(posterPath))
                                    {
                                        tmdbPosterUrl = "https://image.tmdb.org/t/p/w300" + posterPath;
                                    }
                                    else
                                    {
                                        tmdbPosterUrl = "https://placehold.co/200x300?text=No+Image";
                                    }

                                    if (!tmdbRating.Contains(".")) tmdbRating += ".0";

                                    string message = !overwriteData ? "Missing data for: " : "Overwriting data for: ";
                                    DisplayMessage("log", message, 0);
                                    DisplayMessage("info", tmdbTitle);
                                }
                                else
                                {
                                    DisplayMessage("error", "No record found at TMDB for: ", 0);
                                    DisplayMessage("warning", imdbIdValue);

                                    MovieFoundAtTmdb = false;
                                }
                            }
                            catch (Exception ex)
                            {
                                DisplayMessage("error", "No record found at TMDB for: ", 0);
                                DisplayMessage("warning", imdbIdValue);
                                DisplayMessage("harderror", ex.Message);

                                MovieFoundAtTmdb = false;
                            }

                            if (MovieFoundAtTmdb)
                            {
                                if ((tmdbTitle != "" && imdbTitleValue.Equals("")) || (overwriteData && !imdbTitleValue.Equals(tmdbTitle)))
                                {
                                    strCellToPutData = "Movies!" + ColumnNumToLetter(imdbTitleColumnNum) + rowNum;

                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { tmdbTitle } }
                                    });

                                    if (overwriteData)
                                    {
                                        dataOverwritten = true;
                                        DisplayMessage("info", "Title overwritten for: ", 0);
                                        DisplayMessage("default", tmdbTitle);
                                        DisplayMessage("default", imdbTitleValue, 0);
                                        DisplayMessage("info", " updated to: ", 0);
                                        DisplayMessage("default", tmdbTitle);
                                    }

                                }

                                if (tmdbTitle != "" && sortTitleValue.Equals(""))
                                {
                                    strCellToPutData = "Movies!" + ColumnNumToLetter(sortTitleColumnNum) + rowNum;

                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { tmdbSortTitle } }
                                    });

                                    if (overwriteData)
                                    {
                                        dataOverwritten = true;
                                        DisplayMessage("info", "Sort Title overwritten for: ", 0);
                                        DisplayMessage("default", tmdbTitle);
                                        DisplayMessage("default", sortTitleValue, 0);
                                        DisplayMessage("info", " updated to: ", 0);
                                        DisplayMessage("default", tmdbSortTitle);
                                    }
                                }

                                if (tmdbIdValue.Equals("") || (overwriteData && !tmdbIdValue.Equals(tmdbId)))
                                {
                                    if (tmdbId != "")
                                    {
                                        strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbIdColumnNum) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { tmdbId } }
                                        });

                                        if (overwriteData)
                                        {
                                            dataOverwritten = true;
                                            DisplayMessage("info", "TMDB ID overwritten for: ", 0);
                                            DisplayMessage("default", tmdbTitle);
                                            DisplayMessage("default", tmdbIdValue, 0);
                                            DisplayMessage("info", " updated to: ", 0);
                                            DisplayMessage("default", tmdbId);
                                        }
                                    }
                                    else
                                    {
                                        intTmdbIdNotFoundCount++;
                                    }
                                }

                                if (tmdbRatingValue.Equals("") || (overwriteData && !tmdbRatingValue.Equals(tmdbRating)))
                                {
                                    if (tmdbRating != "")
                                    {
                                        strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbRatingColumnNum) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { tmdbRating } }
                                        });

                                        if (overwriteData)
                                        {
                                            dataOverwritten = true;
                                            DisplayMessage("info", "TMDB Rating overwritten for: ", 0);
                                            DisplayMessage("default", tmdbTitle);
                                            DisplayMessage("default", tmdbRatingValue, 0);
                                            DisplayMessage("info", " updated to: ", 0);
                                            DisplayMessage("default", tmdbRating);
                                        }
                                    }
                                    else
                                    {
                                        intTmdbRatingNotFoundCount++;
                                    }
                                }

                                if (plotValue.Equals("") || (overwriteData && !plotValue.Equals(tmdbPlot)))
                                {
                                    if (tmdbPlot != "")
                                    {
                                        strCellToPutData = "Movies!" + ColumnNumToLetter(plotColumnNum) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { tmdbPlot } }
                                        });

                                        if (overwriteData)
                                        {
                                            dataOverwritten = true;
                                            DisplayMessage("info", "Plot overwritten for: ", 0);
                                            DisplayMessage("default", tmdbTitle);
                                            DisplayMessage("default", plotValue, 0);
                                            DisplayMessage("info", " updated to: ", 0);
                                            DisplayMessage("default", tmdbPlot);
                                        }
                                    }
                                    else
                                    {
                                        intPlotNotFoundCount++;
                                    }
                                }

                                if (releaseDateValue.Equals("") || overwriteData)
                                {
                                    if (tmdbReleaseDate != "")
                                    {
                                        if (overwriteData && (!releaseDateValue.Equals("")))
                                        {
                                            // Now that we have confirmed both dates aren't empty. Let's parse the strings to dates in order to correctly compare them.
                                            string date1 = releaseDateValue;
                                            string date2 = tmdbReleaseDate;

                                            // Parse both strings into DateTime objects
                                            DateTime parsedDate1 = DateTime.Parse(date1);
                                            DateTime parsedDate2 = DateTime.Parse(date2);

                                            // If the dates are equal to each other, then we can just exit out of here.
                                            if (parsedDate1 == parsedDate2) { break; }
                                        }

                                        strCellToPutData = "Movies!" + ColumnNumToLetter(releaseDateColumnNum) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { tmdbReleaseDate } }
                                        });

                                        if (overwriteData)
                                        {
                                            dataOverwritten = true;
                                            DisplayMessage("info", "Release Date overwritten for: ", 0);
                                            DisplayMessage("default", tmdbTitle);
                                            DisplayMessage("default", releaseDateValue, 0);
                                            DisplayMessage("info", " updated to: ", 0);
                                            DisplayMessage("default", tmdbReleaseDate);
                                        }
                                    }
                                    else
                                    {
                                        intReleaseDateNotFoundCount++;
                                    }
                                }

                                if (posterValue.Equals("") || (overwriteData && !posterValue.Equals(tmdbPosterUrl)))
                                {
                                    if (tmdbPosterUrl != "")
                                    {
                                        strCellToPutData = "Movies!" + ColumnNumToLetter(posterColumnNum) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { tmdbPosterUrl } }
                                        });

                                        if (overwriteData)
                                        {
                                            dataOverwritten = true;
                                            DisplayMessage("info", "Poster overwritten for: ", 0);
                                            DisplayMessage("default", tmdbTitle);
                                        }
                                    }
                                }
                            }

                            if (overwriteData && !dataOverwritten)
                            {
                                DisplayMessage("warning", "No data to overwrite for: ", 0);
                                DisplayMessage("info", tmdbTitle);
                            }
                        }

                        // Now check for missing trailer ID
                        if (youTubeTrailerIdValue.Equals(""))
                        {
                            logWritten = true;
                            var populatedTitle = !tmdbTitle.Equals("") ? tmdbTitle : imdbTitleValue;
                            DisplayMessage("info", "NOTE: I see that the YouTube ID for: ", 0);
                            DisplayMessage("data", populatedTitle, 0);
                            DisplayMessage("info", " is empty.");
                            DisplayMessage("data", "I will now try to retrieve it.");
                            DisplayMessage("warning", "Making TMDB Call... ", 0);

                            tmdbResponse = TmdbApi.MoviesGetVideos(imdbIdValue);

                            DisplayMessage("success", "DONE");

                            try
                            {
                                if (!tmdbResponse.Equals(""))
                                {
                                    dynamic videoList = tmdbResponse.results;

                                    if (videoList.Count > 0)
                                    {
                                        var officialTrailerId = "";
                                        var trailerId = "";
                                        foreach (var video in videoList)
                                        {
                                            string videoSite = video.site,
                                                videoType = video.type,
                                                videoOfficial = video.official;

                                            if (videoSite.ToUpper() == "YOUTUBE" && videoType.ToUpper() == "TRAILER")
                                            {
                                                if (videoOfficial == "TRUE")
                                                {
                                                    officialTrailerId = video.key;
                                                }
                                                else
                                                {
                                                    trailerId = video.key;
                                                }
                                            }
                                        }

                                        if (!trailerId.Equals("") || !officialTrailerId.Equals(""))
                                        {
                                            strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["YouTube Trailer ID"])) + rowNum;

                                            string trailerIdToWrite = !officialTrailerId.Equals("") ? officialTrailerId : trailerId;

                                            batchRequest.Data.Add(new ValueRange
                                            {
                                                Range = strCellToPutData,
                                                MajorDimension = "ROWS",
                                                Values = new List<IList<object>> { new List<object> { trailerIdToWrite } }
                                            });
                                        }
                                        else
                                        {
                                            strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["YouTube Trailer ID"])) + rowNum;

                                            batchRequest.Data.Add(new ValueRange
                                            {
                                                Range = strCellToPutData,
                                                MajorDimension = "ROWS",
                                                Values = new List<IList<object>> { new List<object> { "N/A" } }
                                            });
                                        }
                                    }
                                    else
                                    {
                                        strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["YouTube Trailer ID"])) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { "N/A" } }
                                        });
                                    }
                                }
                                else
                                {
                                    strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["YouTube Trailer ID"])) + rowNum;

                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { "N/A" } }
                                    });
                                }
                            }
                            catch (Exception ex)
                            {
                                DisplayMessage("error", "Something went wrong getting the YouTube Trailer ID for: ", 0);
                                DisplayMessage("warning", imdbIdValue);
                                DisplayMessage("harderror", ex.Message, 2);

                                MovieFoundAtTmdb = false;
                            }
                        }

                        if (logWritten)
                        {
                            DisplayMessage("log", "-----------------------------", 2);
                        }

                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong while putting in movie data for: " + tmdbTitle, 0, 0, 1, "Red");
                        Type(e.Message, 0, 0, 2, "DarkRed");
                    }

                }
            }

            if (batchRequest.Data.Count > 0)
            {
                var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
            }

            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("TMDB IDs inserted: " + intTmdbIdDoneCount, 0, 0, 1, "Green");
            Type("TMDB Ratings inserted: " + intTmdbRatingDoneCount, 0, 0, 1, "Green");
            Type("Plots inserted: " + intPlotDoneCount, 0, 0, 1, "Green");
            Type("Release Date inserted: " + intReleaseDateDoneCount, 0, 0, 1, "Green");
            Type("TMDB IDs not available: " + intTmdbIdNotFoundCount, 0, 0, 1, "Red");
            Type("TMDB Ratings not available: " + intTmdbRatingNotFoundCount, 0, 0, 1, "Red");
            Type("Plots not available: " + intPlotNotFoundCount, 0, 0, 1, "Red");
            Type("Release Date not available: " + intReleaseDateNotFoundCount, 0, 0, 1, "Red");
        }

        /// <summary>
        /// Copies the auto populated data into the proper column.
        /// </summary>
        /// <param name="data">The movie data to run through.</param>
        /// <param name="sheetVariables">The column names to look at.</param>
        protected static bool CopyAutoPopulatedData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, Boolean overwiteData = false)
        {
            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            int intImdbTitlesCopiedCount = 0,
                intSortTitlesCopiedCount = 0,
                intContentRatingCopiedCount = 0,
                intMpaaRatingCopiedCount = 0,
                intTitlesLoadingCount = 0,
                intSortTitlesSkippedCount = 0,
                intContentRatingLoadingCount = 0,
                intMpaaRatingLoadingCount = 0;

            string rowNum = "",
                autoTitleValue = "",
                imdbTitleValue = "",
                sortTitleValue = "",
                autoContentRatingValue = "",
                contentRatingValue = "",
                autoMpaaRatingValue = "",
                mpaaRatingValue = "",
                movieTitle = "",
                strCellToPutData = "";

            int imdbTitleColumnNum = 0,
                sortTitleColumnNum = 0,
                contentRatingColumnNum = 0,
                mpaaRatingColumnNum = 0;

            foreach (var row in data)
            {
                if (row.Count > 20) // If it's an empty row then it should have less than this.
                {
                    try
                    {
                        rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                        autoTitleValue = row[Convert.ToInt16(sheetVariables["Auto Title"])].ToString();
                        imdbTitleValue = row[Convert.ToInt16(sheetVariables["IMDB Title"])].ToString();
                        imdbTitleColumnNum = Convert.ToInt16(sheetVariables["IMDB Title"]);
                        sortTitleValue = row[Convert.ToInt16(sheetVariables["Sort Title"])].ToString();
                        sortTitleColumnNum = Convert.ToInt16(sheetVariables["Sort Title"]);
                        autoContentRatingValue = row[Convert.ToInt16(sheetVariables["Auto Content Rating"])].ToString();
                        contentRatingValue = row[Convert.ToInt16(sheetVariables["Content Rating"])].ToString();
                        contentRatingColumnNum = Convert.ToInt16(sheetVariables["Content Rating"]);
                        autoMpaaRatingValue = row[Convert.ToInt16(sheetVariables["Auto MPAA"])].ToString();
                        mpaaRatingValue = row[Convert.ToInt16(sheetVariables["MPAA"])].ToString();
                        mpaaRatingColumnNum = Convert.ToInt16(sheetVariables["MPAA"]);

                        if (!imdbTitleValue.Equals(""))
                        {
                            movieTitle = imdbTitleValue;
                        }
                        else if (!autoTitleValue.Equals("") && !autoTitleValue.Equals("Loading..."))
                        {
                            movieTitle = autoTitleValue;
                        }
                        else
                        {
                            movieTitle = "row num " + rowNum;
                        }

                        if (!autoContentRatingValue.Equals("") && contentRatingValue.Equals(""))
                        {
                            if (autoContentRatingValue.Equals("Loading...") || autoContentRatingValue.Equals("#N/A"))
                            {
                                DisplayMessage("warning", "Auto Content Rating has not yet loaded for: ", 0);
                                DisplayMessage("info", movieTitle, 0);
                                DisplayMessage("error", " | Try again later");
                                intContentRatingLoadingCount++;
                            }
                            else if (autoContentRatingValue.Equals("#N/A"))
                            {
                                DisplayMessage("log", "Missing Content Rating for: ", 0);
                                DisplayMessage("info", movieTitle);

                                try
                                {
                                    strCellToPutData = "Movies!" + ColumnNumToLetter(contentRatingColumnNum) + rowNum;

                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { "Not Rated" } }
                                    });
                                }
                                catch (Exception ex)
                                {
                                    DisplayMessage("error", "An error occured saving the Content Rating for: ", 0);
                                    DisplayMessage("warning", movieTitle);
                                    DisplayMessage("harderror", ex.Message);
                                }
                            }
                            else
                            {
                                DisplayMessage("log", "Missing Content Rating for: ", 0);
                                DisplayMessage("info", movieTitle);

                                try
                                {
                                    strCellToPutData = "Movies!" + ColumnNumToLetter(contentRatingColumnNum) + rowNum;

                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { autoContentRatingValue } }
                                    });
                                }
                                catch (Exception ex)
                                {
                                    DisplayMessage("error", "An error occured saving the Content Rating for: ", 0);
                                    DisplayMessage("warning", movieTitle);
                                    DisplayMessage("harderror", ex.Message);
                                }
                            }
                        }

                        if (!autoMpaaRatingValue.Equals("") && mpaaRatingValue.Equals(""))
                        {
                            if (autoMpaaRatingValue.Equals("Loading...") || autoMpaaRatingValue.Equals("#N/A"))
                            {
                                DisplayMessage("warning", "Auto MPAA has not yet loaded for: ", 0);
                                DisplayMessage("info", movieTitle, 0);
                                DisplayMessage("error", " | Try again later");
                                intMpaaRatingLoadingCount++;
                            }
                            else
                            {
                                DisplayMessage("log", "Missing MPAA for: ", 0);
                                DisplayMessage("info", movieTitle);

                                try
                                {
                                    if (autoMpaaRatingValue.Equals("#N/A"))
                                    {
                                        autoMpaaRatingValue = "X";
                                    }
                                    else if (!autoMpaaRatingValue.ToString().Contains("Rated "))
                                    {
                                        autoMpaaRatingValue = "X";
                                    }
                                    strCellToPutData = "Movies!" + ColumnNumToLetter(mpaaRatingColumnNum) + rowNum;

                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { autoMpaaRatingValue } }
                                    });
                                }
                                catch (Exception ex)
                                {
                                    DisplayMessage("error", "An error occured saving the MPAA for: ", 0);
                                    DisplayMessage("warning", movieTitle);
                                    DisplayMessage("harderror", ex.Message);
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong while copying movie data on row: " + rowNum, 0, 0, 1, "Red");
                        Type(e.Message, 0, 0, 2, "DarkRed");
                    }
                }
            }

            if (batchRequest.Data.Count > 0)
            {
                var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
            }

            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("IMDB Titles copied: " + intImdbTitlesCopiedCount, 0, 0, 1, "Green");
            Type("IMDB Titles skipped due to loading: " + intTitlesLoadingCount, 0, 0, 1, "Yellow");
            Type("Sort Titles copied: " + intSortTitlesCopiedCount, 0, 0, 1, "Green");
            Type("Sort Titles skipped due to data filled in: " + intSortTitlesSkippedCount, 0, 0, 1, "Yellow");
            Type("Content Rating copied: " + intContentRatingCopiedCount, 0, 0, 1, "Green");
            Type("Content Rating skipped due to loading: " + intContentRatingLoadingCount, 0, 0, 1, "Yellow");
            Type("MPAA Rating copied: " + intMpaaRatingCopiedCount, 0, 0, 1, "Green");
            Type("MPAA Rating skipped due to loading: " + intMpaaRatingLoadingCount, 0, 0, 2, "Yellow");

            if (intTitlesLoadingCount == 0 && intContentRatingLoadingCount == 0 && intMpaaRatingLoadingCount == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Grabs the cast for a movie from the Google Sheet.
        /// </summary>
        /// <param name="data">The movie data to run through.</param>
        /// <param name="sheetVariables">The column names to look at.</param>
        protected static void InputMovieCredits(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, Boolean overwiteData = false)
        {
            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            int intTmdbCastDoneCount = 0,
                intNoCastCount = 0,
                intPlotDoneCount = 0,
                intTmdbIdNotFoundCount = 0,
                intTmdbRatingNotFoundCount = 0,
                intPlotNotFoundCount = 0;

            string imdbIdValue = "", // Our current TMDB ID value from the Google Sheet.
                rowNum = "", // Holds the row number we are on.
                tmdbDirectorValue = "", // Our current Director value from the Google Sheet.
                tmdbCastValue = "", // Our current Cast value from the Google Sheet.
                strCellToPutData = "", // The string of the location to write the data to.
                imdbTitle = "",
                autoTitle = "",
                movieTitle = "";

            int tmdbIdColumnNum = 0, // Used to input the returned ID back into the Google Sheet.
                tmdbDirectorColumnNum = 0, // Used to input the returned Director into the Google Sheet.
                tmdbCastColumnNum = 0, // Used to input the returned Cast into the Google Sheet.
                quickCreateColumnNum = 0; // Used to mark the movies that Plot gets updated.

            dynamic tmdbResponse; // The API call response.

            foreach (var row in data)
            {
                if (row.Count > 20) // If it's an empty row then it should have less than this.
                {
                    try
                    {
                        bool MovieFoundAtTmdb = true;
                        rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                        imdbIdValue = row[Convert.ToInt16(sheetVariables["IMDB ID"])].ToString();
                        tmdbIdColumnNum = Convert.ToInt16(sheetVariables["TMDB ID"]);
                        tmdbCastValue = row[Convert.ToInt16(sheetVariables["Cast"])].ToString();
                        tmdbCastColumnNum = Convert.ToInt16(sheetVariables["Cast"]);
                        imdbTitle = row[Convert.ToInt16(sheetVariables["IMDB Title"])].ToString();
                        autoTitle = row[Convert.ToInt16(sheetVariables["Auto Title"])].ToString();
                        if (sheetVariables.ContainsKey(QUICK_CREATE)) quickCreateColumnNum = Convert.ToInt16(sheetVariables[QUICK_CREATE]);

                        if (!imdbTitle.Equals(""))
                        {
                            movieTitle = imdbTitle;
                        }
                        else if (!autoTitle.Equals("") && !autoTitle.Equals("Loading..."))
                        {
                            movieTitle = autoTitle;
                        }
                        else
                        {
                            movieTitle = "row num " + rowNum;
                        }

                        if (!imdbIdValue.Equals("") && (tmdbCastValue.Equals("") || overwiteData))
                        {
                            DisplayMessage("log", "Missing data for: ", 0);
                            DisplayMessage("info", imdbTitle);
                            DisplayMessage("warning", "Making TMDB Call... ", 0);

                            tmdbResponse = TmdbApi.MoviesGetCredits(imdbIdValue);

                            DisplayMessage("success", "DONE");

                            try
                            {
                                if (!tmdbResponse.Equals(""))
                                {
                                    DisplayMessage("success", "We found data!");
                                    var castResponse = tmdbResponse.cast;
                                    ArrayList castToPush = new ArrayList();
                                    String castStringToAdd = "";

                                    if (castResponse.Count > 0)
                                    {
                                        for (int i = 0; i < castResponse.Count; i++)
                                        {
                                            castToPush.Add(castResponse[i].name.ToString() + " - " + castResponse[i].id.ToString());
                                        }

                                        if (castToPush.Count > 0)
                                        {
                                            for (int j = 0; j < castToPush.Count; j++)
                                            {
                                                castStringToAdd += castToPush[j];
                                                if (castToPush.Count > 1 && j < castToPush.Count)
                                                {
                                                    castStringToAdd += ", ";
                                                }
                                            }
                                        }

                                        strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbCastColumnNum) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { castStringToAdd } }
                                        });
                                    }
                                    else
                                    {
                                        strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbCastColumnNum) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { "N/A" } }
                                        });

                                        DisplayMessage("warning", "No Cast to save for: " + movieTitle);
                                        intNoCastCount++;
                                    }
                                }
                                else
                                {
                                    strCellToPutData = "Movies!" + ColumnNumToLetter(tmdbCastColumnNum) + rowNum;

                                    batchRequest.Data.Add(new ValueRange
                                    {
                                        Range = strCellToPutData,
                                        MajorDimension = "ROWS",
                                        Values = new List<IList<object>> { new List<object> { "N/A" } }
                                    });

                                    DisplayMessage("error", "No record found at TMDB for: ", 0);
                                    DisplayMessage("warning", movieTitle);

                                    MovieFoundAtTmdb = false;
                                }
                            }
                            catch (Exception ex)
                            {
                                DisplayMessage("error", "No record found at TMDB for: ", 0);
                                DisplayMessage("warning", movieTitle);
                                DisplayMessage("harderror", ex.Message);

                                MovieFoundAtTmdb = false;
                            }
                        }

                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong while putting in movie data for: " + movieTitle, 0, 0, 1, "Red");
                        Type(e.Message, 0, 0, 2, "DarkRed");
                    }
                }
            }

            if (batchRequest.Data.Count > 0)
            {
                var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
            }

            Console.WriteLine();
            Type("It looks like that's the end of it.", 0, 0, 1);
            Type("TMDB Cast inserted: " + intTmdbCastDoneCount, 0, 0, 1, "Green");
            Type("No Cast to insert: " + intNoCastCount, 0, 0, 1, "Yellow");
            //Type("TMDB Ratings inserted: " + intTmdbRatingDoneCount, 0, 0, 1, "Green");
            //Type("Plots inserted: " + intPlotDoneCount, 0, 0, 1, "Green");
            //Type("TMDB IDs not available: " + intTmdbIdNotFoundCount, 0, 0, 1, "Red");
            //Type("TMDB Ratings not available: " + intTmdbRatingNotFoundCount, 0, 0, 1, "Red");
            //Type("Plots not available: " + intPlotNotFoundCount, 0, 0, 1, "Red");

        }

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
            Console.Write(new string(' ', Console.WindowWidth - 1));
            Console.SetCursorPosition(0, currentLineCursor);
        }

        protected static void ClearSelectedRowData(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, IList<IList<Object>> clearList = null)
        {
            DisplayMessage("info", "We will now look for selected rows to clear", 2);

            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            bool writeData = false;
            List<ValueRange> skipAdditions = new List<ValueRange>();

            foreach (var row in data)
            {
                if (row.Count > 55)
                {
                    string sheetQuickCreateValue = "";

                    int sheetQuickCreateInt = 0;

                    if (sheetVariables.ContainsKey(QUICK_CREATE) && row.Count > Convert.ToInt16(sheetVariables[QUICK_CREATE]))
                    {
                        sheetQuickCreateValue = row[Convert.ToInt16(sheetVariables[QUICK_CREATE])].ToString();
                        sheetQuickCreateInt = Convert.ToInt16(sheetVariables[QUICK_CREATE]);
                    }

                    if (sheetQuickCreateValue.ToUpper() == "X")
                    {
                        writeData = true;

                        // --- CLEAR MOVIE SHEET CELLS ---
                        string rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                        foreach (var entry in sheetVariables)
                        {
                            if (!entry.Key.Equals(ROW_NUM))
                            {
                                string strCellToPutData = "Movies!" + ColumnNumToLetter(entry.Value) + rowNum;
                                batchRequest.Data.Add(new ValueRange
                                {
                                    Range = strCellToPutData,
                                    MajorDimension = "ROWS",
                                    Values = new List<IList<object>> { new List<object> { "" } }
                                });
                            }
                        }



                        // --- HANDLE CLEAR LIST ---
                        if (clearList != null && sheetVariables.ContainsKey("TMDB ID"))
                        {
                            string tmdbId = row[sheetVariables["TMDB ID"]].ToString().Trim();

                            // Check if it's already in clearList (including newly added during this method execution)
                            bool existsInClearList = clearList.Any(entry => entry.Count > 0 && entry[0].ToString() == tmdbId);

                            if (!existsInClearList && !string.IsNullOrEmpty(tmdbId))
                            {
                                // Find next empty row index in clearList (starts at row 3)
                                int nextEmptyRowIndex = clearList.Count + 3;

                                // Add to the batch update request
                                skipAdditions.Add(new ValueRange
                                {
                                    Range = $"Autopopulate Actors!D{nextEmptyRowIndex}",
                                    MajorDimension = "ROWS",
                                    Values = new List<IList<object>> { new List<object> { tmdbId } }
                                });

                                // Add to clearList immediately so duplicates are avoided in this run
                                clearList.Add(new List<object> { tmdbId });
                            }
                        }

                    }
                }
            }

            if (writeData)
            {
                var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
                DisplayMessage("info", "Looks like that's it.", 2);
            }
            else
            {
                DisplayMessage("warning", "No row was selected to clear", 2);
            }

            // Append to "Autopopulate Actors" sheet if needed
            if (skipAdditions.Count > 0)
            {
                BatchUpdateValuesRequest skipRequest = new BatchUpdateValuesRequest
                {
                    ValueInputOption = "USER_ENTERED",
                    Data = skipAdditions
                };

                BulkWriteToSheet(skipRequest);
            }
        }

        protected static void FillInStreamingProviders(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, bool overwrite = false)
        {
            DisplayMessage("info", "Filling in the streaming providers", 2);

            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            IList<IList<Object>> filteredData = new List<IList<Object>>();

            bool writeData = false;
            int fullRowCount = 0;
            int currentCheckRow = 1;

            DisplayMessage("data", $"{data.Count} rows found. We will now filter them out.", 2);

            // First loop through the data to filter out the rows we won't be working with.
            foreach (var row in data)
            {
                if (row.Count > 55)
                {
                    fullRowCount++;
                    string tmdbId = row[Convert.ToInt16(sheetVariables["TMDB ID"])].ToString();
                    string streamFab = row[Convert.ToInt16(sheetVariables["StreamFab"])].ToString();
                    string resolution = row[Convert.ToInt16(sheetVariables["Resolution"])].ToString();
                    string fileType = row[Convert.ToInt16(sheetVariables["File Type"])].ToString().ToLower();

                    int resolutionValue = resolution != "" ? int.Parse(resolution.Replace("p", "")) : 0;

                    if ((tmdbId != "" && tmdbId != "N/A") && (streamFab != "Y" || resolutionValue < 1080 || fileType != "mkv"))
                    {
                        filteredData.Add(row);
                    }
                }
            }

            DisplayMessage("data", $"Of the {data.Count} rows found", 1);
            DisplayMessage("data", $"{fullRowCount} rows have been identified as rows with actual data.", 1);
            DisplayMessage("data", $"{filteredData.Count} of those rows have been identified as movies that need to be upgraded to 1080p.", 1);
            DisplayMessage("info", "We will now step through them and check for streaming providers.", 2);

            foreach (var filteredRow in filteredData)
            {
                string tmdbId = filteredRow[Convert.ToInt16(sheetVariables["TMDB ID"])].ToString();
                string streamFab = filteredRow[Convert.ToInt16(sheetVariables["StreamFab"])].ToString();
                string resolution = filteredRow[Convert.ToInt16(sheetVariables["Resolution"])].ToString();
                string recordSource = filteredRow[Convert.ToInt16(sheetVariables["Possible Record Source"])].ToString();
                string imdbTitle = filteredRow[Convert.ToInt16(sheetVariables["IMDB Title"])].ToString();
                string recordedSource = filteredRow[Convert.ToInt16(sheetVariables["Recorded Source"])].ToString();
                string fileType = filteredRow[Convert.ToInt16(sheetVariables["File Type"])].ToString();

                DisplayMessage("default", $"{currentCheckRow} of {filteredData.Count} | ", 0);
                DisplayMessage("log", "Checking for: ", 0);
                DisplayMessage("info", imdbTitle);
                currentCheckRow++;
                dynamic tmdbResponse = TmdbApi.MoviesGetWatchProviders(tmdbId);
                try
                {
                    ArrayList providers = new ArrayList();
                    if (tmdbResponse != null && tmdbResponse.GetType() != typeof(string) && tmdbResponse?.results != null && tmdbResponse.results.Count != 0)
                    {
                        // If recordedSource is not empty, I need it to match up with how it is listed in includedProviders
                        if (recordedSource != "")
                        {
                            switch (recordedSource)
                            {
                                case "Prime":
                                    recordedSource = "Amazon Prime Video";
                                    break;

                                case "Paramount+":
                                    recordedSource = "Paramount Plus Premium";
                                    break;

                                case "Disney+":
                                    recordedSource = "Disney Plus";
                                    break;

                                case "HBO Max":
                                    recordedSource = "HBO Max Amazon Channel";
                                    break;

                                case "Peacock":
                                    recordedSource = "Peacock Premium";
                                    break;

                                case "Apple TV+":
                                    recordedSource = "Apple TV";
                                    break;
                            }
                        }

                        string[] includeProviders =
                        {
                            "Amazon Prime Video",
                            "Angel Studios",
                            "Apple TV",
                            "Crunchyroll",
                            "Disney Plus",
                            "Hulu",
                            "HBO Max Amazon Channel",
                            "Netflix",
                            "Paramount Plus Premium",
                            "Paramount+ with Showtime",
                            "Peacock Premium"
                        };

                        // Other providers that I may have an interest in the future.
                        // Put them here for reference and to add to the includedProviders array if I ever do sign up for them.
                        // BritBox Amazon Channel
                        // A&E Crime Central
                        // Lifetime Movie Club
                        // History Vault
                        // Starz Amazon Channel
                        // MGM+ Amazon Channel

                        var flatrateResponse = tmdbResponse?.results?.US?.flatrate;
                        var freeResponse = tmdbResponse?.results?.US?.free;

                        var streamFabEqualsR = string.Equals(streamFab, "R", StringComparison.OrdinalIgnoreCase);
                        var fileTypeNotEqualMkv = !string.Equals(fileType, "mkv", StringComparison.OrdinalIgnoreCase);
                        var disneyOrHulu = new[] { "Disney+", "Disney Plus", "Hulu" };

                        if (flatrateResponse != null)
                        {
                            foreach (var streamer in flatrateResponse)
                            {
                                string providerName = streamer.provider_name?.ToString().Trim();

                                var sameProviderGroup =
                                    disneyOrHulu.Contains(recordedSource) &&
                                    disneyOrHulu.Contains(providerName);

                                var shouldSkipBecauseSameSource =
                                    string.Equals(providerName, recordedSource, StringComparison.OrdinalIgnoreCase) ||
                                    sameProviderGroup;

                                if ((!shouldSkipBecauseSameSource || streamFabEqualsR || fileTypeNotEqualMkv) &&
                                    includeProviders.Contains(providerName))
                                {
                                    providers.Add(providerName);
                                }
                            }
                        }

                        if (freeResponse != null)
                        {
                            foreach (var streamer in freeResponse)
                            {
                                string providerName = streamer.provider_name?.ToString().Trim();

                                var sameProviderGroup =
                                    disneyOrHulu.Contains(recordedSource) &&
                                    disneyOrHulu.Contains(providerName);

                                var shouldSkipBecauseSameSource =
                                    string.Equals(providerName, recordedSource, StringComparison.OrdinalIgnoreCase) ||
                                    sameProviderGroup;

                                if ((!shouldSkipBecauseSameSource || streamFabEqualsR || fileTypeNotEqualMkv) &&
                                    includeProviders.Contains(providerName))
                                {
                                    providers.Add(providerName);
                                }
                            }
                        }

                        var list = String.Join(",", providers.ToArray());
                        if (recordSource != list)
                        {
                            // TODO: Create a list that we will push this data to and then show the user once this method finishes running.
                            // It will work similar to how the countFiles method works.
                            DisplayMessage("log", "Updated: ", 0);
                            DisplayMessage("warning", recordSource, 0);
                            DisplayMessage("log", " to ", 0);
                            DisplayMessage("success", list, 0);
                            DisplayMessage("log", " for: ", 0);
                            DisplayMessage("data", imdbTitle);
                            string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Possible Record Source"])) + filteredRow[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                            writeData = true;
                            batchRequest.Data.Add(new ValueRange
                            {
                                Range = strCellToPutData,
                                MajorDimension = "ROWS",
                                Values = new List<IList<object>> { new List<object> { list } }
                            });
                        }
                    }
                    else
                    {
                        DisplayMessage("info", "An error occurred when getting TMDB data for: ", 0);
                        DisplayMessage("log", imdbTitle);
                        DisplayMessage("warning", tmdbResponse);
                    }
                }
                catch (Exception e)
                {
                    Type("Something went wrong grabbing the streaming provider for: " + imdbTitle, 0, 0, 1, "Red");
                    Type(e.Message, 0, 0, 2, "DarkRed");
                }
            }

            if (writeData)
            {
                var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
                DisplayMessage("info", "Looks like that's it.", 2);
            }
            else
            {
                DisplayMessage("warning", "No streaming providers needed to be updated", 2);
            }
        }

        protected static void FillInTvShowStreamingProviders(IList<IList<Object>> data, Dictionary<string, int> sheetVariables)
        {
            DisplayMessage("info", "Filling in the TV Show possible record sources", 2);

            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>()
            };

            IList<IList<Object>> filteredData = new List<IList<Object>>();

            bool writeData = false;
            int fullRowCount = 0;
            int currentCheckRow = 1;

            DisplayMessage("data", $"{data.Count} rows found. We will now filter them out.", 2);

            foreach (var row in data)
            {
                string tvdbId = GetSheetCell(row, sheetVariables, "TVDB ID");
                if (string.IsNullOrWhiteSpace(tvdbId))
                {
                    continue;
                }

                fullRowCount++;

                string streamFab = GetSheetCell(row, sheetVariables, "StreamFab");
                string resolution = GetSheetCell(row, sheetVariables, "Resolution");
                string tvShowFormat = GetSheetCell(row, sheetVariables, "Format");

                bool streamFabRecorded = IsStreamFabRecorded(streamFab);
                bool resolutionIs1080p = GetResolutionValue(resolution) == 1080;
                bool formatIsMkv = string.Equals(tvShowFormat, "mkv", StringComparison.OrdinalIgnoreCase);

                if (!streamFabRecorded || !resolutionIs1080p || !formatIsMkv)
                {
                    filteredData.Add(row);
                }
            }

            DisplayMessage("data", $"Of the {data.Count} rows found", 1);
            DisplayMessage("data", $"{fullRowCount} rows have been identified as rows with actual TV Show data.", 1);
            DisplayMessage("data", $"{filteredData.Count} of those rows need possible record sources checked.", 1);
            DisplayMessage("info", "We will now step through them and check for streaming providers.", 2);

            foreach (var filteredRow in filteredData)
            {
                string tvdbId = GetSheetCell(filteredRow, sheetVariables, "TVDB ID");
                string recordSource = GetSheetCell(filteredRow, sheetVariables, "Possible Record Source");
                string seriesName = GetSheetCell(filteredRow, sheetVariables, "Series Name");
                string cleanNameWithYear = GetSheetCell(filteredRow, sheetVariables, "Clean Name with Year");
                string rowNum = GetSheetCell(filteredRow, sheetVariables, ROW_NUM);

                if (string.IsNullOrWhiteSpace(seriesName))
                {
                    seriesName = cleanNameWithYear;
                }

                DisplayMessage("default", $"{currentCheckRow} of {filteredData.Count} | ", 0);
                DisplayMessage("log", "Checking for: ", 0);
                DisplayMessage("info", seriesName);
                currentCheckRow++;

                dynamic tmdbResponse = TmdbApi.TvGetWatchProvidersByTvdbId(tvdbId);
                try
                {
                    if (tmdbResponse != null && tmdbResponse.GetType() != typeof(string) && tmdbResponse?.results != null && tmdbResponse.results.Count != 0)
                    {
                        string list = BuildIncludedStreamingProviderList(tmdbResponse, "", true);

                        if (recordSource != list)
                        {
                            DisplayMessage("log", "Updated: ", 0);
                            DisplayMessage("warning", recordSource, 0);
                            DisplayMessage("log", " to ", 0);
                            DisplayMessage("success", list, 0);
                            DisplayMessage("log", " for: ", 0);
                            DisplayMessage("data", seriesName);

                            string strCellToPutData = "DB!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Possible Record Source"])) + rowNum;
                            writeData = true;
                            batchRequest.Data.Add(new ValueRange
                            {
                                Range = strCellToPutData,
                                MajorDimension = "ROWS",
                                Values = new List<IList<object>> { new List<object> { list } }
                            });
                        }
                    }
                    else
                    {
                        DisplayMessage("info", "An error occurred when getting TMDB data for: ", 0);
                        DisplayMessage("log", seriesName);
                        DisplayMessage("warning", tmdbResponse);
                    }
                }
                catch (Exception e)
                {
                    Type("Something went wrong grabbing the streaming provider for: " + seriesName, 0, 0, 1, "Red");
                    Type(e.Message, 0, 0, 2, "DarkRed");
                }
            }

            if (writeData)
            {
                var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
                DisplayMessage("info", "Looks like that's it.", 2);
            }
            else
            {
                DisplayMessage("warning", "No TV Show possible record sources needed to be updated", 2);
            }
        }

        private static string GetSheetCell(IList<Object> row, Dictionary<string, int> sheetVariables, string columnName)
        {
            if (!sheetVariables.ContainsKey(columnName))
            {
                return "";
            }

            int columnIndex = sheetVariables[columnName];
            if (columnIndex < 0 || columnIndex >= row.Count)
            {
                return "";
            }

            return row[columnIndex]?.ToString() ?? "";
        }

        private static bool IsStreamFabRecorded(string streamFab)
        {
            return !string.IsNullOrWhiteSpace(streamFab) &&
                streamFab.Trim().StartsWith("Y", StringComparison.OrdinalIgnoreCase);
        }

        private static int GetResolutionValue(string resolution)
        {
            if (string.IsNullOrWhiteSpace(resolution))
            {
                return 0;
            }

            string digits = Regex.Replace(resolution, @"[^\d]", "");
            int resolutionValue;
            return int.TryParse(digits, out resolutionValue) ? resolutionValue : 0;
        }

        private static string BuildIncludedStreamingProviderList(dynamic tmdbResponse, string recordedSource, bool allowSameSource)
        {
            ArrayList providers = new ArrayList();
            string[] includeProviders = GetIncludedStreamingProviders();
            string normalizedRecordedSource = NormalizeRecordedSourceForProvider(recordedSource);

            var flatrateResponse = tmdbResponse?.results?.US?.flatrate;
            var freeResponse = tmdbResponse?.results?.US?.free;

            AddIncludedStreamingProviders(flatrateResponse, providers, includeProviders, normalizedRecordedSource, allowSameSource);
            AddIncludedStreamingProviders(freeResponse, providers, includeProviders, normalizedRecordedSource, allowSameSource);

            return String.Join(",", providers.ToArray());
        }

        private static void AddIncludedStreamingProviders(dynamic providerResponse, ArrayList providers, string[] includeProviders, string recordedSource, bool allowSameSource)
        {
            if (providerResponse == null)
            {
                return;
            }

            string[] disneyOrHulu = { "Disney+", "Disney Plus", "Hulu" };

            foreach (var streamer in providerResponse)
            {
                string providerName = streamer.provider_name?.ToString().Trim();

                bool sameProviderGroup =
                    disneyOrHulu.Contains(recordedSource) &&
                    disneyOrHulu.Contains(providerName);

                bool shouldSkipBecauseSameSource =
                    string.Equals(providerName, recordedSource, StringComparison.OrdinalIgnoreCase) ||
                    sameProviderGroup;

                if ((allowSameSource || !shouldSkipBecauseSameSource) &&
                    includeProviders.Contains(providerName))
                {
                    providers.Add(providerName);
                }
            }
        }

        private static string NormalizeRecordedSourceForProvider(string recordedSource)
        {
            if (string.IsNullOrWhiteSpace(recordedSource))
            {
                return "";
            }

            switch (recordedSource.Trim())
            {
                case "Prime":
                    return "Amazon Prime Video";

                case "Paramount+":
                    return "Paramount Plus Premium";

                case "Disney+":
                    return "Disney Plus";

                case "HBO Max":
                    return "HBO Max Amazon Channel";

                case "Peacock":
                    return "Peacock Premium";

                case "Apple TV+":
                    return "Apple TV";

                default:
                    return recordedSource.Trim();
            }
        }

        private static string[] GetIncludedStreamingProviders()
        {
            return new[]
            {
                "Amazon Prime Video",
                "Angel Studios",
                "Apple TV",
                "Crunchyroll",
                "Disney Plus",
                "Hulu",
                "HBO Max Amazon Channel",
                "Netflix",
                "Paramount Plus Premium",
                "Paramount+ with Showtime",
                "Peacock Premium"
            };
        }

        protected static void FillInStreamingProviderForSelectedId(Dictionary<string, int> sheetVariables, string tmdbId, string rowNum, string imdbTitle)
        {
            dynamic tmdbResponse = TmdbApi.MoviesGetWatchProviders(tmdbId);
            try
            {
                ArrayList providers = new ArrayList();
                if (tmdbResponse != null && tmdbResponse.GetType() != typeof(string) && tmdbResponse?.results != null && tmdbResponse.results.Count != 0)
                {
                    string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Possible Record Source"])) + rowNum;

                    try
                    {
                        var myResponse = tmdbResponse.results.US;
                    }
                    catch (Exception)
                    {
                        DisplayMessage("log", "No US");
                        DisplayMessage("log", tmdbResponse.results.ToString());
                    }
                    var flatrateResponse = tmdbResponse?.results?.US?.flatrate;
                    var freeResponse = tmdbResponse?.results?.US?.free;

                    if (flatrateResponse != null)
                    {
                        foreach (var streamer in flatrateResponse)
                        {
                            providers.Add(streamer.provider_name);
                        }
                    }

                    if (freeResponse != null)
                    {
                        foreach (var streamer in freeResponse)
                        {
                            providers.Add(streamer.provider_name);
                        }
                    }

                    var list = String.Join(",", providers.ToArray());
                    if (WriteSingleCellToSheet(list, strCellToPutData))
                    {
                        DisplayMessage("success", "Streamer list saved for: ", 0);
                        DisplayMessage("info", imdbTitle, 0);
                        DisplayMessage("log", " at- ", 0);
                        DisplayMessage("info", strCellToPutData);
                    }
                    else
                    {
                        DisplayMessage("error", "An error occured saving the streaming source for: ", 0);
                        DisplayMessage("warning", imdbTitle);
                    }
                }
                else
                {
                    DisplayMessage("log", "TMDB did not have a result for: " + imdbTitle);
                }
            }
            catch (Exception e)
            {
                Type("Something went wrong grabbing the streaming provider for: " + imdbTitle, 0, 0, 1, "Red");
                Type(e.Message, 0, 0, 2, "DarkRed");
            }
        }
        protected static void FillInVideoResolution(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, bool overwrite)
        {
            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            foreach (var row in data)
            {
                if (row.Count > 25)
                {
                    string rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                    string status = row[Convert.ToInt16(sheetVariables["Status"])].ToString();
                    var movieDirectory = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString();
                    string resolution = row[Convert.ToInt16(sheetVariables["Resolution"])].ToString();
                    string cleanTitle = row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString();
                    string streamFab = row[Convert.ToInt16(sheetVariables["StreamFab"])].ToString().ToUpper();

                    try
                    {
                        // If the directory path in the Google Sheet isn't empty, but the video resolution is, then let's try to find the video and add the resolution to the Google Sheet.
                        if (status != "" && movieDirectory != "" && (resolution == "" || overwrite))
                        {
                            // Let's go ahead and look for the hard drive letter now.
                            var hardDriveLetters = FindDriveLetters(movieDirectory);

                            if (Directory.Exists(movieDirectory))
                            {
                                string[] fileEntries = Directory.GetFiles(movieDirectory);

                                ArrayList videoFiles = GrabMovieFiles(fileEntries, false);

                                if (videoFiles.Count > 0)
                                {
                                    string video = "";

                                    foreach (var videoFile in videoFiles)
                                    {
                                        if (!Path.GetFileName(videoFile.ToString()).Contains("-trailer"))
                                        {
                                            video = videoFile.ToString();
                                        }
                                    }

                                    if (!video.Equals(""))
                                    {
                                        string calculatedResolution = GetVideoResolution(Path.Combine(movieDirectory, video));

                                        if (calculatedResolution != "N/A")
                                        {
                                            if (resolution != calculatedResolution)
                                            {
                                                try
                                                {
                                                    string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Resolution"])) + rowNum;
                                                    batchRequest.Data.Add(new ValueRange
                                                    {
                                                        Range = strCellToPutData,
                                                        MajorDimension = "ROWS",
                                                        Values = new List<IList<object>> { new List<object> { calculatedResolution } }
                                                    });

                                                    DisplayMessage("success", "Resolution ", 0);
                                                    DisplayMessage("info", calculatedResolution, 0);
                                                    if (overwrite)
                                                    {
                                                        DisplayMessage("success", " overwritten from: ", 0);
                                                        DisplayMessage("info", resolution, 0);
                                                        DisplayMessage("success", " for: ", 0);
                                                    }
                                                    else DisplayMessage("success", " saved for: ", 0);

                                                    DisplayMessage("info", cleanTitle, 0);
                                                    DisplayMessage("log", " at- ", 0);
                                                    DisplayMessage("info", strCellToPutData);
                                                }
                                                catch (Exception ex)
                                                {
                                                    DisplayMessage("error", "An error occured saving the resolution for: ", 0);
                                                    DisplayMessage("warning", cleanTitle);
                                                    DisplayMessage("harderror", ex.Message);
                                                    throw;
                                                }
                                            }
                                            else
                                            {
                                                DisplayMessage("info", resolution, 0);
                                                DisplayMessage("default", " Resolution for ", 0);
                                                DisplayMessage("info", cleanTitle, 0);
                                                DisplayMessage("default", " is correct");
                                            }
                                        }
                                        else
                                        {
                                            // calculatedResolution came back "N/A" so show the movie title that failed.
                                            Type("calculatedResolution failed for: ", 0, 0, 0, "Red");
                                            Type(cleanTitle, 0, 0, 2, "DarkRed");
                                        }
                                    }
                                    else
                                    {
                                        Type("We could not find the actual video file for: ", 0, 0, 0, "Red");
                                        Type(cleanTitle, 0, 0, 2, "DarkRed");
                                    }
                                }
                                else
                                {
                                    Type("No videos available in: ", 0, 0, 0, "Blue");
                                    Type(movieDirectory, 0, 0, 1);
                                }
                            }
                            else
                            {
                                Type("We did not find the hard drive for: ", 0, 0, 0, "Red");
                                Type(cleanTitle, 0, 0, 1, "Yellow");
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong grabbing the video resolution for: " + cleanTitle, 0, 0, 1, "Red");
                        Type(e.Message, 0, 0, 2, "DarkRed");
                        break;
                    }
                }
            } // End foreach
            var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
            DisplayMessage("info", "Looks like that's it.", 2);
        }

        private static async Task<string> GetVideoDurationAsync(string videoId)
        {
            var youtubeService = new YouTubeService(new BaseClientService.Initializer()
            {
                ApiKey = YOUTUBE_API_KEY,
                ApplicationName = "MovieTrailerDownloader"
            });

            var videoRequest = youtubeService.Videos.List("contentDetails");
            videoRequest.Id = videoId;
            var videoResponse = await videoRequest.ExecuteAsync();

            if (videoResponse.Items.Count > 0)
            {
                // Extract video duration in ISO 8601 format (e.g., PT2M30S for 2 minutes, 30 seconds)
                var duration = videoResponse.Items[0].ContentDetails.Duration;
                return duration;
            }

            return null;
        }

        public static bool IsTrailerLength(string isoDuration)
        {
            // Convert ISO 8601 duration to TimeSpan
            var duration = XmlConvert.ToTimeSpan(isoDuration);

            // Consider trailers under 4 minutes as valid
            return duration.TotalMinutes < 4;
        }

        async protected static void SearchAndDownloadMovieTrailers(IList<IList<Object>> data, Dictionary<string, int> sheetVariables)
        {
            foreach (var row in data)
            {
                if (row.Count > 25)
                {
                    string rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                    var movieDirectory = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString();
                    string cleanTitle = row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString();
                    string youtubeTrailerId = row[Convert.ToInt16(sheetVariables["YouTube Trailer ID"])].ToString();
                    string status = row[Convert.ToInt16(sheetVariables["Status"])].ToString();

                    try
                    {
                        // If the status, and movie directory are not empty, but the youtube trailer id is N/A then we will look for a trailer to download.
                        if (status != "" && movieDirectory != "" && youtubeTrailerId == "N/A")
                        {
                            Type("It looks like ", 0, 0, 0, "Blue");
                            Type(cleanTitle, 0, 0, 0, "Green");
                            Type(" is missing a trailer.", 0, 0, 1, "Blue");
                            // Let's go ahead and look for the hard drive letter now.
                            var hardDriveLetters = FindDriveLetters(movieDirectory);

                            Type("We will now verify the movie has a directory.", 0, 0, 1, "Blue");
                            // Even though the directory is filled in on the Google sheet, we will verify we can find it on the computer.
                            if (Directory.Exists(movieDirectory))
                            {
                                Type("Directory found. Now checking for an existing trailer.", 0, 0, 1, "Green");
                                SanitizeTrailerFilename(movieDirectory, cleanTitle);
                                string[] fileEntries = Directory.GetFiles(movieDirectory);

                                ArrayList videoFiles = GrabMovieFiles(fileEntries);

                                bool foundTrailer = false;

                                // We will still check for an existing trailer in the directory before going to download one.
                                foreach (var videoFile in videoFiles)
                                {
                                    if (IsTrailerFileForTitle(videoFile.ToString(), cleanTitle))
                                    {
                                        foundTrailer = true;
                                    }
                                }

                                if (!foundTrailer)
                                {
                                    Type("No trailer found. We will now go and get one.", 0, 0, 1, "Blue");
                                    var youtubeService = new YouTubeService(new BaseClientService.Initializer()
                                    {
                                        ApiKey = YOUTUBE_API_KEY,
                                        ApplicationName = "MovieTrailerDownloader"
                                    });

                                    var searchRequest = youtubeService.Search.List("snippet");
                                    searchRequest.Q = cleanTitle + " official trailer";
                                    searchRequest.MaxResults = 2; // Fetch more results
                                    searchRequest.Type = "video";

                                    var searchResponse = await searchRequest.ExecuteAsync();

                                    string trailerURL = "";
                                    var videoId = "";
                                    bool trailerFound = false;
                                    foreach (var item in searchResponse.Items)
                                    {
                                        // Fetch the duration of each video
                                        var duration = await GetVideoDurationAsync(item.Id.VideoId);

                                        // If the title contains 'trailer' and is under 4 minutes, assume it's a valid trailer
                                        if (item.Snippet.Title.ToLower().Contains("trailer") && IsTrailerLength(duration))
                                        {
                                            trailerURL = $"https://www.youtube.com/watch?v={item.Id.VideoId}";
                                            videoId = item.Id.VideoId;
                                            trailerFound = true;
                                        }
                                    }

                                    if (!trailerFound)
                                    {
                                        Type("I was not able to find an official trailer for: ", 0, 0, 0, "Red");
                                        Type(cleanTitle, 0, 0, 1, "Yellow");

                                        string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["YouTube Trailer ID"])) + rowNum;

                                        if (WriteSingleCellToSheet("X", strCellToPutData))
                                        {
                                        }
                                        else
                                        {
                                            DisplayMessage("error", "An error occured saving the YouTube Trailer ID for: ", 0);
                                            DisplayMessage("warning", cleanTitle);
                                        }
                                    }
                                    else
                                    {
                                        var processInfo = new ProcessStartInfo
                                        {
                                            FileName = "yt-dlp",  // Assumes yt-dlp is installed and in PATH
                                            Arguments =
                                                // Select MP4 video+audio (<=1080p), fall back to single MP4 if needed
                                                $"-f \"bv*[ext=mp4][height<=1080]+ba[ext=m4a]/b[ext=mp4]\" " +
                                                // Force final container MP4 after merge
                                                $"--merge-output-format mp4 " +
                                                // Clean, stable filename (final file will be exactly ...-trailer.mp4)
                                                $"-o \"{movieDirectory}/{cleanTitle}-trailer.%(ext)s\" " +
                                                // Progress lines (for your console progress handler)
                                                $"--newline " +
                                                // Avoid partial/fragment leftovers
                                                $"--no-part --no-keep-fragments " +
                                                // Tidy names; avoids odd chars but keeps our extension place
                                                $"--restrict-filenames " +
                                                // Overwrite existing trailer if present
                                                $"--force-overwrites " +
                                                // The URL last
                                                $"{trailerURL}",
                                            RedirectStandardOutput = true,
                                            RedirectStandardError = true,
                                            UseShellExecute = false,
                                            CreateNoWindow = true
                                        };

                                        using (var process = new Process())
                                        {
                                            process.StartInfo = processInfo;
                                            string output = "";
                                            string error = "";

                                            // Capture standard output and errors
                                            process.OutputDataReceived += (sender, args) =>
                                            {
                                                if (!string.IsNullOrWhiteSpace(args.Data))
                                                {
                                                    // yt-dlp progress lines usually contain a % somewhere in them
                                                    if (args.Data.Contains("%") || args.Data.Contains("fragment"))
                                                    {
                                                        // Overwrite same line in console
                                                        Console.Write("\r" + args.Data.PadRight(Console.WindowWidth - 1));
                                                    }
                                                    else
                                                    {
                                                        // For non-progress messages, just print normally
                                                        Console.WriteLine(args.Data);
                                                    }
                                                }
                                            };

                                            process.ErrorDataReceived += (sender, args) =>
                                            {
                                                if (args.Data != null)
                                                {
                                                    error += args.Data + Environment.NewLine;
                                                    Console.WriteLine("ERROR: " + args.Data); // Optional: print to console
                                                }
                                            };

                                            process.Start();
                                            process.BeginOutputReadLine();
                                            process.BeginErrorReadLine();

                                            process.WaitForExit(); // Wait until process is done

                                            // Check for non-zero exit code (failure)
                                            if (process.ExitCode != 0)
                                            {
                                                Console.WriteLine($"Process failed with exit code {process.ExitCode}");
                                                Console.WriteLine("Error: " + error);  // Display error message captured
                                            }
                                            else
                                            {
                                                SanitizeTrailerFilename(movieDirectory, cleanTitle);

                                                string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["YouTube Trailer ID"])) + rowNum;

                                                if (WriteSingleCellToSheet(videoId, strCellToPutData))
                                                {
                                                    DisplayMessage("success", "YouTube Trailer ID ", 0);
                                                    DisplayMessage("info", videoId, 0);
                                                    DisplayMessage("success", " saved for: ", 0);
                                                    DisplayMessage("info", cleanTitle, 0);
                                                    DisplayMessage("log", " at- ", 0);
                                                    DisplayMessage("info", strCellToPutData);
                                                }
                                                else
                                                {
                                                    DisplayMessage("error", "An error occured saving the YouTube Trailer ID for: ", 0);
                                                    DisplayMessage("warning", cleanTitle);
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    Type(cleanTitle, 0, 0, 0);
                                    Type(" already has a movie trailer.", 0, 0, 0, "green");
                                    Type(" We will skip this one.", 0, 0, 1, "blue");
                                }
                            }
                            else
                            {
                                Type("We did not find the hard drive for: ", 0, 0, 0, "Red");
                                Type(cleanTitle, 0, 0, 1, "Yellow");
                            }
                            Type("---------------------------------", 0, 0, 2, "Gray");
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong grabbing the movie trailer for: ", 0, 0, 0, "Red");
                        Type(cleanTitle, 0, 0, 1, "Yellow");
                        Type(e.Message, 0, 0, 2, "DarkRed");
                        break;
                    }
                }
            } // End foreach
        }

        protected static void DownloadMovieTrailers(IList<IList<Object>> data, Dictionary<string, int> sheetVariables)
        {
            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            int movieTrailerDownloadedCount = 0;
            int movieAlreadyHasATrailerCount = 0;
            int errorDownloadingTrailerCount = 0;

            foreach (var row in data)
            {
                if (row.Count > 25)
                {
                    string rowNum = row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString();
                    var movieDirectory = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString();
                    string cleanTitle = row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString();
                    string youtubeTrailerId = row[Convert.ToInt16(sheetVariables["YouTube Trailer ID"])].ToString();
                    string movieHasTrailerData = row[Convert.ToInt16(sheetVariables["Movie Has Trailer"])].ToString();
                    int movieHasTrailerColumnNum = Convert.ToInt16(sheetVariables["Movie Has Trailer"]);
                    string status = row[Convert.ToInt16(sheetVariables["Status"])].ToString();

                    try
                    {
                        // If movieHasTrailerData says that the directory has a trailer, then we will ignore it.
                        if (movieHasTrailerData == "X")
                        {
                            //Type(cleanTitle, 0, 0, 0);
                            //Type(" already has a movie trailer.", 0, 0, 0, "green");
                            //Type(" We will skip this one.", 0, 0, 2, "blue");
                            movieAlreadyHasATrailerCount++;
                        }
                        // If movieHasTrailerData says N/A, then we will not waste the resources to try and download the video.
                        // We will have to fix the YouTube ID to a valid one in order to fix it.
                        else if (movieHasTrailerData == "N/A")
                        {
                            //Type(cleanTitle, 0, 0, 0);
                            //Type(" has errored before.", 0, 0, 0, "yellow");
                            //Type(" We will skip this one.", 0, 0, 2, "blue");
                            errorDownloadingTrailerCount++;
                        }
                        else if (youtubeTrailerId == "X")
                        {
                        }
                        else
                        {
                            // If the status, movie directory and YouTube trailer id are not empty then we will check the directory for a trailer.
                            if (status != "" && movieDirectory != "" && (youtubeTrailerId != "" && youtubeTrailerId != "N/A" && youtubeTrailerId != "X") && movieHasTrailerData == "")
                            {
                                Type("Looking for a trailer in the directory for the following movie: ", 0, 0, 0, "Yellow");
                                Type(cleanTitle, 0, 0, 1, "Blue");
                                // Let's go ahead and look for the hard drive letter now.
                                var hardDriveLetters = FindDriveLetters(movieDirectory);

                                // Even though the directory is filled in on the Google sheet, we will verify we can find it on the computer.
                                if (Directory.Exists(movieDirectory))
                                {
                                    Type("Directory found. Will now check for an existing trailer.", 0, 0, 1, "Green");
                                    SanitizeTrailerFilename(movieDirectory, cleanTitle);
                                    string[] fileEntries = Directory.GetFiles(movieDirectory);

                                    ArrayList videoFiles = GrabMovieFiles(fileEntries);

                                    bool foundTrailer = false;

                                    // Now check for an existing trailer in the directory.
                                    foreach (var videoFile in videoFiles)
                                    {
                                        if (IsTrailerFileForTitle(videoFile.ToString(), cleanTitle))
                                        {
                                            foundTrailer = true;
                                        }
                                    }

                                    if (!foundTrailer)
                                    {
                                        Type("No trailer found. We will now go download it.", 0, 0, 1, "Blue");
                                        string trailerURL = "https://www.youtube.com/watch?v=" + youtubeTrailerId;

                                        var processInfo = new ProcessStartInfo
                                        {
                                            FileName = "yt-dlp",  // Assumes yt-dlp is installed and in PATH
                                            Arguments =
                                                // Select MP4 video+audio (<=1080p), fall back to single MP4 if needed
                                                $"-f \"bv*[ext=mp4][height<=1080]+ba[ext=m4a]/b[ext=mp4]\" " +
                                                // Force final container MP4 after merge
                                                $"--merge-output-format mp4 " +
                                                // Clean, stable filename (final file will be exactly ...-trailer.mp4)
                                                $"-o \"{movieDirectory}/{cleanTitle}-trailer.%(ext)s\" " +
                                                // Progress lines (for your console progress handler)
                                                $"--newline " +
                                                // Avoid partial/fragment leftovers
                                                $"--no-part --no-keep-fragments " +
                                                // Tidy names; avoids odd chars but keeps our extension place
                                                $"--restrict-filenames " +
                                                // Overwrite existing trailer if present
                                                $"--force-overwrites " +
                                                // The URL last
                                                $"{trailerURL}",
                                            RedirectStandardOutput = true,
                                            RedirectStandardError = true,
                                            UseShellExecute = false,
                                            CreateNoWindow = true
                                        };

                                        using (var process = new Process())
                                        {
                                            process.StartInfo = processInfo;
                                            string output = "";
                                            string error = "";

                                            // Capture standard output and errors
                                            process.OutputDataReceived += (sender, args) =>
                                            {
                                                if (!string.IsNullOrWhiteSpace(args.Data))
                                                {
                                                    // yt-dlp progress lines usually contain a % somewhere in them
                                                    if (args.Data.Contains("%") || args.Data.Contains("fragment"))
                                                    {
                                                        // Overwrite same line in console
                                                        Console.Write("\r" + args.Data.PadRight(Console.WindowWidth - 1));
                                                    }
                                                    else
                                                    {
                                                        // For non-progress messages, just print normally
                                                        Console.WriteLine(args.Data);
                                                    }
                                                }
                                            };

                                            process.ErrorDataReceived += (sender, args) =>
                                            {
                                                if (args.Data != null)
                                                {
                                                    error += args.Data + Environment.NewLine;
                                                    Type("Error: " + args.Data, 0, 0, 2, "DarkRed");
                                                }
                                            };

                                            process.Start();
                                            process.BeginOutputReadLine();
                                            process.BeginErrorReadLine();

                                            process.WaitForExit(); // Wait until process is done

                                            // Check for non-zero exit code (failure)
                                            if (process.ExitCode != 0)
                                            {
                                                Type("Something went wrong grabbing the movie trailer for: ", 0, 0, 0, "Red");
                                                Type(cleanTitle, 0, 0, 1, "Yellow");
                                                Type($"Process failed with exit code {process.ExitCode}", 0, 0, 1, "Red");
                                                Type("Error: " + error, 0, 0, 1, "DarkRed");
                                                errorDownloadingTrailerCount++;

                                                string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Movie Has Trailer"])) + rowNum;

                                                batchRequest.Data.Add(new ValueRange
                                                {
                                                    Range = strCellToPutData,
                                                    MajorDimension = "ROWS",
                                                    Values = new List<IList<object>> { new List<object> { "N/A" } }
                                                });

                                                DisplayMessage("success", "Recorded in the Google sheet that the trailer for ", 0);
                                                DisplayMessage("info", cleanTitle, 0);
                                                DisplayMessage("success", " failed to download at- ", 0);
                                                DisplayMessage("log", strCellToPutData, 2);

                                                // Sometimes when the download fails, it will still keep the old broken files.
                                                // We want to make sure those broken files are removed, otherwise we will think that the movie has a trailer when it really doesn't
                                                CleanUpBadTrailerFiles(movieDirectory, cleanTitle);
                                            }
                                            else
                                            {
                                                DisplayMessage("success", "Movie trailer downloaded for: ", 0);
                                                DisplayMessage("info", cleanTitle, 1);
                                                movieTrailerDownloadedCount++;

                                                string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Movie Has Trailer"])) + rowNum;

                                                batchRequest.Data.Add(new ValueRange
                                                {
                                                    Range = strCellToPutData,
                                                    MajorDimension = "ROWS",
                                                    Values = new List<IList<object>> { new List<object> { "X" } }
                                                });

                                                DisplayMessage("success", "Successfully recorded in the Google sheet that ", 0);
                                                DisplayMessage("info", cleanTitle, 0);
                                                DisplayMessage("success", " now has a trailer at- ", 0);
                                                DisplayMessage("log", strCellToPutData, 2);

                                                // Now go and verify yt-dlp didn't save an awful name i.e. 1 Mile To You (2017)-trailer.f137.mp4
                                                SanitizeTrailerFilename(movieDirectory, cleanTitle);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Type(cleanTitle, 0, 0, 0);
                                        Type(" already has a movie trailer.", 0, 0, 0, "green");
                                        Type(" We will skip this one.", 0, 0, 1, "blue");
                                        movieAlreadyHasATrailerCount++;

                                        string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Movie Has Trailer"])) + rowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { "X" } }
                                        });

                                        DisplayMessage("success", "Successfully recorded in the Google sheet that ", 0);
                                        DisplayMessage("info", cleanTitle, 0);
                                        DisplayMessage("success", " already has a trailer at- ", 0);
                                        DisplayMessage("log", strCellToPutData, 2);
                                    }
                                }
                                else
                                {
                                    Type("We did not find the hard drive for: ", 0, 0, 0, "Red");
                                    Type(cleanTitle, 0, 0, 2, "Yellow");
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong grabbing the movie trailer for: ", 0, 0, 0, "Red");
                        Type(cleanTitle, 0, 0, 1, "Yellow");
                        Type(e.Message, 0, 0, 2, "DarkRed");
                        errorDownloadingTrailerCount++;
                    }
                }

                if (batchRequest.Data.Count > 25)
                {
                    var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);
                    batchRequest.Data.Clear();
                }
            } // End foreach

            var BatchUpdateValuesResponse2 = BulkWriteToSheet(batchRequest);

            Type("-----SUMMARY-----", 0, 0, 1);
            Type("Trailers downloaded: " + movieTrailerDownloadedCount, 0, 0, 1, "Green");
            Type("Movies already had a trailer: " + movieAlreadyHasATrailerCount, 0, 0, 1, "Yellow");
            Type("Errors downloading trailer: " + errorDownloadingTrailerCount, 0, 0, 1, "Red");

            Type("It looks like that's the end of it.", 0, 0, 2, "magenta");
        }

        protected static void CleanUpBadTrailerFiles(string movieDirectory, string cleanTitle)
        {
            // Look for any trailer-like files for this movie, including old spaced names.
            var toDelete = FindTrailerFilesForTitle(movieDirectory, cleanTitle).ToList();

            if (toDelete.Count == 0) return; // All good

            foreach (var path in toDelete)
            {
                try
                {
                    File.Delete(path);
                }
                catch (Exception ex)
                {
                    Type($"[CLEANUP] Failed to delete '{path}': {ex.Message}");
                }
            }
        }

        protected static void SanitizeTrailerFilename(string movieDirectory, string cleanTitle)
        {
            var candidates = FindTrailerFilesForTitle(movieDirectory, cleanTitle)
                .Where(IsTrailerVideoFile)
                .ToList();

            if (candidates.Count == 0) return; // nothing to fix

            var correctFiles = candidates
                .Where(p => IsExpectedTrailerFileName(p, cleanTitle))
                .ToList();

            if (correctFiles.Count > 0)
            {
                var extraCandidates = candidates
                    .Where(p => !correctFiles.Contains(p, StringComparer.OrdinalIgnoreCase))
                    .ToList();

                foreach (var path in extraCandidates)
                {
                    try
                    {
                        File.Delete(path);
                        Type($"[SANITIZE] Deleted extra trailer-like file: {Path.GetFileName(path)}");
                    }
                    catch (Exception ex)
                    {
                        Type($"[SANITIZE] Failed to delete extra trailer-like file '{Path.GetFileName(path)}': {ex.Message}");
                    }
                }

                return; // a Plex-friendly trailer already exists
            }

            if (candidates.Count == 1)
            {
                string extension = Path.GetExtension(candidates[0]);
                string expected = Path.Combine(movieDirectory, $"{cleanTitle}-trailer{extension}");

                try
                {
                    File.Move(candidates[0], expected);
                    Type($"[SANITIZE] Trailer renamed to: {Path.GetFileName(expected)}");
                }
                catch (Exception ex)
                {
                    Type($"[SANITIZE] Failed to sanitize trailer filename: {ex.Message}");
                }
            }
            else if (candidates.Count > 1)
            {
                Type($"[SANITIZE] Multiple trailer-like files found; not renaming: {string.Join(", ", candidates.Select(Path.GetFileName))}");
            }
        }

        private static IEnumerable<string> FindTrailerFilesForTitle(string movieDirectory, string cleanTitle)
        {
            if (string.IsNullOrWhiteSpace(movieDirectory) ||
                string.IsNullOrWhiteSpace(cleanTitle) ||
                !Directory.Exists(movieDirectory))
            {
                return Enumerable.Empty<string>();
            }

            Regex trailerRegex = BuildTrailerFileNameRegex(cleanTitle);

            return Directory.EnumerateFiles(movieDirectory, "*.*", SearchOption.TopDirectoryOnly)
                .Where(p => trailerRegex.IsMatch(Path.GetFileNameWithoutExtension(p)));
        }

        private static Regex BuildTrailerFileNameRegex(string cleanTitle)
        {
            string escapedTitle = Regex.Escape(cleanTitle.Trim());

            return new Regex(
                "^" + escapedTitle + @"\s*-\s*trailer(?:[\s._-]*[A-Za-z0-9]+)?$",
                RegexOptions.IgnoreCase);
        }

        private static bool IsTrailerFileForTitle(string path, string cleanTitle)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(cleanTitle))
                return false;

            return BuildTrailerFileNameRegex(cleanTitle)
                .IsMatch(Path.GetFileNameWithoutExtension(path));
        }

        private static bool IsExpectedTrailerFileName(string path, string cleanTitle)
        {
            string baseName = Path.GetFileNameWithoutExtension(path);
            return string.Equals(baseName, cleanTitle + "-trailer", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTrailerVideoFile(string path)
        {
            string extension = Path.GetExtension(path);

            return extension.Equals(".mp4", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".mkv", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".m4v", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".avi", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".webm", StringComparison.OrdinalIgnoreCase);
        }

        protected static Boolean FillInActorMovieCredits(IList<IList<Object>> movieSheetData, Dictionary<string, int> sheetVariables, dynamic actorMovieCredits, ref IList<IList<Object>> skipMovieIdsData, string message)
        {
            BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            BatchUpdateValuesRequest skipMovieIdsBatchRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = new List<ValueRange>() // Initialize the list
            };

            bool movieFound = false;
            bool movieIdAddedToSkipColumn = false, movieAddedToMoviesSheet = false;
            int initialEmptyRowNum = -1;
            int skipMovieIdsDataEmptyRow = skipMovieIdsData.Count + 3;

            // Loop through each movie in the actors list to get the TMDB ID and check if the movie is already in our Movie sheet.
            foreach (dynamic movie in actorMovieCredits)
            {
                string movieTitle = movie.original_title.ToString(),
                        movieId = movie.id.ToString(),
                        releaseDate = movie.release_date.ToString();
                bool continueToGetDetails = true;

                string year = releaseDate.Length >= 4 ? releaseDate.Substring(0, 4) : "Unknown";

                string fullMovieInfo = $"{movieTitle} ({year}) - {movieId}";



                if (skipMovieIdsData.Any(inner => inner.Contains(movieId)))
                {
                    DisplayMessage("default", message, 0);
                    DisplayMessage("info", "We will skip ", 0);
                    DisplayMessage("data", fullMovieInfo, 0);
                    DisplayMessage("info", " because it was found in the list of IDs to skip.");

                    continueToGetDetails = false;
                }

                if (continueToGetDetails)
                {
                    if (movieTitle.Contains("WWE"))
                    {
                        DisplayMessage("default", message, 0);
                        DisplayMessage("info", "We will skip ", 0);
                        DisplayMessage("data", fullMovieInfo, 0);
                        DisplayMessage("info", " because it has been identified as WWE.");
                        continueToGetDetails = false;
                    }
                }

                if (continueToGetDetails)
                {
                    if (movieTitle.Contains("WWF"))
                    {
                        DisplayMessage("default", message, 0);
                        DisplayMessage("info", "We will skip ", 0);
                        DisplayMessage("data", fullMovieInfo, 0);
                        DisplayMessage("info", " because it has been identified as WWF.");
                        continueToGetDetails = false;
                    }
                }

                if (continueToGetDetails)
                {
                    if (movieTitle.ToLower().Contains("wrestlemania"))
                    {
                        DisplayMessage("default", message, 0);
                        DisplayMessage("info", "We will skip ", 0);
                        DisplayMessage("data", fullMovieInfo, 0);
                        DisplayMessage("info", " because it has been identified as wrestlemania.");
                        continueToGetDetails = false;
                    }
                }

                if (continueToGetDetails)
                {
                    foreach (var genre in movie.genre_ids)
                    {
                        if (genre == 99)
                        {
                            DisplayMessage("default", message, 0);
                            DisplayMessage("info", "We will skip ", 0);
                            DisplayMessage("data", fullMovieInfo, 0);
                            DisplayMessage("info", " because it has been identified as a documentary.");
                            continueToGetDetails = false;

                            string strCellToPutData = "Autopopulate Actors!D" + skipMovieIdsDataEmptyRow;
                            skipMovieIdsBatchRequest.Data.Add(new ValueRange
                            {
                                Range = strCellToPutData,
                                MajorDimension = "ROWS",
                                Values = new List<IList<object>> { new List<object> { movieId } }
                            });

                            skipMovieIdsDataEmptyRow++;
                            movieIdAddedToSkipColumn = true;
                        }
                    }
                }

                if (continueToGetDetails)
                {
                    try
                    {
                        movieFound = false;
                        DisplayMessage("default", message, 0);
                        DisplayMessage("warning", "Searching Movie sheet for ", 0);
                        DisplayMessage("info", fullMovieInfo);
                        // Loop through each row of the Movie sheet to check if the ID is in there.
                        foreach (var row in movieSheetData)
                        {
                            if (row.Count > 70)
                            {
                                string tmdbIdValue = row[Convert.ToInt16(sheetVariables["TMDB ID"])].ToString();

                                if (tmdbIdValue == movie.id.ToString())
                                {
                                    movieFound = true;
                                    DisplayMessage("default", message, 0);
                                    DisplayMessage("success", "Movie found! Moving to next movie.");
                                    break;
                                }
                            }
                            else
                            {
                                DisplayMessage("default", message, 0);
                                DisplayMessage("warning", "Movie NOT found! Moving on to add it.");
                                initialEmptyRowNum = initialEmptyRowNum > 0 ? initialEmptyRowNum : int.Parse(row[Convert.ToInt16(sheetVariables[ROW_NUM])].ToString());
                                break;
                            }
                        }

                        // If the movie was not found, movieFound should still be false, so now let's add it to the sheet.
                        if (!movieFound)
                        {
                            int movieSheetRowNum = initialEmptyRowNum - 3;

                            // Verify that the initialEmptyRowNum is actually an empty row.
                            // For now, I am just going to have the app error and break if it tries to write to a row that isn't empty,
                            // But someday I'd like it to dynamically see that it isn't an empty row and automatically go find the next one.
                            if (movieSheetData[movieSheetRowNum].Count > 70)
                            {
                                DisplayMessage("warning", "We caught the app trying to write to a row that isn't empty");
                                DisplayMessage("default", "Row ", 0);
                                DisplayMessage("info", initialEmptyRowNum.ToString(), 0);
                                DisplayMessage("default", " is not empty");
                                DisplayMessage("error", "We will exit the app right now. Please make sure your Google Sheet is sorted");
                                throw new InvalidOperationException();
                            }

                            dynamic movieDetails = TmdbApi.MoviesGetDetailsByTmdbId(movie.id.ToString());
                            if (!movieDetails.Equals(""))
                            {
                                if (movieDetails.imdb_id != null && movieDetails.imdb_id != "")
                                {
                                    if (movieDetails.original_language.ToString() != "en")
                                    {
                                        DisplayMessage("default", message, 0);
                                        DisplayMessage("info", "We will skip ", 0);
                                        DisplayMessage("data", fullMovieInfo, 0);
                                        DisplayMessage("info", " because the original language is not English.");

                                        string strCellToPutData = "Autopopulate Actors!D" + skipMovieIdsDataEmptyRow;
                                        skipMovieIdsBatchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { movieId } }
                                        });

                                        skipMovieIdsDataEmptyRow++;
                                        movieIdAddedToSkipColumn = true;
                                    }
                                    else if (movieDetails.status.ToString() == "Rumored")
                                    {
                                        DisplayMessage("default", message, 0);
                                        DisplayMessage("info", "We will skip ", 0);
                                        DisplayMessage("data", fullMovieInfo, 0);
                                        DisplayMessage("info", " because the status of the movie is rumored.");
                                    }
                                    else if (movieDetails.status.ToString() == "Planned")
                                    {
                                        DisplayMessage("default", message, 0);
                                        DisplayMessage("info", "We will skip ", 0);
                                        DisplayMessage("data", fullMovieInfo, 0);
                                        DisplayMessage("info", " because the status of the movie is planned.");
                                    }
                                    else if (movieDetails.status.ToString() == "Canceled")
                                    {
                                        DisplayMessage("default", message, 0);
                                        DisplayMessage("info", "We will skip ", 0);
                                        DisplayMessage("data", fullMovieInfo, 0);
                                        DisplayMessage("info", " because the status of the movie was canceled.");
                                    }
                                    else if (movieDetails.status.ToString() == "In Production")
                                    {
                                        DisplayMessage("default", message, 0);
                                        DisplayMessage("info", "We will skip ", 0);
                                        DisplayMessage("data", fullMovieInfo, 0);
                                        DisplayMessage("info", " because the status of the movie is in production.");
                                    }
                                    else if (movieDetails.status.ToString() == "Post Production")
                                    {
                                        DisplayMessage("default", message, 0);
                                        DisplayMessage("info", "We will skip ", 0);
                                        DisplayMessage("data", fullMovieInfo, 0);
                                        DisplayMessage("info", " because the status of the movie is in post-production.");
                                    }
                                    else if (movieDetails.release_date.ToString() == "")
                                    {
                                        DisplayMessage("default", message, 0);
                                        DisplayMessage("info", "We will skip ", 0);
                                        DisplayMessage("data", fullMovieInfo, 0);
                                        DisplayMessage("info", " because there is no release date.");
                                    }
                                    else
                                    {
                                        string runTime = movieDetails.runtime.ToString("D3");
                                        string strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Runtime"])) + initialEmptyRowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { runTime } }
                                        });

                                        string title = movieDetails.original_title.ToString() + $" ({year})";
                                        strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["IMDB Title"])) + initialEmptyRowNum;

                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { title } }
                                        });

                                        string sortTitle = title;
                                        if (sortTitle.Substring(0, 4) == "The ")
                                        {
                                            sortTitle = sortTitle.Substring(4);
                                        }

                                        strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["Sort Title"])) + initialEmptyRowNum;
                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { sortTitle } }
                                        });

                                        string imdbUrl = "https://www.imdb.com/title/" + movieDetails.imdb_id.ToString();

                                        strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["IMDB URL"])) + initialEmptyRowNum;
                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { imdbUrl } }
                                        });

                                        string tmdbId = movie.id.ToString();

                                        strCellToPutData = "Movies!" + ColumnNumToLetter(Convert.ToInt16(sheetVariables["TMDB ID"])) + initialEmptyRowNum;
                                        batchRequest.Data.Add(new ValueRange
                                        {
                                            Range = strCellToPutData,
                                            MajorDimension = "ROWS",
                                            Values = new List<IList<object>> { new List<object> { tmdbId } }
                                        });
                                        initialEmptyRowNum = initialEmptyRowNum + 1;
                                        movieAddedToMoviesSheet = true;
                                    }
                                }
                                else
                                {
                                    DisplayMessage("default", message, 0);
                                    DisplayMessage("info", "Skipping " + movieTitle + " because it is missing the IMDB ID");
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong saving the movie details for: " + movieTitle, 0, 0, 1, "Red");
                        Type(e.Message, 0, 0, 2, "DarkRed");
                        break;
                    }
                }
            }

            if (movieIdAddedToSkipColumn)
            {
                DisplayMessage("default", message, 0);
                DisplayMessage("info", "New movie IDs added to skip column - Going to add now");

                // Refresh skipMovieIdsData
                var BatchUpdateValuesResponse = BulkWriteToSheet(skipMovieIdsBatchRequest);
                skipMovieIdsData = CallGetData(new Dictionary<string, int> { { "Skip", -1 } }, SKIP_ACTORS_ID_TITLE_RANGE, SKIP_ACTORS_ID_DATA_RANGE, "Refreshing skip movie IDs...");
            }

            if (movieAddedToMoviesSheet)
            {
                DisplayMessage("default", message, 0);
                DisplayMessage("info", "New movies found to add to the Movies sheet - Going to add now");

                var BatchUpdateValuesResponse = BulkWriteToSheet(batchRequest);

                return true;
            }

            return false;
        }

        protected static string GetVideoResolution(string video)
        {
            string calculatedResolution = "";
            try
            {
                var ffProbe = new NReco.VideoInfo.FFProbe();
                var videoInfo = ffProbe.GetMediaInfo(Path.Combine(video));

                var videoStream = videoInfo.Streams.FirstOrDefault(s => s.CodecType == "video");
                if (videoStream != null && videoStream.Height > 0 && videoStream.Width > 0)
                {
                    calculatedResolution = FindResolution(videoStream.Width, videoStream.Height);
                }
                else
                {
                    Type("No video stream found for: " + video, 0, 0, 1, "Red");
                }
            }
            catch (Exception e)
            {
                Type("Something went wrong using ffProbe for: " + video, 0, 0, 1, "Red");
            }
            return calculatedResolution;
        }

        protected static string FindResolution(int width, int height)
        {
            DisplayMessage("info", $"{width}x{height}");

            if (width == 1280) return "720p";
            if (width == 1920) return "1080p";

            // Otherwise Emby sticks to the height thresholds without rounding up:
            //if (height < 360) return "240p";
            //if (height <= 360) return "360p";
            if (height == 576) return "576p";
            if (height < 720) return "480p";
            if (height < 1080) return "720p";
            if (height < 1440) return "1080p";
            if (height < 2160) return "1440p";
            if (height < 4320) return "2160p";
            if (height >= 4320) return "4320p";
            return "N/A";
        }

        protected static void ConvertVideo(IList<IList<Object>> data, Dictionary<string, int> sheetVariables, string presetChoice)
        {
            // Declare variables.
            int intTotalMoviesCount = 0,
                intImagesCount = 0,
                intAlreadyConvertedFilesCount = 0,
                intNoTitleCount = 0,
                intConvertedFilesCount = 0,
                intSkippedFilesCount = 0;

            foreach (var row in data)
            {
                if (row[Convert.ToInt16(sheetVariables["Show"])].ToString() != "") // If it's an empty row then this cell should be empty.
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
                            convertPath = "",
                            convertDirectory = "",
                            nfoBody = "",
                            skipFile = "",
                            cleanTitle = "";
                    try
                    {
                        i = row[Convert.ToInt16(sheetVariables[ISO_INPUT])].ToString();
                        o = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString() + "\\" + row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString() + ".mp4";
                        showTitle = row[Convert.ToInt16(sheetVariables["Show"])].ToString();
                        if (row[Convert.ToInt16(sheetVariables["Override Show"])].ToString() != "") showTitle = row[Convert.ToInt16(sheetVariables["Show"])].ToString();
                        SeasonNum = row[Convert.ToInt16(sheetVariables["Season #"])].ToString();
                        string pathRoot = Path.GetPathRoot(i.ToString());
                        cleanTitle = row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString();
                        convertPath = pathRoot + "These are finished running through HandBrake\\" + showTitle + "\\Season " + SeasonNum + "\\" + cleanTitle + ".mp4";
                        convertDirectory = pathRoot + "These are finished running through HandBrake\\" + showTitle + "\\Season " + SeasonNum;
                        title = row[Convert.ToInt16(sheetVariables[ISO_TITLE_NUM])].ToString();
                        additionalCommands = " " + row[Convert.ToInt16(sheetVariables[ADDITIONAL_COMMANDS])].ToString();
                        chapter = row[Convert.ToInt16(sheetVariables[ISO_CH_NUM])].ToString();
                        directoryLocation = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString();
                        nfoBody = row[Convert.ToInt16(sheetVariables[NFO_BODY])].ToString();
                        skipFile = row[Convert.ToInt16(sheetVariables["Skip"])].ToString();


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
                                    if (skipFile == "")
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

                                        if (nfoBody != "")
                                        {
                                            var nfoFile = Path.Combine(convertDirectory, cleanTitle) + ".nfo";
                                            WriteNfoFile(nfoFile, nfoBody);
                                        }

                                        intConvertedFilesCount++;
                                    }
                                    else
                                    {
                                        DisplayMessage("warning", "We skipped: ", 0);
                                        DisplayMessage("default", cleanTitle);
                                        DisplayMessage("warning", "Per your request");
                                        intSkippedFilesCount++;
                                    }
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
                            Type("We didn't find " + i, 0, 0, 1, "yellow");
                            Type("We won't be able to convert this one at this time.", 0, 0, 1, "yellow");
                            Type("-------------------------------------------------------------------", 0, 0, 1);
                        }
                    }
                    catch (Exception e)
                    {
                        Type("Something went wrong converting the following video: " + cleanTitle, 0, 0, 1, "Red");
                        Type(e.Message, 0, 0, 2, "DarkRed");
                        break;
                    }
                }
            } // End foreach
            Type("-----SUMMARY-----", 0, 0, 1);
            Type(intTotalMoviesCount + " total episodes in list to convert.", 0, 0, 1);
            Type(intImagesCount + " disc images found.", 0, 0, 1);
            Type(intAlreadyConvertedFilesCount + " episodes already converted and were skipped.", 0, 0, 1);
            Type(intConvertedFilesCount + " episodes converted.", 0, 0, 1);
            Type(intNoTitleCount + " missing titles to convert.", 0, 0, 1);

            Type("It looks like that's the end of it.", 0, 0, 2, "magenta");
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
                Type("Enter your directory", 0, 0, 1);

                // Add that directory to the directory variable.
                var directory = Console.ReadLine();

                // Now get all directories in the given directory and put them in an array.
                string[] subdirectoryEntries = Directory.GetDirectories(directory);

                // Check that there are some subdirectories.
                if (subdirectoryEntries.Length > 0)
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
                    Type("There are no folders to rename in this directory.", 0, 0, 1);
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
                    Type(folderName + " was not split because the dash seems to be part of the movie title.", 0, 0, 1);
                }
                // Else it doesn't contain a '(' and is probably fine to replace.
                else
                {
                    // Now replace the original path name with the split name.
                    string replacedName = path.Replace(folderName, split[0]);

                    // Finally, actually rename the folder.
                    Directory.Move(path, replacedName);

                    // Tell the user what happened.
                    Type(folderName + " was split.", 0, 0, 1);
                }

            }
            // Else if there is more than one dash, I don't want to rename it.
            else if (intDashCount > 1)
            {
                // Tell the user it wasn't split because of too many dashes.
                // Just rename those manually.
                Type(folderName + " has more than one dash and wasn't split.", 0, 0, 1);

            }
            // Else it doesn't have any dashes and won't be renamed.
            else
            {
                // Tell the user it wasn't split because it has no dashes.
                Type(folderName + " has no dashes and was not split.", 0, 0, 1);

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
            foreach (char c in value)
            {
                // If the character in the string is equal to the requested character then add one to our count.
                if (c == ch)
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
                Type("Unable to convert file. | " + objException.Message, 0, 0, 1);
            }
        } // End HandBrake()

        //protected static void GetDataToConvertEpisodes(string itemType, string presetFile)
        //{
        //    UserCredential credential;
        //    Dictionary<string, int> SheetVariables = new Dictionary<string, int>
        //    {
        //        { "Image Location", -1 },
        //        { "Episode Location", -1 },
        //        { ISO_TITLE_NUM, -1 },
        //        { "Chapter", -1 },
        //        { ADDITIONAL_COMMANDS, -1 }
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
        //        Type("We didn't find a column that we were looking for...", 0, 0, 1, "Red");
        //        foreach (KeyValuePair<string, int> variable in NotFoundColumns)
        //        {
        //            Type("Key: " + variable.Key.ToString() + ", Value: " + variable.Value.ToString(), 0, 0, 1, "Red");

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
        //                                title = row[Convert.ToInt16(SheetVariables[ISO_TITLE_NUM])].ToString(),
        //                                additionalCommands = " " + row[Convert.ToInt16(SheetVariables[ADDITIONAL_COMMANDS])].ToString(),
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
        //            Type("-----SUMMARY-----", 0, 0, 1);
        //            Type(intTotalEpisodesCount + " Total Episodes.", 0, 0, 1);
        //            Type(intImagesCount + " Images Found.", 0, 0, 1);
        //            Type(intAlreadyConvertedFilesCount + " Episode Files Found.", 0, 0, 1);
        //            Type(intConvertedFilesCount + " Episodes converted.", 0, 0, 1);
        //            Type(intNoTitleCount + " Missing Titles to convert.", 0, 0, 1);
        //        }
        //        else
        //        {
        //            Console.WriteLine("No data found.");
        //        }
        //        Type("It looks like that's the end of it.", 0, 0, 2, "magenta");
        //    }
        //} // End GetDataToConvertEpisodes()

        public static void verifySrtFileNames(string[] srtLocations)
        {
            foreach (var location in srtLocations)
            {
                string[] fileEntries = Directory.GetFiles(location);
                string[] subdirectoryEntries = Directory.GetDirectories(location);

                if (fileEntries.Length > 0)
                {
                    fileEntriesCount += fileEntries.Count();
                    foreach (string fileName in fileEntries)
                    {
                        if (IsSrtFile(fileName))
                        {
                            srtEntriesCount++;

                            if (!HasEngSrtFlag(fileName))
                            {
                                missingEngCount++;
                                missingEng.Add(fileName);
                            }
                        }

                    }
                }

                // Recurse into subdirectories of this directory.
                foreach (string subdirectory in subdirectoryEntries)
                {
                    // Don't go into the bonus features folders.
                    if (!subdirectory.Contains(@"\Behind The Scenes") &&
                        !subdirectory.Contains(@"\Scenes") &&
                        !subdirectory.Contains(@"\Deleted Scenes") &&
                        !subdirectory.Contains(@"\Shorts") &&
                        !subdirectory.Contains(@"\Featurettes") &&
                        !subdirectory.Contains(@"\Trailers") &&
                        !subdirectory.Contains(@"\Interviews") &&
                        !subdirectory.Contains(@"\Broken apart") &&
                        !subdirectory.Contains(@"\Other") &&
                        !subdirectory.Contains(@"\_Collections") &&
                        !subdirectory.Contains(@"\.sync"))
                    {
                        string[] directory = { subdirectory };
                        verifySrtFileNames(directory);
                    }
                }
            }
        }

        protected static void CountFiles()
        {
            var missingDirectories = new List<List<string>>();
            Dictionary<int, string> missingDirectoriesList = new Dictionary<int, string> { };
            bool keepAskingForDirectory = true, keepAskingForList = true;
            do
            {
                ClearDirectories();

                var directory = AskForDirectory();

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
                        missingDirectoriesList.Add(i, " Empty Directory");
                        i++;
                    }
                    if (res240List.Count() > 0)
                    {
                        directoryPlural = res240List.Count() == 1 ? " directory has " : " directories have ";
                        Type(res240List.Count().ToString() + directoryPlural + "a 240p video.", 0, 0, 1, "Magenta");
                        missingDirectories.Add(res240List);
                        missingDirectoriesList.Add(i, " 240p videos");
                        i++;
                    }
                    if (res360List.Count() > 0)
                    {
                        directoryPlural = res360List.Count() == 1 ? " directory has " : " directories have ";
                        Type(res360List.Count().ToString() + directoryPlural + "a 360p video.", 0, 0, 1, "Magenta");
                        missingDirectories.Add(res360List);
                        missingDirectoriesList.Add(i, " 360p videos");
                        i++;
                    }
                    if (res480List.Count() > 0)
                    {
                        directoryPlural = res480List.Count() == 1 ? " directory has " : " directories have ";
                        Type(res480List.Count().ToString() + directoryPlural + "a 480p video.", 0, 0, 1, "Magenta");
                        missingDirectories.Add(res480List);
                        missingDirectoriesList.Add(i, " 480p videos");
                        i++;
                    }
                    if (res576List.Count() > 0)
                    {
                        directoryPlural = res576List.Count() == 1 ? " directory has " : " directories have ";
                        Type(res576List.Count().ToString() + directoryPlural + "a 576p video.", 0, 0, 1, "Magenta");
                        missingDirectories.Add(res576List);
                        missingDirectoriesList.Add(i, " 576p videos");
                        i++;
                    }
                    if (res720List.Count() > 0)
                    {
                        directoryPlural = res720List.Count() == 1 ? " directory has " : " directories have ";
                        Type(res720List.Count().ToString() + directoryPlural + "a 720p video.", 0, 0, 1, "Magenta");
                        missingDirectories.Add(res720List);
                        missingDirectoriesList.Add(i, " 720p videos");
                        i++;
                    }
                    if (res1080List.Count() > 0)
                    {
                        directoryPlural = res1080List.Count() == 1 ? " directory has " : " directories have ";
                        Type(res1080List.Count().ToString() + directoryPlural + "a 1080p video.", 0, 0, 1, "Magenta");
                        missingDirectories.Add(res1080List);
                        missingDirectoriesList.Add(i, " 1080p videos");
                        i++;
                    }
                    if (res1440List.Count() > 0)
                    {
                        directoryPlural = res1440List.Count() == 1 ? " directory has " : " directories have ";
                        Type(res1440List.Count().ToString() + directoryPlural + "a 1440p video.", 0, 0, 1, "Magenta");
                        missingDirectories.Add(res1440List);
                        missingDirectoriesList.Add(i, " 1440p videos");
                        i++;
                    }
                    if (res2160List.Count() > 0)
                    {
                        directoryPlural = res2160List.Count() == 1 ? " directory has " : " directories have ";
                        Type(res2160List.Count().ToString() + directoryPlural + "a 2160p video.", 0, 0, 1, "Magenta");
                        missingDirectories.Add(res2160List);
                        missingDirectoriesList.Add(i, " 2160p videos");
                        i++;
                    }
                    if (resNAList.Count() > 0)
                    {
                        directoryPlural = resNAList.Count() == 1 ? " directory has " : " directories have ";
                        Type(resNAList.Count().ToString() + directoryPlural + "N/A resolution video.", 0, 0, 1, "Magenta");
                        missingDirectories.Add(resNAList);
                        missingDirectoriesList.Add(i, " N/A resolution videos");
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
                        Type("The following directories have" + missingDirectoriesList[int.Parse(response)] + " files:", 0, 0, 1);
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
                }
                else if (i == 1)
                {
                    fontColor = "Yellow";
                }
                else if (i == 2)
                {
                    fontColor = "Red";
                }
                Type(variable.Key.ToString() + ": " + variable.Value.ToString(), 0, 0, 1, fontColor);
                i++;
            }
            Type("-----------------------------------------------------------", 0, 0, 1);
        } // End DisplayResults()

        protected static string AskForDirectory(string message = "Enter your directory:")
        {
            bool keepAskingForDirectory = true;
            if (chosenDirectory == null || chosenDirectory == "")
            {
                do
                {
                    DisplayMessage("question", message + " (0 to cancel)");
                    chosenDirectory = RemoveCharFromString(Console.ReadLine(), '"');
                    if (chosenDirectory == "0")
                    {
                        keepAskingForDirectory = false;
                    }
                    else if (File.Exists(chosenDirectory))
                    {
                        Type("No, I need the path to a folder location, not a file.", 0, 0, 1, "Red");
                    }
                    else if (Directory.Exists(chosenDirectory))
                    {
                        keepAskingForDirectory = false;
                    }
                } while (keepAskingForDirectory);
            }

            return chosenDirectory;
        } // End AskForDirectory()

        protected static void ConvertHandbrakeList(ArrayList videoFiles)
        {
            Type("Now converting " + videoFiles.Count + " files... ", 0, 0, 1, "Yellow");

            // An ArrayList to hold the files that have finished converting so that we can remove the metadata from them.
            ArrayList outputFiles = new ArrayList();

            try
            {
                int count = 1;
                foreach (var myFile in videoFiles)
                {
                    if (File.Exists(myFile.ToString()))
                    {
                        Type("Converting " + count + " of " + videoFiles.Count + " files", 0, 0, 1, "Blue");

                        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(myFile.ToString());
                        string pathRoot = Path.GetPathRoot(myFile.ToString());
                        string i = myFile.ToString(),
                                o = Path.GetFullPath(Path.Combine(myFile.ToString(), @"..\..\These are finished running through HandBrake\" + Path.GetFileName(myFile.ToString()))),
                                presetChoice = "--preset-import-file MP4_RF22f.json -Z \"MP4 RF22f\"";

                        Directory.CreateDirectory(Path.GetDirectoryName(o));

                        ArrayList inputArrayList = new ArrayList { i };
                        long sizeOfInputFile = SizeOfFiles(inputArrayList);
                        ArrayList outputArrayList = new ArrayList { o };
                        // Since the output file MAY not exist yet we wait to get the size of it.
                        long sizeOfOutputFile = 0;

                        if (!File.Exists(o))
                        {
                            outputFiles.Add(o);

                            string strMyConversionString = "HandBrakeCLI -i \"" + i + "\" -o \"" + o + "\" " + presetChoice;

                            Type("Now converting: " + fileNameWithoutExtension, 0, 0, 1, "Magenta");

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

                        }
                        else
                        {
                            // Now that the output file definitely exists we can grab the size of it.
                            sizeOfOutputFile = SizeOfFiles(outputArrayList);

                            // Display the amount of bytes that conversion saved.
                            DisplaySavings(sizeOfOutputFile, sizeOfInputFile);

                            Type(fileNameWithoutExtension + " already exists at destination. --Skipping to next file.", 0, 0, 1, "Yellow");
                        }
                        // Now delete the input file.
                        File.Delete(i);
                        DisplayMessage("info", "Input file deleted.");

                        count++;
                        DisplayEndOfCurrentProcessLines();

                        Type("DONE", 0, 0, 1, "Green");
                    }
                }

            }
            catch (Exception e)
            {
                Type("Something happened | " + e.Message, 0, 0, 1, "Red");
            }
        } // End ConvertHandbrakeList()

        protected static void MoveVideoFilesToHoldFolder(string directory)
        {
            string[] graysonFiles = Directory.GetFiles(Path.GetFullPath(Path.Combine(directory, @"..\Dead Air Removed\Grayson")));
            string[] carsonFiles = Directory.GetFiles(Path.GetFullPath(Path.Combine(directory, @"..\Dead Air Removed\Carson")));
            string[] emersonFiles = Directory.GetFiles(Path.GetFullPath(Path.Combine(directory, @"..\Dead Air Removed\Emerson")));
            string[] evelynFiles = Directory.GetFiles(Path.GetFullPath(Path.Combine(directory, @"..\Dead Air Removed\Evelyn")));

            string[] fileEntries = new string[graysonFiles.Length + carsonFiles.Length + emersonFiles.Length + evelynFiles.Length];
            graysonFiles.CopyTo(fileEntries, 0);
            carsonFiles.CopyTo(fileEntries, graysonFiles.Length);
            emersonFiles.CopyTo(fileEntries, graysonFiles.Length + carsonFiles.Length);
            evelynFiles.CopyTo(fileEntries, graysonFiles.Length + carsonFiles.Length + emersonFiles.Length);

            // Filter out the files that aren't video files.
            ArrayList videoFilesToMove = GrabMovieFiles(fileEntries);
            if (videoFilesToMove.Count > 0)
            {
                DisplayMessage("info", "We found some new videos that just had some dead air removed, we will move them now");
                foreach (var moveFile in videoFilesToMove)
                {
                    MoveDirectory(moveFile.ToString(), Path.GetFullPath(Path.Combine(moveFile.ToString(), @"..\..\..\Run these through Handbrake\Hold\" + Path.GetFileName(moveFile.ToString()))));
                }
            }
        }

        protected static void DisplaySavings(long oFile, long iFile)
        {
            // Display the amount of bytes that conversion saved.
            long difference = iFile - oFile;
            //Console.WriteLine("iFile: " + iFile.ToString("N"));
            //Console.WriteLine("oFile: " + oFile.ToString("N"));
            //Console.WriteLine("difference: " + difference.ToString("N"));
            if (difference >= 0)
            {
                Type("Conversion savings: ", 0, 0, 0, "Blue");
                Type(FormatSize(difference, true) + " of " + FormatSize(iFile, true) + " -" + FormatPercentage(difference, iFile) + "%", 0, 0, 1, "Yellow");
            }
            else
            {
                Type("Conversion loss: ", 0, 0, 0, "Red");
                Type(FormatSize(difference * -1, true) + " more than " + FormatSize(iFile, true) + " +" + FormatPercentage(difference * -1, oFile) + "%", 0, 0, 1, "Yellow");
            }

            // Add the difference to display the total running difference in bytes.
            runningDifference += difference;
            runningFileSize += iFile;
            Type("Total savings: ", 0, 0, 0, "Blue");
            Type(FormatSize(runningDifference, true) + " of " + FormatSize(runningFileSize, true) + " " + FormatPercentage(runningDifference, runningFileSize) + "% saved", 0, 0, 1, "Cyan");

            // Add the difference to display the total session difference in bytes.
            runningSessionSavings += difference;
            runningSessionFileSize += iFile;
            Type("Session savings: ", 0, 0, 0, "Blue");
            Type(FormatSize(runningSessionSavings, true) + " of " + FormatSize(runningSessionFileSize, true) + " " + FormatPercentage(runningSessionSavings, runningSessionFileSize) + "% saved", 0, 0, 1, "Green");
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
                    }
                    else
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
            Type("Removing Metadata from the video files... ", 0, 0, 0, "Yellow");
            string performersRemovedCount = "Performers Removed Count", titlesRemovedCount = "Titles Removed Count", commentsRemovedCount = "Comments Removed Count";
            Dictionary<string, int> resultVariables = new Dictionary<string, int> { };
            resultVariables.Add(performersRemovedCount, 0);
            resultVariables.Add(titlesRemovedCount, 0);
            resultVariables.Add(commentsRemovedCount, 0);

            foreach (var myFile in videoFiles)
            {
                try
                {
                    string ext = Path.GetExtension(myFile.ToString()).ToUpperInvariant();

                    if (ext == ".MP4" || ext == ".M4V")
                    {
                        Type($"Processing MP4/M4V file: {myFile}", 0, 0, 1, "Yellow");
                        bool saveFile = false;

                        using (TagLib.File videoFile = TagLib.File.Create(myFile.ToString()))
                        {
                            // Performers
                            if (videoFile.Tag.Performers != null && videoFile.Tag.Performers.Length > 0)
                            {
                                videoFile.Tag.Performers = Array.Empty<string>();
                                resultVariables[performersRemovedCount] += 1;
                                saveFile = true;
                                Type("Removed performers metadata", 0, 0, 1, "Green");
                            }

                            // Title
                            if (!string.IsNullOrEmpty(videoFile.Tag.Title))
                            {
                                videoFile.Tag.Title = string.Empty;
                                resultVariables[titlesRemovedCount] += 1;
                                saveFile = true;
                                Type("Removed title metadata", 0, 0, 1, "Green");
                            }

                            // Comment
                            if (!string.IsNullOrEmpty(videoFile.Tag.Comment))
                            {
                                videoFile.Tag.Comment = string.Empty;
                                resultVariables[commentsRemovedCount] += 1;
                                saveFile = true;
                                Type("Removed comment metadata", 0, 0, 1, "Green");
                            }

                            if (saveFile)
                            {
                                videoFile.Save();
                                Type($"Done removing metadata from {myFile}", 0, 0, 1, "Cyan");
                            }
                            else
                            {
                                Type($"No metadata to remove from {myFile}", 0, 0, 1, "Gray");
                            }
                        }
                    }
                    else if (ext.Equals(".MKV", StringComparison.OrdinalIgnoreCase))
                    {
                        Type($"MKV file detected, now going to remove metadata from: {myFile}", 0, 0, 1, "Yellow");

                        bool ok = MkvToolNix.ClearTitleAndComments(
                            myFile.ToString(),
                            msg => Type(msg, 0, 0, 1, "Gray"),
                            msg => Type(msg, 0, 0, 1, "Red")
                        // , @"C:\Path\To\mkvpropedit.exe" // optional override
                        );

                        if (ok)
                        {
                            Type($"Done cleaning metadata from {myFile}", 0, 0, 1, "Cyan");
                            resultVariables[performersRemovedCount] += 1;
                            resultVariables[titlesRemovedCount] += 1;
                            resultVariables[commentsRemovedCount] += 1;
                        }
                        else
                        {
                            Type($"mkvpropedit failed for {myFile}", 0, 0, 1, "Red");
                        }
                    }
                    else
                    {
                        Type($"Skipping unsupported file type: {myFile}", 0, 0, 1, "Gray");
                    }
                }
                catch (Exception e)
                {
                    Type("Unable to remove the metadata on a file", 0, 0, 1, "Red");
                    Type(myFile.ToString(), 0, 0, 1, "Blue");
                    Type(e.Message, 0, 0, 1, "DarkRed");
                }
            }

            Type("DONE", 0, 0, 1, "Green");

            DisplayResults(resultVariables);

        } // End RemoveMetadata()

        protected static void CopyJpgFiles()
        {
            bool keepAskingForDirectory = true;
            do
            {
                Type("Enter your directory", 0, 0, 1);
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
                }
                else if (person == "1" || person == "2")
                {
                    Type("What hard drive am I copying to? (Just the hard drive letter)", 0, 0, 1, "Yellow");

                    chosenDestination = Console.ReadLine().ToUpper();

                    if (!HardDriveHasSpace(chosenDestination))
                    {
                        DisplayMessage("error", "We won't be able to copy any more movies to this hard drive because available space is below 10%");
                        break;
                    }
                    else
                    {
                        Console.WriteLine("We will copy to hard drive " + chosenDestination);
                    }

                    Type("What hard drive am I copying from? (Just the hard drive letter)", 0, 0, 1, "Yellow");

                    sourceHardDriveLetter = Console.ReadLine().ToUpper();

                    if (chosenDestination != sourceHardDriveLetter)
                    {
                        Console.WriteLine("We will copy from the " + sourceHardDriveLetter + " drive.");
                    }
                    else
                    {
                        DisplayMessage("error", "I'm sorry the source hard drive can't be the same as the destination hard drive.");
                        repeatProcess = true;
                    }

                }
                else
                {
                    DisplayMessage("error", "I'm sorry, I don't recognise " + person + " yet. Could you add them to my DB before continuing?");
                    repeatProcess = true;
                }

                if (!repeatProcess)
                {
                    foreach (var row in data)
                    {
                        if (!HardDriveHasSpace(chosenDestination))
                        {
                            DisplayMessage("error", "We have stopped copying movies because available hard drive space is below 10%");
                            break;
                        }

                        if (row.Count > 4) // If it's an empty row then it should have less than this.
                        {
                            var cleanTitle = row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString();
                            var movieDirectory = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString();
                            var status = row[Convert.ToInt16(sheetVariables[STATUS])].ToString();
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
                                Type("Something went wrong when looking for: " + sourceHardDriveLetter + "\\" + movieDirectory + " | " + e.Message, 0, 0, 1, "Red");
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
                            var cleanTitle = row[Convert.ToInt16(sheetVariables[CLEAN_TITLE])].ToString();
                            var movieDirectory = row[Convert.ToInt16(sheetVariables[DIRECTORY])].ToString();
                            var status = row[Convert.ToInt16(sheetVariables[STATUS])].ToString();
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

                                    // Concatenate to the containing directory.
                                    fullDestinationPathToFileToDelete = containingDirectory + "\\" + cleanTitle;

                                    // Loop through the containing directory to see if the movie is already in there.
                                    string[] fileEntries = Directory.GetFiles(containingDirectory, cleanTitle + ".*");
                                    if (fileEntries.Length > 0)
                                    {
                                        foreach (var movie in fileEntries)
                                        {
                                            File.SetAttributes(movie, FileAttributes.Normal);
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
            res240List.Clear();
            res360List.Clear();
            res480List.Clear();
            res576List.Clear();
            res720List.Clear();
            res1080List.Clear();
            res1440List.Clear();
            res2160List.Clear();
            resNAList.Clear();
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
                int nfoCount = 0,
                    jpgCount = 0,
                    mp4Count = 0,
                    mkvCount = 0,
                    m4vCount = 0,
                    aviCount = 0,
                    webmCount = 0,
                    unidentifiedCount = 0,
                    isoCount = 0,
                    partCount = 0,
                    res240Count = 0,
                    res360Count = 0,
                    res480Count = 0,
                    res576Count = 0,
                    res720Count = 0,
                    res1080Count = 0,
                    res1440Count = 0,
                    res2160Count = 0,
                    resNACount = 0;
                foreach (string fileName in fileEntries)
                {
                    if (fileName.ToUpper().Contains(".NFO"))
                        nfoCount++;
                    else if (fileName.ToUpper().Contains(".JPG"))
                        jpgCount++;
                    else if (fileName.ToUpper().Contains(".MP4"))
                    {
                        mp4Count++;
                        string videoResolution = GetVideoResolution(fileName);
                        if (videoResolution == "240p")
                            res240Count++;
                        else if (videoResolution == "360p")
                            res360Count++;
                        else if (videoResolution == "480p")
                            res480Count++;
                        else if (videoResolution == "576p")
                            res576Count++;
                        else if (videoResolution == "720p")
                            res720Count++;
                        else if (videoResolution == "1080p")
                            res1080Count++;
                        else if (videoResolution == "1440p")
                            res1440Count++;
                        else if (videoResolution == "2160p")
                            res2160Count++;
                        else if (videoResolution == "N/A")
                            resNACount++;
                    }
                    else if (fileName.ToUpper().Contains(".MKV"))
                    {
                        mkvCount++;
                        string videoResolution = GetVideoResolution(fileName);
                        if (videoResolution == "240p")
                            res240Count++;
                        else if (videoResolution == "360p")
                            res360Count++;
                        else if (videoResolution == "480p")
                            res480Count++;
                        else if (videoResolution == "576p")
                            res576Count++;
                        else if (videoResolution == "720p")
                            res720Count++;
                        else if (videoResolution == "1080p")
                            res1080Count++;
                        else if (videoResolution == "1440p")
                            res1440Count++;
                        else if (videoResolution == "2160p")
                            res2160Count++;
                        else if (videoResolution == "N/A")
                            resNACount++;
                    }
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
                if (res240Count > 0)
                {
                    res240List.Add(targetDirectory);
                }
                if (res360Count > 0)
                {
                    res360List.Add(targetDirectory);
                }
                if (res480Count > 0)
                {
                    res480List.Add(targetDirectory);
                }
                if (res576Count > 0)
                {
                    res576List.Add(targetDirectory);
                }
                if (res720Count > 0)
                {
                    res720List.Add(targetDirectory);
                }
                if (res1080Count > 0)
                {
                    res1080List.Add(targetDirectory);
                }
                if (res1440Count > 0)
                {
                    res1440List.Add(targetDirectory);
                }
                if (res2160Count > 0)
                {
                    res2160List.Add(targetDirectory);
                }
                if (resNACount > 0)
                {
                    resNAList.Add(targetDirectory);
                }

                //Type(nfoCount + " nfo, " + jpgCount + " jpg, " + mp4Count + " mp4, " + mkvCount + " mkv, " + m4vCount + " m4v, " + isoCount + " iso, " + unidentifiedCount + " unidentified in " + targetDirectory, 0, 0, 1);
            }
            else if (subdirectoryEntries.Length == 0)
            {
                Directory.Delete(targetDirectory);
                emptyDirectory.Add(targetDirectory);
            }

            // Recurse into subdirectories of this directory.
            foreach (string subdirectory in subdirectoryEntries)
            {
                // Don't go into the bonus features folders.
                if (!subdirectory.Contains(@"\Behind The Scenes") &&
                    !subdirectory.Contains(@"\Scenes") &&
                    !subdirectory.Contains(@"\Deleted Scenes") &&
                    !subdirectory.Contains(@"\Shorts") &&
                    !subdirectory.Contains(@"\Featurettes") &&
                    !subdirectory.Contains(@"\Trailers") &&
                    !subdirectory.Contains(@"\Interviews") &&
                    !subdirectory.Contains(@"\Broken apart") &&
                    !subdirectory.Contains(@"\_Collections") &&
                    !subdirectory.Contains(@"\.sync"))
                {
                    ProcessDirectory(subdirectory);
                }
            }
        }

        // Delete all movie files in a folder and sub folders.
        public static void DeleteMoviesInFolder(string targetDirectory)
        {
            Type("Checking for video files in ", 0, 0, 0, "Yellow");
            Type(targetDirectory, 0, 0, 1, "Blue");
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            if (fileEntries.Length > 0)
            {
                ArrayList videoFiles = GrabMovieFiles(fileEntries);

                if (videoFiles.Count == 0)
                {
                    Type("No video files to delete from ", 0, 0, 0, "Yellow");
                    Type(targetDirectory, 0, 0, 1, "Blue");
                }
                else
                {
                    Type(videoFiles.Count.ToString(), 0, 0, 0, "Blue");
                    Type(" videos deleted from ", 0, 0, 0, "Yellow");
                    Type(targetDirectory, 0, 0, 1, "Blue");
                    foreach (var videoFile in videoFiles)
                    {
                        File.Delete(videoFile.ToString());
                    }
                }

            }
            else if (subdirectoryEntries.Length == 0)
            {
                Directory.Delete(targetDirectory);
            }

            // Recurse into subdirectories of this directory.
            foreach (string subdirectory in subdirectoryEntries)
            {
                DeleteMoviesInFolder(subdirectory);
            }
        }

        // Delete all JPG files in a folder and sub folders.
        public static void DeleteJpgsInFolder(string targetDirectory)
        {
            Type("Checking for JPG files in ", 0, 0, 0, "Yellow");
            Type(targetDirectory, 0, 0, 1, "Blue");
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            if (fileEntries.Length > 0)
            {
                ArrayList jpgFiles = GrabJpgFiles(fileEntries);

                if (jpgFiles.Count == 0)
                {
                    Type("No JPG files to delete from ", 0, 0, 0, "Yellow");
                    Type(targetDirectory, 0, 0, 1, "Blue");
                }
                else
                {
                    Type(jpgFiles.Count.ToString(), 0, 0, 0, "Blue");
                    Type(" JPG files deleted from ", 0, 0, 0, "Yellow");
                    Type(targetDirectory, 0, 0, 1, "Blue");
                    foreach (var jpg in jpgFiles)
                    {
                        File.Delete(jpg.ToString());
                    }
                }

            }

            // Recurse into subdirectories of this directory.
            foreach (string subdirectory in subdirectoryEntries)
            {
                DeleteJpgsInFolder(subdirectory);
            }
        }

        // Remove metadata from all movie files in a folder and sub folders.
        public static void RemoveMetadataInFolder(string targetDirectory)
        {
            Type("Checking for video files in ", 0, 0, 0, "Yellow");
            Type(targetDirectory, 0, 0, 1, "Blue");
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            if (fileEntries.Length > 0)
            {
                ArrayList videoFiles = GrabMovieFiles(fileEntries);

                if (videoFiles.Count == 0)
                {
                    Type("No video files to remove metadata from ", 0, 0, 0, "Yellow");
                    Type(targetDirectory, 0, 0, 1, "Blue");
                }
                else
                {
                    //Type(videoFiles.Count.ToString(), 0, 0, 0, "Blue");
                    //Type(" metadata removed from ", 0, 0, 0, "Yellow");
                    //Type(targetDirectory, 0, 0, 1, "Blue");
                    RemoveMetadata(videoFiles);
                }

            }

            // Recurse into subdirectories of this directory.
            foreach (string subdirectory in subdirectoryEntries)
            {
                RemoveMetadataInFolder(subdirectory);
            }
        }

        // Move all contents in the srcDirectory to the targetDirectory.
        public static void MoveFolderContent(string srcDirectory, string targetDirectory)
        {
            Type("Moving content... ", 0, 0, 0, "Yellow");
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(srcDirectory);
            string[] subdirectoryEntries = Directory.GetDirectories(srcDirectory);
            if (fileEntries.Length > 0)
            {
                string directoryName = Path.GetDirectoryName(srcDirectory);
                string directoryName2 = Path.GetFileName(srcDirectory);
                foreach (var file in fileEntries)
                {
                    Type("I am here", 0, 0, 1);
                }







                //ArrayList videoFiles = GrabMovieFiles(fileEntries);

                //if (videoFiles.Count == 0)
                //{
                //    Type("No video files to delete from ", 0, 0, 0, "Yellow");
                //    Type(targetDirectory, 0, 0, 1, "Blue");
                //}
                //else
                //{
                //    Type(videoFiles.Count.ToString(), 0, 0, 0, "Blue");
                //    Type(" videos deleted from ", 0, 0, 0, "Yellow");
                //    Type(targetDirectory, 0, 0, 1, "Blue");
                //    foreach (var videoFile in videoFiles)
                //    {
                //        File.Delete(videoFile.ToString());
                //    }
                //}

            }
            else if (subdirectoryEntries.Length == 0)
            {
                Directory.Delete(targetDirectory);
            }

            // Recurse into subdirectories of this directory.
            foreach (string subdirectory in subdirectoryEntries)
            {
                MoveFolderContent(subdirectory, targetDirectory);
            }
        }

        // Process all files in the directory passed in, recurse on any directories
        // that are found, and move files up one.
        public static void RecurseSameMovieFolder(string topLevelDirectory, string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            if (fileEntries.Length > 0 && targetDirectory != topLevelDirectory)
            {
                foreach (string fileName in fileEntries)
                {
                    File.Move(fileName, Path.GetFullPath(Path.Combine(fileName, @"..\..\" + Path.GetFileName(fileName.ToString()))));
                }
            }
            if (subdirectoryEntries.Length == 0)
            {
                Directory.Delete(targetDirectory);
            }

            // Recurse into subdirectories of this directory.
            foreach (string subdirectory in subdirectoryEntries)
            {
                RecurseSameMovieFolder(topLevelDirectory, subdirectory);
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
        /// <param name="pause">The amount of MS to pause before going to the next line.</param>
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
                    Type(message, speed, pause, numLines, "white");
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
        public static void Type(string myString, int speed = 0, int timeToPauseBeforeNewLine = 0, int numberOfNewLines = 1, string color = "gray")
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
                Type("Something went wrong writing to path: " + path, 0, 0, 1, "Red");
                Type(e.Message, 0, 0, 2, "DarkRed");
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
            try
            {
                string[] myString = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ" };

                return myString[columnNum];
            }
            catch (Exception e)
            {
                Type("Something went wrong converting the following columnNum to a column letter: " + columnNum, 0, 0, 1, "Red");
                Type(e.Message, 0, 0, 2, "DarkRed");
                return "";
            }
        }

        public static bool WriteSingleCellToSheet(string strDataToSave, string strCellToSaveData, int sleep = 500)
        {
            var tryAgain = false;
            do
            {
                try
                {
                    Thread.Sleep(sleep);
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
                            GoogleClientSecrets.FromStream(stream).Secrets,
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
                    }
                    else
                    {
                        DisplayMessage("error", "An error has occurred.");
                        DisplayMessage("harderror", m);
                    }
                }

            } while (tryAgain);
            return true;

        } // End WriteSingleCellToSheet

        public class MoviesSheetBatchContext
        {
            public SheetsService Service { get; set; }
            public IList<object> Headers { get; set; }
            public IList<IList<object>> DataRows { get; set; }
            public Dictionary<string, int> ColIndex { get; set; }
            public List<ValueRange> PendingUpdates { get; set; } = new List<ValueRange>();
        }

        private static MoviesSheetBatchContext BuildMoviesSheetBatchContext()
        {
            var service = GetSheetsService();

            // 1) Get header row
            var headerResponse = service.Spreadsheets.Values.Get(SPREADSHEET_ID, MOVIES_TITLE_RANGE).Execute();
            var headers = headerResponse.Values?.FirstOrDefault();
            if (headers == null)
                throw new Exception("Movies header row not found in Google Sheet.");

            // 2) Get data rows
            var dataResponse = service.Spreadsheets.Values.Get(SPREADSHEET_ID, MOVIES_DATA_RANGE).Execute();
            var data = dataResponse.Values ?? new List<IList<object>>();

            // 3) Map header -> index
            var colIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headers.Count; i++)
            {
                var key = headers[i]?.ToString()?.Trim();
                if (!string.IsNullOrWhiteSpace(key) && !colIndex.ContainsKey(key))
                    colIndex[key] = i;
            }

            if (!colIndex.ContainsKey(CLEAN_TITLE))
                throw new Exception($"Could not find '{CLEAN_TITLE}' column in Movies sheet.");

            return new MoviesSheetBatchContext
            {
                Service = service,
                Headers = headers,
                DataRows = data,
                ColIndex = colIndex
            };
        }

        protected static BatchUpdateValuesResponse BulkWriteToSheet(BatchUpdateValuesRequest batchRequest)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new
                    FileDataStore(credPath, true)).Result;
            }
            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "Google-SheetsSample/0.1",
            });

            return sheetsService.Spreadsheets.Values.BatchUpdate(batchRequest, SPREADSHEET_ID).Execute();
        }
    } // End class

    public class KeyValuePair
    {
        public string Value { get; set; }
        public int CellNumber { get; set; }

        public KeyValuePair(string value, int cellNumber)
        {
            Value = value;
            CellNumber = cellNumber;
        }
    }

    public static class MkvToolNix
    {
        // Common install locations; adjust if needed
        private static readonly string[] DefaultLocations = new[]
        {
        @"C:\Program Files\MKVToolNix\mkvpropedit.exe",
        @"C:\Program Files (x86)\MKVToolNix\mkvpropedit.exe"
    };

        /// <summary>
        /// Clears the MKV container title and all tags (comments/extra metadata) using mkvpropedit.
        /// Returns true on success, false on failure.
        /// </summary>
        public static bool ClearTitleAndComments(
            string mkvFile,
            Action<string> logInfo,
            Action<string> logError,
            string mkvpropeditPath = null)
        {
            if (string.IsNullOrWhiteSpace(mkvFile) || !File.Exists(mkvFile))
            {
                SafeLog(logError, "[mkvpropedit] File not found: " + mkvFile);
                return false;
            }

            string exe = ResolveMkvPropEdit(mkvpropeditPath);
            if (string.IsNullOrEmpty(exe))
            {
                SafeLog(logError, "[mkvpropedit] Could not find mkvpropedit.exe. Install MKVToolNix or add it to PATH.");
                return false;
            }

            // One-shot: clear title and remove all tags
            string args = "\"" + mkvFile + "\" --edit info --set \"title=\" --tags all:";

            var psi = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = args,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            SafeLog(logInfo, "[mkvpropedit] Clearing title and comments: " + mkvFile);

            try
            {
                using (var proc = new Process { StartInfo = psi, EnableRaisingEvents = false })
                {
                    proc.OutputDataReceived += (s, e) => { if (!string.IsNullOrEmpty(e.Data)) SafeLog(logInfo, "[mkvpropedit] " + e.Data); };
                    proc.ErrorDataReceived += (s, e) => { if (!string.IsNullOrEmpty(e.Data)) SafeLog(logError, "[mkvpropedit] " + e.Data); };

                    if (!proc.Start())
                    {
                        SafeLog(logError, "[mkvpropedit] Failed to start process.");
                        return false;
                    }

                    proc.BeginOutputReadLine();
                    proc.BeginErrorReadLine();
                    proc.WaitForExit();

                    if (proc.ExitCode == 0)
                    {
                        SafeLog(logInfo, "[mkvpropedit] OK: " + mkvFile);
                        return true;
                    }

                    SafeLog(logError, "[mkvpropedit] Exit code " + proc.ExitCode + " for file: " + mkvFile);
                    return false;
                }
            }
            catch (Exception ex)
            {
                SafeLog(logError, "[mkvpropedit] Exception for " + mkvFile + ": " + ex.Message);
                return false;
            }
        }

        /// <summary>Overload with simple Console logging.</summary>
        public static bool ClearTitleAndComments(string mkvFile, string mkvpropeditPath = null)
        {
            return ClearTitleAndComments(
                mkvFile,
                Console.WriteLine,
                Console.Error.WriteLine,
                mkvpropeditPath
            );
        }

        private static string ResolveMkvPropEdit(string explicitPath)
        {
            // Explicit path
            if (!string.IsNullOrWhiteSpace(explicitPath) && File.Exists(explicitPath))
                return explicitPath;

            // Defaults
            for (int i = 0; i < DefaultLocations.Length; i++)
                if (File.Exists(DefaultLocations[i])) return DefaultLocations[i];

            // PATH: use `where mkvpropedit`
            try
            {
                var wherePsi = new ProcessStartInfo
                {
                    FileName = "where",
                    Arguments = "mkvpropedit",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (var p = Process.Start(wherePsi))
                {
                    if (p != null)
                    {
                        string line = p.StandardOutput.ReadLine();
                        if (!string.IsNullOrWhiteSpace(line) && File.Exists(line.Trim()))
                            return line.Trim();
                    }
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private static void SafeLog(Action<string> logger, string message)
        {
            try { if (logger != null) logger(message); } catch { /* swallow */ }
        }
    }
} // End namespace
