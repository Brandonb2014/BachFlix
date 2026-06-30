using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Scans filtered rows in the Movies sheet, evaluates the health of matching video files,
    /// and writes the resulting health status back to the File Health column.
    /// </summary>
    public static class VideoHealthSheetRunner
    {
        // Header labels in your Movies sheet
        private const string HEADER_DIRECTORY = "Directory";
        private const string HEADER_CLEAN_TITLE = "Clean Title";
        private const string HEADER_FILE_HEALTH = "File Health";
        private const string HEADER_RECORD_SOURCE = "Recorded Source";

        // Filter headers
        private const string HEADER_RESOLUTION = "Resolution";
        private const string HEADER_STREAMFAB = "StreamFab";
        private const string HEADER_FILE_TYPE = "File Type";

        // Filter values
        private const string TARGET_RESOLUTION = "1080p";
        private const string TARGET_STREAMFAB = "Y";
        private const string TARGET_FILE_TYPE = "mkv";

        /// <summary>
        /// Scans the Movies sheet for rows with an empty File Health value, filtered to
        /// movies recorded from StreamFab in 1080p MKV format, then runs the video health
        /// analyzer and writes OK, WARNING, or BAD back into the sheet.
        /// </summary>
        /// <param name="sheetsService">The authenticated Google Sheets service instance.</param>
        /// <param name="spreadsheetId">The target spreadsheet ID.</param>
        /// <param name="sheetName">The worksheet name to scan, such as "Movies".</param>
        public static void Run(SheetsService sheetsService, string spreadsheetId, string sheetName)
        {
            Console.WriteLine();
            Console.WriteLine("=== FILE HEALTH SHEET SCAN (FINAL SANITY CHECK FILTER) ===");
            Console.WriteLine("Scanning rows where:");
            Console.WriteLine($"  - '{HEADER_FILE_HEALTH}' is empty");
            Console.WriteLine($"  - '{HEADER_RESOLUTION}' = {TARGET_RESOLUTION}");
            Console.WriteLine($"  - '{HEADER_STREAMFAB}'  = {TARGET_STREAMFAB}");
            Console.WriteLine($"  - '{HEADER_FILE_TYPE}'  = {TARGET_FILE_TYPE}");
            Console.WriteLine("Running FULL ffmpeg checks on matching rows only.");
            Console.WriteLine();

            string range = sheetName + "!A:ZZ";
            SpreadsheetsResource.ValuesResource.GetRequest getReq =
                sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = getReq.Execute();
            IList<IList<object>> values = response.Values;

            if (values == null || values.Count == 0)
            {
                Console.WriteLine("No rows found in sheet.");
                return;
            }

            // Assume header row is the same row SRT runner uses (row 2, index 1)
            IList<object> headerRow = values[1];

            int colDirectory = FindColumnIndex(headerRow, HEADER_DIRECTORY);
            int colCleanTitle = FindColumnIndex(headerRow, HEADER_CLEAN_TITLE);
            int colFileHealth = FindColumnIndex(headerRow, HEADER_FILE_HEALTH);
            int colRecordSource = FindColumnIndex(headerRow, HEADER_RECORD_SOURCE);

            int colResolution = FindColumnIndex(headerRow, HEADER_RESOLUTION);
            int colStreamFab = FindColumnIndex(headerRow, HEADER_STREAMFAB);
            int colFileType = FindColumnIndex(headerRow, HEADER_FILE_TYPE);

            if (colDirectory == -1 || colCleanTitle == -1 || colFileHealth == -1 ||
                colResolution == -1 || colStreamFab == -1 || colFileType == -1 || colRecordSource == -1)
            {
                Console.WriteLine("Could not find one or more required headers:");
                Console.WriteLine("  Directory:       " + (colDirectory == -1 ? "NOT FOUND" : ColumnIndexToLetter(colDirectory)));
                Console.WriteLine("  Clean Title:     " + (colCleanTitle == -1 ? "NOT FOUND" : ColumnIndexToLetter(colCleanTitle)));
                Console.WriteLine("  File Health:     " + (colFileHealth == -1 ? "NOT FOUND" : ColumnIndexToLetter(colFileHealth)));
                Console.WriteLine("  Recorded Source: " + (colRecordSource == -1 ? "NOT FOUND" : ColumnIndexToLetter(colRecordSource)));
                Console.WriteLine("  Resolution:      " + (colResolution == -1 ? "NOT FOUND" : ColumnIndexToLetter(colResolution)));
                Console.WriteLine("  StreamFab:       " + (colStreamFab == -1 ? "NOT FOUND" : ColumnIndexToLetter(colStreamFab)));
                Console.WriteLine("  File Type:       " + (colFileType == -1 ? "NOT FOUND" : ColumnIndexToLetter(colFileType)));
                return;
            }

            Console.WriteLine(
                "Using columns: Directory={0}, Clean Title={1}, File Health={2}, Resolution={3}, StreamFab={4}, File Type={5}",
                ColumnIndexToLetter(colDirectory),
                ColumnIndexToLetter(colCleanTitle),
                ColumnIndexToLetter(colFileHealth),
                ColumnIndexToLetter(colResolution),
                ColumnIndexToLetter(colStreamFab),
                ColumnIndexToLetter(colFileType));

            int okCount = 0;
            int warningCount = 0;
            int badCount = 0;
            int matchCount = 0;

            var warningMovies = new List<string>();
            var badMovies = new List<string>();

            int maxIndex = new[]
            {
                colDirectory,
                colCleanTitle,
                colFileHealth,
                colResolution,
                colStreamFab,
                colFileType,
                colRecordSource
            }.Max();

            // Process data rows (start at rowIndex 2 because 1 is header row)
            for (int rowIndex = 2; rowIndex < values.Count; rowIndex++)
            {
                IList<object> row = values[rowIndex];

                while (row.Count <= maxIndex)
                {
                    row.Add(string.Empty);
                }

                string currentHealth = SafeGet(row, colFileHealth);
                if (!string.IsNullOrWhiteSpace(currentHealth))
                    continue;

                string resolution = NormalizeCell(SafeGet(row, colResolution));
                string streamFab = NormalizeCell(SafeGet(row, colStreamFab));
                string fileType = NormalizeCell(SafeGet(row, colFileType));

                if (!string.Equals(resolution, TARGET_RESOLUTION, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!string.Equals(streamFab, TARGET_STREAMFAB, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (fileType.StartsWith(".", StringComparison.OrdinalIgnoreCase))
                    fileType = fileType.Substring(1);

                if (!string.Equals(fileType, TARGET_FILE_TYPE, StringComparison.OrdinalIgnoreCase))
                    continue;

                string directory = SafeGet(row, colDirectory);
                string cleanTitle = SafeGet(row, colCleanTitle);
                string recordSource = SafeGet(row, colRecordSource);

                if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(cleanTitle))
                    continue;

                matchCount++;

                string videoPath = FindVideoFile(directory, cleanTitle, expectedExtension: ".mkv");
                if (videoPath == null)
                {
                    Console.WriteLine(
                        "[Row {0}] No .mkv video file found for title '{1}' in '{2}'",
                        rowIndex + 1, cleanTitle, directory);
                    continue;
                }

                if (!string.Equals(Path.GetExtension(videoPath), ".mkv", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine(
                        "[Row {0}] SKIP: Found file is not .mkv: {1}",
                        rowIndex + 1, videoPath);
                    continue;
                }

                Console.WriteLine("[Row {0}] Checking video health: {1}", rowIndex + 1, videoPath);

                string statusText = "BAD";

                try
                {
                    VideoHealthCheck.VideoHealthResult result = VideoHealthCheck.AnalyzeVideoFile(videoPath);
                    statusText = result.Status.ToString();

                    if (result.Status == VideoHealthCheck.VideoHealthStatus.OK)
                    {
                        okCount++;
                    }
                    else if (result.Status == VideoHealthCheck.VideoHealthStatus.WARNING)
                    {
                        warningCount++;
                        warningMovies.Add(BuildSummaryTitle(cleanTitle, recordSource));
                    }
                    else
                    {
                        badCount++;
                        badMovies.Add(BuildSummaryTitle(cleanTitle, recordSource));
                    }

                    Console.Write("   -> Result: ");
                    WriteColoredStatus(statusText);
                    Console.WriteLine(" (hadErrors={0}, exitCode={1})", result.HadErrors, result.ExitCode);

                    if (result.MatchedWarningPatterns.Count > 0)
                        Console.WriteLine("   -> Warning patterns: " + string.Join(", ", result.MatchedWarningPatterns.Distinct(StringComparer.OrdinalIgnoreCase)));

                    if (result.MatchedBadPatterns.Count > 0)
                        Console.WriteLine("   -> Bad patterns: " + string.Join(", ", result.MatchedBadPatterns.Distinct(StringComparer.OrdinalIgnoreCase)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine("   -> Exception while checking video: " + ex.Message);
                    statusText = "BAD";
                    badCount++;
                    badMovies.Add(BuildSummaryTitle(cleanTitle, recordSource));
                }

                string colLetter = ColumnIndexToLetter(colFileHealth);
                string cellRange = sheetName + "!" + colLetter + (rowIndex + 1);

                try
                {
                    ValueRange valueRange = new ValueRange
                    {
                        Values = new List<IList<object>>
                        {
                            new List<object> { statusText }
                        }
                    };

                    var updateRequest = sheetsService.Spreadsheets.Values.Update(
                        valueRange,
                        spreadsheetId,
                        cellRange);

                    updateRequest.ValueInputOption =
                        SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

                    updateRequest.Execute();

                    Console.Write("   -> Wrote File Health '");
                    WriteColoredStatus(statusText);
                    Console.WriteLine("' to {0}", cellRange);
                    Console.WriteLine();
                }
                catch (Exception exUpdate)
                {
                    Console.WriteLine("   -> ERROR writing File Health to sheet: " + exUpdate.Message);
                    // Leave the cell blank so this row can be retried on a future run
                }
            }

            Console.WriteLine();
            Console.WriteLine("File Health scan complete.");
            Console.Write("Summary: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("{0} OK", okCount);
            Console.ResetColor();
            Console.Write(", ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("{0} WARNING", warningCount);
            Console.ResetColor();
            Console.Write(", ");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("{0} BAD", badCount);
            Console.ResetColor();
            Console.WriteLine();

            if (warningMovies.Count > 0)
            {
                Console.WriteLine("WARNING movies:");
                foreach (string warningTitle in warningMovies)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(" - " + warningTitle);
                    Console.ResetColor();
                }
            }

            if (badMovies.Count > 0)
            {
                Console.WriteLine("BAD movies:");
                foreach (string badTitle in badMovies)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(" - " + badTitle);
                    Console.ResetColor();
                }
            }

            Console.WriteLine("Matched rows (before file existence checks): {0}", matchCount);
        }

        /// <summary>
        /// Finds a video file in the specified directory whose name matches the clean title
        /// and whose extension matches the expected extension.
        /// </summary>
        /// <param name="directory">The directory to search.</param>
        /// <param name="cleanTitle">The expected filename without extension.</param>
        /// <param name="expectedExtension">The required file extension, such as ".mkv".</param>
        /// <returns>The matched file path, or null if no match is found.</returns>
        private static string FindVideoFile(string directory, string cleanTitle, string expectedExtension)
        {
            try
            {
                if (!Directory.Exists(directory))
                    return null;

                if (string.IsNullOrWhiteSpace(expectedExtension))
                    expectedExtension = string.Empty;

                string searchPattern = cleanTitle + ".*";
                string[] files = Directory.GetFiles(directory, searchPattern, SearchOption.TopDirectoryOnly);

                var matches = files
                    .Where(f => string.Equals(Path.GetExtension(f), expectedExtension, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (matches.Count == 0)
                    return null;

                matches.Sort(StringComparer.OrdinalIgnoreCase);
                return matches[0];
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Finds the zero-based index of a header name in a sheet header row.
        /// </summary>
        /// <param name="headerRow">The sheet header row.</param>
        /// <param name="headerName">The header text to find.</param>
        /// <returns>The zero-based column index, or -1 if not found.</returns>
        private static int FindColumnIndex(IList<object> headerRow, string headerName)
        {
            if (headerRow == null)
                return -1;

            for (int i = 0; i < headerRow.Count; i++)
            {
                string cell = headerRow[i] == null ? string.Empty : headerRow[i].ToString();
                if (string.Equals(cell.Trim(), headerName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }

            return -1;
        }

        /// <summary>
        /// Safely gets a string value from a sheet row at the specified column index.
        /// </summary>
        /// <param name="row">The row to read.</param>
        /// <param name="index">The zero-based column index.</param>
        /// <returns>The cell value as a string, or an empty string if unavailable.</returns>
        private static string SafeGet(IList<object> row, int index)
        {
            if (index < 0 || index >= row.Count)
                return string.Empty;

            return row[index] == null ? string.Empty : row[index].ToString();
        }

        /// <summary>
        /// Normalizes a cell value by converting null to an empty string and trimming whitespace.
        /// </summary>
        /// <param name="value">The cell value to normalize.</param>
        /// <returns>The trimmed string value.</returns>
        private static string NormalizeCell(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        /// <summary>
        /// Builds a summary label using the clean title and record source when available.
        /// </summary>
        /// <param name="cleanTitle">The movie title from the sheet.</param>
        /// <param name="recordSource">The recorded source from the sheet.</param>
        /// <returns>A formatted summary string for console output.</returns>
        private static string BuildSummaryTitle(string cleanTitle, string recordSource)
        {
            if (string.IsNullOrWhiteSpace(recordSource))
                return cleanTitle;

            return cleanTitle + " - " + recordSource;
        }

        /// <summary>
        /// Writes a color-coded status token to the console without adding a newline.
        /// </summary>
        /// <param name="statusText">The status text to write.</param>
        private static void WriteColoredStatus(string statusText)
        {
            ConsoleColor originalColor = Console.ForegroundColor;

            if (string.Equals(statusText, "OK", StringComparison.OrdinalIgnoreCase))
                Console.ForegroundColor = ConsoleColor.Green;
            else if (string.Equals(statusText, "WARNING", StringComparison.OrdinalIgnoreCase))
                Console.ForegroundColor = ConsoleColor.Yellow;
            else
                Console.ForegroundColor = ConsoleColor.Red;

            Console.Write(statusText);
            Console.ForegroundColor = originalColor;
        }

        /// <summary>
        /// Converts a zero-based column index to Excel-style column letters.
        /// </summary>
        /// <param name="colIndex">The zero-based column index.</param>
        /// <returns>The Excel-style column letter value.</returns>
        private static string ColumnIndexToLetter(int colIndex)
        {
            int dividend = colIndex + 1;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}