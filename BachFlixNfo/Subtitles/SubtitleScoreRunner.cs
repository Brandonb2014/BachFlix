using BachFlixNfo.Subtitles;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BachFlixNfo.Subtitles
{
    public static class SubtitleScoreRunner
    {
        // Header labels in your sheet
        private const string HEADER_DIRECTORY = "Directory";
        private const string HEADER_CLEAN_TITLE = "Clean Title";
        private const string HEADER_SRT_SCORE = "SRT Score";

        /// <summary>
        /// Main entry point. Call this from your BachFlixNfo menu with your existing SheetsService + SPREADSHEET_ID.
        /// </summary>
        public static void Run(SheetsService sheetsService, string spreadsheetId, string sheetName)
        {
            var engine = new SrtScoringEngine();

            // 1. Read all rows
            string range = sheetName + "!A:ZZ";
            var getReq = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = getReq.Execute();
            IList<IList<object>> values = response.Values;

            if (values == null || values.Count == 0)
            {
                Console.WriteLine("No rows found in sheet.");
                return;
            }

            // 2. Find header row
            var headerRow = values[1];

            int colDirectory = FindColumnIndex(headerRow, HEADER_DIRECTORY);
            int colCleanTitle = FindColumnIndex(headerRow, HEADER_CLEAN_TITLE);
            int colSrtScore = FindColumnIndex(headerRow, HEADER_SRT_SCORE);

            if (colDirectory == -1 || colCleanTitle == -1 || colSrtScore == -1)
            {
                Console.WriteLine("Could not find one or more required headers:");
                Console.WriteLine($"  Directory:  {(colDirectory == -1 ? "NOT FOUND" : colDirectory.ToString())}");
                Console.WriteLine($"  Clean Title: {(colCleanTitle == -1 ? "NOT FOUND" : colCleanTitle.ToString())}");
                Console.WriteLine($"  SRT Score:  {(colSrtScore == -1 ? "NOT FOUND" : colSrtScore.ToString())}");
                return;
            }

            Console.WriteLine($"Using columns: Directory={ColumnIndexToLetter(colDirectory)}, Clean Title={ColumnIndexToLetter(colCleanTitle)}, SRT Score={ColumnIndexToLetter(colSrtScore)}");

            var updates = new List<ValueRange>();

            // 3. Process data rows (start at rowIndex 2 because 1 is header)
            for (int rowIndex = 2; rowIndex < values.Count; rowIndex++)
            {
                IList<object> row = values[rowIndex];

                // Pad row if needed so indexes are safe
                while (row.Count <= Math.Max(Math.Max(colDirectory, colCleanTitle), colSrtScore))
                    row.Add(string.Empty);

                string currentScore = SafeGet(row, colSrtScore);
                if (!string.IsNullOrWhiteSpace(currentScore))
                    continue; // already scored

                string directory = SafeGet(row, colDirectory);
                string cleanTitle = SafeGet(row, colCleanTitle);

                if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(cleanTitle))
                    continue;

                string videoPath = FindVideoFile(directory, cleanTitle);
                if (videoPath == null)
                {
                    //Console.WriteLine($"[Row {rowIndex + 1}] No video file found for title '{cleanTitle}' in '{directory}'");
                    continue;
                }

                int? score = engine.ScoreSubtitleForVideo(videoPath);
                if (score == null)
                {
                    //Console.WriteLine("   -> No SRT found or could not analyze. Skipping.");
                    continue;
                }


                Console.WriteLine($"[Row {rowIndex + 1}] Scanning: {videoPath}");
                Console.WriteLine($"   -> SRT Score: {score}/100");
                Console.WriteLine();

                string colLetter = ColumnIndexToLetter(colSrtScore);
                string cellRange = sheetName + "!" + colLetter + (rowIndex + 1); // +1 because sheet rows are 1-based

                updates.Add(new ValueRange
                {
                    Range = cellRange,
                    Values = new List<IList<object>>
                    {
                        new List<object> { score.Value }
                    }
                });
            }

            if (updates.Count == 0)
            {
                Console.WriteLine("No SRT scores to update.");
                return;
            }

            Console.WriteLine($"Updating {updates.Count} SRT scores in Google Sheet...");

            var batchRequest = new BatchUpdateValuesRequest
            {
                Data = updates,
                ValueInputOption = "RAW"
            };

            var batchUpdate = sheetsService.Spreadsheets.Values.BatchUpdate(batchRequest, spreadsheetId);
            batchUpdate.Execute();

            Console.WriteLine("SRT scores updated.");
        }

        /// <summary>
        /// Find a video file in the given directory whose file name (without extension)
        /// matches the given cleanTitle (exact match).
        /// </summary>
        private static string FindVideoFile(string directory, string cleanTitle)
        {
            try
            {
                if (!Directory.Exists(directory))
                {
                    return null;
                }

                // Typical video extensions – extend as needed
                var videoExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    ".mkv", ".mp4", ".avi", ".m4v", ".mov", ".wmv"
                };

                // Get all files whose base name matches cleanTitle
                var files = Directory.GetFiles(directory, cleanTitle + ".*", SearchOption.TopDirectoryOnly);

                // Prefer known video extensions
                var videoCandidates = files
                    .Where(f => videoExtensions.Contains(Path.GetExtension(f)))
                    .ToList();

                if (videoCandidates.Count > 0)
                    return videoCandidates.OrderBy(f => f).First();

                // Fallback: any file whose filename without extension exactly matches cleanTitle
                var anyCandidate = files.FirstOrDefault(f =>
                    string.Equals(Path.GetFileNameWithoutExtension(f), cleanTitle, StringComparison.OrdinalIgnoreCase));

                return anyCandidate;
            }
            catch
            {
                return null;
            }
        }

        private static int FindColumnIndex(IList<object> headerRow, string headerName)
        {
            if (headerRow == null) return -1;

            for (int i = 0; i < headerRow.Count; i++)
            {
                string cell = headerRow[i]?.ToString() ?? string.Empty;
                if (string.Equals(cell.Trim(), headerName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static string SafeGet(IList<object> row, int index)
        {
            if (index < 0 || index >= row.Count) return string.Empty;
            return row[index]?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Convert 0-based column index to Excel-style letter (0 -> A, 1 -> B, ... 25 -> Z, 26 -> AA, etc.)
        /// </summary>
        private static string ColumnIndexToLetter(int colIndex)
        {
            int dividend = colIndex + 1; // convert to 1-based
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
