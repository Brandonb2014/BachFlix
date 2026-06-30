using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Scans Movies sheet directories for movie files whose base filename does not match
    /// the Movies sheet title, shows suggested renames, and applies only after approval.
    /// </summary>
    public static class MovieFilenameSheetScanner
    {
        private const string HeaderDirectory = "Directory";
        private const string HeaderCleanTitle = "Clean Title";
        private const string HeaderImdbTitle = "IMDB Title";
        private const string HeaderImdbId = "IMDB ID";
        private const string HeaderAutoTitle = "Auto Title";
        private const string HeaderStatus = "Status";

        private static readonly string[] VideoExtensions =
        {
            ".mkv", ".mp4", ".m4v", ".avi", ".mov", ".wmv", ".mpg", ".mpeg", ".iso", ".webm"
        };

        private static readonly Regex TitleYearPrefixRegex =
            new Regex(@"^(?<base>.+\(\d{4}\))(?<suffix>(?:[.\- ].*)?)$", RegexOptions.Compiled);

        private static readonly Regex TrailerMarkerSuffixRegex =
            new Regex(@"^(?<title>.+?)(?:\s*-\s*|\s+)(?:official[\s._-]+)?trailer(?:[\s._-]*[A-Za-z0-9]+)?$",
                RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex TrailerSuffixOnlyRegex =
            new Regex(@"^(?:\s*-\s*|\s+)(?:official[\s._-]+)?trailer(?:[\s._-]*[A-Za-z0-9]+)?$",
                RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static void RunInteractive(
            SheetsService sheetsService,
            string spreadsheetId,
            string sheetName,
            Action<string, string, int> log)
        {
            if (sheetsService == null)
                throw new ArgumentNullException(nameof(sheetsService));
            if (string.IsNullOrWhiteSpace(spreadsheetId))
                throw new ArgumentException("Spreadsheet ID is required.", nameof(spreadsheetId));
            if (string.IsNullOrWhiteSpace(sheetName))
                sheetName = "Movies";

            log?.Invoke("info", "Movie filename scanner: reading Movies sheet...", 2);

            SheetData sheet = ReadSheet(sheetsService, spreadsheetId, sheetName);
            if (sheet.Rows.Count == 0)
            {
                log?.Invoke("warning", "No Movies sheet rows were found.", 2);
                return;
            }

            int directoryCol = FindColumnIndex(sheet.Headers, HeaderDirectory);
            int cleanTitleCol = FindColumnIndex(sheet.Headers, HeaderCleanTitle);
            int imdbTitleCol = FindColumnIndex(sheet.Headers, HeaderImdbTitle);
            int imdbIdCol = FindColumnIndex(sheet.Headers, HeaderImdbId);
            int autoTitleCol = FindColumnIndex(sheet.Headers, HeaderAutoTitle);
            int statusCol = FindColumnIndex(sheet.Headers, HeaderStatus);

            if (directoryCol < 0)
            {
                log?.Invoke("error", "Movies sheet is missing the Directory column.", 2);
                return;
            }

            if (cleanTitleCol < 0 && imdbTitleCol < 0 && autoTitleCol < 0)
            {
                log?.Invoke("error", "Movies sheet needs Clean Title, IMDB Title, or Auto Title for filename matching.", 2);
                return;
            }

            string rootFilter = "";
            log?.Invoke("info", "Folder filter: scanning every Movies sheet directory.", 1);
            bool activeOnly = statusCol >= 0;
            if (statusCol >= 0)
            {
                log?.Invoke("info", "Status filter: scanning Movies rows where Status = n.", 1);
            }
            else
            {
                log?.Invoke("warning", "Status column was not found, so all Movies rows will be scanned.", 1);
            }

            log?.Invoke("info", "Scanning movie folders for filename mismatches...", 2);

            ScanSummary summary;
            List<RenameFinding> findings = ScanRows(
                sheet,
                directoryCol,
                cleanTitleCol,
                imdbTitleCol,
                imdbIdCol,
                autoTitleCol,
                statusCol,
                rootFilter,
                activeOnly,
                out summary);

            if (findings.Count == 0)
            {
                DisplayScanSummary(summary, findings, log);
                log?.Invoke("success", "No filename mismatches were found.", 2);
                return;
            }

            DisplayFindings(findings, log);
            DisplayScanSummary(summary, findings, log);
            ConfirmAndApply(findings, log);
        }

        private static SheetData ReadSheet(SheetsService sheetsService, string spreadsheetId, string sheetName)
        {
            string range = QuoteSheetName(sheetName) + "!A:ZZ";
            SpreadsheetsResource.ValuesResource.GetRequest getRequest =
                sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = getRequest.Execute();
            IList<IList<object>> values = response.Values ?? new List<IList<object>>();

            if (values.Count < 2)
                return new SheetData { Headers = new List<object>(), Rows = new List<SheetRow>() };

            var rows = new List<SheetRow>();
            for (int i = 2; i < values.Count; i++)
            {
                rows.Add(new SheetRow
                {
                    SheetRowNumber = i + 1,
                    Values = values[i]
                });
            }

            return new SheetData
            {
                Headers = values[1],
                Rows = rows
            };
        }

        private static List<RenameFinding> ScanRows(
            SheetData sheet,
            int directoryCol,
            int cleanTitleCol,
            int imdbTitleCol,
            int imdbIdCol,
            int autoTitleCol,
            int statusCol,
            string rootFilter,
            bool activeOnly,
            out ScanSummary summary)
        {
            summary = new ScanSummary();
            var findings = new List<RenameFinding>();
            var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var duplicateTitleKeys = BuildDuplicateTitleKeyIndex(
                sheet,
                directoryCol,
                cleanTitleCol,
                imdbTitleCol,
                autoTitleCol,
                statusCol,
                rootFilter,
                activeOnly);

            foreach (SheetRow row in sheet.Rows)
            {
                if (!RowHasData(row.Values))
                    continue;

                if (activeOnly)
                {
                    string status = GetCell(row.Values, statusCol);
                    if (!status.Equals("n", StringComparison.OrdinalIgnoreCase))
                    {
                        summary.FilteredRows++;
                        continue;
                    }
                }

                string directory = GetCell(row.Values, directoryCol);
                string sheetTitle = FirstNonBlank(
                    GetCell(row.Values, cleanTitleCol),
                    GetCell(row.Values, imdbTitleCol),
                    GetCell(row.Values, autoTitleCol));
                string imdbId = GetCell(row.Values, imdbIdCol);

                if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(sheetTitle))
                {
                    summary.SkippedRows++;
                    continue;
                }

                if (!IsWithinRootFilter(directory, rootFilter))
                {
                    summary.FilteredRows++;
                    continue;
                }

                summary.RowsScanned++;

                if (!Directory.Exists(directory))
                {
                    summary.MissingDirectories++;
                    summary.MissingDirectoryRows.Add(new ScanIssue
                    {
                        SheetRowNumber = row.SheetRowNumber,
                        SheetTitle = sheetTitle,
                        Directory = directory,
                        Detail = "Directory does not exist."
                    });
                    continue;
                }

                string expectedBase = BuildExpectedBase(sheetTitle, imdbId, duplicateTitleKeys);
                if (string.IsNullOrWhiteSpace(expectedBase))
                {
                    summary.SkippedRows++;
                    continue;
                }

                FolderSnapshot snapshot = ReadFolder(directory);
                if (snapshot.ReadError != null)
                {
                    summary.FolderReadErrors++;
                    continue;
                }

                expectedBase = BuildExpectedBase(sheetTitle, imdbId, duplicateTitleKeys, snapshot);

                if (snapshot.PrimaryVideoFiles.Count == 0)
                {
                    summary.NoVideoFolders++;
                    summary.NoPrimaryVideoRows.Add(new ScanIssue
                    {
                        SheetRowNumber = row.SheetRowNumber,
                        SheetTitle = sheetTitle,
                        Directory = directory,
                        Detail = BuildNoPrimaryVideoDetail(snapshot)
                    });
                }

                bool rowHadMismatch = false;
                bool rowHadMatch = false;

                foreach (FolderItem item in snapshot.RelatedItems)
                {
                    RenameFinding trailerFinding;
                    if (TryBuildTrailerFinding(row, directory, sheetTitle, expectedBase, item, snapshot, out trailerFinding))
                    {
                        AddFindingIfNew(findings, trailerFinding, seenKeys);
                        rowHadMismatch = true;
                        continue;
                    }

                    NameParts parts;
                    if (!TrySplitRelatedName(item.Name, expectedBase, out parts))
                        continue;

                    if (string.Equals(parts.BaseName, expectedBase, StringComparison.Ordinal))
                    {
                        rowHadMatch = true;
                        continue;
                    }

                    RenameFinding finding = BuildFinding(
                        row,
                        directory,
                        sheetTitle,
                        expectedBase,
                        parts.BaseName,
                        item.ItemType,
                        parts.Suffix,
                        snapshot,
                        item.Path);

                    AddFindingIfNew(findings, finding, seenKeys);
                    rowHadMismatch = true;
                }

                string folderName = snapshot.FolderName ?? "";
                if (!string.Equals(folderName, expectedBase, StringComparison.Ordinal))
                {
                    RenameFinding finding = BuildFinding(
                        row,
                        directory,
                        sheetTitle,
                        expectedBase,
                        folderName,
                        "Movie Folder",
                        "",
                        snapshot,
                        directory);

                    AddFindingIfNew(findings, finding, seenKeys);
                    rowHadMismatch = true;
                }

                if (rowHadMismatch)
                    summary.MismatchRows++;
                else if (rowHadMatch || snapshot.NfoFiles.Any(n => string.Equals(Path.GetFileNameWithoutExtension(n), expectedBase, StringComparison.Ordinal)))
                    summary.MatchRows++;
            }

            for (int i = 0; i < findings.Count; i++)
                findings[i].Number = i + 1;

            return findings;
        }

        private static HashSet<string> BuildDuplicateTitleKeyIndex(
            SheetData sheet,
            int directoryCol,
            int cleanTitleCol,
            int imdbTitleCol,
            int autoTitleCol,
            int statusCol,
            string rootFilter,
            bool activeOnly)
        {
            var titleCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (SheetRow row in sheet.Rows)
            {
                if (!RowHasData(row.Values))
                    continue;

                if (activeOnly)
                {
                    string status = GetCell(row.Values, statusCol);
                    if (!status.Equals("n", StringComparison.OrdinalIgnoreCase))
                        continue;
                }

                string directory = GetCell(row.Values, directoryCol);
                if (!IsWithinRootFilter(directory, rootFilter))
                    continue;

                string sheetTitle = FirstNonBlank(
                    GetCell(row.Values, cleanTitleCol),
                    GetCell(row.Values, imdbTitleCol),
                    GetCell(row.Values, autoTitleCol));

                string key = NormalizeTitleKeyForDuplicateCheck(sheetTitle);
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                if (!titleCounts.ContainsKey(key))
                    titleCounts[key] = 0;

                titleCounts[key]++;
            }

            return new HashSet<string>(
                titleCounts.Where(kv => kv.Value > 1).Select(kv => kv.Key),
                StringComparer.OrdinalIgnoreCase);
        }

        private static string BuildExpectedBase(string sheetTitle, string imdbId, HashSet<string> duplicateTitleKeys, FolderSnapshot snapshot = null)
        {
            string expectedBase = SanitizeFileName(sheetTitle);
            if (string.IsNullOrWhiteSpace(expectedBase))
                return "";

            string duplicateKey = NormalizeTitleKeyForDuplicateCheck(sheetTitle);
            string normalizedImdbId = NormalizeImdbId(imdbId);
            bool needsImdbSuffix = duplicateTitleKeys != null &&
                                   duplicateTitleKeys.Contains(duplicateKey) &&
                                   !ContainsImdbDisambiguator(expectedBase);

            if (!needsImdbSuffix &&
                !string.IsNullOrWhiteSpace(normalizedImdbId) &&
                SnapshotHasMatchingImdbDisambiguator(snapshot, normalizedImdbId) &&
                !ContainsImdbDisambiguator(expectedBase))
            {
                needsImdbSuffix = true;
            }

            if (!needsImdbSuffix)
                return expectedBase;

            if (string.IsNullOrWhiteSpace(normalizedImdbId))
                return expectedBase;

            return SanitizeFileName(expectedBase + " {imdb-" + normalizedImdbId + "}");
        }

        private static RenameFinding BuildFinding(
            SheetRow row,
            string directory,
            string sheetTitle,
            string expectedBase,
            string currentBase,
            string itemType,
            string itemExtension,
            FolderSnapshot snapshot,
            string sourcePath)
        {
            var finding = new RenameFinding
            {
                SheetRowNumber = row.SheetRowNumber,
                Directory = directory,
                SheetTitle = sheetTitle,
                ExpectedBase = expectedBase,
                CurrentBase = currentBase,
                ItemType = itemType,
                ItemExtension = itemExtension ?? "",
                SourcePath = sourcePath,
                Reason = BuildReason(currentBase, expectedBase, snapshot),
                IncludeTrailerFilesInPrefixRename = false,
                Operations = BuildRenameOperations(directory, currentBase, expectedBase, false)
            };

            finding.HasConflicts = finding.Operations.Any(o => o.TargetExists);
            return finding;
        }

        private static bool TryBuildTrailerFinding(
            SheetRow row,
            string directory,
            string sheetTitle,
            string expectedBase,
            FolderItem item,
            FolderSnapshot snapshot,
            out RenameFinding finding)
        {
            finding = null;

            if (item == null || string.IsNullOrWhiteSpace(item.Path) || !IsVideoFile(item.Path))
                return false;

            string currentTrailerBase;
            string expectedTrailerBase;
            string extension;
            if (!TryBuildTrailerRenameBase(item.Name, expectedBase, out currentTrailerBase, out expectedTrailerBase, out extension))
                return false;

            finding = new RenameFinding
            {
                SheetRowNumber = row.SheetRowNumber,
                Directory = directory,
                SheetTitle = sheetTitle,
                ExpectedBase = expectedTrailerBase,
                CurrentBase = currentTrailerBase,
                ItemType = "Trailer",
                ItemExtension = extension ?? "",
                SourcePath = item.Path,
                Reason = BuildTrailerReason(currentTrailerBase, expectedTrailerBase, snapshot, item.Path),
                IsTrailerFinding = true,
                IncludeTrailerFilesInPrefixRename = true,
                Operations = BuildTrailerRenameOperations(directory, currentTrailerBase, expectedTrailerBase)
            };

            finding.HasConflicts = finding.Operations.Any(o => o.TargetExists);
            return true;
        }

        private static void AddFindingIfNew(List<RenameFinding> findings, RenameFinding finding, HashSet<string> seenKeys)
        {
            string key = finding.Directory + "|" + finding.CurrentBase + "|" + finding.ExpectedBase;
            if (seenKeys.Contains(key))
                return;

            seenKeys.Add(key);
            findings.Add(finding);
        }

        private static string BuildReason(string currentBase, string expectedBase, FolderSnapshot snapshot)
        {
            if (string.Equals(currentBase, expectedBase, StringComparison.OrdinalIgnoreCase))
                return "Case-only mismatch; the sheet title differs by capitalization.";

            if (snapshot.NfoFiles.Any(n => string.Equals(Path.GetFileNameWithoutExtension(n), expectedBase, StringComparison.Ordinal)))
                return "The NFO already matches the sheet title.";

            if (NormalizeForCompare(currentBase) == NormalizeForCompare(expectedBase))
                return "The title matches after spacing/punctuation cleanup.";

            if (string.Equals(snapshot.FolderName, expectedBase, StringComparison.OrdinalIgnoreCase) ||
                NormalizeForCompare(snapshot.FolderName) == NormalizeForCompare(expectedBase))
                return "The containing folder matches the sheet title.";

            return "Suggested from the Movies sheet row for this directory.";
        }

        private static FolderSnapshot ReadFolder(string directory)
        {
            var snapshot = new FolderSnapshot { FolderName = Path.GetFileName(directory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)) };

            try
            {
                string[] files = Directory.GetFiles(directory, "*", SearchOption.TopDirectoryOnly);
                string[] folders = Directory.GetDirectories(directory, "*", SearchOption.TopDirectoryOnly);

                snapshot.PrimaryVideoFiles = files
                    .Where(IsVideoFile)
                    .Where(f => !IsTrailerFileName(Path.GetFileNameWithoutExtension(f)))
                    .Where(f => !IsSampleFileName(Path.GetFileNameWithoutExtension(f)))
                    .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                snapshot.NfoFiles = files
                    .Where(f => string.Equals(Path.GetExtension(f), ".nfo", StringComparison.OrdinalIgnoreCase))
                    .Where(f => !IsSampleFileName(Path.GetFileNameWithoutExtension(f)))
                    .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                snapshot.RelatedItems = files
                    .Where(IsRelatedMovieFile)
                    .Where(f => !IsSampleFileName(Path.GetFileNameWithoutExtension(f)))
                    .Select(f => new FolderItem
                    {
                        Path = f,
                        Name = Path.GetFileName(f),
                        ItemType = GetRelatedFileType(f)
                    })
                    .Concat(folders
                        .Where(IsRelatedMovieFolder)
                        .Select(f => new FolderItem
                        {
                            Path = f,
                            Name = Path.GetFileName(f),
                            ItemType = "Folder"
                        }))
                    .OrderBy(i => i.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch (Exception ex)
            {
                snapshot.ReadError = ex.Message;
            }

            return snapshot;
        }

        private static List<RenameOperation> BuildRenameOperations(string directory, string oldBase, string newBase, bool includeTrailerFiles)
        {
            var operations = new List<RenameOperation>();
            if (string.IsNullOrWhiteSpace(directory) ||
                string.IsNullOrWhiteSpace(oldBase) ||
                string.IsNullOrWhiteSpace(newBase) ||
                !Directory.Exists(directory))
            {
                return operations;
            }

            string containingFolderName = Path.GetFileName(directory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            string containingFolderTargetName;
            if (!ShouldSkipNameBecauseAlreadyUsesBase(containingFolderName, newBase) &&
                TryBuildRenamedPrefixName(containingFolderName, oldBase, newBase, out containingFolderTargetName))
            {
                string parentDirectory = Path.GetDirectoryName(directory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                if (!string.IsNullOrWhiteSpace(parentDirectory))
                {
                    operations.Add(new RenameOperation
                    {
                        IsDirectory = true,
                        IsContainingDirectory = true,
                        SourcePath = directory,
                        TargetPath = Path.Combine(parentDirectory, containingFolderTargetName)
                    });
                }
            }

            foreach (string file in Directory.GetFiles(directory, "*", SearchOption.TopDirectoryOnly))
            {
                string currentName = Path.GetFileName(file);
                if (!includeTrailerFiles && IsTrailerFileName(Path.GetFileNameWithoutExtension(file)))
                    continue;

                if (ShouldSkipBecauseAlreadyUsesBase(file, currentName, newBase, includeTrailerFiles))
                    continue;

                string newName;
                if (!TryBuildRenamedPrefixName(currentName, oldBase, newBase, out newName))
                    continue;

                operations.Add(new RenameOperation
                {
                    IsDirectory = false,
                    SourcePath = file,
                    TargetPath = Path.Combine(directory, newName)
                });
            }

            foreach (string folder in Directory.GetDirectories(directory, "*", SearchOption.TopDirectoryOnly))
            {
                string currentName = Path.GetFileName(folder);
                if (ShouldSkipNameBecauseAlreadyUsesBase(currentName, newBase))
                    continue;

                string newName;
                if (!TryBuildRenamedPrefixName(currentName, oldBase, newBase, out newName))
                    continue;

                operations.Add(new RenameOperation
                {
                    IsDirectory = true,
                    SourcePath = folder,
                    TargetPath = Path.Combine(directory, newName)
                });
            }

            foreach (RenameOperation operation in operations)
            {
                bool samePath = PathsEqualIgnoreCase(operation.SourcePath, operation.TargetPath);
                operation.TargetExists = PathExists(operation.TargetPath) && !samePath;
            }

            return operations
                .GroupBy(o => o.SourcePath, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderBy(o => o.IsDirectory)
                .ThenBy(o => o.IsContainingDirectory)
                .ThenBy(o => o.SourcePath, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool ShouldSkipBecauseAlreadyUsesBase(string path, string currentName, string newBase, bool includeTrailerFiles)
        {
            if (!ShouldSkipNameBecauseAlreadyUsesBase(currentName, newBase))
                return false;

            if (includeTrailerFiles &&
                IsVideoFile(path) &&
                IsTrailerFileName(Path.GetFileNameWithoutExtension(path)) &&
                !string.Equals(Path.GetFileNameWithoutExtension(path), newBase, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return true;
        }

        private static bool ShouldSkipNameBecauseAlreadyUsesBase(string currentName, string newBase)
        {
            if (!NameAlreadyUsesBase(currentName, newBase))
                return false;

            return NameAlreadyUsesBaseExactly(currentName, newBase);
        }

        private static List<RenameOperation> BuildTrailerRenameOperations(string directory, string oldBase, string newBase)
        {
            List<RenameOperation> operations = BuildRenameOperations(directory, oldBase, newBase, true);

            foreach (RenameOperation operation in operations)
            {
                if (operation.TargetExists &&
                    !operation.IsDirectory &&
                    IsTrailerFileName(Path.GetFileNameWithoutExtension(operation.SourcePath)))
                {
                    operation.DeleteSource = true;
                    operation.TargetExists = false;
                }
            }

            return operations;
        }

        private static bool NameAlreadyUsesBase(string currentName, string expectedBase)
        {
            if (string.IsNullOrWhiteSpace(currentName) || string.IsNullOrWhiteSpace(expectedBase))
                return false;

            if (!currentName.StartsWith(expectedBase, StringComparison.OrdinalIgnoreCase))
                return false;

            string suffix = currentName.Substring(expectedBase.Length);
            return suffix.Length == 0 || suffix[0] == '.' || suffix[0] == '-' || suffix[0] == ' ';
        }

        private static bool NameAlreadyUsesBaseExactly(string currentName, string expectedBase)
        {
            if (string.IsNullOrWhiteSpace(currentName) || string.IsNullOrWhiteSpace(expectedBase))
                return false;

            if (!currentName.StartsWith(expectedBase, StringComparison.Ordinal))
                return false;

            string suffix = currentName.Substring(expectedBase.Length);
            return suffix.Length == 0 || suffix[0] == '.' || suffix[0] == '-' || suffix[0] == ' ';
        }

        private static bool TryBuildRenamedPrefixName(string currentName, string oldBase, string newBase, out string newName)
        {
            newName = "";

            if (string.IsNullOrWhiteSpace(currentName) ||
                string.IsNullOrWhiteSpace(oldBase) ||
                !currentName.StartsWith(oldBase, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string suffix = currentName.Substring(oldBase.Length);
            if (suffix.Length > 0 && suffix[0] != '.' && suffix[0] != '-' && suffix[0] != ' ')
                return false;

            string normalizedTrailerSuffix;
            newName = newBase + (TryNormalizeTrailerSuffix(currentName, suffix, out normalizedTrailerSuffix)
                ? normalizedTrailerSuffix
                : suffix);
            return !string.Equals(currentName, newName, StringComparison.Ordinal);
        }

        private static bool TryBuildTrailerRenameBase(
            string itemName,
            string expectedBase,
            out string currentTrailerBase,
            out string expectedTrailerBase,
            out string extension)
        {
            currentTrailerBase = "";
            expectedTrailerBase = "";
            extension = "";

            if (string.IsNullOrWhiteSpace(itemName) || string.IsNullOrWhiteSpace(expectedBase))
                return false;

            string trailerExtension = Path.GetExtension(itemName);
            if (string.IsNullOrWhiteSpace(trailerExtension) ||
                !VideoExtensions.Any(e => e.Equals(trailerExtension, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            extension = trailerExtension;

            string baseName = Path.GetFileNameWithoutExtension(itemName);
            string titleBase;
            if (!TryStripTrailerSuffix(baseName, out titleBase))
                return false;

            if (!TrailerTitleMatchesExpectedBase(titleBase, expectedBase))
                return false;

            currentTrailerBase = baseName;
            expectedTrailerBase = expectedBase + "-trailer";

            return !string.Equals(currentTrailerBase, expectedTrailerBase, StringComparison.Ordinal);
        }

        private static bool TryNormalizeTrailerSuffix(string currentName, string suffix, out string normalizedSuffix)
        {
            normalizedSuffix = "";

            string extension = Path.GetExtension(currentName);
            if (string.IsNullOrWhiteSpace(extension) ||
                !VideoExtensions.Any(e => e.Equals(extension, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(suffix) || suffix.Length <= extension.Length)
                return false;

            string suffixWithoutExtension = suffix.Substring(0, suffix.Length - extension.Length);
            if (!TrailerSuffixOnlyRegex.IsMatch(suffixWithoutExtension))
                return false;

            normalizedSuffix = "-trailer" + extension;
            return true;
        }

        private static bool TryStripTrailerSuffix(string baseName, out string titleBase)
        {
            titleBase = "";

            if (string.IsNullOrWhiteSpace(baseName))
                return false;

            Match match = TrailerMarkerSuffixRegex.Match(baseName.Trim());
            if (!match.Success)
                return false;

            titleBase = match.Groups["title"].Value.Trim().Trim('-', '_', ' ', '.');
            return !string.IsNullOrWhiteSpace(titleBase);
        }

        private static bool TrailerTitleMatchesExpectedBase(string trailerTitleBase, string expectedBase)
        {
            if (string.IsNullOrWhiteSpace(trailerTitleBase) || string.IsNullOrWhiteSpace(expectedBase))
                return false;

            return string.Equals(trailerTitleBase, expectedBase, StringComparison.OrdinalIgnoreCase) ||
                   NormalizeForCompare(trailerTitleBase) == NormalizeForCompare(expectedBase) ||
                   NormalizeTitleKeyForDuplicateCheck(trailerTitleBase) == NormalizeTitleKeyForDuplicateCheck(expectedBase);
        }

        private static string BuildTrailerReason(string currentTrailerBase, string expectedTrailerBase, FolderSnapshot snapshot, string sourcePath)
        {
            if (string.Equals(currentTrailerBase, expectedTrailerBase, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(currentTrailerBase, expectedTrailerBase, StringComparison.Ordinal))
            {
                return "Case-only trailer mismatch; the sheet title differs by capitalization.";
            }

            if (snapshot != null &&
                snapshot.RelatedItems != null &&
                snapshot.RelatedItems.Any(i =>
                    !string.IsNullOrWhiteSpace(i.Name) &&
                    !PathsEqualIgnoreCase(i.Path, sourcePath) &&
                    string.Equals(
                        Path.GetFileNameWithoutExtension(i.Name),
                        expectedTrailerBase,
                        StringComparison.Ordinal)))
            {
                return "A Plex-friendly trailer already exists, so the extra trailer-like file will be cleaned up.";
            }

            if (NormalizeForCompare(currentTrailerBase) == NormalizeForCompare(expectedTrailerBase))
                return "Trailer filename differs only by spacing/punctuation cleanup.";

            return "Trailer files must use the Plex-friendly -trailer suffix so they are not scanned as movies.";
        }

        private static string BuildNoPrimaryVideoDetail(FolderSnapshot snapshot)
        {
            if (snapshot == null || snapshot.RelatedItems == null)
                return "No primary movie video was found.";

            List<string> videoNames = snapshot.RelatedItems
                .Where(i => !string.IsNullOrWhiteSpace(i.Path) && IsVideoFile(i.Path))
                .Select(i => i.Name)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (videoNames.Count == 0)
                return "No video files were found in the folder.";

            return "Video-like files found: " + string.Join(", ", videoNames);
        }

        private static bool TrySplitRelatedName(string itemName, string expectedBase, out NameParts parts)
        {
            parts = null;

            if (string.IsNullOrWhiteSpace(itemName))
                return false;

            if (!string.IsNullOrWhiteSpace(expectedBase) &&
                itemName.StartsWith(expectedBase, StringComparison.OrdinalIgnoreCase))
            {
                string suffix = itemName.Substring(expectedBase.Length);
                if (suffix.Length == 0 || suffix[0] == '.' || suffix[0] == '-' || suffix[0] == ' ')
                {
                    parts = new NameParts
                    {
                        BaseName = itemName.Substring(0, expectedBase.Length).Trim(),
                        Suffix = suffix
                    };
                    return !string.IsNullOrWhiteSpace(parts.BaseName);
                }
            }

            Match yearMatch = TitleYearPrefixRegex.Match(itemName);
            if (yearMatch.Success)
            {
                parts = new NameParts
                {
                    BaseName = yearMatch.Groups["base"].Value.Trim(),
                    Suffix = yearMatch.Groups["suffix"].Value
                };
                return !string.IsNullOrWhiteSpace(parts.BaseName);
            }

            int trailerIndex = itemName.IndexOf("-trailer", StringComparison.OrdinalIgnoreCase);
            if (trailerIndex > 0)
            {
                parts = new NameParts
                {
                    BaseName = itemName.Substring(0, trailerIndex).Trim(),
                    Suffix = itemName.Substring(trailerIndex)
                };
                return !string.IsNullOrWhiteSpace(parts.BaseName);
            }

            string extension = Path.GetExtension(itemName);
            if (!string.IsNullOrWhiteSpace(extension))
            {
                parts = new NameParts
                {
                    BaseName = itemName.Substring(0, itemName.Length - extension.Length).Trim(),
                    Suffix = extension
                };
                return !string.IsNullOrWhiteSpace(parts.BaseName);
            }

            parts = new NameParts
            {
                BaseName = itemName.Trim(),
                Suffix = ""
            };

            return !string.IsNullOrWhiteSpace(parts.BaseName);
        }

        private static void DisplayScanSummary(ScanSummary summary, List<RenameFinding> findings, Action<string, string, int> log)
        {
            log?.Invoke("info", "Scan summary:", 1);
            log?.Invoke("data", "Rows scanned: " + summary.RowsScanned, 1);
            log?.Invoke("data", "Rows filtered/skipped: " + (summary.FilteredRows + summary.SkippedRows), 1);
            log?.Invoke("success", "Rows already matching: " + summary.MatchRows, 1);
            log?.Invoke("warning", "Rows with mismatches: " + summary.MismatchRows, 1);
            log?.Invoke("warning", "Folders with no primary video: " + summary.NoVideoFolders, 1);
            log?.Invoke("warning", "Missing directories: " + summary.MissingDirectories, 1);
            log?.Invoke("warning", "Folder read errors: " + summary.FolderReadErrors, 1);
            log?.Invoke("warning", "Rename suggestions found: " + findings.Count, 2);

            DisplayScanIssueList("Folders with no primary video:", summary.NoPrimaryVideoRows, log);
            DisplayScanIssueList("Missing directory details:", summary.MissingDirectoryRows, log);
        }

        private static void DisplayScanIssueList(string header, List<ScanIssue> issues, Action<string, string, int> log)
        {
            if (issues == null || issues.Count == 0)
                return;

            log?.Invoke("warning", header, 1);

            foreach (ScanIssue issue in issues.OrderBy(i => i.SheetRowNumber))
            {
                log?.Invoke("warning", "  Row " + issue.SheetRowNumber + " - " + issue.SheetTitle, 1);
                log?.Invoke("data", "    Directory: " + issue.Directory, 1);

                if (!string.IsNullOrWhiteSpace(issue.Detail))
                    log?.Invoke("info", "    " + issue.Detail, 1);
            }

            log?.Invoke("default", "", 1);
        }

        private static void DisplayFindings(List<RenameFinding> findings, Action<string, string, int> log)
        {
            log?.Invoke("info", "Suggested rename report:", 1);

            foreach (RenameFinding finding in findings)
            {
                DisplayFinding(finding, log, true);
            }
        }

        private static void DisplayFinding(RenameFinding finding, Action<string, string, int> log, bool includeOperations)
        {
            string suggestedName = finding.ExpectedBase + finding.ItemExtension;

            log?.Invoke("warning", "[" + finding.Number + "] Row " + finding.SheetRowNumber + " - " + finding.ItemType + " mismatch", 1);
            log?.Invoke("data", "Folder: " + finding.Directory, 1);
            log?.Invoke("data", "Sheet:  " + finding.SheetTitle, 1);
            log?.Invoke("data", "Found:  " + Path.GetFileName(finding.SourcePath), 1);
            log?.Invoke("success", "Suggest: " + suggestedName, 1);
            log?.Invoke("info", "Why: " + finding.Reason, 1);

            if (includeOperations)
            {
                if (finding.Operations.Count == 0)
                {
                    log?.Invoke("warning", "No file operations could be built for this suggestion.", 1);
                }
                else
                {
                    log?.Invoke("log", "Files/folders that would be changed:", 1);
                    foreach (RenameOperation operation in finding.Operations)
                    {
                        if (operation.DeleteSource)
                        {
                            string duplicateTarget = string.IsNullOrWhiteSpace(operation.TargetPath)
                                ? ""
                                : " (target already exists: " + Path.GetFileName(operation.TargetPath) + ")";
                            log?.Invoke("log", "  " + Path.GetFileName(operation.SourcePath) + " -> [delete duplicate trailer]" + duplicateTarget, 1);
                            continue;
                        }

                        string conflict = operation.TargetExists ? "  [TARGET EXISTS - will skip]" : "";
                        log?.Invoke("log", "  " + Path.GetFileName(operation.SourcePath) + " -> " + Path.GetFileName(operation.TargetPath) + conflict, 1);
                    }

                    int trailerFilenameFixCount = finding.Operations.Count(o =>
                        OperationTouchesTrailer(o) && !o.DeleteSource);

                    int duplicateTrailerCleanupCount = finding.Operations.Count(o =>
                        OperationTouchesTrailer(o) && o.DeleteSource);

                    int srtRenameCount = finding.Operations.Count(o =>
                        !o.IsDirectory &&
                        string.Equals(Path.GetExtension(o.SourcePath), ".srt", StringComparison.OrdinalIgnoreCase));

                    bool renamesMovieFolder = finding.Operations.Any(o => o.IsContainingDirectory);

                    log?.Invoke(
                        trailerFilenameFixCount > 0 ? "success" : "info",
                        "Trailer filename fix(es) in this group: " + trailerFilenameFixCount,
                        1);
                    log?.Invoke(
                        duplicateTrailerCleanupCount > 0 ? "success" : "info",
                        "Duplicate trailer cleanup(s) in this group: " + duplicateTrailerCleanupCount,
                        1);
                    log?.Invoke(
                        srtRenameCount > 0 ? "success" : "info",
                        "SRT rename(s) in this group: " + srtRenameCount,
                        1);
                    log?.Invoke(
                        renamesMovieFolder ? "success" : "info",
                        "Movie folder rename in this group: " + (renamesMovieFolder ? "yes" : "no"),
                        1);
                }
            }

            log?.Invoke("default", "", 1);
        }

        private static void ConfirmAndApply(List<RenameFinding> findings, Action<string, string, int> log)
        {
            int operationCount = findings.Sum(f => f.Operations.Count);
            if (operationCount == 0)
            {
                log?.Invoke("warning", "Nothing can be renamed from these suggestions.", 2);
                return;
            }

            var sessionLogLines = new List<string>();

            while (true)
            {
                log?.Invoke("question", "Apply suggestions? r = review one by one, a = apply all, n = no changes", 1);
                Console.Write("> ");
                string input = (Console.ReadLine() ?? "").Trim();

                if (string.IsNullOrWhiteSpace(input) || input.Equals("r", StringComparison.OrdinalIgnoreCase))
                {
                    ReviewOneByOne(findings, log, sessionLogLines);
                    WriteRenameLog(sessionLogLines, log);
                    return;
                }

                if (input.Equals("a", StringComparison.OrdinalIgnoreCase))
                {
                    log?.Invoke("question", "Type APPLY to run every suggested rename/delete operation that has no target conflict.", 1);
                    Console.Write("> ");
                    string confirm = (Console.ReadLine() ?? "").Trim();
                    if (confirm.Equals("APPLY", StringComparison.Ordinal))
                    {
                        ApplyFindings(findings, log, sessionLogLines);
                        WriteRenameLog(sessionLogLines, log);
                    }
                    else
                    {
                        log?.Invoke("warning", "No changes made.", 2);
                    }

                    return;
                }

                if (input.Equals("n", StringComparison.OrdinalIgnoreCase) ||
                    input.Equals("no", StringComparison.OrdinalIgnoreCase) ||
                    input.Equals("0", StringComparison.OrdinalIgnoreCase))
                {
                    log?.Invoke("warning", "No changes made.", 2);
                    return;
                }

                int findingNumber;
                if (int.TryParse(input, out findingNumber))
                {
                    RenameFinding finding = findings.FirstOrDefault(f => f.Number == findingNumber);
                    if (finding != null)
                    {
                        ReviewSingleFinding(finding, log, sessionLogLines);
                        WriteRenameLog(sessionLogLines, log);
                        return;
                    }
                }

                log?.Invoke("warning", "I did not recognize that option.", 1);
            }
        }

        private static void ReviewOneByOne(List<RenameFinding> findings, Action<string, string, int> log, List<string> sessionLogLines)
        {
            foreach (RenameFinding finding in findings)
            {
                RefreshFindingOperations(finding);
                if (finding.Operations.Count == 0)
                    continue;

                DisplayFinding(finding, log, true);
                if (!ReviewSingleFinding(finding, log, sessionLogLines))
                    break;
            }
        }

        private static bool ReviewSingleFinding(RenameFinding finding, Action<string, string, int> log, List<string> sessionLogLines)
        {
            RefreshFindingOperations(finding);

            while (true)
            {
                log?.Invoke("question", "Rename this group? y = yes, n = skip, e = edit target, q = quit review", 1);
                Console.Write("> ");
                string input = (Console.ReadLine() ?? "").Trim();

                if (input.Equals("y", StringComparison.OrdinalIgnoreCase) ||
                    input.Equals("yes", StringComparison.OrdinalIgnoreCase))
                {
                    ApplyFinding(finding, log, sessionLogLines);
                    return true;
                }

                if (input.Equals("n", StringComparison.OrdinalIgnoreCase) ||
                    input.Equals("no", StringComparison.OrdinalIgnoreCase) ||
                    string.IsNullOrWhiteSpace(input))
                {
                    log?.Invoke("warning", "Skipped.", 1);
                    return true;
                }

                if (input.Equals("q", StringComparison.OrdinalIgnoreCase) ||
                    input.Equals("quit", StringComparison.OrdinalIgnoreCase))
                {
                    log?.Invoke("warning", "Stopped review.", 2);
                    return false;
                }

                if (input.Equals("e", StringComparison.OrdinalIgnoreCase) ||
                    input.Equals("edit", StringComparison.OrdinalIgnoreCase))
                {
                    RenameFinding edited = BuildEditedFinding(finding, log);
                    if (edited == null)
                        continue;

                    DisplayFinding(edited, log, true);
                    log?.Invoke("question", "Apply this edited suggestion? (y/N)", 1);
                    Console.Write("> ");
                    string confirm = (Console.ReadLine() ?? "").Trim();
                    if (confirm.Equals("y", StringComparison.OrdinalIgnoreCase) ||
                        confirm.Equals("yes", StringComparison.OrdinalIgnoreCase))
                    {
                        ApplyFinding(edited, log, sessionLogLines);
                        return true;
                    }

                    log?.Invoke("warning", "Edited suggestion skipped.", 1);
                    return true;
                }

                log?.Invoke("warning", "I did not recognize that option.", 1);
            }
        }

        private static RenameFinding BuildEditedFinding(RenameFinding finding, Action<string, string, int> log)
        {
            log?.Invoke("question", "Enter the new base filename without extension. You can also paste a full video filename.", 1);
            Console.Write("> ");
            string input = (Console.ReadLine() ?? "").Trim();
            input = CleanConsolePath(input);

            if (string.IsNullOrWhiteSpace(input))
            {
                log?.Invoke("warning", "Empty target; returning to this suggestion.", 1);
                return null;
            }

            string customBase = StripKnownVideoExtension(Path.GetFileName(input));
            customBase = SanitizeFileName(customBase);
            if (string.IsNullOrWhiteSpace(customBase))
            {
                log?.Invoke("warning", "That target becomes empty after filename cleanup.", 1);
                return null;
            }

            return new RenameFinding
            {
                Number = finding.Number,
                SheetRowNumber = finding.SheetRowNumber,
                Directory = finding.Directory,
                SheetTitle = finding.SheetTitle,
                ExpectedBase = customBase,
                CurrentBase = finding.CurrentBase,
                ItemType = finding.ItemType,
                ItemExtension = finding.ItemExtension,
                SourcePath = finding.SourcePath,
                Reason = "Edited manually during review.",
                IncludeTrailerFilesInPrefixRename = true,
                Operations = BuildRenameOperations(finding.Directory, finding.CurrentBase, customBase, true)
            };
        }

        private static void ApplyFindings(List<RenameFinding> findings, Action<string, string, int> log, List<string> sessionLogLines)
        {
            foreach (RenameFinding finding in findings)
            {
                ApplyFinding(finding, log, sessionLogLines);
            }
        }

        private static void ApplyFinding(RenameFinding finding, Action<string, string, int> log, List<string> sessionLogLines)
        {
            RefreshFindingOperations(finding);

            int renamed = 0;
            int deleted = 0;
            int skipped = 0;

            sessionLogLines.Add("Row " + finding.SheetRowNumber + ": " + finding.CurrentBase + " -> " + finding.ExpectedBase);

            foreach (RenameOperation operation in finding.Operations)
            {
                if (operation.DeleteSource)
                {
                    try
                    {
                        DeletePath(operation);
                        deleted++;

                        string message = "DELETED duplicate trailer: " + operation.SourcePath;
                        sessionLogLines.Add(message);
                        log?.Invoke("success", "DELETED duplicate trailer: " + Path.GetFileName(operation.SourcePath), 1);
                    }
                    catch (Exception ex)
                    {
                        skipped++;
                        string message = "ERROR deleting duplicate trailer: " + operation.SourcePath + " | " + ex.Message;
                        sessionLogLines.Add(message);
                        log?.Invoke("error", message, 1);
                    }

                    continue;
                }

                bool targetExistsNow =
                    operation.TargetExists ||
                    (PathExists(operation.TargetPath) && !PathsEqualIgnoreCase(operation.SourcePath, operation.TargetPath));

                if (targetExistsNow && IsDuplicateTrailerOperation(operation))
                {
                    try
                    {
                        DeletePath(operation);
                        deleted++;

                        string message = "DELETED duplicate trailer: " + operation.SourcePath;
                        sessionLogLines.Add(message);
                        log?.Invoke("success", "DELETED duplicate trailer: " + Path.GetFileName(operation.SourcePath), 1);
                    }
                    catch (Exception ex)
                    {
                        skipped++;
                        string message = "ERROR deleting duplicate trailer: " + operation.SourcePath + " | " + ex.Message;
                        sessionLogLines.Add(message);
                        log?.Invoke("error", message, 1);
                    }

                    continue;
                }

                if (targetExistsNow)
                {
                    skipped++;
                    string message = "SKIP target exists: " + operation.TargetPath;
                    sessionLogLines.Add(message);
                    log?.Invoke("warning", message, 1);
                    continue;
                }

                try
                {
                    MovePath(operation);
                    renamed++;

                    string message = "RENAMED: " + operation.SourcePath + " -> " + operation.TargetPath;
                    sessionLogLines.Add(message);
                    log?.Invoke("success", Path.GetFileName(operation.SourcePath) + " -> " + Path.GetFileName(operation.TargetPath), 1);
                }
                catch (Exception ex)
                {
                    skipped++;
                    string message = "ERROR: " + operation.SourcePath + " -> " + operation.TargetPath + " | " + ex.Message;
                    sessionLogLines.Add(message);
                    log?.Invoke("error", message, 1);
                }
            }

            sessionLogLines.Add("Result: Renamed=" + renamed + ", Deleted=" + deleted + ", Skipped=" + skipped);
            sessionLogLines.Add("");
            log?.Invoke("info", "Rename group complete. Renamed=" + renamed + ", Deleted=" + deleted + ", Skipped=" + skipped, 2);
        }

        private static void RefreshFindingOperations(RenameFinding finding)
        {
            if (finding == null)
                return;

            string currentDirectory = ResolveCurrentDirectoryForFinding(finding);
            if (string.IsNullOrWhiteSpace(currentDirectory))
                return;

            finding.Directory = currentDirectory;
            finding.Operations = finding.IsTrailerFinding
                ? BuildTrailerRenameOperations(currentDirectory, finding.CurrentBase, finding.ExpectedBase)
                : BuildRenameOperations(
                    currentDirectory,
                    finding.CurrentBase,
                    finding.ExpectedBase,
                    finding.IncludeTrailerFilesInPrefixRename);

            finding.HasConflicts = finding.Operations.Any(o => o.TargetExists);

            RenameOperation sourceOperation = finding.Operations
                .FirstOrDefault(o => !o.IsContainingDirectory &&
                                     string.Equals(
                                         Path.GetFileNameWithoutExtension(o.SourcePath),
                                         finding.CurrentBase,
                                         StringComparison.OrdinalIgnoreCase));

            if (sourceOperation == null)
                sourceOperation = finding.Operations.FirstOrDefault(o => !o.IsContainingDirectory);

            if (sourceOperation != null)
                finding.SourcePath = sourceOperation.SourcePath;
        }

        private static string ResolveCurrentDirectoryForFinding(RenameFinding finding)
        {
            if (finding == null || string.IsNullOrWhiteSpace(finding.Directory))
                return "";

            if (Directory.Exists(finding.Directory))
                return finding.Directory;

            string originalDirectory = finding.Directory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string parentDirectory = Path.GetDirectoryName(originalDirectory);
            if (string.IsNullOrWhiteSpace(parentDirectory) || !Directory.Exists(parentDirectory))
                return finding.Directory;

            string expectedFolderName = GetExpectedMovieFolderBase(finding);
            if (!string.IsNullOrWhiteSpace(expectedFolderName))
            {
                string directCandidate = Path.Combine(parentDirectory, expectedFolderName);
                if (Directory.Exists(directCandidate))
                    return directCandidate;
            }

            try
            {
                string expectedCompare = NormalizeForCompare(expectedFolderName);
                foreach (string candidate in Directory.GetDirectories(parentDirectory, "*", SearchOption.TopDirectoryOnly))
                {
                    string candidateName = Path.GetFileName(candidate);
                    if (string.Equals(candidateName, expectedFolderName, StringComparison.OrdinalIgnoreCase) ||
                        NormalizeForCompare(candidateName) == expectedCompare)
                    {
                        return candidate;
                    }
                }
            }
            catch
            {
            }

            return finding.Directory;
        }

        private static string GetExpectedMovieFolderBase(RenameFinding finding)
        {
            if (finding == null)
                return "";

            if (finding.IsTrailerFinding)
            {
                string titleBase;
                if (TryStripTrailerSuffix(finding.ExpectedBase, out titleBase))
                    return titleBase;
            }

            return finding.ExpectedBase ?? "";
        }

        private static void WriteRenameLog(List<string> sessionLogLines, Action<string, string, int> log)
        {
            if (sessionLogLines == null || sessionLogLines.Count == 0)
                return;

            string error;
            string path = global::BachFlixLog.WriteBachFlixLog(
                sessionLogLines,
                "Movie Filename Scanner",
                "MovieFilenameScanner",
                out error);

            if (!string.IsNullOrWhiteSpace(path))
                log?.Invoke("success", "Rename log saved to: " + path, 1);
            else if (!string.IsNullOrWhiteSpace(error))
                log?.Invoke("warning", "Could not write rename log: " + error, 1);
        }

        private static void MovePath(RenameOperation operation)
        {
            if (PathsEqual(operation.SourcePath, operation.TargetPath))
                return;

            if (PathsEqualIgnoreCase(operation.SourcePath, operation.TargetPath))
            {
                string tempPath = BuildTemporarySiblingPath(operation.SourcePath);
                if (operation.IsDirectory)
                {
                    Directory.Move(operation.SourcePath, tempPath);
                    Directory.Move(tempPath, operation.TargetPath);
                }
                else
                {
                    File.Move(operation.SourcePath, tempPath);
                    File.Move(tempPath, operation.TargetPath);
                }

                return;
            }

            if (operation.IsDirectory)
                Directory.Move(operation.SourcePath, operation.TargetPath);
            else
                File.Move(operation.SourcePath, operation.TargetPath);
        }

        private static void DeletePath(RenameOperation operation)
        {
            if (operation.IsDirectory)
                Directory.Delete(operation.SourcePath, false);
            else
                File.Delete(operation.SourcePath);
        }

        private static bool IsDuplicateTrailerOperation(RenameOperation operation)
        {
            return operation != null &&
                   !operation.IsDirectory &&
                   IsTrailerFileName(Path.GetFileNameWithoutExtension(operation.SourcePath)) &&
                   IsTrailerFileName(Path.GetFileNameWithoutExtension(operation.TargetPath));
        }

        private static bool OperationTouchesTrailer(RenameOperation operation)
        {
            return operation != null &&
                   !operation.IsDirectory &&
                   (IsTrailerFileName(Path.GetFileNameWithoutExtension(operation.SourcePath)) ||
                    IsTrailerFileName(Path.GetFileNameWithoutExtension(operation.TargetPath)));
        }

        private static string BuildTemporarySiblingPath(string sourcePath)
        {
            string directory = Path.GetDirectoryName(sourcePath) ?? "";
            string name = Path.GetFileName(sourcePath);
            string candidate;

            do
            {
                candidate = Path.Combine(directory, ".bachflix-rename-" + Guid.NewGuid().ToString("N") + "-" + name);
            }
            while (PathExists(candidate));

            return candidate;
        }

        private static bool IsVideoFile(string path)
        {
            string ext = Path.GetExtension(path);
            return VideoExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsRelatedMovieFile(string path)
        {
            string ext = Path.GetExtension(path);
            if (VideoExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase)))
                return true;

            return ext.Equals(".nfo", StringComparison.OrdinalIgnoreCase) ||
                   ext.Equals(".srt", StringComparison.OrdinalIgnoreCase) ||
                   ext.Equals(".bif", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRelatedMovieFolder(string path)
        {
            string name = Path.GetFileName(path) ?? "";
            return name.EndsWith(".trickplay", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetRelatedFileType(string path)
        {
            string ext = Path.GetExtension(path);
            string baseName = Path.GetFileNameWithoutExtension(path);

            if (VideoExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase)))
                return IsTrailerFileName(baseName) ? "Trailer" : "Video";

            if (ext.Equals(".nfo", StringComparison.OrdinalIgnoreCase))
                return "NFO";

            if (ext.Equals(".srt", StringComparison.OrdinalIgnoreCase))
                return "SRT";

            if (ext.Equals(".bif", StringComparison.OrdinalIgnoreCase))
                return "BIF";

            return "File";
        }

        private static bool IsTrailerFileName(string baseName)
        {
            if (string.IsNullOrWhiteSpace(baseName))
                return false;

            string titleBase;
            return TryStripTrailerSuffix(baseName, out titleBase);
        }

        private static bool IsSampleFileName(string baseName)
        {
            if (string.IsNullOrWhiteSpace(baseName))
                return false;

            return baseName.ToLowerInvariant().Contains("sample");
        }

        private static bool IsWithinRootFilter(string directory, string rootFilter)
        {
            if (string.IsNullOrWhiteSpace(rootFilter))
                return true;

            string fullDirectory = NormalizePathForCompare(directory);
            string fullRoot = NormalizePathForCompare(rootFilter).TrimEnd('\\');

            return fullDirectory.Equals(fullRoot, StringComparison.OrdinalIgnoreCase) ||
                   fullDirectory.StartsWith(fullRoot + "\\", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizePathForCompare(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return "";

            try
            {
                return Path.GetFullPath(path).TrimEnd('\\', '/');
            }
            catch
            {
                return path.Trim().TrimEnd('\\', '/');
            }
        }

        private static bool PathExists(string path)
        {
            return File.Exists(path) || Directory.Exists(path);
        }

        private static bool PathsEqual(string left, string right)
        {
            return string.Equals(
                NormalizePathForCompare(left),
                NormalizePathForCompare(right),
                StringComparison.Ordinal);
        }

        private static bool PathsEqualIgnoreCase(string left, string right)
        {
            return string.Equals(
                NormalizePathForCompare(left),
                NormalizePathForCompare(right),
                StringComparison.OrdinalIgnoreCase);
        }

        private static string StripKnownVideoExtension(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return "";

            string ext = Path.GetExtension(fileName);
            if (!string.IsNullOrWhiteSpace(ext) &&
                VideoExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase)))
            {
                return fileName.Substring(0, fileName.Length - ext.Length);
            }

            return fileName;
        }

        private static string CleanConsolePath(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "";

            input = input.Trim();
            while (input.Length >= 2 &&
                   ((input.StartsWith("\"") && input.EndsWith("\"")) ||
                    (input.StartsWith("'") && input.EndsWith("'"))))
            {
                input = input.Substring(1, input.Length - 2).Trim();
            }

            return input;
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "";

            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);

            foreach (char ch in name)
            {
                if (!invalid.Contains(ch))
                    sb.Append(ch);
            }

            return Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
        }

        private static string NormalizeForCompare(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            value = value.Replace('_', ' ');
            value = Regex.Replace(value, @"[^\w\s]", "");
            value = Regex.Replace(value, @"\s+", " ").Trim();
            return value.ToLowerInvariant();
        }

        private static string NormalizeTitleKeyForDuplicateCheck(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            value = Regex.Replace(value, @"\s*\{imdb-tt\d+\}\s*", " ", RegexOptions.IgnoreCase);
            return NormalizeForCompare(value);
        }

        private static bool ContainsImdbDisambiguator(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            return Regex.IsMatch(value, @"\{imdb-tt\d+\}", RegexOptions.IgnoreCase);
        }

        private static bool SnapshotHasMatchingImdbDisambiguator(FolderSnapshot snapshot, string normalizedImdbId)
        {
            if (snapshot == null || string.IsNullOrWhiteSpace(normalizedImdbId))
                return false;

            string token = "{imdb-" + normalizedImdbId + "}";

            if (!string.IsNullOrWhiteSpace(snapshot.FolderName) &&
                snapshot.FolderName.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return snapshot.RelatedItems != null &&
                   snapshot.RelatedItems.Any(i =>
                       !string.IsNullOrWhiteSpace(i.Name) &&
                       i.Name.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static string NormalizeImdbId(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            value = value.Trim();

            int ttIndex = value.IndexOf("tt", StringComparison.OrdinalIgnoreCase);
            if (ttIndex >= 0)
                value = value.Substring(ttIndex);

            value = new string(value.Where(char.IsLetterOrDigit).ToArray());
            if (!value.StartsWith("tt", StringComparison.OrdinalIgnoreCase))
                return "";

            string digits = new string(value.Substring(2).Where(char.IsDigit).ToArray());
            if (string.IsNullOrWhiteSpace(digits))
                return "";

            return "tt" + digits;
        }

        private static int FindColumnIndex(IList<object> headers, params string[] names)
        {
            if (headers == null || names == null)
                return -1;

            for (int i = 0; i < headers.Count; i++)
            {
                string header = (headers[i] ?? "").ToString().Trim();
                foreach (string name in names)
                {
                    if (header.Equals(name, StringComparison.OrdinalIgnoreCase))
                        return i;
                }
            }

            for (int i = 0; i < headers.Count; i++)
            {
                string header = NormalizeHeader((headers[i] ?? "").ToString());
                foreach (string name in names)
                {
                    if (header.Equals(NormalizeHeader(name), StringComparison.OrdinalIgnoreCase))
                        return i;
                }
            }

            return -1;
        }

        private static string NormalizeHeader(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            return new string(value.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
        }

        private static string GetCell(IList<object> row, int index)
        {
            if (row == null || index < 0 || index >= row.Count)
                return "";

            return (row[index] ?? "").ToString().Trim();
        }

        private static string FirstNonBlank(params string[] values)
        {
            if (values == null)
                return "";

            foreach (string value in values)
            {
                if (!string.IsNullOrWhiteSpace(value))
                    return value.Trim();
            }

            return "";
        }

        private static bool RowHasData(IList<object> row)
        {
            return row != null && row.Any(v => !string.IsNullOrWhiteSpace((v ?? "").ToString()));
        }

        private static string QuoteSheetName(string sheetName)
        {
            return "'" + (sheetName ?? "").Replace("'", "''") + "'";
        }

        private sealed class SheetData
        {
            public IList<object> Headers { get; set; }
            public List<SheetRow> Rows { get; set; }
        }

        private sealed class SheetRow
        {
            public int SheetRowNumber { get; set; }
            public IList<object> Values { get; set; }
        }

        private sealed class FolderSnapshot
        {
            public string FolderName { get; set; }
            public string ReadError { get; set; }
            public List<string> PrimaryVideoFiles { get; set; } = new List<string>();
            public List<string> NfoFiles { get; set; } = new List<string>();
            public List<FolderItem> RelatedItems { get; set; } = new List<FolderItem>();
        }

        private sealed class FolderItem
        {
            public string Path { get; set; }
            public string Name { get; set; }
            public string ItemType { get; set; }
        }

        private sealed class NameParts
        {
            public string BaseName { get; set; }
            public string Suffix { get; set; }
        }

        private sealed class RenameFinding
        {
            public int Number { get; set; }
            public int SheetRowNumber { get; set; }
            public string Directory { get; set; }
            public string SheetTitle { get; set; }
            public string ExpectedBase { get; set; }
            public string CurrentBase { get; set; }
            public string ItemType { get; set; }
            public string ItemExtension { get; set; }
            public string SourcePath { get; set; }
            public string Reason { get; set; }
            public bool HasConflicts { get; set; }
            public bool IsTrailerFinding { get; set; }
            public bool IncludeTrailerFilesInPrefixRename { get; set; }
            public List<RenameOperation> Operations { get; set; }
        }

        private sealed class RenameOperation
        {
            public bool IsDirectory { get; set; }
            public bool IsContainingDirectory { get; set; }
            public string SourcePath { get; set; }
            public string TargetPath { get; set; }
            public bool TargetExists { get; set; }
            public bool DeleteSource { get; set; }
        }

        private sealed class ScanIssue
        {
            public int SheetRowNumber { get; set; }
            public string SheetTitle { get; set; }
            public string Directory { get; set; }
            public string Detail { get; set; }
        }

        private sealed class ScanSummary
        {
            public int RowsScanned { get; set; }
            public int FilteredRows { get; set; }
            public int SkippedRows { get; set; }
            public int MatchRows { get; set; }
            public int MismatchRows { get; set; }
            public int MissingDirectories { get; set; }
            public int NoVideoFolders { get; set; }
            public int FolderReadErrors { get; set; }
            public List<ScanIssue> MissingDirectoryRows { get; set; } = new List<ScanIssue>();
            public List<ScanIssue> NoPrimaryVideoRows { get; set; } = new List<ScanIssue>();
        }
    }
}
