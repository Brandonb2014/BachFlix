using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace BachFlixNfo.Features
{
    /// <summary>
    /// Adds an English AC3 5.1 compatibility track to MKV files from an English EAC3 5.1 source
    /// while copying all existing streams and leaving video untouched.
    /// </summary>
    public static class AudioCompatibilityConverter
    {
        private static readonly string[] SupportedVideoExtensions = { ".mkv" };
        private const string TempWorkFolderName = "BachFlixNfo-AudioCompatibility-Work";

        private enum InteractiveMode
        {
            Cancel,
            DryRunFolder,
            SafeFolder,
            ReplaceFolderWithBackup,
            DryRunLibrary,
            SafeLibrary,
            ReplaceLibraryWithBackup
        }

        public sealed class Options
        {
            public IEnumerable<string> RootPaths { get; set; }
            public bool Recursive { get; set; }
            public bool DryRun { get; set; }
            public bool ReplaceOriginal { get; set; }
            public bool BackupOriginal { get; set; }
            public int Ac3BitrateKbps { get; set; }
            public string FfmpegPath { get; set; }
            public string FfprobePath { get; set; }
            public Action<string, string, int> Log { get; set; }

            public Options()
            {
                RootPaths = new List<string>();
                Recursive = true;
                DryRun = true;
                ReplaceOriginal = false;
                BackupOriginal = true;
                Ac3BitrateKbps = 640;
            }
        }

        public sealed class Summary
        {
            public int FilesFound { get; set; }
            public int FilesInspected { get; set; }
            public int Converted { get; set; }
            public int DryRunMatches { get; set; }
            public int Skipped { get; set; }
            public int Failed { get; set; }
            public List<string> LogLines { get; private set; }

            public Summary()
            {
                LogLines = new List<string>();
            }
        }

        private sealed class AudioStreamInfo
        {
            public int StreamIndex { get; set; }
            public int AudioIndex { get; set; }
            public string CodecName { get; set; }
            public string Language { get; set; }
            public string Title { get; set; }
            public int Channels { get; set; }
            public string ChannelLayout { get; set; }

            public string Describe()
            {
                string language = string.IsNullOrWhiteSpace(Language) ? "(none)" : Language;
                string title = string.IsNullOrWhiteSpace(Title) ? "" : ", title=" + Title;
                string layout = string.IsNullOrWhiteSpace(ChannelLayout) ? "(none)" : ChannelLayout;

                return string.Format(
                    "stream={0}, audioIndex={1}, codec={2}, language={3}, channels={4}, layout={5}{6}",
                    StreamIndex,
                    AudioIndex,
                    CodecName,
                    language,
                    Channels,
                    layout,
                    title);
            }
        }

        private sealed class FileDecision
        {
            public bool ShouldConvert { get; set; }
            public string Reason { get; set; }
            public AudioStreamInfo SourceEac3Stream { get; set; }
            public List<AudioStreamInfo> AudioStreams { get; set; }

            public FileDecision()
            {
                Reason = "";
                AudioStreams = new List<AudioStreamInfo>();
            }
        }

        private sealed class ProcessResult
        {
            public int ExitCode { get; set; }
            public string StandardOutput { get; set; }
            public string StandardError { get; set; }

            public ProcessResult()
            {
                StandardOutput = "";
                StandardError = "";
            }
        }

        public static void RunInteractive(Action<string, string, int> log, IEnumerable<string> defaultLibraryRoots)
        {
            Write(log, "info", "=== AUDIO COMPATIBILITY CONVERTER ===", 1);
            Write(log, "default", "Adds an English AC3 5.1 track from English EAC3 5.1 while copying video/subtitles.", 1);
            Write(log, "warning", "Default is dry-run or safe output. Original files are only replaced in the explicit replace-with-backup modes.", 2);

            InteractiveMode mode = PromptMode(log);
            if (mode == InteractiveMode.Cancel)
            {
                Write(log, "warning", "Audio compatibility conversion cancelled.", 2);
                return;
            }

            bool libraryMode = mode == InteractiveMode.DryRunLibrary ||
                               mode == InteractiveMode.SafeLibrary ||
                               mode == InteractiveMode.ReplaceLibraryWithBackup;

            List<string> roots = libraryMode
                ? GetExistingDefaultRoots(defaultLibraryRoots, log)
                : new List<string>();

            if (!libraryMode)
            {
                string folder = PromptForFolder(log, "Enter the folder to scan recursively");
                if (string.IsNullOrWhiteSpace(folder))
                    return;

                roots.Add(folder);
            }
            else if (roots.Count == 0)
            {
                Write(log, "warning", "No existing default library roots were found. You can enter a single library root instead.", 1);
                string folder = PromptForFolder(log, "Enter the library root to scan recursively");
                if (string.IsNullOrWhiteSpace(folder))
                    return;

                roots.Add(folder);
            }

            int bitrate = PromptBitrate(log);

            var options = new Options
            {
                RootPaths = roots,
                Recursive = true,
                DryRun = mode == InteractiveMode.DryRunFolder || mode == InteractiveMode.DryRunLibrary,
                ReplaceOriginal = mode == InteractiveMode.ReplaceFolderWithBackup || mode == InteractiveMode.ReplaceLibraryWithBackup,
                BackupOriginal = true,
                Ac3BitrateKbps = bitrate,
                Log = log
            };

            Run(options);
        }

        public static Summary Run(Options options)
        {
            if (options == null)
                options = new Options();

            Summary summary = new Summary();
            AddLog(summary, "RUN", "", "Started " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            AddLog(summary, "RUN", "", "DryRun=" + options.DryRun + ", ReplaceOriginal=" + options.ReplaceOriginal + ", BackupOriginal=" + options.BackupOriginal + ", Bitrate=" + options.Ac3BitrateKbps + "k");

            string ffprobe = ResolveExecutable(options.FfprobePath, "ffprobe");
            string ffmpeg = ResolveExecutable(options.FfmpegPath, "ffmpeg");

            if (!VerifyTool(ffprobe, "ffprobe", options.Log, summary) ||
                !VerifyTool(ffmpeg, "ffmpeg", options.Log, summary))
            {
                summary.Failed++;
                WriteLogFile(summary, options.Log);
                return summary;
            }

            List<string> roots = NormalizeRootPaths(options.RootPaths);
            if (roots.Count == 0)
            {
                Write(options.Log, "warning", "No folders were selected for scanning.", 2);
                AddLog(summary, "SKIP", "", "No folders were selected for scanning.");
                WriteLogFile(summary, options.Log);
                return summary;
            }

            Write(options.Log, "info", "Scanning for MKV files...", 1);

            List<string> files = new List<string>();
            foreach (string root in roots)
            {
                AddLog(summary, "ROOT", root, "Scanning root");
                files.AddRange(EnumerateVideoFiles(root, options.Recursive, options.Log, summary));
            }

            files = files
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();

            summary.FilesFound = files.Count;
            Write(options.Log, "data", "Video files found: " + summary.FilesFound, 2);

            for (int i = 0; i < files.Count; i++)
            {
                string file = files[i];
                Write(options.Log, "default", "[" + (i + 1) + "/" + files.Count + "] " + file, 1);

                try
                {
                    ProcessFile(file, ffprobe, ffmpeg, options, summary);
                }
                catch (Exception ex)
                {
                    summary.Failed++;
                    Write(options.Log, "error", "FAILED: " + Path.GetFileName(file) + " | " + ex.Message, 1);
                    AddLog(summary, "FAILED", file, ex.ToString());
                }
            }

            WriteSummary(summary, options.Log);
            WriteLogFile(summary, options.Log);
            return summary;
        }

        private static void ProcessFile(string file, string ffprobe, string ffmpeg, Options options, Summary summary)
        {
            summary.FilesInspected++;

            FileDecision decision = InspectFile(file, ffprobe);
            AddLog(summary, "INSPECT", file, "Audio streams found: " + decision.AudioStreams.Count);
            foreach (AudioStreamInfo stream in decision.AudioStreams)
                AddLog(summary, "AUDIO", file, stream.Describe());

            if (!decision.ShouldConvert)
            {
                summary.Skipped++;
                Write(options.Log, "warning", "SKIP: " + decision.Reason, 1);
                AddLog(summary, "SKIPPED", file, decision.Reason);
                return;
            }

            if (decision.SourceEac3Stream == null)
            {
                summary.Failed++;
                Write(options.Log, "error", "FAILED: Converter selected no EAC3 source stream.", 1);
                AddLog(summary, "FAILED", file, "Converter selected no EAC3 source stream.");
                return;
            }

            string safeOutputPath = BuildSafeOutputPath(file);
            if (!options.ReplaceOriginal && File.Exists(safeOutputPath))
            {
                summary.Skipped++;
                string reason = "Safe output already exists: " + safeOutputPath;
                Write(options.Log, "warning", "SKIP: " + reason, 1);
                AddLog(summary, "SKIPPED", file, reason);
                return;
            }

            string sourceDescription = "English EAC3 5.1 " + decision.SourceEac3Stream.Describe();

            if (options.DryRun)
            {
                summary.DryRunMatches++;
                Write(options.Log, "success", "DRY RUN: would add AC3 5.1 from " + sourceDescription, 1);
                AddLog(summary, "DRY-RUN", file, "Would add AC3 5.1 from " + sourceDescription);
                return;
            }

            string tempPath = BuildTempOutputPath(file);
            int newAudioOutputIndex = decision.AudioStreams.Count;
            string ffmpegArguments = BuildFfmpegArguments(
                file,
                tempPath,
                decision.SourceEac3Stream.AudioIndex,
                newAudioOutputIndex,
                options.Ac3BitrateKbps);

            AddLog(summary, "FFMPEG", file, ffmpeg + " " + ffmpegArguments);
            Write(options.Log, "info", "Converting: " + Path.GetFileName(file), 1);

            ProcessResult ffmpegResult = RunProcess(ffmpeg, ffmpegArguments);
            if (ffmpegResult.ExitCode != 0)
            {
                summary.Failed++;
                TryDelete(tempPath);
                TryDeleteEmptyTempWorkFolder(tempPath);
                string error = FirstUsefulLine(ffmpegResult.StandardError, ffmpegResult.StandardOutput);
                Write(options.Log, "error", "FAILED: FFmpeg exit code " + ffmpegResult.ExitCode + " | " + error, 1);
                AddLog(summary, "FAILED", file, "FFmpeg exit code " + ffmpegResult.ExitCode + ". " + LastLines(ffmpegResult.StandardError, 20));
                return;
            }

            FileDecision outputDecision = InspectFile(tempPath, ffprobe);
            bool outputHasEnglishAc351 = outputDecision.AudioStreams.Any(IsEnglishAc351);
            bool outputKeptEnglishEac3 = outputDecision.AudioStreams.Any(IsEnglishEac3);

            if (!outputHasEnglishAc351 || !outputKeptEnglishEac3)
            {
                summary.Failed++;
                TryDelete(tempPath);
                TryDeleteEmptyTempWorkFolder(tempPath);
                string validationReason = "Output validation failed. English AC3 5.1 found=" + outputHasEnglishAc351 + ", English EAC3 kept=" + outputKeptEnglishEac3;
                Write(options.Log, "error", "FAILED: " + validationReason, 1);
                AddLog(summary, "FAILED", file, validationReason);
                return;
            }

            string finalOutputPath;
            string backupPath;
            string commitError;
            if (!CommitOutput(file, tempPath, safeOutputPath, options, out finalOutputPath, out backupPath, out commitError))
            {
                summary.Failed++;
                TryDelete(tempPath);
                TryDeleteEmptyTempWorkFolder(tempPath);
                Write(options.Log, "error", "FAILED: " + commitError, 1);
                AddLog(summary, "FAILED", file, commitError);
                return;
            }

            TryDeleteEmptyTempWorkFolder(tempPath);

            summary.Converted++;
            string convertedMessage = options.ReplaceOriginal
                ? "Converted and replaced original. Backup: " + backupPath
                : "Converted safe output: " + finalOutputPath;

            Write(options.Log, "success", "CONVERTED: " + convertedMessage, 1);
            AddLog(summary, "CONVERTED", file, convertedMessage);
        }

        private static FileDecision InspectFile(string file, string ffprobe)
        {
            List<AudioStreamInfo> audioStreams = ProbeAudioStreams(file, ffprobe);

            AudioStreamInfo existingEnglishAc351 = audioStreams.FirstOrDefault(IsEnglishAc351);
            if (existingEnglishAc351 != null)
            {
                return new FileDecision
                {
                    ShouldConvert = false,
                    Reason = "Already has English AC3 5.1 (" + existingEnglishAc351.Describe() + ")",
                    AudioStreams = audioStreams
                };
            }

            AudioStreamInfo sourceEac3 = audioStreams
                .Where(IsEnglishEac351OrBetter)
                .OrderBy(s => s.AudioIndex)
                .FirstOrDefault();

            if (sourceEac3 == null)
            {
                return new FileDecision
                {
                    ShouldConvert = false,
                    Reason = "No English EAC3 5.1 audio stream found.",
                    AudioStreams = audioStreams
                };
            }

            return new FileDecision
            {
                ShouldConvert = true,
                Reason = "English EAC3 5.1 found and no English AC3 5.1 exists.",
                SourceEac3Stream = sourceEac3,
                AudioStreams = audioStreams
            };
        }

        private static List<AudioStreamInfo> ProbeAudioStreams(string file, string ffprobe)
        {
            string args =
                "-v error -select_streams a " +
                "-show_entries stream=index,codec_type,codec_name,channels,channel_layout:stream_tags=language,title " +
                "-of json " +
                QuoteArg(file);

            ProcessResult result = RunProcess(ffprobe, args);
            if (result.ExitCode != 0)
                throw new Exception("ffprobe failed with exit code " + result.ExitCode + ": " + FirstUsefulLine(result.StandardError, result.StandardOutput));

            JObject root = JObject.Parse(result.StandardOutput);
            JArray streams = root["streams"] as JArray;
            List<AudioStreamInfo> audioStreams = new List<AudioStreamInfo>();

            if (streams == null)
                return audioStreams;

            int audioIndex = 0;
            foreach (JToken token in streams)
            {
                string codecType = ReadString(token, "codec_type");
                if (!string.IsNullOrWhiteSpace(codecType) &&
                    !codecType.Equals("audio", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                JObject tags = token["tags"] as JObject;
                audioStreams.Add(new AudioStreamInfo
                {
                    StreamIndex = ReadInt(token, "index"),
                    AudioIndex = audioIndex,
                    CodecName = ReadString(token, "codec_name"),
                    Channels = ReadInt(token, "channels"),
                    ChannelLayout = ReadString(token, "channel_layout"),
                    Language = tags == null ? "" : ReadString(tags, "language"),
                    Title = tags == null ? "" : ReadString(tags, "title")
                });

                audioIndex++;
            }

            return audioStreams;
        }

        private static string BuildFfmpegArguments(
            string inputPath,
            string outputPath,
            int eac3InputAudioIndex,
            int newAudioOutputIndex,
            int ac3BitrateKbps)
        {
            List<string> args = new List<string>();
            args.Add("-hide_banner");
            args.Add("-nostdin");
            args.Add("-y");
            args.Add("-v");
            args.Add("warning");
            args.Add("-i");
            args.Add(QuoteArg(inputPath));

            // Map everything first, then append the selected EAC3 audio stream once more.
            // Because all original audio streams are copied first, the appended track's
            // output audio index is the original audio stream count.
            args.Add("-map");
            args.Add("0");
            args.Add("-map");
            args.Add("0:a:" + eac3InputAudioIndex);
            args.Add("-map_metadata");
            args.Add("0");
            args.Add("-map_chapters");
            args.Add("0");

            args.Add("-c");
            args.Add("copy");
            args.Add("-c:a:" + newAudioOutputIndex);
            args.Add("ac3");
            args.Add("-b:a:" + newAudioOutputIndex);
            args.Add(ac3BitrateKbps + "k");
            args.Add("-ac:a:" + newAudioOutputIndex);
            args.Add("6");
            args.Add("-metadata:s:a:" + newAudioOutputIndex);
            args.Add("language=eng");
            args.Add("-metadata:s:a:" + newAudioOutputIndex);
            args.Add(QuoteArg("title=AC3 5.1 Compatibility"));
            args.Add("-disposition:a:" + newAudioOutputIndex);
            args.Add("0");
            args.Add(QuoteArg(outputPath));

            return string.Join(" ", args);
        }

        private static bool CommitOutput(
            string originalPath,
            string tempPath,
            string safeOutputPath,
            Options options,
            out string finalOutputPath,
            out string backupPath,
            out string error)
        {
            finalOutputPath = "";
            backupPath = "";
            error = "";

            try
            {
                if (!options.ReplaceOriginal)
                {
                    File.Move(tempPath, safeOutputPath);
                    finalOutputPath = safeOutputPath;
                    return true;
                }

                if (options.BackupOriginal)
                {
                    backupPath = BuildBackupPath(originalPath);
                    File.Move(originalPath, backupPath);
                }
                else
                {
                    File.Delete(originalPath);
                }

                try
                {
                    File.Move(tempPath, originalPath);
                    finalOutputPath = originalPath;
                    return true;
                }
                catch
                {
                    if (!string.IsNullOrWhiteSpace(backupPath) &&
                        File.Exists(backupPath) &&
                        !File.Exists(originalPath))
                    {
                        File.Move(backupPath, originalPath);
                    }

                    throw;
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static List<string> EnumerateVideoFiles(string root, bool recursive, Action<string, string, int> log, Summary summary)
        {
            List<string> files = new List<string>();

            if (File.Exists(root))
            {
                if (IsSupportedVideoFile(root) && !IsGeneratedOutputOrBackup(root))
                    files.Add(root);

                return files;
            }

            if (!Directory.Exists(root))
            {
                Write(log, "warning", "Scan root does not exist: " + root, 1);
                AddLog(summary, "SKIPPED", root, "Scan root does not exist.");
                return files;
            }

            Stack<string> pending = new Stack<string>();
            pending.Push(root);

            while (pending.Count > 0)
            {
                string directory = pending.Pop();
                string[] directoryFiles;

                try
                {
                    directoryFiles = Directory.GetFiles(directory);
                }
                catch (Exception ex)
                {
                    Write(log, "warning", "Could not read files in: " + directory + " | " + ex.Message, 1);
                    AddLog(summary, "SKIPPED", directory, "Could not read files: " + ex.Message);
                    continue;
                }

                foreach (string file in directoryFiles)
                {
                    if (IsSupportedVideoFile(file) && !IsGeneratedOutputOrBackup(file))
                        files.Add(file);
                }

                if (!recursive)
                    continue;

                string[] subdirectories;
                try
                {
                    subdirectories = Directory.GetDirectories(directory);
                }
                catch (Exception ex)
                {
                    Write(log, "warning", "Could not read folders in: " + directory + " | " + ex.Message, 1);
                    AddLog(summary, "SKIPPED", directory, "Could not read folders: " + ex.Message);
                    continue;
                }

                foreach (string subdirectory in subdirectories)
                {
                    if (!IsReparsePoint(subdirectory))
                        pending.Push(subdirectory);
                }
            }

            return files;
        }

        private static bool IsSupportedVideoFile(string path)
        {
            string extension = Path.GetExtension(path);
            return SupportedVideoExtensions.Any(e => e.Equals(extension, StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsGeneratedOutputOrBackup(string path)
        {
            string fileName = Path.GetFileName(path) ?? "";
            return fileName.IndexOf(".bfac3compat.", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   fileName.IndexOf(".ac3compat", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   fileName.IndexOf(".backup-", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsReparsePoint(string path)
        {
            try
            {
                return (new DirectoryInfo(path).Attributes & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsEnglishAc351(AudioStreamInfo stream)
        {
            return IsAc3(stream) && IsEnglish(stream) && Is51OrBetter(stream);
        }

        private static bool IsEnglishEac3(AudioStreamInfo stream)
        {
            return IsEac3(stream) && IsEnglish(stream);
        }

        private static bool IsEnglishEac351OrBetter(AudioStreamInfo stream)
        {
            return IsEnglishEac3(stream) && Is51OrBetter(stream);
        }

        private static bool IsAc3(AudioStreamInfo stream)
        {
            return NormalizeCodec(stream.CodecName) == "ac3";
        }

        private static bool IsEac3(AudioStreamInfo stream)
        {
            return NormalizeCodec(stream.CodecName) == "eac3";
        }

        private static string NormalizeCodec(string codec)
        {
            if (string.IsNullOrWhiteSpace(codec))
                return "";

            return new string(codec.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
        }

        private static bool IsEnglish(AudioStreamInfo stream)
        {
            string language = NormalizeLanguage(stream.Language);
            if (language == "eng" || language == "en" || language == "english")
                return true;

            string title = (stream.Title ?? "").Trim();
            return title.IndexOf("english", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   title.Equals("eng", StringComparison.OrdinalIgnoreCase) ||
                   title.StartsWith("eng ", StringComparison.OrdinalIgnoreCase) ||
                   title.EndsWith(" eng", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeLanguage(string language)
        {
            if (string.IsNullOrWhiteSpace(language))
                return "";

            return new string(language.Trim().Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
        }

        private static bool Is51OrBetter(AudioStreamInfo stream)
        {
            if (stream.Channels >= 6)
                return true;

            string layout = stream.ChannelLayout ?? "";
            return layout.IndexOf("5.1", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   layout.IndexOf("6 channels", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string BuildSafeOutputPath(string originalPath)
        {
            string directory = Path.GetDirectoryName(originalPath) ?? "";
            string fileName = Path.GetFileNameWithoutExtension(originalPath);
            string extension = Path.GetExtension(originalPath);
            return Path.Combine(directory, fileName + ".ac3compat" + extension);
        }

        private static string BuildTempOutputPath(string originalPath)
        {
            string directory = BuildTempWorkDirectory(originalPath);
            string fileName = Path.GetFileNameWithoutExtension(originalPath);
            string extension = Path.GetExtension(originalPath);

            Directory.CreateDirectory(directory);

            string tempPath;
            do
            {
                tempPath = Path.Combine(directory, fileName + ".bfac3compat." + Guid.NewGuid().ToString("N") + extension);
            }
            while (File.Exists(tempPath));

            return tempPath;
        }

        private static string BuildTempWorkDirectory(string originalPath)
        {
            string fullPath = originalPath;
            try
            {
                fullPath = Path.GetFullPath(originalPath);
            }
            catch
            {
                // Use the provided path below if it cannot be normalized.
            }

            string root = Path.GetPathRoot(fullPath);
            if (string.IsNullOrWhiteSpace(root))
                return Path.GetDirectoryName(originalPath) ?? "";

            return Path.Combine(root, TempWorkFolderName);
        }

        private static string BuildBackupPath(string originalPath)
        {
            string directory = Path.GetDirectoryName(originalPath) ?? "";
            string fileName = Path.GetFileNameWithoutExtension(originalPath);
            string extension = Path.GetExtension(originalPath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string backupPath = Path.Combine(directory, fileName + ".backup-" + timestamp + extension);

            int counter = 1;
            while (File.Exists(backupPath))
            {
                backupPath = Path.Combine(directory, fileName + ".backup-" + timestamp + "-" + counter + extension);
                counter++;
            }

            return backupPath;
        }

        private static InteractiveMode PromptMode(Action<string, string, int> log)
        {
            Write(log, "question", "Choose mode:", 1);
            Write(log, "default", "0 - Cancel", 1);
            Write(log, "default", "1 - Dry-run one folder", 1);
            Write(log, "default", "2 - Convert one folder, safe output only", 1);
            Write(log, "default", "3 - Convert one folder, replace originals with backups", 1);
            Write(log, "default", "4 - Dry-run entire library", 1);
            Write(log, "default", "5 - Convert entire library, safe output only", 1);
            Write(log, "default", "6 - Convert entire library, replace originals with backups", 1);
            Console.Write("> ");

            string choice = (Console.ReadLine() ?? "").Trim();
            switch (choice)
            {
                case "1": return InteractiveMode.DryRunFolder;
                case "2": return InteractiveMode.SafeFolder;
                case "3": return InteractiveMode.ReplaceFolderWithBackup;
                case "4": return InteractiveMode.DryRunLibrary;
                case "5": return InteractiveMode.SafeLibrary;
                case "6": return InteractiveMode.ReplaceLibraryWithBackup;
                default: return InteractiveMode.Cancel;
            }
        }

        private static string PromptForFolder(Action<string, string, int> log, string message)
        {
            while (true)
            {
                Write(log, "question", message + " (0 to cancel)", 1);
                Console.Write("> ");
                string input = NormalizeUserPath(Console.ReadLine());

                if (input == "0")
                {
                    Write(log, "warning", "Cancelled.", 1);
                    return "";
                }

                if (Directory.Exists(input))
                    return input;

                Write(log, "error", "That folder does not exist: " + input, 1);
            }
        }

        private static int PromptBitrate(Action<string, string, int> log)
        {
            Write(log, "question", "AC3 bitrate in kbps (Enter for 640)", 1);
            Console.Write("> ");
            string input = (Console.ReadLine() ?? "").Trim();
            if (string.IsNullOrWhiteSpace(input))
                return 640;

            int bitrate;
            if (int.TryParse(input.TrimEnd('k', 'K'), out bitrate) && bitrate > 0)
                return bitrate;

            Write(log, "warning", "Invalid bitrate. Using 640k.", 1);
            return 640;
        }

        private static List<string> GetExistingDefaultRoots(IEnumerable<string> defaultLibraryRoots, Action<string, string, int> log)
        {
            List<string> roots = NormalizeRootPaths(defaultLibraryRoots);
            List<string> existing = new List<string>();

            foreach (string root in roots)
            {
                if (Directory.Exists(root))
                {
                    existing.Add(root);
                }
                else
                {
                    Write(log, "warning", "Default library root not found, skipping: " + root, 1);
                }
            }

            if (existing.Count > 0)
            {
                Write(log, "info", "Library roots selected:", 1);
                foreach (string root in existing)
                    Write(log, "data", root, 1);
            }

            return existing;
        }

        private static List<string> NormalizeRootPaths(IEnumerable<string> rootPaths)
        {
            if (rootPaths == null)
                return new List<string>();

            return rootPaths
                .Select(NormalizeUserPath)
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string NormalizeUserPath(string rawInput)
        {
            if (string.IsNullOrWhiteSpace(rawInput))
                return "";

            string s = rawInput.Trim();
            bool changed = true;

            while (changed && s.Length > 1)
            {
                changed = false;

                if ((s.StartsWith("\"") && s.EndsWith("\"")) ||
                    (s.StartsWith("'") && s.EndsWith("'")))
                {
                    s = s.Substring(1, s.Length - 2).Trim();
                    changed = true;
                }
            }

            return s;
        }

        private static string ResolveExecutable(string explicitPath, string toolName)
        {
            if (!string.IsNullOrWhiteSpace(explicitPath) && File.Exists(explicitPath))
                return explicitPath;

            string exeName = toolName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                ? toolName
                : toolName + ".exe";

            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string currentDirectory = Environment.CurrentDirectory;

            string[] candidates =
            {
                Path.Combine(baseDirectory, exeName),
                Path.Combine(currentDirectory, exeName),
                toolName
            };

            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                    return candidate;
            }

            try
            {
                ProcessResult whereResult = RunProcess("where", toolName);
                if (whereResult.ExitCode == 0)
                {
                    string firstPath = whereResult.StandardOutput
                        .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(l => l.Trim())
                        .FirstOrDefault(File.Exists);

                    if (!string.IsNullOrWhiteSpace(firstPath))
                        return firstPath;
                }
            }
            catch
            {
                // Fall through and let process startup report a useful error later.
            }

            return toolName;
        }

        private static bool VerifyTool(string toolPath, string toolName, Action<string, string, int> log, Summary summary)
        {
            try
            {
                ProcessResult result = RunProcess(toolPath, "-version");
                if (result.ExitCode == 0)
                {
                    AddLog(summary, "TOOL", toolName, "Using " + toolPath);
                    return true;
                }

                string message = toolName + " failed version check: " + FirstUsefulLine(result.StandardError, result.StandardOutput);
                Write(log, "error", message, 1);
                AddLog(summary, "FAILED", toolName, message);
                return false;
            }
            catch (Exception ex)
            {
                string message = "Could not start " + toolName + ". Install FFmpeg or add it to PATH. " + ex.Message;
                Write(log, "error", message, 1);
                AddLog(summary, "FAILED", toolName, message);
                return false;
            }
        }

        private static ProcessResult RunProcess(string fileName, string arguments)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            StringBuilder stdout = new StringBuilder();
            StringBuilder stderr = new StringBuilder();

            using (Process process = new Process())
            {
                process.StartInfo = psi;
                process.OutputDataReceived += delegate (object sender, DataReceivedEventArgs e)
                {
                    if (e.Data != null)
                        stdout.AppendLine(e.Data);
                };
                process.ErrorDataReceived += delegate (object sender, DataReceivedEventArgs e)
                {
                    if (e.Data != null)
                        stderr.AppendLine(e.Data);
                };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                process.WaitForExit();
                process.WaitForExit();

                return new ProcessResult
                {
                    ExitCode = process.ExitCode,
                    StandardOutput = stdout.ToString(),
                    StandardError = stderr.ToString()
                };
            }
        }

        private static string QuoteArg(string value)
        {
            if (value == null)
                return "\"\"";

            return "\"" + value.Replace("\"", "\\\"") + "\"";
        }

        private static string ReadString(JToken token, string name)
        {
            JToken value = token == null ? null : token[name];
            return value == null ? "" : value.ToString();
        }

        private static int ReadInt(JToken token, string name)
        {
            int value;
            return int.TryParse(ReadString(token, name), out value) ? value : 0;
        }

        private static string FirstUsefulLine(string primary, string secondary)
        {
            string line = (primary ?? "")
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .FirstOrDefault(l => !string.IsNullOrWhiteSpace(l));

            if (!string.IsNullOrWhiteSpace(line))
                return line;

            line = (secondary ?? "")
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .FirstOrDefault(l => !string.IsNullOrWhiteSpace(l));

            return string.IsNullOrWhiteSpace(line) ? "(no output)" : line;
        }

        private static string LastLines(string text, int count)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            return string.Join(
                Environment.NewLine,
                text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Skip(Math.Max(0, text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length - count)));
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
                // Best effort cleanup only.
            }
        }

        private static void TryDeleteEmptyTempWorkFolder(string tempPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(tempPath))
                    return;

                string directory = Path.GetDirectoryName(tempPath);
                if (string.IsNullOrWhiteSpace(directory) ||
                    !string.Equals(Path.GetFileName(directory), TempWorkFolderName, StringComparison.OrdinalIgnoreCase) ||
                    !Directory.Exists(directory) ||
                    Directory.EnumerateFileSystemEntries(directory).Any())
                {
                    return;
                }

                Directory.Delete(directory, recursive: false);
            }
            catch
            {
                // Best effort cleanup only.
            }
        }

        private static void WriteSummary(Summary summary, Action<string, string, int> log)
        {
            Write(log, "info", "=== AUDIO COMPATIBILITY SUMMARY ===", 1);
            Write(log, "data", "Found: " + summary.FilesFound, 1);
            Write(log, "data", "Inspected: " + summary.FilesInspected, 1);
            Write(log, "success", "Converted: " + summary.Converted, 1);
            Write(log, "success", "Dry-run matches: " + summary.DryRunMatches, 1);
            Write(log, "warning", "Skipped: " + summary.Skipped, 1);
            Write(log, "error", "Failed: " + summary.Failed, 2);

            AddLog(summary, "SUMMARY", "", "Found=" + summary.FilesFound +
                                      ", Inspected=" + summary.FilesInspected +
                                      ", Converted=" + summary.Converted +
                                      ", DryRunMatches=" + summary.DryRunMatches +
                                      ", Skipped=" + summary.Skipped +
                                      ", Failed=" + summary.Failed);
        }

        private static void WriteLogFile(Summary summary, Action<string, string, int> log)
        {
            string error;
            string logPath = global::BachFlixLog.WriteBachFlixLog(
                summary.LogLines,
                "Audio Compatibility",
                "AudioCompatibility",
                out error);

            if (!string.IsNullOrWhiteSpace(logPath))
                Write(log, "success", "Audio compatibility log written: " + logPath, 2);
            else if (!string.IsNullOrWhiteSpace(error))
                Write(log, "warning", "Could not write audio compatibility log: " + error, 2);
        }

        private static void AddLog(Summary summary, string status, string path, string message)
        {
            if (summary == null)
                return;

            summary.LogLines.Add(
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                " [" + status + "] " +
                (string.IsNullOrWhiteSpace(path) ? "" : path + " | ") +
                (message ?? ""));
        }

        private static void Write(Action<string, string, int> log, string type, string message, int indent)
        {
            if (log != null)
            {
                log(type, message, indent);
                return;
            }

            Console.WriteLine(message);
        }
    }
}
