using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace BachFlixNfo.Features
{
    public interface IWordTranscriptionService
    {
        ExternalToolResolution ResolveTranscriptionTool(
            string explicitPath,
            Action<string, string, int> log,
            IList<string> commandLog);

        WordTranscriptionResult Transcribe(WordTranscriptionRequest request);
    }

    public sealed class WhisperXTranscriptionService : IWordTranscriptionService
    {
        private const string WorkDirectoryPrefix = "BachFlixNfo-AudioCensor-";

        public ExternalToolResolution ResolveTranscriptionTool(
            string explicitPath,
            Action<string, string, int> log,
            IList<string> commandLog)
        {
            string path = explicitPath;
            if (string.IsNullOrWhiteSpace(path))
                path = Environment.GetEnvironmentVariable("AUDIO_CENSOR_WHISPERX_PATH");

            if (string.IsNullOrWhiteSpace(path))
                path = Environment.GetEnvironmentVariable("WHISPERX_PATH");

            return ExternalToolResolver.Resolve(path, "whisperx", log, commandLog);
        }

        public WordTranscriptionResult Transcribe(WordTranscriptionRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            if (string.IsNullOrWhiteSpace(request.MediaPath) || !File.Exists(request.MediaPath))
                throw new FileNotFoundException("Input media file was not found.", request.MediaPath ?? "");

            if (string.IsNullOrWhiteSpace(request.FfmpegPath))
                throw new ArgumentException("FFmpeg path is required for WhisperX audio extraction.", nameof(request));

            if (string.IsNullOrWhiteSpace(request.TranscriptionToolPath))
                throw new ArgumentException("WhisperX path is required for transcription.", nameof(request));

            bool createdWorkDirectory = string.IsNullOrWhiteSpace(request.WorkDirectory);
            string workDirectory = createdWorkDirectory
                ? Path.Combine(Path.GetTempPath(), WorkDirectoryPrefix + Guid.NewGuid().ToString("N"))
                : request.WorkDirectory;

            Directory.CreateDirectory(workDirectory);

            string mediaBaseName = Path.GetFileNameWithoutExtension(request.MediaPath);
            if (string.IsNullOrWhiteSpace(mediaBaseName))
                mediaBaseName = "audio";

            int audioTrackIndex = request.AudioTrackIndex < 0 ? 0 : request.AudioTrackIndex;
            string audioPath = Path.Combine(workDirectory, mediaBaseName + ".audio-" + audioTrackIndex + ".wav");
            string outputDirectory = Path.Combine(workDirectory, "whisperx-" + audioTrackIndex);
            Directory.CreateDirectory(outputDirectory);

            var result = new WordTranscriptionResult
            {
                EngineName = "WhisperX",
                SourceAudioTrackIndex = audioTrackIndex,
                WorkDirectory = workDirectory
            };

            try
            {
                ExtractAudioTrack(request, audioPath, audioTrackIndex);
                string jsonPath = RunWhisperX(request, audioPath, outputDirectory);

                string finalJsonPath = GetOutputPath(
                    request.OutputJsonPath,
                    request.MediaPath,
                    audioTrackIndex,
                    ".whisperx.json",
                    request.UseCanonicalOutputNames);

                string finalSrtPath = GetOutputPath(
                    request.OutputSrtPath,
                    request.MediaPath,
                    audioTrackIndex,
                    ".whisperx.srt",
                    request.UseCanonicalOutputNames);

                CopyFile(jsonPath, finalJsonPath);
                string whisperSrtPath = LocateWhisperOutput(outputDirectory, audioPath, ".srt");
                if (!string.IsNullOrWhiteSpace(whisperSrtPath) && File.Exists(whisperSrtPath))
                    CopyFile(whisperSrtPath, finalSrtPath);
                else
                    WriteSrtFromJsonSegments(finalJsonPath, finalSrtPath);

                result.TranscriptJsonPath = finalJsonPath;
                result.TranscriptSrtPath = finalSrtPath;
                result.Words.AddRange(LoadTranscriptWords(finalJsonPath));
                result.Segments.AddRange(LoadTranscriptSegments(finalJsonPath));
                result.TranscriptWordCount = result.Words.Count;

                return result;
            }
            finally
            {
                if (!request.KeepWorkFiles && createdWorkDirectory)
                    TryDeleteDirectory(workDirectory, result);
            }
        }

        private static void ExtractAudioTrack(WordTranscriptionRequest request, string audioPath, int audioTrackIndex)
        {
            string args =
                "-hide_banner -nostdin -y -v error " +
                "-i " + Quote(request.MediaPath) + " " +
                "-map 0:a:" + audioTrackIndex + " " +
                "-vn -sn -dn -ac 1 -ar 16000 -f wav " +
                Quote(audioPath);

            ExternalCommandResult result = ExternalToolResolver.RunProcess(request.FfmpegPath, args, request.Log, request.CommandLog);
            if (result.ExitCode != 0)
                throw new InvalidOperationException("FFmpeg audio extraction failed with exit code " + result.ExitCode + ": " + FirstUsefulLine(result.StandardError, result.StandardOutput));
        }

        private static string RunWhisperX(WordTranscriptionRequest request, string audioPath, string outputDirectory)
        {
            var args = new List<string>();
            args.Add(Quote(audioPath));
            AddOption(args, "--model", request.Model);
            AddOption(args, "--device", request.Device);
            AddOption(args, "--compute_type", request.ComputeType);
            AddOption(args, "--language", request.Language);
            args.Add("--output_dir");
            args.Add(Quote(outputDirectory));
            args.Add("--output_format");
            args.Add("all");

            ExternalCommandResult result = ExternalToolResolver.RunProcess(request.TranscriptionToolPath, string.Join(" ", args), request.Log, request.CommandLog);
            if (result.ExitCode != 0)
                throw new InvalidOperationException("WhisperX transcription failed with exit code " + result.ExitCode + ": " + FirstUsefulLine(result.StandardError, result.StandardOutput));

            string jsonPath = LocateWhisperOutput(outputDirectory, audioPath, ".json");
            if (!string.IsNullOrWhiteSpace(jsonPath) && File.Exists(jsonPath))
                return jsonPath;

            throw new FileNotFoundException("WhisperX completed, but no JSON transcript was written.", Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(audioPath) + ".json"));
        }

        private static string LocateWhisperOutput(string outputDirectory, string audioPath, string extension)
        {
            string expected = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(audioPath) + extension);
            if (File.Exists(expected))
                return expected;

            return Directory.GetFiles(outputDirectory, "*" + extension, SearchOption.TopDirectoryOnly)
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault();
        }

        private static string GetOutputPath(string explicitPath, string mediaPath, int audioTrackIndex, string suffix, bool canonical)
        {
            if (!string.IsNullOrWhiteSpace(explicitPath))
                return explicitPath;

            string directory = Path.GetDirectoryName(mediaPath) ?? "";
            string baseName = Path.GetFileNameWithoutExtension(mediaPath);
            string trackSuffix = canonical ? "" : ".audio-" + audioTrackIndex;
            return Path.Combine(directory, baseName + trackSuffix + suffix);
        }

        private static void CopyFile(string source, string destination)
        {
            string directory = Path.GetDirectoryName(destination);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            File.Copy(source, destination, true);
        }

        public static List<TimedTranscriptWord> LoadTranscriptWords(string jsonPath)
        {
            JObject root = JObject.Parse(File.ReadAllText(jsonPath));
            var words = new List<TimedTranscriptWord>();

            JArray wordSegments = root["word_segments"] as JArray;
            if (wordSegments != null)
                AddWords(words, wordSegments);

            JArray segments = root["segments"] as JArray;
            if (segments != null)
            {
                foreach (JToken segment in segments)
                {
                    JArray segmentWords = segment["words"] as JArray;
                    if (segmentWords != null)
                        AddWords(words, segmentWords);
                }
            }

            return words
                .Where(w => w.End > w.Start && !string.IsNullOrWhiteSpace(w.NormalizedWord))
                .OrderBy(w => w.Start)
                .ThenBy(w => w.End)
                .ToList();
        }

        public static List<TranscriptSegment> LoadTranscriptSegments(string jsonPath)
        {
            JObject root = JObject.Parse(File.ReadAllText(jsonPath));
            var segments = new List<TranscriptSegment>();
            JArray jsonSegments = root["segments"] as JArray;
            if (jsonSegments == null)
                return segments;

            foreach (JToken segment in jsonSegments)
            {
                TimeSpan? start = ReadTimeSpanSeconds(segment, "start");
                TimeSpan? end = ReadTimeSpanSeconds(segment, "end");
                if (!start.HasValue || !end.HasValue || end.Value <= start.Value)
                    continue;

                segments.Add(new TranscriptSegment
                {
                    Start = start.Value,
                    End = end.Value,
                    Text = ReadString(segment, "text").Trim()
                });
            }

            return segments;
        }

        private static void AddWords(List<TimedTranscriptWord> words, JArray tokens)
        {
            foreach (JToken token in tokens)
            {
                string word = ReadString(token, "word").Trim();
                TimeSpan? start = ReadTimeSpanSeconds(token, "start");
                TimeSpan? end = ReadTimeSpanSeconds(token, "end");

                if (!start.HasValue || !end.HasValue)
                    continue;

                words.Add(new TimedTranscriptWord
                {
                    Word = word,
                    NormalizedWord = ProfanityDictionary.NormalizeToken(word),
                    Start = start.Value,
                    End = end.Value
                });
            }
        }

        private static void WriteSrtFromJsonSegments(string jsonPath, string srtPath)
        {
            List<TranscriptSegment> segments = LoadTranscriptSegments(jsonPath);
            var lines = new List<string>();

            int sequence = 1;
            foreach (TranscriptSegment segment in segments)
            {
                if (string.IsNullOrWhiteSpace(segment.Text))
                    continue;

                lines.Add(sequence.ToString(CultureInfo.InvariantCulture));
                lines.Add(FormatSrtTime(segment.Start) + " --> " + FormatSrtTime(segment.End));
                lines.Add(segment.Text.Trim());
                lines.Add("");
                sequence++;
            }

            File.WriteAllLines(srtPath, lines);
        }

        private static TimeSpan? ReadTimeSpanSeconds(JToken token, string name)
        {
            double seconds;
            if (!double.TryParse(ReadString(token, name), NumberStyles.Float, CultureInfo.InvariantCulture, out seconds))
                return null;

            if (seconds < 0)
                return null;

            return TimeSpan.FromSeconds(seconds);
        }

        private static string ReadString(JToken token, string name)
        {
            JToken value = token == null ? null : token[name];
            return value == null ? "" : value.ToString();
        }

        private static void AddOption(List<string> args, string name, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return;

            args.Add(name);
            args.Add(value.IndexOfAny(new[] { ' ', '\t', '"' }) >= 0 ? Quote(value) : value);
        }

        private static string Quote(string value)
        {
            return "\"" + (value ?? "").Replace("\"", "\\\"") + "\"";
        }

        private static string FirstUsefulLine(params string[] values)
        {
            foreach (string value in values)
            {
                if (string.IsNullOrWhiteSpace(value))
                    continue;

                string line = value
                    .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(l => l.Trim())
                    .FirstOrDefault(l => !string.IsNullOrWhiteSpace(l));

                if (!string.IsNullOrWhiteSpace(line))
                    return line;
            }

            return "(no output)";
        }

        private static string FormatSrtTime(TimeSpan value)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0:D2}:{1:D2}:{2:D2},{3:D3}",
                (int)value.TotalHours,
                value.Minutes,
                value.Seconds,
                value.Milliseconds);
        }

        private static void TryDeleteDirectory(string path, WordTranscriptionResult result)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
                    return;

                string directoryName = Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                if (string.IsNullOrWhiteSpace(directoryName) || !directoryName.StartsWith(WorkDirectoryPrefix, StringComparison.OrdinalIgnoreCase))
                    return;

                Directory.Delete(path, recursive: true);
                if (result != null)
                    result.TemporaryFilesCleaned = true;
            }
            catch
            {
                if (result != null)
                    result.TemporaryFilesCleaned = false;
            }
        }
    }

    public sealed class WordTranscriptionRequest
    {
        public string MediaPath { get; set; }
        public int AudioTrackIndex { get; set; }
        public string FfmpegPath { get; set; }
        public string TranscriptionToolPath { get; set; }
        public string WorkDirectory { get; set; }
        public string OutputJsonPath { get; set; }
        public string OutputSrtPath { get; set; }
        public bool UseCanonicalOutputNames { get; set; }
        public string Model { get; set; }
        public string Device { get; set; }
        public string ComputeType { get; set; }
        public string Language { get; set; }
        public bool KeepWorkFiles { get; set; }
        public Action<string, string, int> Log { get; set; }
        public IList<string> CommandLog { get; set; }

        public WordTranscriptionRequest()
        {
            MediaPath = "";
            FfmpegPath = "";
            TranscriptionToolPath = "";
            WorkDirectory = "";
            OutputJsonPath = "";
            OutputSrtPath = "";
            Model = "small";
            Device = "cpu";
            ComputeType = "int8";
            Language = "en";
            CommandLog = new List<string>();
        }
    }

    public sealed class WordTranscriptionResult
    {
        public string EngineName { get; set; }
        public int SourceAudioTrackIndex { get; set; }
        public List<TimedTranscriptWord> Words { get; private set; }
        public List<TranscriptSegment> Segments { get; private set; }
        public int TranscriptWordCount { get; set; }
        public string WorkDirectory { get; set; }
        public string TranscriptJsonPath { get; set; }
        public string TranscriptSrtPath { get; set; }
        public bool TemporaryFilesCleaned { get; set; }

        public WordTranscriptionResult()
        {
            EngineName = "";
            Words = new List<TimedTranscriptWord>();
            Segments = new List<TranscriptSegment>();
            WorkDirectory = "";
            TranscriptJsonPath = "";
            TranscriptSrtPath = "";
        }
    }

    public sealed class TimedTranscriptWord
    {
        public string Word { get; set; }
        public string NormalizedWord { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }

        public TimedTranscriptWord()
        {
            Word = "";
            NormalizedWord = "";
        }
    }

    public sealed class TranscriptSegment
    {
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public string Text { get; set; }

        public TranscriptSegment()
        {
            Text = "";
        }
    }
}
