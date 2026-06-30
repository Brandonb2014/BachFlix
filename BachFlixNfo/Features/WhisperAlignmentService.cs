using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BachFlixNfo.Features
{
    public interface IWhisperAlignmentService
    {
        ExternalToolResolution ResolveWhisperX(
            string explicitPath,
            Action<string, string, int> log,
            IList<string> commandLog);

        WhisperAlignmentResult AlignWords(WhisperAlignmentRequest request);
    }

    public sealed class WhisperAlignmentService : IWhisperAlignmentService
    {
        private const string WorkDirectoryPrefix = "BachFlixNfo-AudioCensor-";
        private static readonly TimeSpan PrimaryCueTolerance = TimeSpan.FromSeconds(3);
        private static readonly TimeSpan ExpandedCueTolerance = TimeSpan.FromSeconds(8);

        public ExternalToolResolution ResolveWhisperX(
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

        public WhisperAlignmentResult AlignWords(WhisperAlignmentRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            if (string.IsNullOrWhiteSpace(request.MkvPath) || !File.Exists(request.MkvPath))
                throw new FileNotFoundException("Input MKV was not found.", request.MkvPath ?? "");

            if (string.IsNullOrWhiteSpace(request.FfmpegPath))
                throw new ArgumentException("FFmpeg path is required for WhisperX alignment audio extraction.", nameof(request));

            if (string.IsNullOrWhiteSpace(request.WhisperXPath))
                throw new ArgumentException("WhisperX path is required for alignment.", nameof(request));

            IReadOnlyList<ProfanityOccurrence> occurrences = request.OccurrencesToAlign ?? new List<ProfanityOccurrence>();
            var alignmentResult = new WhisperAlignmentResult();
            if (occurrences.Count == 0)
                return alignmentResult;

            bool createdWorkDirectory = string.IsNullOrWhiteSpace(request.WorkDirectory);
            string workDirectory = createdWorkDirectory
                ? Path.Combine(Path.GetTempPath(), WorkDirectoryPrefix + Guid.NewGuid().ToString("N"))
                : request.WorkDirectory;

            Directory.CreateDirectory(workDirectory);
            alignmentResult.WorkDirectory = workDirectory;

            string safeBaseName = Path.GetFileNameWithoutExtension(request.MkvPath);
            if (string.IsNullOrWhiteSpace(safeBaseName))
                safeBaseName = "audio";

            string audioPath = Path.Combine(workDirectory, safeBaseName + ".audio-" + Math.Max(0, request.AudioTrackIndex) + ".wav");
            string outputDirectory = Path.Combine(workDirectory, "whisperx");
            Directory.CreateDirectory(outputDirectory);

            try
            {
                ExtractAudioTrack(request, audioPath);
                string jsonPath = RunWhisperX(request, audioPath, outputDirectory);
                alignmentResult.TranscriptJsonPath = jsonPath;

                List<TimedTranscriptWord> transcriptWords = LoadTranscriptWords(jsonPath);
                alignmentResult.TranscriptWordCount = transcriptWords.Count;

                List<AlignedProfanityWord> alignedWords = MatchOccurrencesToTranscriptWords(occurrences, transcriptWords, alignmentResult.UnalignedOccurrences);
                alignmentResult.AlignedWords.AddRange(alignedWords);

                return alignmentResult;
            }
            finally
            {
                if (!request.KeepWorkFiles && createdWorkDirectory)
                    TryDeleteDirectory(workDirectory, alignmentResult);
            }
        }

        private static void ExtractAudioTrack(WhisperAlignmentRequest request, string audioPath)
        {
            int audioTrackIndex = request.AudioTrackIndex < 0 ? 0 : request.AudioTrackIndex;
            string args =
                "-hide_banner -nostdin -y -v error " +
                "-i " + Quote(request.MkvPath) + " " +
                "-map 0:a:" + audioTrackIndex + " " +
                "-vn -sn -dn -ac 1 -ar 16000 -f wav " +
                Quote(audioPath);

            ExternalCommandResult result = ExternalToolResolver.RunProcess(request.FfmpegPath, args, request.Log, request.CommandLog);
            if (result.ExitCode != 0)
                throw new InvalidOperationException("FFmpeg audio extraction failed with exit code " + result.ExitCode + ": " + FirstUsefulLine(result.StandardError, result.StandardOutput));
        }

        private static string RunWhisperX(WhisperAlignmentRequest request, string audioPath, string outputDirectory)
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
            args.Add("json");

            ExternalCommandResult result = ExternalToolResolver.RunProcess(request.WhisperXPath, string.Join(" ", args), request.Log, request.CommandLog);
            if (result.ExitCode != 0)
                throw new InvalidOperationException("WhisperX alignment failed with exit code " + result.ExitCode + ": " + FirstUsefulLine(result.StandardError, result.StandardOutput));

            string expectedJson = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(audioPath) + ".json");
            if (File.Exists(expectedJson))
                return expectedJson;

            string latestJson = Directory.GetFiles(outputDirectory, "*.json", SearchOption.TopDirectoryOnly)
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(latestJson) && File.Exists(latestJson))
                return latestJson;

            throw new FileNotFoundException("WhisperX completed, but no JSON transcript was written.", expectedJson);
        }

        private static List<TimedTranscriptWord> LoadTranscriptWords(string jsonPath)
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

        private static List<AlignedProfanityWord> MatchOccurrencesToTranscriptWords(
            IReadOnlyList<ProfanityOccurrence> occurrences,
            List<TimedTranscriptWord> transcriptWords,
            List<ProfanityOccurrence> unalignedOccurrences)
        {
            var aligned = new List<AlignedProfanityWord>();
            var usedTranscriptIndexes = new HashSet<int>();

            foreach (ProfanityOccurrence occurrence in occurrences.OrderBy(o => o.ReviewNumber))
            {
                TimedWordMatch match = FindBestMatch(occurrence, transcriptWords, usedTranscriptIndexes, PrimaryCueTolerance);
                if (match == null)
                    match = FindBestMatch(occurrence, transcriptWords, usedTranscriptIndexes, ExpandedCueTolerance);

                if (match == null)
                {
                    if (unalignedOccurrences != null)
                        unalignedOccurrences.Add(occurrence);
                    continue;
                }

                usedTranscriptIndexes.Add(match.Index);
                aligned.Add(new AlignedProfanityWord
                {
                    SourceOccurrence = occurrence,
                    Word = match.Word.Word,
                    Start = match.Word.Start,
                    End = match.Word.End
                });
            }

            return aligned;
        }

        private static TimedWordMatch FindBestMatch(
            ProfanityOccurrence occurrence,
            List<TimedTranscriptWord> transcriptWords,
            HashSet<int> usedTranscriptIndexes,
            TimeSpan cueTolerance)
        {
            if (occurrence == null || occurrence.Cue == null || transcriptWords == null)
                return null;

            string occurrenceWord = ProfanityDictionary.NormalizeToken(occurrence.Word);
            string dictionaryTerm = ProfanityDictionary.NormalizeToken(occurrence.DictionaryTerm);
            if (string.IsNullOrWhiteSpace(occurrenceWord) && string.IsNullOrWhiteSpace(dictionaryTerm))
                return null;

            TimeSpan earliest = occurrence.Cue.Start - cueTolerance;
            if (earliest < TimeSpan.Zero)
                earliest = TimeSpan.Zero;

            TimeSpan latest = occurrence.Cue.End + cueTolerance;
            TimeSpan expectedTime = EstimateOccurrenceTime(occurrence);
            var matches = new List<TimedWordMatch>();

            for (int i = 0; i < transcriptWords.Count; i++)
            {
                if (usedTranscriptIndexes.Contains(i))
                    continue;

                TimedTranscriptWord word = transcriptWords[i];
                if (!WordMatches(word.NormalizedWord, occurrenceWord, dictionaryTerm))
                    continue;

                if (word.End < earliest || word.Start > latest)
                    continue;

                TimeSpan wordMidpoint = TimeSpan.FromTicks((word.Start.Ticks + word.End.Ticks) / 2);
                matches.Add(new TimedWordMatch
                {
                    Index = i,
                    Word = word,
                    DistanceFromExpectedTime = Math.Abs((wordMidpoint - expectedTime).TotalMilliseconds)
                });
            }

            return matches
                .OrderBy(m => m.DistanceFromExpectedTime)
                .ThenBy(m => m.Word.Start)
                .FirstOrDefault();
        }

        private static TimeSpan EstimateOccurrenceTime(ProfanityOccurrence occurrence)
        {
            TimeSpan cueDuration = occurrence.Cue.End - occurrence.Cue.Start;
            if (cueDuration <= TimeSpan.Zero || string.IsNullOrWhiteSpace(occurrence.Cue.Text))
                return TimeSpan.FromTicks((occurrence.Cue.Start.Ticks + occurrence.Cue.End.Ticks) / 2);

            double fraction = Math.Max(0, Math.Min(1, occurrence.CharacterIndex / (double)Math.Max(1, occurrence.Cue.Text.Length)));
            return occurrence.Cue.Start + TimeSpan.FromTicks((long)(cueDuration.Ticks * fraction));
        }

        private static bool WordMatches(string normalizedTranscriptWord, string occurrenceWord, string dictionaryTerm)
        {
            if (string.IsNullOrWhiteSpace(normalizedTranscriptWord))
                return false;

            return (!string.IsNullOrWhiteSpace(occurrenceWord) && normalizedTranscriptWord == occurrenceWord) ||
                   (!string.IsNullOrWhiteSpace(dictionaryTerm) && normalizedTranscriptWord == dictionaryTerm);
        }

        private static TimeSpan? ReadTimeSpanSeconds(JToken token, string name)
        {
            double seconds;
            if (!double.TryParse(ReadString(token, name), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out seconds))
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

        private static void TryDeleteDirectory(string path, WhisperAlignmentResult result)
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

        private sealed class TimedTranscriptWord
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

        private sealed class TimedWordMatch
        {
            public int Index { get; set; }
            public TimedTranscriptWord Word { get; set; }
            public double DistanceFromExpectedTime { get; set; }
        }
    }

    public sealed class WhisperAlignmentRequest
    {
        public string MkvPath { get; set; }
        public string SubtitlePath { get; set; }
        public int AudioTrackIndex { get; set; }
        public IReadOnlyList<ProfanityOccurrence> OccurrencesToAlign { get; set; }
        public string FfmpegPath { get; set; }
        public string WhisperXPath { get; set; }
        public string WorkDirectory { get; set; }
        public string Model { get; set; }
        public string Device { get; set; }
        public string ComputeType { get; set; }
        public string Language { get; set; }
        public bool KeepWorkFiles { get; set; }
        public Action<string, string, int> Log { get; set; }
        public IList<string> CommandLog { get; set; }

        public WhisperAlignmentRequest()
        {
            MkvPath = "";
            SubtitlePath = "";
            OccurrencesToAlign = new List<ProfanityOccurrence>();
            FfmpegPath = "";
            WhisperXPath = "";
            WorkDirectory = "";
            Model = "small";
            Device = "cpu";
            ComputeType = "int8";
            Language = "en";
            KeepWorkFiles = false;
            CommandLog = new List<string>();
        }
    }

    public sealed class WhisperAlignmentResult
    {
        public List<AlignedProfanityWord> AlignedWords { get; private set; }
        public List<ProfanityOccurrence> UnalignedOccurrences { get; private set; }
        public int TranscriptWordCount { get; set; }
        public string WorkDirectory { get; set; }
        public string TranscriptJsonPath { get; set; }
        public bool TemporaryFilesCleaned { get; set; }

        public WhisperAlignmentResult()
        {
            AlignedWords = new List<AlignedProfanityWord>();
            UnalignedOccurrences = new List<ProfanityOccurrence>();
            WorkDirectory = "";
            TranscriptJsonPath = "";
        }
    }

    public sealed class AlignedProfanityWord
    {
        public ProfanityOccurrence SourceOccurrence { get; set; }
        public string Word { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }

        public AlignedProfanityWord()
        {
            Word = "";
        }
    }
}