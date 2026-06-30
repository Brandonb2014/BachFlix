using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BachFlixNfo.Features
{
    public sealed class FfmpegMuxService
    {
        public ExternalToolResolution ResolveFfmpeg(
            string explicitPath,
            Action<string, string, int> log,
            IList<string> commandLog)
        {
            return ExternalToolResolver.Resolve(explicitPath, "ffmpeg", log, commandLog);
        }

        public ExternalToolResolution ResolveFfprobe(
            string explicitPath,
            Action<string, string, int> log,
            IList<string> commandLog)
        {
            return ExternalToolResolver.Resolve(explicitPath, "ffprobe", log, commandLog);
        }

        public List<AudioTrackInfo> ProbeAudioTracks(
            string ffprobePath,
            string inputMkvPath,
            Action<string, string, int> log,
            IList<string> commandLog)
        {
            if (string.IsNullOrWhiteSpace(ffprobePath))
                throw new ArgumentException("FFprobe path is required.", nameof(ffprobePath));

            if (string.IsNullOrWhiteSpace(inputMkvPath))
                throw new ArgumentException("Input MKV path is required.", nameof(inputMkvPath));

            string arguments =
                "-v error -select_streams a " +
                "-show_entries stream=index,codec_name,channels,channel_layout:stream_tags=language,title " +
                "-of json " +
                Quote(inputMkvPath);

            ExternalCommandResult result = ExternalToolResolver.RunProcess(ffprobePath, arguments, log, commandLog);
            if (result.ExitCode != 0)
                throw new InvalidOperationException("ffprobe failed with exit code " + result.ExitCode + ": " + FirstUsefulLine(result.StandardError, result.StandardOutput));

            JObject root = JObject.Parse(result.StandardOutput);
            JArray streams = root["streams"] as JArray;
            var tracks = new List<AudioTrackInfo>();

            if (streams == null)
                return tracks;

            int audioTrackIndex = 0;
            foreach (JToken stream in streams)
            {
                JObject tags = stream["tags"] as JObject;
                tracks.Add(new AudioTrackInfo
                {
                    AudioTrackIndex = audioTrackIndex,
                    StreamIndex = ReadInt(stream, "index"),
                    CodecName = ReadString(stream, "codec_name"),
                    Channels = ReadInt(stream, "channels"),
                    ChannelLayout = ReadString(stream, "channel_layout"),
                    Language = tags == null ? "" : ReadString(tags, "language"),
                    Title = tags == null ? "" : ReadString(tags, "title")
                });

                audioTrackIndex++;
            }

            return tracks;
        }

        public string BuildCleanOutputPath(string inputMkvPath)
        {
            if (string.IsNullOrWhiteSpace(inputMkvPath))
                throw new ArgumentException("Input MKV path is required.", nameof(inputMkvPath));

            string directory = Path.GetDirectoryName(inputMkvPath) ?? "";
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(inputMkvPath);
            string extension = Path.GetExtension(inputMkvPath);

            if (string.IsNullOrWhiteSpace(extension))
                extension = ".mkv";

            return Path.Combine(directory, fileNameWithoutExtension + "_Clean" + extension);
        }

        public FfmpegMuxPlan BuildMuxPlan(FfmpegMuxRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            if (string.IsNullOrWhiteSpace(request.FfmpegPath))
                throw new ArgumentException("FFmpeg path is required.", nameof(request));

            if (string.IsNullOrWhiteSpace(request.InputMkvPath))
                throw new ArgumentException("Input MKV path is required.", nameof(request));

            string outputPath = string.IsNullOrWhiteSpace(request.OutputMkvPath)
                ? BuildCleanOutputPath(request.InputMkvPath)
                : request.OutputMkvPath;

            if (string.Equals(Path.GetFullPath(request.InputMkvPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Refusing to build an FFmpeg plan that writes over the original file.");

            if (string.IsNullOrWhiteSpace(request.AudioFilter))
                throw new ArgumentException("Audio filter is required.", nameof(request));

            string audioCodec = string.IsNullOrWhiteSpace(request.AudioCodec) ? "aac" : request.AudioCodec;
            int audioTrackIndex = request.AudioTrackIndex < 0 ? 0 : request.AudioTrackIndex;

            var args = new List<string>();
            args.Add("-hide_banner");
            args.Add("-nostdin");
            args.Add("-n");
            args.Add("-v");
            args.Add("warning");
            args.Add("-i");
            args.Add(Quote(request.InputMkvPath));
            args.Add("-map");
            args.Add("0");
            args.Add("-map_metadata");
            args.Add("0");
            args.Add("-map_chapters");
            args.Add("0");
            args.Add("-c");
            args.Add("copy");
            args.Add("-c:v");
            args.Add("copy");
            args.Add("-c:s");
            args.Add("copy");
            args.Add("-c:t");
            args.Add("copy");
            args.Add("-filter:a:" + audioTrackIndex);
            args.Add(Quote(request.AudioFilter));
            args.Add("-c:a:" + audioTrackIndex);
            args.Add(audioCodec);
            args.Add(Quote(outputPath));

            return new FfmpegMuxPlan
            {
                FfmpegPath = request.FfmpegPath,
                Arguments = string.Join(" ", args),
                InputMkvPath = request.InputMkvPath,
                OutputMkvPath = outputPath,
                AudioTrackIndex = audioTrackIndex
            };
        }

        public void MuxCensoredAudio(FfmpegMuxPlan plan)
        {
            throw new NotSupportedException("FFmpeg mux execution is intentionally not implemented in this phase.");
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
    }

    public sealed class FfmpegMuxRequest
    {
        public string FfmpegPath { get; set; }
        public string InputMkvPath { get; set; }
        public string OutputMkvPath { get; set; }
        public int AudioTrackIndex { get; set; }
        public string AudioFilter { get; set; }
        public string AudioCodec { get; set; }

        public FfmpegMuxRequest()
        {
            FfmpegPath = "";
            InputMkvPath = "";
            OutputMkvPath = "";
            AudioFilter = "";
            AudioCodec = "aac";
        }
    }

    public sealed class FfmpegMuxPlan
    {
        public string FfmpegPath { get; set; }
        public string Arguments { get; set; }
        public string InputMkvPath { get; set; }
        public string OutputMkvPath { get; set; }
        public int AudioTrackIndex { get; set; }

        public FfmpegMuxPlan()
        {
            FfmpegPath = "";
            Arguments = "";
            InputMkvPath = "";
            OutputMkvPath = "";
        }

        public string CommandLine
        {
            get { return "\"" + FfmpegPath + "\" " + Arguments; }
        }
    }

    public sealed class AudioTrackInfo
    {
        public int AudioTrackIndex { get; set; }
        public int StreamIndex { get; set; }
        public string CodecName { get; set; }
        public string Language { get; set; }
        public string Title { get; set; }
        public int Channels { get; set; }
        public string ChannelLayout { get; set; }

        public AudioTrackInfo()
        {
            CodecName = "";
            Language = "";
            Title = "";
            ChannelLayout = "";
        }

        public string Describe()
        {
            string language = string.IsNullOrWhiteSpace(Language) ? "(none)" : Language;
            string title = string.IsNullOrWhiteSpace(Title) ? "" : ", title=" + Title;
            string layout = string.IsNullOrWhiteSpace(ChannelLayout) ? "" : ", layout=" + ChannelLayout;

            return string.Format(
                "audioTrack={0}, stream={1}, codec={2}, language={3}, channels={4}{5}{6}",
                AudioTrackIndex,
                StreamIndex,
                string.IsNullOrWhiteSpace(CodecName) ? "(unknown)" : CodecName,
                language,
                Channels,
                layout,
                title);
        }
    }
}