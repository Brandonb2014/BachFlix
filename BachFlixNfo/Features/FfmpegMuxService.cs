using System;
using System.Collections.Generic;
using System.IO;

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

            string arguments =
                "-i " + Quote(request.InputMkvPath) +
                " -map 0" +
                " -map_metadata 0" +
                " -map_chapters 0" +
                " -c copy" +
                " -c:v copy" +
                " -c:s copy" +
                " -c:t copy" +
                " -filter:a:" + audioTrackIndex + " " + Quote(request.AudioFilter) +
                " -c:a:" + audioTrackIndex + " " + audioCodec +
                " " + Quote(outputPath);

            return new FfmpegMuxPlan
            {
                FfmpegPath = request.FfmpegPath,
                Arguments = arguments,
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

        public AudioTrackInfo()
        {
            CodecName = "";
            Language = "";
            Title = "";
        }
    }
}
