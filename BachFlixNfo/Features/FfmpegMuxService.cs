using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
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

        public MediaProbeResult ProbeMedia(
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
                "-v error " +
                "-show_entries stream=index,codec_type,codec_name,channels,channel_layout,duration:stream_tags=language,title:stream_disposition=default:format=duration:format_tags " +
                "-show_chapters " +
                "-of json " +
                Quote(inputMkvPath);

            ExternalCommandResult result = ExternalToolResolver.RunProcess(ffprobePath, arguments, log, commandLog);
            if (result.ExitCode != 0)
                throw new InvalidOperationException("ffprobe failed with exit code " + result.ExitCode + ": " + FirstUsefulLine(result.StandardError, result.StandardOutput));

            JObject root = JObject.Parse(result.StandardOutput);
            var probe = new MediaProbeResult { InputPath = inputMkvPath };
            JObject format = root["format"] as JObject;
            if (format != null)
                probe.Duration = ReadTimeSpanSeconds(format, "duration");

            JArray streams = root["streams"] as JArray;
            if (streams != null)
            {
                int audioTrackIndex = 0;
                int subtitleTrackIndex = 0;
                int attachmentIndex = 0;

                foreach (JToken stream in streams)
                {
                    JObject tags = stream["tags"] as JObject;
                    JObject disposition = stream["disposition"] as JObject;
                    string codecType = ReadString(stream, "codec_type");
                    var streamInfo = new MediaStreamInfo
                    {
                        StreamIndex = ReadInt(stream, "index"),
                        CodecType = codecType,
                        CodecName = ReadString(stream, "codec_name"),
                        Language = tags == null ? "" : ReadString(tags, "language"),
                        Title = tags == null ? "" : ReadString(tags, "title"),
                        Duration = ReadTimeSpanSeconds(stream, "duration"),
                        IsDefault = disposition != null && ReadInt(disposition, "default") == 1
                    };
                    probe.Streams.Add(streamInfo);

                    if (string.Equals(codecType, "audio", StringComparison.OrdinalIgnoreCase))
                    {
                        probe.AudioTracks.Add(new AudioTrackInfo
                        {
                            AudioTrackIndex = audioTrackIndex,
                            StreamIndex = streamInfo.StreamIndex,
                            CodecName = streamInfo.CodecName,
                            Channels = ReadInt(stream, "channels"),
                            ChannelLayout = ReadString(stream, "channel_layout"),
                            Language = streamInfo.Language,
                            Title = streamInfo.Title,
                            Duration = streamInfo.Duration,
                            IsDefault = streamInfo.IsDefault
                        });
                        audioTrackIndex++;
                    }
                    else if (string.Equals(codecType, "subtitle", StringComparison.OrdinalIgnoreCase))
                    {
                        probe.SubtitleTracks.Add(new SubtitleTrackInfo
                        {
                            SubtitleTrackIndex = subtitleTrackIndex,
                            StreamIndex = streamInfo.StreamIndex,
                            CodecName = streamInfo.CodecName,
                            Language = streamInfo.Language,
                            Title = streamInfo.Title,
                            Duration = streamInfo.Duration,
                            IsDefault = streamInfo.IsDefault
                        });
                        subtitleTrackIndex++;
                    }
                    else if (string.Equals(codecType, "attachment", StringComparison.OrdinalIgnoreCase))
                    {
                        probe.AttachmentStreams.Add(new AttachmentStreamInfo
                        {
                            AttachmentIndex = attachmentIndex,
                            StreamIndex = streamInfo.StreamIndex,
                            CodecName = streamInfo.CodecName,
                            Title = streamInfo.Title
                        });
                        attachmentIndex++;
                    }
                }
            }

            JArray chapters = root["chapters"] as JArray;
            probe.ChapterCount = chapters == null ? 0 : chapters.Count;
            return probe;
        }

        public List<AudioTrackInfo> ProbeAudioTracks(
            string ffprobePath,
            string inputMkvPath,
            Action<string, string, int> log,
            IList<string> commandLog)
        {
            return ProbeMedia(ffprobePath, inputMkvPath, log, commandLog).AudioTracks;
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

            var audioOutputs = request.CensoredAudioTracks ?? new List<AudioTrackCensorOutput>();
            var subtitleOutputs = request.CensoredSubtitleTracks ?? new List<CensoredSubtitleOutput>();
            if (audioOutputs.Count == 0 && !string.IsNullOrWhiteSpace(request.AudioFilter))
            {
                audioOutputs.Add(new AudioTrackCensorOutput
                {
                    SourceAudioTrackIndex = request.AudioTrackIndex < 0 ? 0 : request.AudioTrackIndex,
                    AudioFilter = request.AudioFilter,
                    AudioCodec = request.AudioCodec,
                    Title = "English Clean",
                    IsDefault = true
                });
            }

            if (audioOutputs.Count == 0 && subtitleOutputs.Count == 0)
                throw new ArgumentException("At least one censored audio or subtitle output is required.", nameof(request));

            var args = new List<string>();
            args.Add("-hide_banner");
            args.Add("-nostdin");
            args.Add("-n");
            args.Add("-v");
            args.Add("warning");
            args.Add("-i");
            args.Add(Quote(request.InputMkvPath));

            foreach (CensoredSubtitleOutput subtitle in subtitleOutputs)
            {
                args.Add("-i");
                args.Add(Quote(subtitle.Path));
            }

            List<string> filterChains = BuildAudioFilterChains(audioOutputs);
            if (filterChains.Count > 0)
            {
                args.Add("-filter_complex");
                args.Add(Quote(string.Join(";", filterChains)));
            }

            args.Add("-map");
            args.Add("0");

            for (int i = 0; i < audioOutputs.Count; i++)
            {
                args.Add("-map");
                args.Add("[censa" + i + "]");
            }

            for (int i = 0; i < subtitleOutputs.Count; i++)
            {
                args.Add("-map");
                args.Add((i + 1).ToString(CultureInfo.InvariantCulture) + ":0");
            }

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

            int originalAudioCount = request.MediaProbe == null ? 0 : request.MediaProbe.AudioTracks.Count;
            for (int i = 0; i < originalAudioCount; i++)
            {
                args.Add("-disposition:a:" + i);
                args.Add("0");
            }

            for (int i = 0; i < audioOutputs.Count; i++)
            {
                AudioTrackCensorOutput output = audioOutputs[i];
                int outputAudioOrdinal = originalAudioCount + i;
                args.Add("-c:a:" + outputAudioOrdinal);
                args.Add(string.IsNullOrWhiteSpace(output.AudioCodec) ? "flac" : output.AudioCodec);
                args.Add("-metadata:s:a:" + outputAudioOrdinal);
                args.Add("language=eng");
                args.Add("-metadata:s:a:" + outputAudioOrdinal);
                args.Add(Quote("title=" + (string.IsNullOrWhiteSpace(output.Title) ? "English Clean" : output.Title)));
                args.Add("-disposition:a:" + outputAudioOrdinal);
                args.Add(output.IsDefault ? "default" : "0");
            }

            int originalSubtitleCount = request.MediaProbe == null ? 0 : request.MediaProbe.SubtitleTracks.Count;
            for (int i = 0; i < originalSubtitleCount; i++)
            {
                args.Add("-disposition:s:" + i);
                args.Add("0");
            }

            for (int i = 0; i < subtitleOutputs.Count; i++)
            {
                CensoredSubtitleOutput output = subtitleOutputs[i];
                int outputSubtitleOrdinal = originalSubtitleCount + i;
                args.Add("-c:s:" + outputSubtitleOrdinal);
                args.Add(string.IsNullOrWhiteSpace(output.Codec) ? "srt" : output.Codec);
                args.Add("-metadata:s:s:" + outputSubtitleOrdinal);
                args.Add("language=" + (string.IsNullOrWhiteSpace(output.Language) ? "eng" : output.Language));
                args.Add("-metadata:s:s:" + outputSubtitleOrdinal);
                args.Add(Quote("title=" + (string.IsNullOrWhiteSpace(output.Title) ? "English Clean" : output.Title)));
                args.Add("-disposition:s:" + outputSubtitleOrdinal);
                args.Add(output.IsDefault ? "default" : "0");
            }

            args.Add(Quote(outputPath));

            return new FfmpegMuxPlan
            {
                FfmpegPath = request.FfmpegPath,
                Arguments = string.Join(" ", args),
                InputMkvPath = request.InputMkvPath,
                OutputMkvPath = outputPath,
                AudioTrackIndex = request.AudioTrackIndex < 0 ? 0 : request.AudioTrackIndex,
                CensoredAudioTrackCount = audioOutputs.Count,
                CensoredSubtitleTrackCount = subtitleOutputs.Count
            };
        }

        public void MuxCensoredMedia(FfmpegMuxPlan plan, Action<string, string, int> log, IList<string> commandLog)
        {
            if (plan == null)
                throw new ArgumentNullException(nameof(plan));

            ExternalCommandResult result = ExternalToolResolver.RunProcess(plan.FfmpegPath, plan.Arguments, log, commandLog);
            if (result.ExitCode != 0)
                throw new InvalidOperationException("FFmpeg mux failed with exit code " + result.ExitCode + ": " + FirstUsefulLine(result.StandardError, result.StandardOutput));
        }

        public void MuxCensoredAudio(FfmpegMuxPlan plan)
        {
            MuxCensoredMedia(plan, null, null);
        }

        private static List<string> BuildAudioFilterChains(IReadOnlyList<AudioTrackCensorOutput> audioOutputs)
        {
            var filters = new List<string>();
            if (audioOutputs == null)
                return filters;

            for (int i = 0; i < audioOutputs.Count; i++)
            {
                AudioTrackCensorOutput output = audioOutputs[i];
                if (output == null || string.IsNullOrWhiteSpace(output.AudioFilter))
                    continue;

                int sourceAudioTrackIndex = output.SourceAudioTrackIndex < 0 ? 0 : output.SourceAudioTrackIndex;
                filters.Add("[0:a:" + sourceAudioTrackIndex + "]" + output.AudioFilter + "[censa" + i + "]");
            }

            return filters;
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

        private static TimeSpan ReadTimeSpanSeconds(JToken token, string name)
        {
            double seconds;
            if (!double.TryParse(ReadString(token, name), NumberStyles.Float, CultureInfo.InvariantCulture, out seconds) || seconds < 0)
                return TimeSpan.Zero;

            return TimeSpan.FromSeconds(seconds);
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
        public MediaProbeResult MediaProbe { get; set; }
        public List<AudioTrackCensorOutput> CensoredAudioTracks { get; set; }
        public List<CensoredSubtitleOutput> CensoredSubtitleTracks { get; set; }

        public FfmpegMuxRequest()
        {
            FfmpegPath = "";
            InputMkvPath = "";
            OutputMkvPath = "";
            AudioFilter = "";
            AudioCodec = "flac";
            CensoredAudioTracks = new List<AudioTrackCensorOutput>();
            CensoredSubtitleTracks = new List<CensoredSubtitleOutput>();
        }
    }

    public sealed class FfmpegMuxPlan
    {
        public string FfmpegPath { get; set; }
        public string Arguments { get; set; }
        public string InputMkvPath { get; set; }
        public string OutputMkvPath { get; set; }
        public int AudioTrackIndex { get; set; }
        public int CensoredAudioTrackCount { get; set; }
        public int CensoredSubtitleTrackCount { get; set; }

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

    public sealed class MediaProbeResult
    {
        public string InputPath { get; set; }
        public TimeSpan Duration { get; set; }
        public int ChapterCount { get; set; }
        public List<MediaStreamInfo> Streams { get; private set; }
        public List<AudioTrackInfo> AudioTracks { get; private set; }
        public List<SubtitleTrackInfo> SubtitleTracks { get; private set; }
        public List<AttachmentStreamInfo> AttachmentStreams { get; private set; }

        public MediaProbeResult()
        {
            InputPath = "";
            Streams = new List<MediaStreamInfo>();
            AudioTracks = new List<AudioTrackInfo>();
            SubtitleTracks = new List<SubtitleTrackInfo>();
            AttachmentStreams = new List<AttachmentStreamInfo>();
        }
    }

    public sealed class MediaStreamInfo
    {
        public int StreamIndex { get; set; }
        public string CodecType { get; set; }
        public string CodecName { get; set; }
        public string Language { get; set; }
        public string Title { get; set; }
        public TimeSpan Duration { get; set; }
        public bool IsDefault { get; set; }

        public MediaStreamInfo()
        {
            CodecType = "";
            CodecName = "";
            Language = "";
            Title = "";
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
        public TimeSpan Duration { get; set; }
        public bool IsDefault { get; set; }

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
            string duration = Duration > TimeSpan.Zero ? ", duration=" + Duration.ToString(@"hh\:mm\:ss") : "";
            string defaultFlag = IsDefault ? ", default=yes" : "";

            return string.Format(
                "audioTrack={0}, stream={1}, codec={2}, language={3}, channels={4}{5}{6}{7}{8}",
                AudioTrackIndex,
                StreamIndex,
                string.IsNullOrWhiteSpace(CodecName) ? "(unknown)" : CodecName,
                language,
                Channels,
                layout,
                duration,
                defaultFlag,
                title);
        }
    }

    public sealed class SubtitleTrackInfo
    {
        public int SubtitleTrackIndex { get; set; }
        public int StreamIndex { get; set; }
        public string CodecName { get; set; }
        public string Language { get; set; }
        public string Title { get; set; }
        public TimeSpan Duration { get; set; }
        public bool IsDefault { get; set; }

        public SubtitleTrackInfo()
        {
            CodecName = "";
            Language = "";
            Title = "";
        }
    }

    public sealed class AttachmentStreamInfo
    {
        public int AttachmentIndex { get; set; }
        public int StreamIndex { get; set; }
        public string CodecName { get; set; }
        public string Title { get; set; }

        public AttachmentStreamInfo()
        {
            CodecName = "";
            Title = "";
        }
    }

    public sealed class AudioTrackCensorOutput
    {
        public int SourceAudioTrackIndex { get; set; }
        public AudioTrackInfo SourceTrack { get; set; }
        public string AudioFilter { get; set; }
        public string AudioCodec { get; set; }
        public string Title { get; set; }
        public bool IsDefault { get; set; }

        public AudioTrackCensorOutput()
        {
            AudioFilter = "";
            AudioCodec = "flac";
            Title = "English Clean";
            IsDefault = true;
        }
    }
}
