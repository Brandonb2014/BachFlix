using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Features
{
    public sealed class SubtitleSourceDetector
    {
        private static readonly string[] SupportedExtensions = new[] { ".srt", ".ass", ".ssa", ".vtt" };

        public List<SubtitleSourceInfo> DetectAllSubtitleSources(MediaProbeResult media, string mkvPath)
        {
            var sources = new List<SubtitleSourceInfo>();

            if (media != null)
            {
                foreach (SubtitleTrackInfo track in media.SubtitleTracks)
                {
                    sources.Add(new SubtitleSourceInfo
                    {
                        SourceKind = SubtitleSourceKind.Embedded,
                        StreamIndex = track.StreamIndex,
                        SubtitleTrackIndex = track.SubtitleTrackIndex,
                        CodecName = track.CodecName,
                        Language = track.Language,
                        Title = track.Title,
                        IsDefault = track.IsDefault,
                        HasUnknownLanguage = string.IsNullOrWhiteSpace(track.Language)
                    });
                }
            }

            sources.AddRange(FindExternalSubtitleFiles(mkvPath));
            MarkEnglishSubtitleSources(sources);
            return sources;
        }

        public List<SubtitleSourceInfo> DetectEnglishSubtitleSources(MediaProbeResult media, string mkvPath)
        {
            return DetectAllSubtitleSources(media, mkvPath)
                .Where(s => s.IsEnglish)
                .ToList();
        }

        private static IEnumerable<SubtitleSourceInfo> FindExternalSubtitleFiles(string mkvPath)
        {
            if (string.IsNullOrWhiteSpace(mkvPath) || !File.Exists(mkvPath))
                return Enumerable.Empty<SubtitleSourceInfo>();

            string directory = Path.GetDirectoryName(mkvPath);
            string baseName = Path.GetFileNameWithoutExtension(mkvPath);
            if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(baseName))
                return Enumerable.Empty<SubtitleSourceInfo>();

            string[] files;
            try
            {
                files = Directory.GetFiles(directory, baseName + "*.*", SearchOption.TopDirectoryOnly);
            }
            catch
            {
                return Enumerable.Empty<SubtitleSourceInfo>();
            }

            return files
                .Where(path => SupportedExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
                .Where(path => !LooksGenerated(path))
                .Select(path =>
                {
                    string language = DetectLanguageFromFileName(path);
                    return new SubtitleSourceInfo
                    {
                        SourceKind = SubtitleSourceKind.External,
                        Path = path,
                        WorkingSubtitlePath = path,
                        CodecName = Path.GetExtension(path).TrimStart('.').ToLowerInvariant(),
                        Language = language,
                        Title = Path.GetFileName(path),
                        HasUnknownLanguage = string.IsNullOrWhiteSpace(language)
                    };
                })
                .ToList();
        }

        private static bool LooksGenerated(string path)
        {
            string lower = Path.GetFileName(path).ToLowerInvariant();
            return lower.Contains(".whisperx.") || lower.Contains(".clean.") || lower.Contains("_clean.");
        }

        private static void MarkEnglishSubtitleSources(List<SubtitleSourceInfo> sources)
        {
            int unknownCount = sources.Count(s => s.HasUnknownLanguage);
            bool onlyOneSubtitleIsUnknown = sources.Count == 1 && unknownCount == 1;

            foreach (SubtitleSourceInfo source in sources)
            {
                if (IsEnglishLanguage(source.Language))
                {
                    source.IsEnglish = true;
                    continue;
                }

                if (onlyOneSubtitleIsUnknown)
                {
                    source.IsEnglish = true;
                    source.IsAssumedEnglish = true;
                }
            }
        }

        public static bool IsEnglishLanguage(string language)
        {
            if (string.IsNullOrWhiteSpace(language))
                return false;

            string value = language.Trim().ToLowerInvariant();
            return value == "eng" || value == "en" || value == "english";
        }

        private static string DetectLanguageFromFileName(string path)
        {
            string fileName = Path.GetFileNameWithoutExtension(path) ?? "";
            string[] tokens = Regex.Split(fileName.ToLowerInvariant(), @"[\s._\-\[\](){}]+")
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .ToArray();

            foreach (string token in tokens)
            {
                if (token == "eng" || token == "en" || token == "english")
                    return token;
            }

            return "";
        }
    }

    public sealed class EmbeddedSubtitleExtractor
    {
        public void ExtractEmbeddedSubtitles(
            IEnumerable<SubtitleSourceInfo> sources,
            string mkvPath,
            string ffmpegPath,
            string workDirectory,
            Action<string, string, int> log,
            IList<string> commandLog)
        {
            if (sources == null)
                return;

            Directory.CreateDirectory(workDirectory);
            string baseName = Path.GetFileNameWithoutExtension(mkvPath);

            foreach (SubtitleSourceInfo source in sources.Where(s => s.SourceKind == SubtitleSourceKind.Embedded))
            {
                string extension = ExtensionForSubtitleCodec(source.CodecName);
                string outputPath = Path.Combine(workDirectory, baseName + ".subtitle-stream-" + source.StreamIndex + extension);
                string args =
                    "-hide_banner -nostdin -y -v error " +
                    "-i " + Quote(mkvPath) + " " +
                    "-map 0:s:" + source.SubtitleTrackIndex + " " +
                    Quote(outputPath);

                ExternalCommandResult result = ExternalToolResolver.RunProcess(ffmpegPath, args, log, commandLog);
                if (result.ExitCode == 0 && File.Exists(outputPath))
                {
                    source.WorkingSubtitlePath = outputPath;
                }
                else
                {
                    source.ScanError = "Could not extract embedded subtitle stream " + source.StreamIndex + ": " + FirstUsefulLine(result.StandardError, result.StandardOutput);
                }
            }
        }

        private static string ExtensionForSubtitleCodec(string codec)
        {
            string value = (codec ?? "").Trim().ToLowerInvariant();
            if (value.Contains("ass"))
                return ".ass";
            if (value.Contains("ssa"))
                return ".ssa";
            if (value.Contains("webvtt") || value == "vtt")
                return ".vtt";
            return ".srt";
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

        private static string Quote(string value)
        {
            return "\"" + (value ?? "").Replace("\"", "\\\"") + "\"";
        }
    }
}
