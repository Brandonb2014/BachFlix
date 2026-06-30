using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BachFlixNfo.Subtitles
{
    public class SrtScoringEngine
    {
        private static readonly Regex TimeLineRegex =
            new Regex(@"^\s*(\d{2}):(\d{2}):(\d{2}),(\d{3})\s*-->\s*(\d{2}):(\d{2}):(\d{2}),(\d{3})\s*$",
                      RegexOptions.Compiled);

        public class SrtCue
        {
            public TimeSpan Start { get; set; }
            public TimeSpan End { get; set; }
        }

        /// <summary>
        /// Main entry – returns 0–100 score. Returns null if SRT or video cannot be analyzed.
        /// </summary>
        public int? ScoreSubtitleForVideo(string videoPath)
        {
            if (string.IsNullOrWhiteSpace(videoPath) || !File.Exists(videoPath))
                return null;

            string srtPath = FindBestMatchingSrt(videoPath);
            if (srtPath == null || !File.Exists(srtPath))
                return null;

            TimeSpan? videoDuration = GetVideoDuration(videoPath);
            if (videoDuration == null || videoDuration.Value.TotalSeconds < 60)
                return null; // no useful duration

            List<SrtCue> cues = ParseSrtFile(srtPath);
            if (cues == null || cues.Count == 0)
                return 0;

            return ComputeScore(cues, videoDuration.Value);
        }

        /// <summary>
        /// Try to find the "main" SRT for a given video.
        /// Priority:
        ///   1) Same basename + *.eng*.srt
        ///   2) Same basename + *.en*.srt
        ///   3) Same basename + *.srt
        /// </summary>
        private string FindBestMatchingSrt(string videoPath)
        {
            string dir = Path.GetDirectoryName(videoPath);
            if (dir == null) return null;

            string baseName = Path.GetFileNameWithoutExtension(videoPath);
            if (baseName == null) return null;

            // Candidates limited to same basename so we don't pick random SRTs
            string[] all = Directory.GetFiles(dir, baseName + "*.srt");

            if (all.Length == 0)
                return null;

            // Prefer files with language hints
            string best = all
                .OrderByDescending(p =>
                {
                    string lower = Path.GetFileName(p).ToLowerInvariant();
                    int score = 0;
                    if (lower.Contains(".eng")) score += 3;
                    if (lower.Contains(".en")) score += 2;
                    if (lower.Contains(".forced")) score -= 1; // maybe not full subs
                    if (lower.Contains(".cc")) score += 1;
                    return score;
                })
                .FirstOrDefault();

            return best;
        }

        private TimeSpan? GetVideoDuration(string videoPath)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "ffprobe",
                    Arguments = "-v error -show_entries format=duration -of default=nk=1:nw=1 \"" + videoPath + "\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var proc = Process.Start(psi))
                {
                    if (proc == null) return null;
                    string output = proc.StandardOutput.ReadToEnd();
                    proc.WaitForExit();

                    output = output.Trim();
                    if (double.TryParse(output, NumberStyles.Any, CultureInfo.InvariantCulture, out double seconds))
                    {
                        if (seconds > 0 && seconds < 24 * 60 * 60) // sanity
                            return TimeSpan.FromSeconds(seconds);
                    }
                }
            }
            catch
            {
                // ignore, just return null
            }

            return null;
        }

        private List<SrtCue> ParseSrtFile(string srtPath)
        {
            var cues = new List<SrtCue>();

            try
            {
                string[] lines = File.ReadAllLines(srtPath);
                for (int i = 0; i < lines.Length; i++)
                {
                    Match m = TimeLineRegex.Match(lines[i]);
                    if (!m.Success) continue;

                    TimeSpan start = ParseTimeFromMatch(m, 1);
                    TimeSpan end = ParseTimeFromMatch(m, 5);

                    if (start < TimeSpan.Zero || end <= start)
                        continue;

                    cues.Add(new SrtCue { Start = start, End = end });
                }
            }
            catch
            {
                return null;
            }

            return cues;
        }

        private TimeSpan ParseTimeFromMatch(Match m, int groupOffset)
        {
            int h = int.Parse(m.Groups[groupOffset + 0].Value);
            int min = int.Parse(m.Groups[groupOffset + 1].Value);
            int s = int.Parse(m.Groups[groupOffset + 2].Value);
            int ms = int.Parse(m.Groups[groupOffset + 3].Value);

            return new TimeSpan(0, h, min, s, ms);
        }

        private int ComputeScore(List<SrtCue> cues, TimeSpan videoDuration)
        {
            if (cues == null || cues.Count == 0)
                return 0;

            double videoSeconds = videoDuration.TotalSeconds;
            double videoMinutes = videoDuration.TotalMinutes;

            // basic metrics
            TimeSpan firstStart = cues.Min(c => c.Start);
            TimeSpan lastEnd = cues.Max(c => c.End);
            int cueCount = cues.Count;
            double cuesPerMinute = cueCount / Math.Max(1.0, videoMinutes);

            int score = 100;

            // 1) Too few cues overall
            if (cueCount < 10)
            {
                score -= 70; // basically useless
            }
            else if (cueCount < videoMinutes * 1.5) // < 1.5 cues / minute
            {
                score -= 25;
            }

            // 2) Cues per minute sanity range
            if (cuesPerMinute < 3.0)
                score -= 20;
            else if (cuesPerMinute > 40.0)
                score -= 10; // likely some weird timing

            // 3) First subtitle too late
            double firstSec = firstStart.TotalSeconds;
            if (firstSec > 3600) score -= 80; // 1h+ late
            else if (firstSec > 1800) score -= 60; // 30m+
            else if (firstSec > 1200) score -= 40; // 20m+
            else if (firstSec > 600) score -= 25; // 10m+
            else if (firstSec > 300) score -= 15; // 5m+
            else if (firstSec > 120) score -= 5;  // 2m+

            // 4) Last subtitle too early compared to video end
            double lastEndSec = lastEnd.TotalSeconds;
            if (lastEndSec < videoSeconds * 0.6)
                score -= 30;
            else if (lastEndSec < videoSeconds * 0.75)
                score -= 15;
            else if (lastEndSec < videoSeconds * 0.85)
                score -= 5;

            // 5) Any cues beyond video duration + tolerance?
            bool anyBeyond = cues.Any(c => c.Start.TotalSeconds > videoSeconds + 30 ||
                                            c.End.TotalSeconds > videoSeconds + 30);
            if (anyBeyond)
                score -= 40;

            // 6) Any decreasing times (broken ordering)?
            int badOrderCount = 0;
            TimeSpan last = TimeSpan.Zero;
            foreach (var cue in cues.OrderBy(c => c.Start))
            {
                if (cue.Start < last)
                    badOrderCount++;
                last = cue.Start;
            }

            if (badOrderCount > 0)
                score -= Math.Min(30, badOrderCount * 5);

            if (score < 0) score = 0;
            if (score > 100) score = 100;

            return score;
        }
    }
}
