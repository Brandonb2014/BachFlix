using System;
using System.Diagnostics;
using System.IO;
using BachFlixNfo.Features;

internal static class VideoHealthProgressTests
{
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    public static int Main()
    {
        string root = Path.Combine(Path.GetTempPath(), "BachFlixHealthTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var original = Console.Out;
        var log = new StringWriter();
        try
        {
            Console.SetOut(log);
            var progress = new FfmpegHealthProgress(100);
            progress.ReadLine("out_time_us=42500000");
            progress.ReadLine("speed=2.5x");
            Check(progress.CurrentText.Contains("42.5%") && progress.CurrentText.Contains("00:00:42 / 00:01:40"), "Timestamp parsing or percentage incorrect");
            Check(progress.CurrentText.Contains("Speed: 2.5x") && progress.CurrentText.Contains("Remaining: ~00:00:23"), "Speed or remaining time incorrect");
            progress.ReadLine("out_time_us=N/A");
            progress.ReadLine("out_time_us=-9223372036854775808");
            Check(progress.CurrentText.Contains("42.5%"), "Invalid timestamps should not change progress");
            progress.ReadLine("out_time_ms=50000000");
            Check(progress.CurrentText.Contains("50.0%"), "Legacy microsecond timestamp parsing incorrect");
            Check(FfmpegHealthProgress.Format(150, 100, 2, 1).Contains("99.9%"), "Must not report 100% before successful exit");
            Check(FfmpegHealthProgress.Format(42, null, 2, 10).Contains("duration unknown"), "Unknown duration needs a fallback");

            string video = Path.Combine(root, "healthy.mkv");
            using (var generator = Process.Start(new ProcessStartInfo
            {
                FileName = "ffmpeg",
                Arguments = "-v error -nostdin -f lavfi -i testsrc2=size=64x64:rate=10 -t 3 -c:v ffv1 \"" + video + "\"",
                UseShellExecute = false,
                CreateNoWindow = true
            }))
            {
                generator.WaitForExit();
                Check(generator.ExitCode == 0, "Fixture generation failed");
            }
            var duration = FfmpegHealthProgress.ReadDuration(video);
            Check(duration.HasValue && Math.Abs(duration.Value - 3) < .1, "ffprobe duration incorrect");
            log.GetStringBuilder().Clear();
            var result = VideoHealthCheck.AnalyzeVideoFile(video);
            Check(result.ExitCode == 0 && !result.HadErrors && result.Status == VideoHealthCheck.VideoHealthStatus.OK, "Progress output must not classify a healthy video as corrupt");
            Check(log.ToString().Contains("Scan:") && log.ToString().Contains("Scan finished: 100%"), "Real ffmpeg progress missing");

            string invalid = Path.Combine(root, "invalid.mkv");
            File.WriteAllText(invalid, "not a video");
            log.GetStringBuilder().Clear();
            result = VideoHealthCheck.AnalyzeVideoFile(invalid);
            Check(result.ExitCode != 0 && result.HadErrors && result.Status == VideoHealthCheck.VideoHealthStatus.BAD, "Decode errors must still classify as BAD");
            Check(log.ToString().Contains("Scan stopped with errors") && !log.ToString().Contains("Scan finished: 100%"), "Failed scan incorrectly shows success");
            Console.SetOut(original);
            Console.WriteLine("Health progress tests passed: real ffmpeg/ffprobe scan, timestamps, percentage/ETA, unknown duration, unchanged health classification.");
            return 0;
        }
        finally
        {
            Console.SetOut(original);
            Directory.Delete(root, true);
        }
    }
}
