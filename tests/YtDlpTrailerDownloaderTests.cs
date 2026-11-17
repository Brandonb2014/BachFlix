using System;
using System.Collections.Generic;
using System.IO;
using BachFlixNfo.Features;

internal static class YtDlpTrailerDownloaderTests
{
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    private static TrailerCommandResult Result(int code, string output = "", string error = "")
    {
        return new TrailerCommandResult { ExitCode = code, Output = output, Error = error };
    }

    private static TrailerCommandResult Scenario(TrailerCommandResult[] results, params string[] expectedArguments)
    {
        int call = 0;
        var result = YtDlpTrailerDownloader.DownloadWithRecovery("C:\\Fake Tools\\yt-dlp.exe", "download arguments",
            (path, arguments, display) =>
            {
                Check(path == "C:\\Fake Tools\\yt-dlp.exe", "Selected executable must be used for update and retry");
                Check(call < expectedArguments.Length && arguments == expectedArguments[call], "Unexpected command or extra retry: " + arguments);
                Check(display == !arguments.Contains("--version"), "Version queries should be quiet");
                return results[call++];
            }, (message, color) => { });
        Check(call == expectedArguments.Length, "Expected commands not run");
        return result;
    }

    public static int Main()
    {
        string root = Path.Combine(Path.GetTempPath(), "BachFlixTrailerTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            string executable = Path.Combine(root, "yt-dlp.exe");
            File.WriteAllText(executable, "test placeholder; never executed");
            Check(YtDlpTrailerDownloader.FindInFolder("\"" + root + "\"") == executable, "Quoted folder path should work");
            Check(YtDlpTrailerDownloader.FindInFolder(Path.Combine(root, "missing")) == null, "Missing executable accepted");
            var inputs = new Queue<string>(new[] { Path.Combine(root, "missing"), "  \"" + root + "\"  " });
            var messages = new List<string>();
            Check(YtDlpTrailerDownloader.PromptForExecutable(() => inputs.Dequeue(), messages.Add) == executable, "Prompt did not accept corrected folder path");
            Check(messages.Exists(message => message.Contains("not found in that folder")), "Invalid folder needs a clear message");
            Check(YtDlpTrailerDownloader.PromptForExecutable(() => "q", messages.Add) == null, "Cancel should stop prompting");
            Check(YtDlpTrailerDownloader.PromptForExecutable(() => null, messages.Add) == null, "Closed input should stop prompting");

            var result = Scenario(new[] { Result(0) }, "download arguments");
            Check(result.ExitCode == 0, "Successful download should not check for updates");

            result = Scenario(new[] { Result(1, error: "download failed"), Result(0, "2026.01.01"), Result(0), Result(0, "2026.10.01"), Result(0) },
                "download arguments", "--ignore-config --version", "--ignore-config -U", "--ignore-config --version", "download arguments");
            Check(result.ExitCode == 0, "Successful update/retry should return success");

            result = Scenario(new[] { Result(1, error: "original failure"), Result(0, "2026.10.01"), Result(0), Result(0, "2026.10.01") },
                "download arguments", "--ignore-config --version", "--ignore-config -U", "--ignore-config --version");
            Check(result.ExitCode == 1 && result.Error == "original failure", "Already-current downloader must keep original failure without retry");

            result = Scenario(new[] { Result(1, error: "original failure"), Result(0, "old"), Result(1, error: "update failed") },
                "download arguments", "--ignore-config --version", "--ignore-config -U");
            Check(result.ExitCode == 1 && result.Error == "original failure", "Failed updater must retain download failure");

            result = Scenario(new[] { Result(1), Result(0, "old"), Result(0), Result(0, "new"), Result(2, error: "retry failed") },
                "download arguments", "--ignore-config --version", "--ignore-config -U", "--ignore-config --version", "download arguments");
            Check(result.ExitCode == 2 && result.Error == "retry failed", "Failed retry must return final failure without looping");

            result = Scenario(new[] { Result(1), Result(1), Result(0), Result(0, "new") },
                "download arguments", "--ignore-config --version", "--ignore-config -U", "--ignore-config --version");
            Check(result.ExitCode == 1, "Unverified version must not trigger retry");

            Console.WriteLine("Trailer downloader tests passed: folder prompt/validation/cancel, success, update and retry, no update, failed update, failed retry, version verification.");
            return 0;
        }
        finally { Directory.Delete(root, true); }
    }
}
