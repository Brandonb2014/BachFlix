using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace BachFlixNfo.Features
{
    public sealed class ExternalToolResolution
    {
        public string ToolName { get; set; }
        public string Path { get; set; }
        public bool Found { get; set; }
        public string Message { get; set; }
        public string VersionLine { get; set; }

        public ExternalToolResolution()
        {
            ToolName = "";
            Path = "";
            Message = "";
            VersionLine = "";
        }
    }

    internal sealed class ExternalCommandResult
    {
        public int ExitCode { get; set; }
        public string StandardOutput { get; set; }
        public string StandardError { get; set; }

        public ExternalCommandResult()
        {
            StandardOutput = "";
            StandardError = "";
        }
    }

    internal static class ExternalToolResolver
    {
        public static ExternalToolResolution Resolve(
            string explicitPath,
            string toolName,
            Action<string, string, int> log,
            IList<string> commandLog)
        {
            if (string.IsNullOrWhiteSpace(toolName))
                throw new ArgumentException("Tool name is required.", nameof(toolName));

            if (!string.IsNullOrWhiteSpace(explicitPath))
            {
                string cleanExplicitPath = NormalizeUserPath(explicitPath);
                if (File.Exists(cleanExplicitPath))
                    return VerifyCandidate(cleanExplicitPath, toolName, log, commandLog);

                Write(log, "warning", toolName + " was not found at explicit path: " + cleanExplicitPath, 1);
            }

            string appDirectoryCandidate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, AddExecutableExtension(toolName));
            if (File.Exists(appDirectoryCandidate))
                return VerifyCandidate(appDirectoryCandidate, toolName, log, commandLog);

            string lookupCommand = IsWindows() ? "where" : "which";
            ExternalCommandResult lookupResult = RunProcess(lookupCommand, toolName, log, commandLog);

            if (lookupResult.ExitCode == 0)
            {
                string resolvedPath = SplitLines(lookupResult.StandardOutput)
                    .Select(NormalizeUserPath)
                    .FirstOrDefault(File.Exists);

                if (!string.IsNullOrWhiteSpace(resolvedPath))
                    return VerifyCandidate(resolvedPath, toolName, log, commandLog);
            }

            return new ExternalToolResolution
            {
                ToolName = toolName,
                Found = false,
                Message = toolName + " was not found. Add it to PATH or provide an explicit path in a future options object."
            };
        }

        internal static ExternalCommandResult RunProcess(
            string fileName,
            string arguments,
            Action<string, string, int> log,
            IList<string> commandLog)
        {
            string commandText = BuildDisplayCommand(fileName, arguments);
            if (commandLog != null)
                commandLog.Add(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " | " + commandText);

            Write(log, "log", "Command: " + commandText, 1);

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments ?? "",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(psi))
                {
                    if (process == null)
                    {
                        return new ExternalCommandResult
                        {
                            ExitCode = -1,
                            StandardError = "Process failed to start."
                        };
                    }

                    string stdout = process.StandardOutput.ReadToEnd();
                    string stderr = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    return new ExternalCommandResult
                    {
                        ExitCode = process.ExitCode,
                        StandardOutput = stdout ?? "",
                        StandardError = stderr ?? ""
                    };
                }
            }
            catch (Exception ex)
            {
                return new ExternalCommandResult
                {
                    ExitCode = -1,
                    StandardError = ex.Message
                };
            }
        }

        private static ExternalToolResolution VerifyCandidate(
            string path,
            string toolName,
            Action<string, string, int> log,
            IList<string> commandLog)
        {
            ExternalCommandResult versionResult = RunProcess(path, "-version", log, commandLog);
            string versionLine = FirstUsefulLine(versionResult.StandardOutput, versionResult.StandardError);

            string message = versionResult.ExitCode == 0
                ? "Using " + toolName + ": " + path
                : "Found " + toolName + ", but version check returned exit code " + versionResult.ExitCode + ": " + path;

            Write(log, versionResult.ExitCode == 0 ? "success" : "warning", message, 1);

            return new ExternalToolResolution
            {
                ToolName = toolName,
                Path = path,
                Found = true,
                Message = message,
                VersionLine = versionLine
            };
        }

        private static string FirstUsefulLine(params string[] values)
        {
            foreach (string value in values)
            {
                string line = SplitLines(value).FirstOrDefault();
                if (!string.IsNullOrWhiteSpace(line))
                    return line.Trim();
            }

            return "";
        }

        private static IEnumerable<string> SplitLines(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return Enumerable.Empty<string>();

            return value
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(line => line.Trim())
                .Where(line => !string.IsNullOrWhiteSpace(line));
        }

        private static string NormalizeUserPath(string path)
        {
            if (path == null)
                return "";

            return path.Trim().Trim('"');
        }

        private static string AddExecutableExtension(string toolName)
        {
            if (!IsWindows())
                return toolName;

            return toolName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                ? toolName
                : toolName + ".exe";
        }

        private static bool IsWindows()
        {
            PlatformID platform = Environment.OSVersion.Platform;
            return platform == PlatformID.Win32NT ||
                   platform == PlatformID.Win32S ||
                   platform == PlatformID.Win32Windows ||
                   platform == PlatformID.WinCE;
        }

        private static string BuildDisplayCommand(string fileName, string arguments)
        {
            string command = QuoteIfNeeded(fileName);
            if (!string.IsNullOrWhiteSpace(arguments))
                command += " " + arguments;

            return command;
        }

        private static string QuoteIfNeeded(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            return value.IndexOf(' ') >= 0 ? "\"" + value + "\"" : value;
        }

        private static void Write(Action<string, string, int> log, string type, string message, int lines)
        {
            try
            {
                if (log != null)
                    log(type, message, lines);
            }
            catch
            {
                // Logging must never break a scan.
            }
        }
    }
}
