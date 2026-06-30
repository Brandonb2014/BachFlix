using System;
using System.Collections.Generic;
using System.IO;

public static class BachFlixLog
{
    /// <summary>
    /// Writes a BachFlixNfo log file to the user's Desktop under:
    /// Desktop\BachFlixNfo Logs\<category>\
    /// </summary>
    /// <param name="lines">Log lines to write. If null/empty, nothing is written and null is returned.</param>
    /// <param name="category">Subfolder under "BachFlixNfo Logs" (e.g., "File Renamer", "SRT Score").</param>
    /// <param name="fileNamePrefix">Filename prefix after "BachFlixNfo_" (e.g., "FileRenamer", "SrtScore").</param>
    /// <param name="error">If write fails, contains the exception message; otherwise null.</param>
    /// <returns>Full path to the written log file, or null if no lines or failed to write.</returns>
    public static string WriteBachFlixLog(
        IReadOnlyCollection<string> lines,
        string category,
        string fileNamePrefix,
        out string error)
    {
        error = null;

        if (lines == null || lines.Count == 0)
            return null;

        if (string.IsNullOrWhiteSpace(category))
            category = "General";

        if (string.IsNullOrWhiteSpace(fileNamePrefix))
            fileNamePrefix = "Log";

        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string logRootDirectory = Path.Combine(desktopPath, "BachFlixNfo Logs", category);

        string logFileName = $"BachFlixNfo_{fileNamePrefix}_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
        string logPath = Path.Combine(logRootDirectory, logFileName);

        try
        {
            Directory.CreateDirectory(logRootDirectory);
            File.WriteAllLines(logPath, lines);
            return logPath;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return null;
        }
    }

    /// <summary>
    /// Convenience overload when you don't care about the error string.
    /// </summary>
    public static string WriteBachFlixLog(
        IReadOnlyCollection<string> lines,
        string category,
        string fileNamePrefix)
    {
        return WriteBachFlixLog(lines, category, fileNamePrefix, out _);
    }
}
