using System;
using System.Collections.Generic;
using System.IO;

public static class LocalEnvironment
{
    private static readonly Lazy<Dictionary<string, string>> LocalValues =
        new Lazy<Dictionary<string, string>>(LoadLocalValues);

    public static string Get(string name, string defaultValue = "")
    {
        string value = Environment.GetEnvironmentVariable(name);
        if (!string.IsNullOrWhiteSpace(value))
            return value;

        if (LocalValues.Value.TryGetValue(name, out value) && !string.IsNullOrWhiteSpace(value))
            return value;

        return defaultValue;
    }

    public static string GetRequired(string name)
    {
        string value = Get(name);
        if (string.IsNullOrWhiteSpace(value))
            throw new InvalidOperationException($"Missing required local configuration value '{name}'. Set it as an environment variable or in BachFlixNfo.local.env.");

        return value;
    }

    private static Dictionary<string, string> LoadLocalValues()
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (string filePath in FindLocalEnvironmentFiles())
            LoadFile(filePath, values);

        return values;
    }

    private static IEnumerable<string> FindLocalEnvironmentFiles()
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (string root in new[] { Environment.CurrentDirectory, AppDomain.CurrentDomain.BaseDirectory })
        {
            string directory = root;
            while (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
            {
                foreach (string fileName in new[] { "BachFlixNfo.local.env", ".env" })
                {
                    string filePath = Path.Combine(directory, fileName);
                    if (seen.Add(filePath) && File.Exists(filePath))
                        yield return filePath;
                }

                DirectoryInfo parent = Directory.GetParent(directory);
                directory = parent == null ? null : parent.FullName;
            }
        }
    }

    private static void LoadFile(string filePath, Dictionary<string, string> values)
    {
        foreach (string rawLine in File.ReadAllLines(filePath))
        {
            string line = rawLine.Trim();
            if (line.Length == 0 || line.StartsWith("#"))
                continue;

            int equalsIndex = line.IndexOf('=');
            if (equalsIndex <= 0)
                continue;

            string name = line.Substring(0, equalsIndex).Trim();
            string value = line.Substring(equalsIndex + 1).Trim();

            if (value.Length >= 2 &&
                ((value[0] == '"' && value[value.Length - 1] == '"') ||
                 (value[0] == '\'' && value[value.Length - 1] == '\'')))
            {
                value = value.Substring(1, value.Length - 2);
            }

            values[name] = value;
        }
    }
}