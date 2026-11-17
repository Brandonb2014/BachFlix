using System;
using System.Collections.Generic;
using System.IO;

namespace BachFlixNfo.Features
{
    internal static class OutputMediaFolders
    {
        // StreamFab folder name -> Recorded Source value used in the sheet.
        // Unlisted names already use the sheet's spelling and pass through.
        private static readonly Dictionary<string, string> SourceAliases =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Amazon", "Prime" }
            };

        public static string GetRecordSource(string providerFolder)
        {
            string name = Path.GetFileName(providerFolder.TrimEnd(
                Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            string source;
            return SourceAliases.TryGetValue(name, out source) ? source : name;
        }

        private static bool IsExcluded(string folder)
        {
            string name = Path.GetFileName(folder);
            return string.Equals(name, "Skip", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, ".sync", StringComparison.OrdinalIgnoreCase)
                || (File.GetAttributes(folder) & FileAttributes.ReparsePoint) != 0;
        }

        public static IEnumerable<string> GetProviderFolders(string outputRoot)
        {
            foreach (string folder in Directory.GetDirectories(outputRoot))
            {
                if (IsExcluded(folder)) continue;
                // Accept both Output/<provider> and the old Output/Ready/<provider>.
                if (string.Equals(Path.GetFileName(folder), "Ready", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (string provider in GetProviderFolders(folder))
                        yield return provider;
                }
                else
                {
                    yield return folder;
                }
            }
        }

        public static IEnumerable<string> GetMovableFolders(string providerFolder)
        {
            foreach (string folder in Directory.GetDirectories(providerFolder))
            {
                // Folder transfers include the entire tree. Hold the whole title if
                // it contains a protected directory, so Skip is never moved indirectly.
                bool hasFiles = false;
                if (!ContainsExcludedFolder(folder, ref hasFiles) && hasFiles) yield return folder;
            }
        }

        private static bool ContainsExcludedFolder(string folder, ref bool hasFiles)
        {
            if (IsExcluded(folder)) return true;
            if (Directory.GetFiles(folder).Length > 0) hasFiles = true;
            foreach (string child in Directory.GetDirectories(folder))
                if (ContainsExcludedFolder(child, ref hasFiles)) return true;
            return false;
        }

        public static void DeleteEmptyFolders(string outputRoot, Action<string, Exception> onError)
        {
            // Keep the configured root itself, even when completely empty.
            string root = Path.GetFullPath(outputRoot).TrimEnd(
                Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (Directory.Exists(root)) DeleteEmptyChildren(root, root, onError);
        }

        private static void DeleteEmptyChildren(string folder, string root, Action<string, Exception> onError)
        {
            try
            {
                foreach (string child in Directory.GetDirectories(folder))
                {
                    string fullPath = Path.GetFullPath(child);
                    if (!fullPath.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                        throw new IOException("Cleanup path is outside the configured output root: " + fullPath);
                    if (IsExcluded(child)) continue;
                    DeleteEmptyChildren(child, root, onError);
                }
                if (!string.Equals(folder, root, StringComparison.OrdinalIgnoreCase)
                    && Directory.GetFileSystemEntries(folder).Length == 0)
                {
                    // Never delete recursively: newly arrived files prevent deletion.
                    Directory.Delete(folder, false);
                }
            }
            catch (IOException ex) { onError?.Invoke(folder, ex); }
            catch (UnauthorizedAccessException ex) { onError?.Invoke(folder, ex); }
        }
    }
}
