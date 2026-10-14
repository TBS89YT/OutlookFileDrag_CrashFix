using System;
using System.IO;
using log4net;

namespace OutlookFileDrag
{
    static class FileUtility
    {
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static string GetTempPath()
        {
            log.Debug("Getting temp path");
            string path = Path.Combine(Path.GetTempPath(), "OutlookFileDrag", Guid.NewGuid().ToString());
            if (!System.IO.Directory.Exists(path))
                System.IO.Directory.CreateDirectory(path);
            log.DebugFormat("Temp path: {0}", path);
            return path;
        }

        public static void CleanupTempPath(int tempFileExpiration)
        {
            log.Debug("Cleaning up temp path");
            string path = Path.Combine(Path.GetTempPath(), "OutlookFileDrag");
            log.InfoFormat("Temp path: {0}", path);
            if (!System.IO.Directory.Exists(path))
            {
                log.Info("Temp path does not exist");
                return;
            }

            var dirInfo = new DirectoryInfo(path);
            foreach(DirectoryInfo subfolder in dirInfo.GetDirectories())
            {
                //If folder was created before expiration window, delete it
                if (subfolder.CreationTimeUtc < DateTime.UtcNow.AddMinutes(-tempFileExpiration))
                    try
                    {
                        log.InfoFormat("Deleting temp folder: {0}", subfolder.FullName);
                        subfolder.Delete(true);
                    }
                    catch
                    {
                        log.WarnFormat("Could not delete temp folder: {0}", subfolder.FullName);
                    }
            }
        }

        public static string GetUniqueFilename(string filename)
        {
            string directory = Path.GetDirectoryName(Path.GetFullPath(filename));
            string name = Path.GetFileName(filename);
            ValidateExportName(name);
            string stem = Path.GetFileNameWithoutExtension(name);
            string ext = Path.GetExtension(name);
            // Never truncate a directory or the collision suffix.
            for (int index = 0; index < 1024; index++)
            {
                string suffix = index == 0 ? "" : string.Format(" ({0})", index);
                int available = Math.Min(255, NativeMethods.MAX_PATH - 2 - directory.Length) - ext.Length - suffix.Length;
                if (available < 1) throw new PathTooLongException("Export path leaves no space for a filename.");
                string shortStem = stem.Length > available ? stem.Substring(0, available) : stem;
                if (shortStem.Length > 0 && char.IsHighSurrogate(shortStem[shortStem.Length - 1]))
                    shortStem = shortStem.Substring(0, shortStem.Length - 1);
                string candidateName = shortStem + suffix + ext;
                ValidateExportName(candidateName);
                string candidate = Path.Combine(directory, candidateName);
                if (!File.Exists(candidate) && !Directory.Exists(candidate)) return candidate;
            }
            throw new IOException("Could not generate a unique export filename.");
        }

        internal static string GetSafeExportPath(string exportDirectory, string name)
        {
            ValidateExportName(name);
            string root = Path.GetFullPath(exportDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string target = Path.GetFullPath(GetUniqueFilename(Path.Combine(root, name)));
            if (!string.Equals(Path.GetDirectoryName(target), root, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Export filename escapes the temporary directory.");
            return target;
        }

        private static void ValidateExportName(string name)
        {
            if (string.IsNullOrWhiteSpace(name) || name == "." || name == ".." ||
                name.Length > 255 || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                name.IndexOf('/') >= 0 || name.IndexOf('\\') >= 0 || Path.IsPathRooted(name) ||
                name.EndsWith(".", StringComparison.Ordinal) || name.EndsWith(" ", StringComparison.Ordinal))
                throw new InvalidDataException("Invalid export filename.");
            string device = name.Split('.')[0].TrimEnd(' ').ToUpperInvariant();
            if (device == "CON" || device == "PRN" || device == "AUX" || device == "NUL" ||
                device == "CONIN$" || device == "CONOUT$" ||
                (device.Length == 4 && (device.StartsWith("COM", StringComparison.Ordinal) || device.StartsWith("LPT", StringComparison.Ordinal)) &&
                 ("123456789¹²³".IndexOf(device[3]) >= 0)))
                throw new InvalidDataException("Reserved Windows device filename.");
        }

    }
}
