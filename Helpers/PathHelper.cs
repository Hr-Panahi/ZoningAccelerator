using System;
using System.IO;
using System.Linq;

namespace ZoningAccelerator.Helpers
{
    public static class PathHelper
    {
        public static string GetUniqueFilePath(string prefix, string cityPath, string extension)
        {
            var exeFolder = AppContext.BaseDirectory;

            var saveDirectory = Path.Combine(exeFolder, "Saved Files");

            if (!Directory.Exists(saveDirectory))
                Directory.CreateDirectory(saveDirectory);

            var cityFileName = Path.GetFileNameWithoutExtension(cityPath);
            var safeCityName = string.Concat(cityFileName.Select(c => char.IsLetterOrDigit(c) ? c : '_'));

            var baseName = $"{prefix}_{safeCityName}";
            var fileName = $"{baseName}{extension}";
            var fullPath = Path.Combine(saveDirectory, fileName);

            int counter = 1;
            while (File.Exists(fullPath))
            {
                fileName = $"{baseName}_{counter}{extension}";
                fullPath = Path.Combine(saveDirectory, fileName);
                counter++;
            }

            return fullPath;
        }
    }
}
