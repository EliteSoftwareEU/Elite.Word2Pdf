using System.Collections.Generic;
using System.IO;
using System.Linq;
namespace Elite.Word2Pdf
{
    public static class FileUtils
    {
        public static bool IsWordFile(this string path)
        {
            return path.Contains(".doc") || path.Contains(".docx");
        }

        public static List<string> WordFiles(this string path)
        {
            var dirInfo = new DirectoryInfo(path);
            var docFiles = dirInfo.EnumerateFiles().Where(fi => IsWordFile(fi.FullName));
            return docFiles.Select(fi => fi.FullName).ToList();
        }
    }
}
