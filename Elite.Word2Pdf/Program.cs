using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
                   
namespace Elite.Word2Pdf
{
    class Program
    {
        static string OS_X_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice";
        static string LINUX_PATH = "/usr/bin/soffice";
        static string SOFFICE_ARGS = " --convert-to pdf --nologo {0}";

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("You must provide file or directory to convert!");
                return;
            }

            string arguments = string.Format(SOFFICE_ARGS, args[0]);

            if (arguments.IsWordFile())
            {
                Convert(arguments);
                return;
            } else 
            {
                foreach(string wordDocumentPath in args[0].WordFiles())
                {
                    Convert(wordDocumentPath);
                }
            }
        }

        private static void Convert(string arguments)
        {
            ProcessStartInfo procStartInfo = new ProcessStartInfo(GetLibreOfficePath(), arguments);
            procStartInfo.RedirectStandardOutput = true;
            procStartInfo.UseShellExecute = false;
            procStartInfo.CreateNoWindow = true;
            procStartInfo.WorkingDirectory = Environment.CurrentDirectory;

            Process process = new Process() { StartInfo = procStartInfo, };
            process.Start();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                throw new LibreOfficeFailedException(process.ExitCode);
            }
        }

        static string GetLibreOfficePath()
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return OS_X_PATH;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) return LINUX_PATH;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                string currentDir = Assembly.GetExecutingAssembly().Location;
                string binaryDirectory = System.IO.Path.GetDirectoryName(currentDir);
                return binaryDirectory + "\\Windows\\program\\soffice.exe";
            }
            else
            {
                throw new PlatformNotSupportedException("Your OS is not supported");
            } 
        }
    }
}
