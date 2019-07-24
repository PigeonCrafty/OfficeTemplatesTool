using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TemplatesTool;
using TemplatesTool.Models;

namespace ConsoleTester
{
    class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine(DateTime.Now + " " + "Developed by Moravia Publishing & Media automation team. All rights reserved." + "\n");
            Console.WriteLine("Please input the root directory of loc templates folders >> ");
            Console.WriteLine("========================================");

            Input:
            var dirInput = Console.ReadLine();

            Console.WriteLine("========================================");
            Console.WriteLine("In progress of processing, please hold on and wait for a while :) ");
            Console.WriteLine("========================================");

            // Check if input path is empty
            if (string.IsNullOrEmpty(dirInput) && string.IsNullOrWhiteSpace(dirInput))
            {
                Console.WriteLine("Empty or invalid directory, please double check and re-enter it!");
                goto Input;
            }

            // Transfer input directory string to DirectoryInfo
            if (!File.Exists(dirInput) && Directory.Exists(dirInput))
            {
                if (!dirInput.EndsWith("\\")) dirInput = dirInput + "\\";

                var dirInfo = new DirectoryInfo(dirInput);

                // Get all subfolder names and add into LocalLanguage List
                try
                {
                    ListLangFolders(dirInfo);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error when create folders list： " + ex.Message);
                    return;
                }

                foreach (var lang in ListLang)
                    try
                    {
                        var langDirInfo = new DirectoryInfo(dirInput + lang + "\\");
                        SortLocFiles(langDirInfo);
                    }
                    catch (IOException e)
                    {
                        Console.WriteLine("Fail to sort files: " + e.Message);
                        return;
                    }
            }
            else if (File.Exists(dirInput))
            {
                var fileInfo = new FileInfo(dirInput);

                switch (fileInfo.Extension)
                {
                    case ".dotx":
                    case ".docx":
                    case ".dotm":
                        ListWord.Add(fileInfo.FullName);
                        break;

                    case ".xltx":
                    case ".xlsx":
                    case ".xltm":
                        ListExcel.Add(fileInfo.FullName);
                        break;

                    case ".potx":
                    case ".pptx":
                    case ".potm":
                        ListPpt.Add(fileInfo.FullName);
                        break;
                }
            }

            KillProcess();

            // PowerPointHandler
            if (ListPpt.Count > 0)
                foreach (var f in ListPpt)
                {
                    if (f.Contains("~$")) continue;
                    Console.Write("\r\n" + "Processing: ");
                    Common.WriteLine("\r\n" + f + "\r\n");
                    var objPpt = new PowerPointHandler(f);
                    objPpt.PptMain(f);
                    Common.WriteSglText(f);
                }

            // WordHandler
            if (ListWord.Count > 0)
                foreach (var f in ListWord)
                {
                    if (f.Contains("~$")) continue;
                    Console.Write("\r\n" + "Processing: ");
                    Common.WriteLine("\r\n" + f + "\r\n");
                    var objWordHandler = new WordHandler(f);
                    objWordHandler.WordMain(f);
                    Common.WriteSglText(f);
                }

            // Process ExcelHandler files
            if (ListExcel.Count > 0)
                foreach (var f in ListExcel)
                {
                    if (f.Contains("~$")) continue;
                    Console.Write("\r\n" + "Processing: ");
                    Common.WriteLine("\r\n" + f + "\r\n");
                    var objExcelHandler = new ExcelHandler(f);
                    objExcelHandler.ExcelMain(f);
                    Common.WriteSglText(f);
                }

            // Write all Console content to text
            Common.WriteAllText(dirInput);

            // End
            Console.WriteLine("========================================");
            Console.WriteLine("All files process complete! You are good to go!");
            Console.ReadLine();
        }

        #region Methods

        public static void ListLangFolders(DirectoryInfo rootInput)
        {
            try
            {
                var folders = rootInput.GetDirectories(); // To get folder list
                foreach (var fd in folders)
                    ListLang.Add(fd.Name);
                // Console.WriteLine(fd.Name);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error!" + ex.Message);
                throw;
            }
        }

        public static void SortLocFiles(DirectoryInfo langFolder)
        {
            if (!langFolder.Exists) return;

            var files = langFolder.GetFiles();

            foreach (var fil in files)
                switch (fil.Extension)
                {
                    case ".dotx":
                    case ".docx":
                    case ".dotm":
                        ListWord.Add(fil.FullName);
                        break;

                    case ".xltx":
                    case ".xlsx":
                    case ".xltm":
                        ListExcel.Add(fil.FullName);
                        break;

                    case ".potx":
                    case ".pptx":
                    case ".potm":
                        ListPpt.Add(fil.FullName);
                        // Console.WriteLine(fil.Directory.Name); // To get the file language folder name
                        break;
                }
        }

        public static void KillProcess()
        {
            if (Process.GetProcessesByName("POWERPNT").Any())
                foreach (var proc in Process.GetProcessesByName("POWERPNT"))
                    proc.Kill();

            if (Process.GetProcessesByName("WINWORD").Any())
                foreach (var proc in Process.GetProcessesByName("WINWORD"))
                    proc.Kill();

            if (Process.GetProcessesByName("EXCEL").Any())
                foreach (var proc in Process.GetProcessesByName("EXCEL"))
                    proc.Kill();
        }

        #endregion

        #region Lists Used        

        public static List<string> ListLang { get; set; } = new List<string>();
        public static List<string> ListWord { get; set; } = new List<string>();
        public static List<string> ListPpt { get; set; } = new List<string>();
        public static List<string> ListExcel { get; set; } = new List<string>();

        #endregion
    }
}
