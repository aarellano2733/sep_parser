using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelConversion
{
    class Program
    {
        public static void Main(string[] args)
        {
            string mapInPath = @"D:\ExcelConversion\ExcelConversion\InputMapForXlsm.xlsx";
            string fileOutPath = "";
            string fileRootPath = Directory.GetParent(Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).ToString()).ToString();
            List<string> files = System.IO.Directory.GetFiles(fileRootPath, "*.xlsm").ToList();

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                if (File.Exists(file))
                {
                    // This path is a file
                    ProcessFile(file);
                }
                else if (Directory.Exists(file))
                {
                    // This path is a directory
                    ProcessDirectory(file);
                }
                else
                {
                    Console.WriteLine("{0} is not a valid file or directory.", file);
                }
            }
            //instaniate class
            ConvertExcel convert = new ConvertExcel();
            //read map
            List<MapVal> map = convert.ReadMapFile(mapInPath);
            //read from excel
            StringBuilder output = convert.ReadInfoFromExcel(fileRootPath, map);
            //write out to CSV
            convert.WriteToCSV(output, fileOutPath);
        }
        public static void ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessFile(fileName);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory);
        }

        // Insert logic for processing found files here.
        public static void ProcessFile(string file)
        {
            Console.WriteLine("Processed file '{0}'.", file);
        }
    }
}