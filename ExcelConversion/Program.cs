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
            string test = Directory.GetCurrentDirectory().ToString();
            string mapInPath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString()).ToString() + @"\InputMapForXlsm.xlsx";
            string fileOutPath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString()).ToString() + @"\SpliceCsvs\";
            string fileRootPath = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString()).ToString() + @"\Splices";
            List<string> files = Directory.GetFiles(fileRootPath, "*.xlsm").ToList();
            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Directory);
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
                string spliceName = "";
                //string[] spliceFile;
                foreach (string splice in files)
                {
                    string[] fileNameStringArray = file.Split('\\');
                    fileNameStringArray = fileNameStringArray[fileNameStringArray.Length - 1].Split('.');
                    spliceName = fileNameStringArray[0] + ".csv";
                }
                //spliceFile = fileOutPath + 
                //instaniate class
                ConvertExcel convert = new ConvertExcel();
                //read map
                List<MapVal> map = convert.ReadMapFile(mapInPath);
                //read from excel
                StringBuilder output = convert.ReadInfoFromExcel(file, map);
                convert.WriteToCSV(output, fileOutPath, spliceName);
            }
        }
        public static void ProcessDirectory(string fileRootPath)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(fileRootPath);
            foreach (string fileName in fileEntries)
                ProcessFile(fileName);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(fileRootPath);
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