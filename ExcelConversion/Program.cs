using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            string mapInPath = @"D:\ExcelConversion\ExcelConversion\InputMapForXlsm.xlsx";
            string fileInPath = @"D:\ExcelConversion\ExcelConversion\Splice-26979.xlsm";
            string fileOutPath = @"D:\ExcelConversion\ExcelConversion\Splice-26979.csv";
            //instaniate class
            ConvertExcel convert = new ConvertExcel();
            //read map
            List<MapVal> map = convert.ReadMapFile(mapInPath);
            //read from excel
            StringBuilder output = convert.ReadInfoFromExcel(fileInPath, map);
            //write out to CSV
            convert.WriteToCSV(output, fileOutPath);
        }
    }
}
