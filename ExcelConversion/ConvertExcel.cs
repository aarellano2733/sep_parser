using ExcelDataReader;
using Nortal.Utilities.Csv;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace ExcelConversion
{
    public class ConvertExcel
    {
        public StringBuilder ReadInfoFromExcel(string fileInPath, List<MapVal> map)
        {

            StringBuilder output = new StringBuilder();


            //using (var stream = File.Open(fileInPath, FileMode.Open, FileAccess.Read))
            using (var stream = File.Open(fileInPath, FileMode.Open, FileAccess.ReadWrite))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsm)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                    //in xlsm sheet[0] is macro, so the first sheet is index 1
                    //System.Data.DataTable workSheetCoverLocInfo = dataSet.Tables["Cover - Location Info"];

                    System.Data.DataTable workSheetCoverLocInfo = dataSet.Tables["Cover - Location Info"];

                    switch (workSheetCoverLocInfo.TableName.ToString())
                    {
                        case "Cover - Location Info":
                            workSheetCoverLocInfo = dataSet.Tables["Cover - Location Info"];
                            break;
                        case "Cover - MH Info":
                            workSheetCoverLocInfo = dataSet.Tables["Cover - MH Info"];
                            break;
                        case "Cover Page":
                            workSheetCoverLocInfo = dataSet.Tables["Cover Page"];
                            break;
                        default:
                            break;
                    }

                    foreach (var field in map)
                    {
                        output.Append(GetFieldValue(workSheetCoverLocInfo, field.fieldName, field.fieldLabel, field.relativePos, field.offset) + ",");
                    }
                }
            }
            return output;
        }

        public List<MapVal> ReadMapFile(string mapFilePath)
        {
            List<MapVal> input = new List<MapVal>();
            using (var streamMap = File.Open(mapFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(streamMap))
                {
                    var inputMap = reader.AsDataSet().Tables[0].AsEnumerable();
                    bool skippedLabels = false;
                    foreach (DataRow row in inputMap)
                    {
                        if (skippedLabels)
                        {
                            var rowItems = row.ItemArray;
                            input.Add(new MapVal { fieldName = rowItems[0].ToString(), fieldLabel = rowItems[1].ToString(), relativePos = rowItems[2].ToString(), offset = rowItems[3].ToString() });
                        } else
                        {
                            skippedLabels = true;
                        }
                    }
                }
            }
            return input;
        }

        private string GetFieldValue(System.Data.DataTable workSheetCoverLocInfo, string fieldName, string fieldLabel, string relativePos, string offset)
        {
            if(fieldLabel == "City,State")
            {
                fieldName.Replace(fieldName, "'City/State'");
            }

            string returnVal = "";

            var xy = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, fieldLabel);
            int offsetNum = int.Parse(offset);

            switch (relativePos)
            {
                case "0": // down
                    returnVal = workSheetCoverLocInfo.AsEnumerable().AsDataView().Table.Rows[xy.rowIndex + offsetNum][xy.colIndex].ToString().Trim();
                    break;
                case "1": // right
                    returnVal = workSheetCoverLocInfo.AsEnumerable().AsDataView().Table.Rows[xy.rowIndex][xy.colIndex + offsetNum].ToString().Trim();
                    break;
                case "2": // up
                    returnVal = workSheetCoverLocInfo.AsEnumerable().AsDataView().Table.Rows[xy.rowIndex - offsetNum][xy.colIndex].ToString().Trim();
                    break;
                case "3": // left
                    returnVal = workSheetCoverLocInfo.AsEnumerable().AsDataView().Table.Rows[xy.rowIndex][xy.colIndex - offsetNum].ToString().Trim();
                    break;
                default:
                    break;
            }
            return returnVal;
        }

        private RowColIndexes GetTableRowColIndexesForExactMatch(System.Data.DataTable workSheetCoverLocInfo, string searchText)
        {
            RowColIndexes returnRCIndexes = new RowColIndexes();
            int rowIndex = -1; //return -1 if no match found

            var rowIndexArray = workSheetCoverLocInfo
             .Rows
             .Cast<DataRow>()
             .Where(r => r.ItemArray.Any(c => Regex.IsMatch(c.ToString().Trim(), Regex.Escape(searchText.Trim()), RegexOptions.IgnoreCase)))
             .Select(r => r.Table.Rows.IndexOf(r)).ToArray();

            if (rowIndexArray.Length > 0)
            {
                //return the row index
                var rowCol = rowIndexArray[0];
                rowIndex = rowCol;
            }

            int colIndex = 0;
            if (rowIndex >= 0)
            {
                //loop through row until column index is found
                foreach (var dc in workSheetCoverLocInfo.Rows[rowIndex].ItemArray)
                {
                    if (dc != DBNull.Value)
                    {
                        //if (dc.ToString() == searchText)
                        if(Regex.IsMatch(dc.ToString().Trim(), Regex.Escape(searchText.Trim()), RegexOptions.IgnoreCase))
                        {
                            break;
                        }
                    }
                    colIndex++;
                }
            }
            else
            {
                colIndex = -1; //return -1 if no match found
            }
            if (rowIndex <= 0 || colIndex <= 0)
            {
                rowIndex = 0;
                colIndex = 0;
                Console.WriteLine(searchText + " not found");
            }
            return new RowColIndexes { rowIndex = rowIndex, colIndex = colIndex };
        }


        private int GetTableRowIndexForContainsText(System.Data.DataTable workSheetCoverLocInfo, string searchText)
        {
            int returnRowIndex = -1; //return -1 if no match found
            var rowIndex = workSheetCoverLocInfo
             .Rows
             .Cast<DataRow>()
             //c.interiorColor color from css
             .Where(r => r.ItemArray.Any(c => Regex.IsMatch(c.ToString().Trim(), searchText, RegexOptions.IgnoreCase)))
             .Select(r => r.Table.Rows.IndexOf(r)).ToArray();

            if (rowIndex.Length > 0)
            {
                //return the row index
                returnRowIndex = rowIndex[0];
            }
            return returnRowIndex;
        }

        public void WriteToCSV(StringBuilder output, string fileOutPath)
        {
            using (var writer = new StringWriter())
            {
                var csv = new CsvWriter(writer, new CsvSettings());
                //csv.Write("MyValue");                    // writing one value at a time
                //csv.Write(2, "N2");                      // or with explicit format
                //csv.WriteLine(DateTime.Now);             // or with automatic formatting
                //csv.WriteLine(1, 2, 3, 4, DateTime.Now);    // another line with many values at once
                csv.WriteLine(output.ToString());
                File.WriteAllText(fileOutPath, writer.ToString());
            }
        }
    }
}
    