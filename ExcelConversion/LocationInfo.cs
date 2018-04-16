namespace ExcelConversion
{
    public class RowColIndexes
    {
        public int rowIndex { get; set; }
        public int colIndex { get; set; }
    }

    public class MapVal
    {
        public string fieldName { get; set; }
        public string fieldLabel { get; set; }

        public string relativePos { get; set; }

        public string offset { get; set; }
    }
}