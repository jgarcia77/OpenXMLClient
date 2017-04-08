namespace OpenXMLClient.Charts
{
    public class MarkerTypeOptions
    {
        private static readonly int EMUsPerInch = 914400;
        private static readonly int PixelsPerInch = 96;

        public string RowId { get; set; }
        public string RowOffset { get; set; }
        public string ColumnId { get; set; }
        public string ColumnOffset { get; set; }

        public static int TotalEnglishMetricUnits(int pixels)
        {
            return pixels * EMUsPerInch / PixelsPerInch;
        }
    }
}
