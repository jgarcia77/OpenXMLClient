
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLClient.Charts
{
    public class MarkerOptions
    {
        public MarkerStyleValues? Symbol { get; set; }
        public byte? Size { get; set; }

        public ChartShapePropertiesOptions ChartShapePropertiesOptions { get; set; }
    }
}
