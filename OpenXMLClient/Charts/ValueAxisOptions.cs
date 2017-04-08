using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLClient.Charts
{
    public class ValueAxisOptions
    {
        public uint Id { get; set; }
        public OrientationValues Orientation { get; set; }
        public MinMaxValues? MinMaxBounds { get; set; }
        public MinMaxValues? MinMaxUnits { get; set; }
        public AxisPositionValues AxisPosition { get; set; }
        public bool ShowMajorGridlines { get; set; }
        public MajorGridlinesOptions MajorGridlinesOptions { get; set; }
        public bool HasNumberingFormat { get; set; }
        public string FormatCode { get; set; }
        public bool SourceLinked { get; set; }
        public TickMarkValues MajorTickMark { get; set; }
        public TickMarkValues MinorTickMark { get; set; }
        public TickLabelPositionValues TickLabelPosition { get; set; }
        public bool HasChartShapeProperties { get; set; }
        public ChartShapePropertiesOptions ChartShapePropertiesOptions { get; set; }
        public bool IsTextStylish { get; set; }
        public uint CrossingAxisVal { get; set; }
        public CrossesValues? Crosses { get; set; }
        public double? CrossesAtVal { get; set; }
        public CrossBetweenValues CrossBetween { get; set; }
    }
}
