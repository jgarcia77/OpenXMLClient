using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLClient.Charts
{
    public abstract class ChartOptions
    {
        public uint Id { get; set; }
        public string Name { get; set; }
        public bool ShowLegend { get; set; }
        public bool PlotVisibleOnly { get; set; }
        public bool ShowDataLabelsOverMaximum { get; set; }
        public DisplayBlanksAsValues DisplayBlanksAs { get; set; }
        public ChartShapeOptions ChartShapeOptions { get; set; }
        public ChartShapePropertiesOptions ChartSpaceChartShapePropertiesOptions { get; set; }
        public ChartShapePropertiesOptions PlotChartShapePropertiesOptions { get; set; }
        public MarkerTypeOptions FromMarkerOptions { get; set; }
        public MarkerTypeOptions ToMarkerOptions { get; set; }
        public GraphicFrameOptions GraphicFrameOptions { get; set; }
    }
}
