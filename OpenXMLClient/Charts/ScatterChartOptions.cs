using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLClient.Charts
{
    public class ScatterChartOptions : ChartOptions
    {
        public ScatterStyleValues ScatterStyle { get; set; }
        public bool VaryColors { get; set; }
        public uint Index { get; set; }
        public uint Order { get; set; }
        public NumberReferenceOptions XNumberReferenceOptions { get; set; }
        public NumberReferenceOptions YNumberReferenceOptions { get; set; }
        public ValueAxisOptions XValueAxisOptions { get; set; }
        public ValueAxisOptions YValueAxisOptions { get; set; }
        public ScatterChartSeriesOptions ScatterChartSeriesOptions { get; set; }
    }
}
