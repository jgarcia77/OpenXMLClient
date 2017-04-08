using DocumentFormat.OpenXml.Drawing;

namespace OpenXMLClient.Charts
{
    public class ChartShapePropertiesOptions
    {
        public BlackWhiteModeValues? BlackWhiteMode { get; set; }
        public FillTypes? FillType { get; set; }
        public SchemaColorOptions SchemaColorOptions { get; set; }
        public bool HasLine { get; set; }
        public OutlineOptions OutlineOptions { get; set; }
        public bool HasEffectList { get; set; }
        public EffectListOptions EffectListOptions { get; set; }
    }
}
