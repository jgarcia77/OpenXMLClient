using DocumentFormat.OpenXml.Drawing;

namespace OpenXMLClient.Charts
{
    public class OutlineOptions
    {
        public JoinTypes? JoinType { get; set; }
        public decimal? Width { get; set; }
        public LineCapValues? LineCap { get; set; }
        public CompoundLineValues? CompoundLine { get; set; }
        public PenAlignmentValues? PenAlignment { get; set; }
        public FillTypes FillType { get; set; }
        public SchemaColorOptions SchemaColorOptions { get; set; }
    }
}
