using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLClient.Charts
{
    public class MajorGridlinesBuilder
    {
        public MajorGridlinesBuilder(MajorGridlinesOptions options)
        {
            this.Options = options;
        }

        public MajorGridlinesOptions Options { get; private set; }

        public MajorGridlines Build()
        {
            var majorGridlines = new MajorGridlines();

            if (this.Options.HasChartShapeProperties)
            {
                var chartShapePropertiesBuilder = new ChartShapePropertiesBuilder(this.Options.ChartShapePropertiesOptions);

                majorGridlines.ChartShapeProperties = chartShapePropertiesBuilder.Build();
            }

            return majorGridlines;
        }
    }
}
