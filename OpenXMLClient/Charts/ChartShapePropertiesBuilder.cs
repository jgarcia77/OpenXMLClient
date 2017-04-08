using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLClient.Charts
{
    public class ChartShapePropertiesBuilder
    {
        public ChartShapePropertiesBuilder(ChartShapePropertiesOptions options)
        {
            this.Options = options;
        }

        public ChartShapePropertiesOptions Options { get; private set; }

        public ChartShapeProperties Build()
        {
            var chartShapeProperties = new ChartShapeProperties();

            if (this.Options.BlackWhiteMode.HasValue)
            {
                chartShapeProperties.BlackWhiteMode = this.Options.BlackWhiteMode;
            }

            if (this.Options.FillType.HasValue)
            {
                switch (this.Options.FillType)
                {
                    case FillTypes.NoFill:
                        chartShapeProperties.Append(new NoFill());

                        break;
                }
            }

            if (this.Options.HasLine)
            {
                var outlineBuilder = new OutlineBuilder(this.Options.OutlineOptions);

                chartShapeProperties.Append(outlineBuilder.Build());
            }

            if (this.Options.HasEffectList)
            {
                var effectListBuilder = new EffectListBuilder(this.Options.EffectListOptions);

                chartShapeProperties.Append(effectListBuilder.Build());
            }

            return chartShapeProperties;
        }
    }
}
