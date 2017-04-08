using DocumentFormat.OpenXml.Drawing;
using System;

namespace OpenXMLClient.Charts
{
    public class SchemaColorBuilder
    {
        public SchemaColorBuilder(SchemaColorOptions options)
        {
            this.Options = options;
        }

        public SchemaColorOptions Options { get; private set; }

        public SchemeColor Build()
        {
            var schemaColor = new SchemeColor();

            schemaColor.Val = this.Options.SchemeColor;

            decimal decTint = this.Options.Tint;

            // we don't have to do anything extra if the tint's zero.
            if (decTint < 0.0m)
            {
                decTint += 1.0m;
                decTint *= 100000m;
                schemaColor.Append(new LuminanceModulation() { Val = Convert.ToInt32(decTint) });
            }
            else if (decTint > 0.0m)
            {
                decTint *= 100000m;
                decTint = decimal.Floor(decTint);
                schemaColor.Append(new LuminanceModulation() { Val = Convert.ToInt32(100000m - decTint) });
                schemaColor.Append(new LuminanceOffset() { Val = Convert.ToInt32(decTint) });
            }

            var alpha = CalculateAlpha();

            if (alpha < 100000)
            {
                schemaColor.Append(new Alpha() { Val = alpha });
            }

            return schemaColor;
        }

        private int CalculateAlpha()
        {
            if (this.Options.Transparency > 100m) this.Options.Transparency = 100m;
            if (this.Options.Transparency < 0m) this.Options.Transparency = 0m;
            return Convert.ToInt32((100m - this.Options.Transparency) * 1000m);
        }
    }
}
