using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace OpenXMLClient.Charts
{
    public class MarkerBuilder
    {
        private MarkerOptions Options { get; set; }

        public MarkerBuilder(MarkerOptions options)
        {
            this.Options = options;
        }

        public Marker Build()
        {
            var marker = new Marker();

            if (this.Options.Symbol.HasValue)
            {
                marker.Symbol = new Symbol { Val = this.Options.Symbol };
            }

            if (this.Options.Size.HasValue)
            {
                marker.Size = new Size { Val = this.Options.Size };
            }

            if (this.Options.ChartShapePropertiesOptions != null)
            {
                var builder = new ChartShapePropertiesBuilder(this.Options.ChartShapePropertiesOptions);

                marker.ChartShapeProperties = builder.Build();
            }

            return marker;
        }
    }
}
