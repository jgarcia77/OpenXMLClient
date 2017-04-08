using DocumentFormat.OpenXml.Drawing;
using System;

namespace OpenXMLClient.Charts
{
    public class OutlineBuilder
    {
        public OutlineBuilder(OutlineOptions options)
        {
            this.Options = options;
        }

        public OutlineOptions Options { get; private set; }

        public Outline Build()
        {
            var outline = new Outline();

            switch (this.Options.FillType)
            {
                case FillTypes.NoFill:
                    outline.Append(new NoFill());

                    break;

                case FillTypes.Solid:
                    var solidFill = new SolidFill();

                    if (this.Options.SchemaColorOptions != null)
                    {
                        var schemaColorBuilder = new SchemaColorBuilder(this.Options.SchemaColorOptions);

                        solidFill.SchemeColor = schemaColorBuilder.Build();
                    }

                    outline.Append(solidFill);

                    break;

                case FillTypes.Gradient:
                    break;
            }

            if (this.Options.JoinType.HasValue)
            {
                switch (this.Options.JoinType)
                {
                    case JoinTypes.Round:
                        outline.Append(new Round());
                        break;
                    case JoinTypes.Bevel:
                        outline.Append(new Bevel());
                        break;
                    case JoinTypes.Miter:
                        outline.Append(new Miter() { Limit = 800000 });
                        break;
                }
            }

            if (this.Options.Width.HasValue)
            {
                outline.Width = Convert.ToInt32(this.Options.Width * 12700);
            }

            if (this.Options.LineCap.HasValue)
            {
                outline.CapType = this.Options.LineCap;
            }

            if (this.Options.CompoundLine.HasValue)
            {
                outline.CompoundLineType = this.Options.CompoundLine;
            }

            if (this.Options.PenAlignment.HasValue)
            {
                outline.Alignment = this.Options.PenAlignment;
            }

            return outline;
        }
    }
}
