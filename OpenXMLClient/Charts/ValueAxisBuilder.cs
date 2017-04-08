using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLClient.Charts
{
    public class ValueAxisBuilder
    {
        public ValueAxisBuilder(ValueAxisOptions options)
        {
            this.Options = options;
        }

        public ValueAxisOptions Options { get; private set; }

        public ValueAxis Build()
        {
            var valueAxis = new ValueAxis();

            valueAxis.AxisId = new AxisId { Val = this.Options.Id };

            valueAxis.Scaling = new Scaling();

            valueAxis.Scaling.Orientation = new Orientation { Val = this.Options.Orientation };

            if (this.Options.MinMaxBounds.HasValue)
            {
                var minAxisValue = new MinAxisValue { Val = this.Options.MinMaxBounds.Value.MinValue };

                var maxAxisValue = new MaxAxisValue { Val = this.Options.MinMaxBounds.Value.MaxValue };

                valueAxis.Scaling.Append(minAxisValue);

                valueAxis.Scaling.Append(maxAxisValue);
            }

            if (this.Options.MinMaxUnits.HasValue)
            {
                var minorUnit = new MinorUnit { Val = this.Options.MinMaxUnits.Value.MinValue };

                var majorUnit = new MajorUnit { Val = this.Options.MinMaxUnits.Value.MaxValue };

                valueAxis.Append(minorUnit);

                valueAxis.Append(majorUnit);
            }

            valueAxis.Delete = new Delete { Val = false };

            valueAxis.AxisPosition = new AxisPosition { Val = this.Options.AxisPosition };

            if (this.Options.ShowMajorGridlines)
            {
                var majorGridlinesBuilder = new MajorGridlinesBuilder(this.Options.MajorGridlinesOptions);

                valueAxis.MajorGridlines = majorGridlinesBuilder.Build();
            }

            if (this.Options.HasNumberingFormat)
            {
                valueAxis.NumberingFormat = new NumberingFormat
                {
                    FormatCode = this.Options.FormatCode,
                    SourceLinked = this.Options.SourceLinked
                };
            }

            valueAxis.MajorTickMark = new MajorTickMark { Val = this.Options.MajorTickMark };

            valueAxis.MinorTickMark = new MinorTickMark { Val = this.Options.MinorTickMark };

            valueAxis.TickLabelPosition = new TickLabelPosition { Val = this.Options.TickLabelPosition };

            if (this.Options.HasChartShapeProperties)
            {
                var chartShapePropertiesBuilder = new ChartShapePropertiesBuilder(this.Options.ChartShapePropertiesOptions);

                valueAxis.ChartShapeProperties = chartShapePropertiesBuilder.Build();
            }

            if (this.Options.IsTextStylish)
            {
                valueAxis.TextProperties = new TextProperties();
                valueAxis.TextProperties.BodyProperties = new BodyProperties()
                {
                    Rotation = -60000000,
                    UseParagraphSpacing = true,
                    VerticalOverflow = TextVerticalOverflowValues.Ellipsis,
                    Vertical = TextVerticalValues.Horizontal,
                    Wrap = TextWrappingValues.Square,
                    Anchor = TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                valueAxis.TextProperties.ListStyle = new ListStyle();

                var para = new Paragraph();
                para.ParagraphProperties = new ParagraphProperties();

                var defrunprops = new DefaultRunProperties();
                defrunprops.FontSize = 900;
                defrunprops.Bold = false;
                defrunprops.Italic = false;
                defrunprops.Underline = TextUnderlineValues.None;
                defrunprops.Strike = TextStrikeValues.NoStrike;
                defrunprops.Kerning = 1200;
                defrunprops.Baseline = 0;

                var schclr = new SchemeColor() { Val = SchemeColorValues.Text1 };
                schclr.Append(new LuminanceModulation() { Val = 65000 });
                schclr.Append(new LuminanceOffset() { Val = 35000 });
                defrunprops.Append(new SolidFill()
                {
                    SchemeColor = schclr
                });

                defrunprops.Append(new LatinFont() { Typeface = "+mn-lt" });
                defrunprops.Append(new EastAsianFont() { Typeface = "+mn-ea" });
                defrunprops.Append(new ComplexScriptFont() { Typeface = "+mn-cs" });

                para.ParagraphProperties.Append(defrunprops);
                para.Append(new EndParagraphRunProperties() { Language = System.Globalization.CultureInfo.CurrentCulture.Name });

                valueAxis.TextProperties.Append(para);
            }
            
            valueAxis.CrossingAxis = new CrossingAxis { Val = this.Options.CrossingAxisVal };

            if (this.Options.Crosses.HasValue)
            {
                valueAxis.Append(new Crosses { Val = this.Options.Crosses });
            }
            else if (this.Options.CrossesAtVal.HasValue)
            {
                valueAxis.Append(new CrossesAt() { Val = this.Options.CrossesAtVal });
            }

            valueAxis.Append(new CrossBetween { Val = this.Options.CrossBetween });

            return valueAxis;
        }
    }
}
