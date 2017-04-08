using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;

namespace OpenXMLClient.Charts
{
    public class ScatterChartBuilder : ChartBuilder
    {
        private ScatterChart scatterChart { get; set; }
        private ScatterChartSeries scatterChartSeries { get; set; }

        public ScatterChartBuilder(WorksheetPart worksheetPart, ScatterChartOptions options) : base(options, worksheetPart)
        {
            this.Options = options;
        }

        public ScatterChartOptions Options { get; private set; }

        public override void Build()
        {
            this.AppendChart();
            
            this.AppendValues();

            chartPart.ChartSpace.Save();

            base.AppendTwoCellAnchor();

            drawingsPart.WorksheetDrawing.Save();
        }

        protected override void AppendChart()
        {
            scatterChart = plotArea.AppendChild<ScatterChart>(new ScatterChart());

            scatterChart.ScatterStyle = new ScatterStyle { Val = this.Options.ScatterStyle };

            scatterChart.VaryColors = new VaryColors { Val = this.Options.VaryColors };

            scatterChartSeries = scatterChart.AppendChild<ScatterChartSeries>(new ScatterChartSeries());

            scatterChartSeries.Index = new Index { Val = this.Options.Index };

            scatterChartSeries.Order = new Order { Val = this.Options.Order };

            this.AppendOptions();
        }

        protected override void AppendValues()
        {
            this.AppendXValues();

            this.AppendYValues();

            this.AppendAxisId();

            this.AppendValueAxisX();

            this.AppendValueAxisY();
        }

        private void AppendXValues()
        {
            var XValues = new XValues();

            var numberReferenceBuilder = new NumberReferenceBuilder(this.Options.XNumberReferenceOptions);

            XValues.NumberReference = numberReferenceBuilder.Build();

            scatterChartSeries.Append(XValues);
        }

        private void AppendYValues()
        {
            var yValues = new YValues();

            var numberReferenceBuilder = new NumberReferenceBuilder(this.Options.YNumberReferenceOptions);

            yValues.NumberReference = numberReferenceBuilder.Build();

            scatterChartSeries.Append(yValues);
        }

        private void AppendAxisId()
        {
            scatterChart.Append(new AxisId() { Val = this.Options.XNumberReferenceOptions.Id });

            scatterChart.Append(new AxisId() { Val = this.Options.YNumberReferenceOptions.Id });
        }

        private void AppendValueAxisX()
        {
            var builder = new ValueAxisBuilder(this.Options.XValueAxisOptions);

            var valueAxis = builder.Build();

            plotArea.Append(valueAxis);
        }

        private void AppendValueAxisY()
        {
            var builder = new ValueAxisBuilder(this.Options.YValueAxisOptions);

            var valueAxis = builder.Build();

            plotArea.Append(valueAxis);
        }

        private void AppendOptions()
        {
            this.AppendScatterChartSeriesOptions();
        }
        
        private void AppendScatterChartSeriesOptions()
        {
            if (this.Options.ScatterChartSeriesOptions != null)
            {
                if (this.Options.ScatterChartSeriesOptions.MarkerOptions != null)
                {
                    var builder = new MarkerBuilder(this.Options.ScatterChartSeriesOptions.MarkerOptions);

                    scatterChartSeries.Marker = builder.Build();

                    scatterChartSeries.Append(new Smooth { Val = this.Options.ScatterChartSeriesOptions.Smooth });
                }
            }
        }
    }
}
