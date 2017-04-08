using DocumentFormat.OpenXml.Spreadsheet;
using SS = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using System.Linq;
using DocumentFormat.OpenXml;

namespace OpenXMLClient.Charts
{
    public class ChartBuilder
    {
        private const string NamespaceSpreadsheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        private const string NamespaceMain = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string NamespaceChart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string NamespaceRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private ChartOptions Options { get; set; }

        protected DrawingsPart drawingsPart { get; set; }
        protected ChartPart chartPart { get; set; }
        protected C.Chart chart { get; set; }
        protected C.PlotArea plotArea { get; set; }

        public ChartBuilder(ChartOptions options, WorksheetPart worksheetPart)
        {
            this.Options = options;

            if (worksheetPart.DrawingsPart == null)
            {
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            }
            else
            {
                drawingsPart = worksheetPart.DrawingsPart;
            }

            if (drawingsPart.WorksheetDrawing == null)
            {
                // WorksheetDrawing is used to position the chart on the worksheet
                drawingsPart.WorksheetDrawing = new SS.WorksheetDrawing();

                drawingsPart.WorksheetDrawing.AddNamespaceDeclaration("xdr", NamespaceSpreadsheetDrawing);

                drawingsPart.WorksheetDrawing.AddNamespaceDeclaration("a", NamespaceMain);

                var drawing = new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) };

                worksheetPart.Worksheet.Append(drawing);

                worksheetPart.Worksheet.Save();
            }

            // Add a new chart part, chart space, and chart
            chartPart = drawingsPart.AddNewPart<ChartPart>();

            chartPart.ChartSpace = new C.ChartSpace();

            chartPart.ChartSpace.AddNamespaceDeclaration("c", NamespaceChart);

            chartPart.ChartSpace.AddNamespaceDeclaration("a", NamespaceMain);

            chartPart.ChartSpace.AddNamespaceDeclaration("r", NamespaceRelationships);

            chart = chartPart.ChartSpace.AppendChild<C.Chart>(new C.Chart());

            // Add a new plot area and layout
            plotArea = chart.AppendChild<C.PlotArea>(new C.PlotArea());

            plotArea.AppendChild<C.Layout>(new C.Layout());

            this.AppendOptions();
        }

        public virtual void Build()
        {
            /*
                this.AppendChart();

                this.AppendValues();

                chartPart.ChartSpace.Save();

                base.AppendTwoCellAnchor();

                drawingsPart.WorksheetDrawing.Save();
             */
        }

        protected virtual void AppendChart()
        {
            // Append a specific chart
        }

        protected virtual void AppendValues()
        {
            // Append specifiv axis and values
        }

        protected void AppendTwoCellAnchor()
        {
            var twoCellAnchor = new SS.TwoCellAnchor();

            twoCellAnchor.FromMarker = new SS.FromMarker
            {
                RowId = new SS.RowId(this.Options.FromMarkerOptions.RowId),
                RowOffset = new SS.RowOffset(this.Options.FromMarkerOptions.RowOffset),
                ColumnId = new SS.ColumnId(this.Options.FromMarkerOptions.ColumnId),
                ColumnOffset = new SS.ColumnOffset(this.Options.FromMarkerOptions.ColumnOffset)
            };

            twoCellAnchor.ToMarker = new SS.ToMarker
            {
                RowId = new SS.RowId(this.Options.ToMarkerOptions.RowId),
                RowOffset = new SS.RowOffset(this.Options.ToMarkerOptions.RowOffset),
                ColumnId = new SS.ColumnId(this.Options.ToMarkerOptions.ColumnId),
                ColumnOffset = new SS.ColumnOffset(this.Options.ToMarkerOptions.ColumnOffset)
            };

            twoCellAnchor.Append(CreateGraphicFrame());

            twoCellAnchor.Append(new SS.ClientData());

            drawingsPart.WorksheetDrawing.Append(twoCellAnchor);
        }

        private SS.GraphicFrame CreateGraphicFrame()
        {
            var graphicFrame = new SS.GraphicFrame();

            graphicFrame.Macro = string.Empty;

            graphicFrame.NonVisualGraphicFrameProperties = new SS.NonVisualGraphicFrameProperties();

            graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties = new SS.NonVisualDrawingProperties();

            graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Id = this.Options.Id;

            graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name = this.Options.Name;

            graphicFrame.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties = new SS.NonVisualGraphicFrameDrawingProperties();

            graphicFrame.Transform = new SS.Transform();

            graphicFrame.Transform.Offset = new Offset() { X = this.Options.GraphicFrameOptions.TransformOffset.X, Y = this.Options.GraphicFrameOptions.TransformOffset.Y };

            graphicFrame.Transform.Extents = new Extents() { Cx = this.Options.GraphicFrameOptions.TransformExtents.X, Cy = this.Options.GraphicFrameOptions.TransformExtents.Y };

            graphicFrame.Graphic = new Graphic();

            graphicFrame.Graphic.GraphicData = new GraphicData();

            graphicFrame.Graphic.GraphicData.Uri = NamespaceChart;

            graphicFrame.Graphic.GraphicData.Append(new C.ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) });

            return graphicFrame;
        }

        private void AppendOptions()
        {
            this.AppendChartOptions();

            this.AppendPlotOptions();

            this.AppendChartSpaceOptions();
        }

        private void AppendChartOptions()
        {
            chart.PlotVisibleOnly = new C.PlotVisibleOnly { Val = this.Options.PlotVisibleOnly };

            chart.DisplayBlanksAs = new C.DisplayBlanksAs { Val = this.Options.DisplayBlanksAs };

            chart.ShowDataLabelsOverMaximum = new C.ShowDataLabelsOverMaximum { Val = this.Options.ShowDataLabelsOverMaximum };
            
            chart.AutoTitleDeleted = new C.AutoTitleDeleted { Val = true };
        }

        private void AppendPlotOptions()
        {
            if (this.Options.PlotChartShapePropertiesOptions != null)
            {
                var builder = new ChartShapePropertiesBuilder(this.Options.PlotChartShapePropertiesOptions);

                plotArea.Append(builder.Build());
            }
        }

        private void AppendChartSpaceOptions()
        {
            chartPart.ChartSpace.RoundedCorners = new C.RoundedCorners { Val = this.Options.ChartShapeOptions.RoundedCorners };

            if (this.Options.ChartSpaceChartShapePropertiesOptions != null)
            {
                var builder = new ChartShapePropertiesBuilder(this.Options.ChartSpaceChartShapePropertiesOptions);

                chartPart.ChartSpace.Append(builder.Build());
            }
        }
    }
}
