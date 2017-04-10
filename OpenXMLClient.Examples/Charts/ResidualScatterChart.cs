using OpenXMLClient.Charts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLClient.Examples.Charts
{
    public static class ResidualScatterChart
    {
        private static uint XId = 1;
        private static uint YId = 2;
        private static string FormatCode = "General";

        public static ScatterChartOptions GetScatterChartOptions(uint id)
        {
            var scatterChartOptions = new ScatterChartOptions();

            scatterChartOptions.Id = id;

            scatterChartOptions.Name = "Residuals";

            scatterChartOptions.ShowLegend = false;

            scatterChartOptions.PlotVisibleOnly = true;

            scatterChartOptions.DisplayBlanksAs = DisplayBlanksAsValues.Zero;

            scatterChartOptions.ShowDataLabelsOverMaximum = false;

            scatterChartOptions.ScatterStyle = ScatterStyleValues.SmoothMarker;

            scatterChartOptions.VaryColors = false;

            scatterChartOptions.Index = 0;

            scatterChartOptions.Order = 0;

            scatterChartOptions.XNumberReferenceOptions = GetXNumberReferenceOptions();

            scatterChartOptions.YNumberReferenceOptions = GetYNumberReferenceOptions();

            scatterChartOptions.XValueAxisOptions = GetXValueAxisOptions();

            scatterChartOptions.YValueAxisOptions = GetYValueAxisOptions();

            scatterChartOptions.FromMarkerOptions = GetFromMarkerOptions();

            scatterChartOptions.ToMarkerOptions = GetToMarkerOptions();

            scatterChartOptions.GraphicFrameOptions = GetGraphicFrameOptions();

            scatterChartOptions.ChartShapeOptions = GetChartShapeOptions();

            scatterChartOptions.ChartSpaceChartShapePropertiesOptions = GetChartShapePropertiesOptions();

            scatterChartOptions.PlotChartShapePropertiesOptions = GetPlotChartShapePropertiesOptions();

            scatterChartOptions.ScatterChartSeriesOptions = GetScatterChartSeriesOptions();
            
            return scatterChartOptions;
        }

        private static NumberReferenceOptions GetXNumberReferenceOptions()
        {
            var numberReferenceOptions = new NumberReferenceOptions
            {
                Id = XId,
                Letter = "E",
                RowStart = 11,
                FormatCode = FormatCode
            };

            return numberReferenceOptions;
        }

        private static NumberReferenceOptions GetYNumberReferenceOptions()
        {
            var numberReferenceOptions = new NumberReferenceOptions
            {
                Id = YId,
                Letter = "A",
                RowStart = 11,
                FormatCode = FormatCode
            };

            return numberReferenceOptions;
        }

        private static ValueAxisOptions GetXValueAxisOptions()
        {
            var valueAxisOptions = new ValueAxisOptions
            {
                Id = XId,
                Orientation = OrientationValues.MinMax,
                AxisPosition = AxisPositionValues.Top,
                ShowMajorGridlines = true,
                MajorGridlinesOptions = new MajorGridlinesOptions
                {
                    HasChartShapeProperties = true,
                    ChartShapePropertiesOptions = new ChartShapePropertiesOptions
                    {
                        BlackWhiteMode = BlackWhiteModeValues.Auto,
                        FillType = null,
                        HasLine = true,
                        OutlineOptions = new OutlineOptions
                        {
                            FillType = FillTypes.Solid,
                            SchemaColorOptions = new SchemaColorOptions
                            {
                                SchemeColor = SchemeColorValues.Text1,
                                Tint = 0.85M,
                                Transparency = 0
                            },
                            JoinType = JoinTypes.Round,
                            Width = 0.75M,
                            LineCap = LineCapValues.Flat,
                            CompoundLine = CompoundLineValues.Single,
                            PenAlignment = PenAlignmentValues.Center
                        },
                        HasEffectList = true,
                        EffectListOptions = new EffectListOptions()
                    }
                },
                HasNumberingFormat = true,
                FormatCode = FormatCode,
                SourceLinked = true,
                MajorTickMark = TickMarkValues.Outside,
                MinorTickMark = TickMarkValues.None,
                TickLabelPosition = TickLabelPositionValues.Low,
                HasChartShapeProperties = true,
                ChartShapePropertiesOptions = new ChartShapePropertiesOptions
                {
                    BlackWhiteMode = BlackWhiteModeValues.Auto,
                    FillType = FillTypes.NoFill,
                    HasLine = true,
                    OutlineOptions = new OutlineOptions
                    {
                        FillType = FillTypes.Solid,
                        SchemaColorOptions = new SchemaColorOptions
                        {
                            SchemeColor = SchemeColorValues.Text1,
                            Tint = 0.85M,
                            Transparency = 0
                        },
                        JoinType = JoinTypes.Round,
                        Width = 0.75M,
                        LineCap = LineCapValues.Flat,
                        CompoundLine = CompoundLineValues.Single,
                        PenAlignment = PenAlignmentValues.Center
                    },
                    HasEffectList = true,
                    EffectListOptions = new EffectListOptions()
                },
                IsTextStylish = true,
                CrossingAxisVal = YId,
                Crosses = CrossesValues.AutoZero,
                CrossesAtVal = null,
                CrossBetween = CrossBetweenValues.MidpointCategory
            };

            return valueAxisOptions;
        }

        private static ValueAxisOptions GetYValueAxisOptions()
        {
            var valueAxisOptions = new ValueAxisOptions
            {
                Id = YId,
                Orientation = OrientationValues.MaxMin,
                AxisPosition = AxisPositionValues.Left,
                ShowMajorGridlines = true,
                MajorGridlinesOptions = new MajorGridlinesOptions
                {
                    HasChartShapeProperties = true,
                    ChartShapePropertiesOptions = new ChartShapePropertiesOptions
                    {
                        BlackWhiteMode = BlackWhiteModeValues.Auto,
                        FillType = null,
                        HasLine = true,
                        OutlineOptions = new OutlineOptions
                        {
                            FillType = FillTypes.Solid,
                            SchemaColorOptions = new SchemaColorOptions
                            {
                                SchemeColor = SchemeColorValues.Text1,
                                Tint = 0.85M,
                                Transparency = 0
                            },
                            JoinType = JoinTypes.Round,
                            Width = 0.75M,
                            LineCap = LineCapValues.Flat,
                            CompoundLine = CompoundLineValues.Single,
                            PenAlignment = PenAlignmentValues.Center
                        },
                        HasEffectList = true,
                        EffectListOptions = new EffectListOptions()
                    }
                },
                HasNumberingFormat = true,
                FormatCode = FormatCode,
                SourceLinked = true,
                MajorTickMark = TickMarkValues.None,
                MinorTickMark = TickMarkValues.None,
                TickLabelPosition = TickLabelPositionValues.Low,
                HasChartShapeProperties = true,
                ChartShapePropertiesOptions = new ChartShapePropertiesOptions
                {
                    BlackWhiteMode = BlackWhiteModeValues.Auto,
                    FillType = FillTypes.NoFill,
                    HasLine = true,
                    OutlineOptions = new OutlineOptions
                    {
                        FillType = FillTypes.NoFill,
                        SchemaColorOptions = null,
                        JoinType = null,
                        Width = null,
                        LineCap = null,
                        CompoundLine = null,
                        PenAlignment = null
                    },
                    HasEffectList = true,
                    EffectListOptions = new EffectListOptions()
                },
                IsTextStylish = true,
                CrossingAxisVal = XId,
                Crosses = CrossesValues.AutoZero,
                CrossesAtVal = null,
                CrossBetween = CrossBetweenValues.MidpointCategory
            };

            return valueAxisOptions;
        }

        private static MarkerTypeOptions GetFromMarkerOptions()
        {
            var markerOptions = new MarkerTypeOptions
            {
                RowId = "9",
                RowOffset = "0",
                ColumnId = "5",
                ColumnOffset = "0"
            };

            return markerOptions;
        }

        private static MarkerTypeOptions GetToMarkerOptions()
        {
            var markerOptions = new MarkerTypeOptions
            {
                RowId = string.Empty,
                RowOffset = (MarkerTypeOptions.TotalEnglishMetricUnits(29) / 4).ToString(),
                ColumnId = "9",
                ColumnOffset = "0"
            };

            return markerOptions;
        }

        private static GraphicFrameOptions GetGraphicFrameOptions()
        {
            var graphicFrameOptions = new GraphicFrameOptions
            {
                TransformOffset = new OpenXMLClient.Charts.Point(0, 0),
                TransformExtents = new OpenXMLClient.Charts.Point(0, 0)
            };

            return graphicFrameOptions;
        }

        private static ChartShapeOptions GetChartShapeOptions()
        {
            var chartShapeOptions = new ChartShapeOptions
            {
                RoundedCorners = true
            };

            return chartShapeOptions;
        }

        private static ChartShapePropertiesOptions GetChartShapePropertiesOptions()
        {
            var chartShapePropertiesOptions = new ChartShapePropertiesOptions
            {
                BlackWhiteMode = null,
                FillType = FillTypes.Solid,
                SchemaColorOptions = new SchemaColorOptions
                {
                    SchemeColor = SchemeColorValues.Background1,
                    Tint = 0
                },
                HasLine = true,
                OutlineOptions = new OutlineOptions
                {
                    FillType = FillTypes.Solid,
                    SchemaColorOptions = new SchemaColorOptions
                    {
                        SchemeColor = SchemeColorValues.Text1,
                        Tint = 0.85M
                    },
                    JoinType = JoinTypes.Round,
                    Width = 0.75M,
                    LineCap = LineCapValues.Flat,
                    CompoundLine = CompoundLineValues.Single,
                    PenAlignment = PenAlignmentValues.Center
                },
                HasEffectList = true,
                EffectListOptions = new EffectListOptions()
            };

            return chartShapePropertiesOptions;
        }

        private static ChartShapePropertiesOptions GetPlotChartShapePropertiesOptions()
        {
            var chartShapePropertiesOptions = new ChartShapePropertiesOptions
            {
                BlackWhiteMode = null,
                FillType = FillTypes.NoFill,
                HasLine = true,
                OutlineOptions = new OutlineOptions
                {
                    FillType = FillTypes.NoFill,
                    JoinType = null,
                    Width = null,
                    LineCap = null,
                    CompoundLine = null,
                    PenAlignment = null
                },
                HasEffectList = true,
                EffectListOptions = new EffectListOptions()
            };

            return chartShapePropertiesOptions;
        }

        private static ScatterChartSeriesOptions GetScatterChartSeriesOptions()
        {
            var scatterChartSeriesOptions = new ScatterChartSeriesOptions
            {
                MarkerOptions = new MarkerOptions
                {
                    Symbol = MarkerStyleValues.Circle,
                    Size = 5,
                    ChartShapePropertiesOptions = null
                },
                Smooth = true
            };

            return scatterChartSeriesOptions;
        }
    }
}
