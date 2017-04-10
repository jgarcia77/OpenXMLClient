namespace OpenXMLClient.Excel
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
    using A = DocumentFormat.OpenXml.Drawing;
    using OpenXMLClient.Common;
    using System.Linq;

    public class WorksheetBase
    {
        private const int maxColumns = 16384;
        protected string[] ColumnHeaderNames = new string[maxColumns];

        protected WorkbookPart WorkbookPart { get; set; }
        protected WorksheetPart WorksheetPart { get; private set; }
        protected DrawingsPart DrawingsPart { get; private set; }
        protected ImagePart ImagePart { get; private set; }

        protected List<int> RowsToExcludeMaxCharacters { get; private set; }
        protected int CurrentRowNumber { get; set; }
        protected int CurrentColumnIndex { get; set; }
        protected Hyperlinks Hyperlinks { get; set; }
        protected MergeCells MergeCells { get; private set; }

        public WorksheetBase(WorkbookPart workbookPart, int sequence, bool multipleReports)
        {
            WorkbookPart = workbookPart;
            Sequence = sequence;
            MultipleReports = multipleReports;
            RowsToExcludeMaxCharacters = new List<int>();
            Hyperlinks = new Hyperlinks();
            MergeCells = new MergeCells();
            InitColumnHeaderNames();
        }

        public bool MultipleReports { get; protected set; }
        public int Sequence { get; protected set; }

        public virtual void Append() { }

        protected virtual void InitWorksheetPart() { }

        protected void AddWorksheetPart(string id)
        {
            WorksheetPart = WorkbookPart.AddNewPart<WorksheetPart>(string.Concat("Sequence", Sequence, "_", id));
        }

        protected void AddDrawingsPart(string id)
        {
            DrawingsPart = WorksheetPart.AddNewPart<DrawingsPart>(id);
        }

        protected void AddImagePart(string id, string base64Image)
        {
            ImagePart = DrawingsPart.AddNewPart<ImagePart>("image/tiff", id);
            StreamImagePart(base64Image);
        }

        protected void AppendCells(Row row, int numberOfCells, uint styleIndex, uint? lastStyleIndex = null)
        {
            if (!lastStyleIndex.HasValue)
            {
                lastStyleIndex = styleIndex;
            }

            for (int cellCount = 1; cellCount <= numberOfCells; cellCount++)
            {
                CurrentColumnIndex++;

                if (cellCount == numberOfCells)
                {
                    styleIndex = lastStyleIndex.Value;
                }

                var cellReference = string.Concat(ColumnHeaderNames[CurrentColumnIndex], CurrentRowNumber);

                row.Append(new Cell() { CellReference = cellReference, StyleIndex = styleIndex });
            }
        }

        protected virtual void InitDrawingsPart(int logoToMarkerColOffset = 824661, 
            int titleFromMarkerColumn = 7,
            int titleToMarkerColumn = 9,
            int titleToMarkerColOffset = 285450)
        {
            Xdr.WorksheetDrawing worksheetDrawing = new Xdr.WorksheetDrawing();
            worksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = "0";
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "0";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = "0";
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "82404";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = "2";
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = logoToMarkerColOffset.ToString();
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = "2";
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "126030";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.Picture picture1 = new Xdr.Picture();

            Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 3" };

            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Xdr.BlipFill blipFill1 = new Xdr.BlipFill() { RotateWithShape = true };

            A.Blip blip1 = new A.Blip() { Embed = "rId1" };
            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            A.SourceRectangle sourceRectangle1 = new A.SourceRectangle() { Top = 32250, Bottom = 34913 };
            A.Stretch stretch1 = new A.Stretch();

            blipFill1.Append(blip1);
            blipFill1.Append(sourceRectangle1);
            blipFill1.Append(stretch1);

            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 82404L };
            A.Extents extents1 = new A.Extents() { Cx = 1827961L, Cy = 450026L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(picture1);
            twoCellAnchor1.Append(clientData1);

            Xdr.TwoCellAnchor twoCellAnchor2 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

            Xdr.FromMarker fromMarker2 = new Xdr.FromMarker();
            Xdr.ColumnId columnId3 = new Xdr.ColumnId();
            columnId3.Text = titleFromMarkerColumn.ToString();
            Xdr.ColumnOffset columnOffset3 = new Xdr.ColumnOffset();
            columnOffset3.Text = "685800";
            Xdr.RowId rowId3 = new Xdr.RowId();
            rowId3.Text = "0";
            Xdr.RowOffset rowOffset3 = new Xdr.RowOffset();
            rowOffset3.Text = "114300";

            fromMarker2.Append(columnId3);
            fromMarker2.Append(columnOffset3);
            fromMarker2.Append(rowId3);
            fromMarker2.Append(rowOffset3);

            Xdr.ToMarker toMarker2 = new Xdr.ToMarker();
            Xdr.ColumnId columnId4 = new Xdr.ColumnId();
            columnId4.Text = titleToMarkerColumn.ToString();
            Xdr.ColumnOffset columnOffset4 = new Xdr.ColumnOffset();
            columnOffset4.Text = titleToMarkerColOffset.ToString();
            Xdr.RowId rowId4 = new Xdr.RowId();
            rowId4.Text = "2";
            Xdr.RowOffset rowOffset4 = new Xdr.RowOffset();
            rowOffset4.Text = "88900";

            toMarker2.Append(columnId4);
            toMarker2.Append(columnOffset4);
            toMarker2.Append(rowId4);
            toMarker2.Append(rowOffset4);

            Xdr.Shape shape1 = new Xdr.Shape() { Macro = "", TextLink = "" };

            Xdr.NonVisualShapeProperties nonVisualShapeProperties1 = new Xdr.NonVisualShapeProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)8U, Name = "TextBox 7" };
            Xdr.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new Xdr.NonVisualShapeDrawingProperties() { TextBox = true };

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);

            Xdr.ShapeProperties shapeProperties2 = new Xdr.ShapeProperties();

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset2 = new A.Offset() { X = 6769100L, Y = 114300L };
            A.Extents extents2 = new A.Extents() { Cx = 1631650L, Cy = 381000L };

            transform2D2.Append(offset2);
            transform2D2.Append(extents2);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill1.Append(schemeColor1);

            A.Outline outline1 = new A.Outline() { Width = 9525, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill1 = new A.NoFill();

            outline1.Append(noFill1);

            shapeProperties2.Append(transform2D2);
            shapeProperties2.Append(presetGeometry2);
            shapeProperties2.Append(solidFill1);
            shapeProperties2.Append(outline1);

            Xdr.ShapeStyle shapeStyle1 = new Xdr.ShapeStyle();

            A.LineReference lineReference1 = new A.LineReference() { Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage1 = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference1.Append(rgbColorModelPercentage1);

            A.FillReference fillReference1 = new A.FillReference() { Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage2 = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference1.Append(rgbColorModelPercentage2);

            A.EffectReference effectReference1 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage3 = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference1.Append(rgbColorModelPercentage3);

            A.FontReference fontReference1 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference1.Append(schemeColor2);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);

            Xdr.TextBody textBody1 = new Xdr.TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties() { VerticalOverflow = A.TextVerticalOverflowValues.Clip, HorizontalOverflow = A.TextHorizontalOverflowValues.Clip, Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.Run run1 = new A.Run();

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", FontSize = 1600, Bold = true };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill2.Append(schemeColor3);

            runProperties1.Append(solidFill2);
            A.Text text1 = new A.Text();
            text1.Text = "Deloitte Reveal";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties2);
            shape1.Append(shapeStyle1);
            shape1.Append(textBody1);
            Xdr.ClientData clientData2 = new Xdr.ClientData();

            twoCellAnchor2.Append(fromMarker2);
            twoCellAnchor2.Append(toMarker2);
            twoCellAnchor2.Append(shape1);
            twoCellAnchor2.Append(clientData2);

            worksheetDrawing.Append(twoCellAnchor1);
            worksheetDrawing.Append(twoCellAnchor2);

            DrawingsPart.WorksheetDrawing = worksheetDrawing;
        }

        protected InlineString FormatMessage(string value, string cellReference)
        {
            var returnObject = new InlineString();

            var anchors = Anchor.Find(value);

            if (!anchors.Any())
            {
                returnObject.Append(new Run(new Text(value.TrimEnd(Environment.NewLine.ToCharArray())) { Space = SpaceProcessingModeValues.Preserve }));
            }
            else
            {
                // Use the first instance because only one hyperlink per cell is allowed
                var anchor = anchors[0];

                var hyperLinkId = string.Concat(cellReference, "HyperLink");

                WorksheetPart.AddHyperlinkRelationship(new System.Uri(anchor.Href, System.UriKind.Absolute), true, hyperLinkId);

                Hyperlinks.Append(new Hyperlink() { Reference = cellReference, Id = hyperLinkId });

                // The entire cell will be a hyperlink but only the anchor text will look like a link
                var valueWithoutAnchor = value.Replace(anchor.Value, string.Empty).TrimEnd('.').TrimEnd(Environment.NewLine.ToCharArray());

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                FontSize fontSize3 = new FontSize() { Val = 11D };
                Color color3 = new Color() { Theme = (UInt32Value)1U };
                RunFont runFont1 = new RunFont() { Val = "Calibri" };
                FontFamily fontFamily1 = new FontFamily() { Val = 2 };
                FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

                runProperties1.Append(fontSize3);
                runProperties1.Append(color3);
                runProperties1.Append(runFont1);
                runProperties1.Append(fontFamily1);
                runProperties1.Append(fontScheme4);
                Text text1 = new Text();
                text1.Text = valueWithoutAnchor;
                text1.Space = SpaceProcessingModeValues.Preserve;

                run1.Append(runProperties1);
                run1.Append(text1);

                Run run2 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Underline underline2 = new Underline();
                FontSize fontSize4 = new FontSize() { Val = 11D };
                Color color4 = new Color() { Theme = (UInt32Value)10U };
                RunFont runFont2 = new RunFont() { Val = "Calibri" };
                FontFamily fontFamily2 = new FontFamily() { Val = 2 };
                FontScheme fontScheme5 = new FontScheme() { Val = FontSchemeValues.Minor };

                runProperties2.Append(underline2);
                runProperties2.Append(fontSize4);
                runProperties2.Append(color4);
                runProperties2.Append(runFont2);
                runProperties2.Append(fontFamily2);
                runProperties2.Append(fontScheme5);
                Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text2.Text = string.Concat(anchor.Text, ".");

                run2.Append(runProperties2);
                run2.Append(text2);

                returnObject.Append(run1);
                returnObject.Append(run2);
            }

            return returnObject;
        }

        protected DoubleValue CalculateRowHeight(string value, int charactersPerLine, int unitWidthPerLine)
        {
            decimal numberOfLines = 0;

            if (value.IndexOf(Environment.NewLine) == -1)
            {
                numberOfLines = Math.Ceiling((decimal)value.Length / (decimal)charactersPerLine);
            }
            else
            {
                var linesArray = string.IsNullOrWhiteSpace(value) ? null : value.Replace(Environment.NewLine, "|").Split('|');

                if (linesArray != null)
                {
                    foreach (var line in linesArray)
                    {
                        if (string.IsNullOrEmpty(line))
                        {
                            numberOfLines++;
                        }
                        else
                        {
                            numberOfLines += Math.Ceiling((decimal)line.Length / (decimal)charactersPerLine);
                        }
                    }
                }
            }

            if (numberOfLines == 0)
            {
                numberOfLines = 1;
            }
                        
            return (DoubleValue)(unitWidthPerLine * (int)numberOfLines);
        }

        protected Column CalculateColumnWidth(KeyValuePair<int, int> item, int minWidth = 13)
        {
            //this is the width of the font
            double maxWidth = 4;

            double width = Math.Ceiling(Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256) + 1;

            if (width < minWidth)
            {
                width = minWidth;
            }

            double pixels = Math.Truncate(((256 * width + Math.Truncate(128 / maxWidth)) / 256) * maxWidth);
            
            double charWidth = Math.Truncate((pixels - 5) / maxWidth * 100 + 0.5) / 100;

            var column = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width };

            return column;
        }

        protected Dictionary<int, int> GetColumnMaxCharacters(SheetData sheetData)
        {
            //iterate over all cells getting a max char value for each column
            Dictionary<int, int> returnCollection = new Dictionary<int, int>();
            var rows = sheetData.Elements<Row>();
            UInt32[] numberStyles = new UInt32[] { 5, 6, 7, 8 }; //styles that will add extra chars
            UInt32[] boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8 }; //styles that will bold

            var rowNumber = 0;

            foreach (var r in rows)
            {
                rowNumber++;

                if (RowsToExcludeMaxCharacters.Contains(rowNumber))
                {
                    continue;
                }

                var cells = r.Elements<Cell>().ToArray();

                //using cell index as my column
                for (int i = 0; i < cells.Length; i++)
                {
                    var cell = cells[i];
                    var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;

                    var cellTextLength = cellValue.Length;

                    if (cellValue.IndexOf(Environment.NewLine) != -1)
                    {
                        cellTextLength = 0;

                        var values = cellValue.Replace(Environment.NewLine, "|").Split('|');

                        foreach (var value in values)
                        {
                            var valueLength = value.Length;

                            if (valueLength > cellTextLength)
                            {
                                cellTextLength = valueLength;
                            }
                        }
                    }

                    if (cell.StyleIndex != null && numberStyles.Contains(cell.StyleIndex))
                    {
                        int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                        //add 3 for '.00' 
                        cellTextLength += (3 + thousandCount);
                    }

                    if (cell.StyleIndex != null && boldStyles.Contains(cell.StyleIndex))
                    {
                        //add an extra char for bold - not 100% acurate but good enough for what i need.
                        cellTextLength += 1;
                    }

                    if (returnCollection.ContainsKey(i))
                    {
                        var current = returnCollection[i];
                        if (cellTextLength > current)
                        {
                            returnCollection[i] = cellTextLength;
                        }
                    }
                    else
                    {
                        returnCollection.Add(i, cellTextLength);
                    }
                }
            }

            return returnCollection;
        }

        private void StreamImagePart(string base64Image)
        {
            Stream data = GetStream(base64Image);
            ImagePart.FeedData(data);
            data.Close();
        }

        private Stream GetStream(string base64String)
        {
            return new MemoryStream(Convert.FromBase64String(base64String));
        }
        
        private void InitColumnHeaderNames()
        {
            var names = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            var columnName = string.Empty;

            int firstLetterIndex, secondLetterIndex, thirdLetterIndex, columnHeaderIndex;

            firstLetterIndex = secondLetterIndex = thirdLetterIndex = -1;

            for (columnHeaderIndex = 0; columnHeaderIndex < maxColumns; ++columnHeaderIndex)
            {
                columnName = string.Empty;

                ++firstLetterIndex;

                if (firstLetterIndex == 26)
                {
                    firstLetterIndex = 0;

                    ++secondLetterIndex;

                    if (secondLetterIndex == 26)
                    {
                        secondLetterIndex = 0;

                        ++thirdLetterIndex;
                    }
                }

                if (thirdLetterIndex >= 0) columnName += names[thirdLetterIndex];

                if (secondLetterIndex >= 0) columnName += names[secondLetterIndex];

                if (firstLetterIndex >= 0) columnName += names[firstLetterIndex];

                ColumnHeaderNames[columnHeaderIndex] = columnName;
            }
        }
    }
}
