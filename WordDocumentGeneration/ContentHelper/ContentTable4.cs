using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace GeneratedCode
{
    public static class ContentTable4
    {
        // Creates an Table instance and adds its children.
        public static Table GenerateTable()
        {
            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 10, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)10U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)10U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)10U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)10U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)10U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)10U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 10, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 10, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);
            TableLook tableLook1 = new TableLook() { Val = "0000", FirstRow = false, LastRow = false, FirstColumn = false, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = false };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableIndentation1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableCellMarginDefault1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "2550" };
            GridColumn gridColumn2 = new GridColumn() { Width = "5700" };
            GridColumn gridColumn3 = new GridColumn() { Width = "360" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);

            TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "009B2C1D", RsidTableRowProperties = "0007641E", ParagraphId = "4B076042", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "8610", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 3 };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(topBorder2);
            tableCellBorders1.Append(leftBorder2);
            tableCellBorders1.Append(bottomBorder2);
            tableCellBorders1.Append(rightBorder2);

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(gridSpan1);
            tableCellProperties1.Append(tableCellBorders1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "7E6D8800", TextId = "77777777" };

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            Bold bold1 = new Bold();
            FontSize fontSize1 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "22" };

            runProperties1.Append(bold1);
            runProperties1.Append(fontSize1);
            runProperties1.Append(fontSizeComplexScript1);
            Text text1 = new Text();
            text1.Text = "EDUCATION AND QUALIFICATIONS";

            run1.Append(runProperties1);
            run1.Append(text1);

            Run run2 = new Run();

            RunProperties runProperties2 = new RunProperties();
            FontSize fontSize2 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "22" };

            runProperties2.Append(fontSize2);
            runProperties2.Append(fontSizeComplexScript2);
            Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text2.Text = "  ";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph1.Append(run1);
            paragraph1.Append(run2);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            tableRow1.Append(tableCell1);

            TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "009B2C1D", ParagraphId = "1EE453E3", TextId = "77777777" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            GridAfter gridAfter1 = new GridAfter() { Val = 1 };
            WidthAfterTableRow widthAfterTableRow1 = new WidthAfterTableRow() { Width = "360", Type = TableWidthUnitValues.Dxa };

            tableRowProperties1.Append(gridAfter1);
            tableRowProperties1.Append(widthAfterTableRow1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2550", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder3);
            tableCellBorders2.Append(leftBorder3);
            tableCellBorders2.Append(bottomBorder3);
            tableCellBorders2.Append(rightBorder3);

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(tableCellBorders2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "575EC095", TextId = "74BCD266" };

            Run run3 = new Run();

            RunProperties runProperties3 = new RunProperties();
            FontSize fontSize3 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "22" };

            runProperties3.Append(fontSize3);
            runProperties3.Append(fontSizeComplexScript3);
            Text text3 = new Text();
            text3.Text = "199";

            run3.Append(runProperties3);
            run3.Append(text3);

            Run run4 = new Run() { RsidRunAddition = "0007641E" };

            RunProperties runProperties4 = new RunProperties();
            FontSize fontSize4 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "22" };

            runProperties4.Append(fontSize4);
            runProperties4.Append(fontSizeComplexScript4);
            Text text4 = new Text();
            text4.Text = "8";

            run4.Append(runProperties4);
            run4.Append(text4);

            Run run5 = new Run();

            RunProperties runProperties5 = new RunProperties();
            FontSize fontSize5 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "22" };

            runProperties5.Append(fontSize5);
            runProperties5.Append(fontSizeComplexScript5);
            Text text5 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text5.Text = " - 200";

            run5.Append(runProperties5);
            run5.Append(text5);

            Run run6 = new Run() { RsidRunAddition = "0007641E" };

            RunProperties runProperties6 = new RunProperties();
            FontSize fontSize6 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "22" };

            runProperties6.Append(fontSize6);
            runProperties6.Append(fontSizeComplexScript6);
            Text text6 = new Text();
            text6.Text = "1";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph2.Append(run3);
            paragraph2.Append(run4);
            paragraph2.Append(run5);
            paragraph2.Append(run6);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "5700", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)1U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(topBorder4);
            tableCellBorders3.Append(leftBorder4);
            tableCellBorders3.Append(bottomBorder4);
            tableCellBorders3.Append(rightBorder4);

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellBorders3);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "690EE799", TextId = "77777777" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation1 = new Indentation() { Left = "144" };

            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(indentation1);

            Run run7 = new Run();

            RunProperties runProperties7 = new RunProperties();
            Bold bold2 = new Bold();
            FontSize fontSize7 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "22" };

            runProperties7.Append(bold2);
            runProperties7.Append(fontSize7);
            runProperties7.Append(fontSizeComplexScript7);
            Text text7 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text7.Text = "UNIVERSITY OF LATVIA ";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph3.Append(paragraphProperties1);
            paragraph3.Append(run7);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "7BD7BFC1", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation2 = new Indentation() { Left = "144" };

            paragraphProperties2.Append(spacingBetweenLines2);
            paragraphProperties2.Append(indentation2);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            FontSize fontSize8 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "22" };

            runProperties8.Append(fontSize8);
            runProperties8.Append(fontSizeComplexScript8);
            Text text8 = new Text();
            text8.Text = "Master of Law";

            run8.Append(runProperties8);
            run8.Append(text8);

            paragraph4.Append(paragraphProperties2);
            paragraph4.Append(run8);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);
            tableCell3.Append(paragraph4);

            tableRow2.Append(tableRowProperties1);
            tableRow2.Append(tableCell2);
            tableRow2.Append(tableCell3);

            TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "009B2C1D", ParagraphId = "7395F855", TextId = "77777777" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            GridAfter gridAfter2 = new GridAfter() { Val = 1 };
            WidthAfterTableRow widthAfterTableRow2 = new WidthAfterTableRow() { Width = "360", Type = TableWidthUnitValues.Dxa };

            tableRowProperties2.Append(gridAfter2);
            tableRowProperties2.Append(widthAfterTableRow2);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "2550", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(topBorder5);
            tableCellBorders4.Append(leftBorder5);
            tableCellBorders4.Append(bottomBorder5);
            tableCellBorders4.Append(rightBorder5);

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(tableCellBorders4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "38730036", TextId = "7E95603F" };

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            FontSize fontSize9 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "22" };

            runProperties9.Append(fontSize9);
            runProperties9.Append(fontSizeComplexScript9);
            Text text9 = new Text();
            text9.Text = "199";

            run9.Append(runProperties9);
            run9.Append(text9);

            Run run10 = new Run() { RsidRunAddition = "0007641E" };

            RunProperties runProperties10 = new RunProperties();
            FontSize fontSize10 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "22" };

            runProperties10.Append(fontSize10);
            runProperties10.Append(fontSizeComplexScript10);
            Text text10 = new Text();
            text10.Text = "4";

            run10.Append(runProperties10);
            run10.Append(text10);

            Run run11 = new Run();

            RunProperties runProperties11 = new RunProperties();
            FontSize fontSize11 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "22" };

            runProperties11.Append(fontSize11);
            runProperties11.Append(fontSizeComplexScript11);
            Text text11 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text11.Text = " - 199";

            run11.Append(runProperties11);
            run11.Append(text11);

            Run run12 = new Run() { RsidRunAddition = "0007641E" };

            RunProperties runProperties12 = new RunProperties();
            FontSize fontSize12 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "22" };

            runProperties12.Append(fontSize12);
            runProperties12.Append(fontSizeComplexScript12);
            Text text12 = new Text();
            text12.Text = "8";

            run12.Append(runProperties12);
            run12.Append(text12);

            Run run13 = new Run();

            RunProperties runProperties13 = new RunProperties();
            FontSize fontSize13 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "22" };

            runProperties13.Append(fontSize13);
            runProperties13.Append(fontSizeComplexScript13);
            Text text13 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text13.Text = " ";

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph5.Append(run9);
            paragraph5.Append(run10);
            paragraph5.Append(run11);
            paragraph5.Append(run12);
            paragraph5.Append(run13);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph5);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "5700", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            TopBorder topBorder6 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)1U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders5.Append(topBorder6);
            tableCellBorders5.Append(leftBorder6);
            tableCellBorders5.Append(bottomBorder6);
            tableCellBorders5.Append(rightBorder6);

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(tableCellBorders5);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "2EDBDBCE", TextId = "77777777" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation3 = new Indentation() { Left = "144" };

            paragraphProperties3.Append(spacingBetweenLines3);
            paragraphProperties3.Append(indentation3);

            Run run14 = new Run();

            RunProperties runProperties14 = new RunProperties();
            Bold bold3 = new Bold();
            FontSize fontSize14 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "22" };

            runProperties14.Append(bold3);
            runProperties14.Append(fontSize14);
            runProperties14.Append(fontSizeComplexScript14);
            Text text14 = new Text();
            text14.Text = "RIGA INTERNATIONAL SCHOOL OF ECONOMICS AND BUSINESS ADMINISTRATION";

            run14.Append(runProperties14);
            run14.Append(text14);

            paragraph6.Append(paragraphProperties3);
            paragraph6.Append(run14);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "438C2C8D", TextId = "77777777" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation4 = new Indentation() { Left = "144" };

            paragraphProperties4.Append(spacingBetweenLines4);
            paragraphProperties4.Append(indentation4);

            Run run15 = new Run();

            RunProperties runProperties15 = new RunProperties();
            FontSize fontSize15 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "22" };

            runProperties15.Append(fontSize15);
            runProperties15.Append(fontSizeComplexScript15);
            Text text15 = new Text();
            text15.Text = "Bachelor of Business Administration";

            run15.Append(runProperties15);
            run15.Append(text15);

            paragraph7.Append(paragraphProperties4);
            paragraph7.Append(run15);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph6);
            tableCell5.Append(paragraph7);

            tableRow3.Append(tableRowProperties2);
            tableRow3.Append(tableCell4);
            tableRow3.Append(tableCell5);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            return table1;
        }


    }
}
