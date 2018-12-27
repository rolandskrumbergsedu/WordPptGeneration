using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordDocumentGeneration.ContentHelper
{
    public static class ContentTable5
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

            TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "009B2C1D", ParagraphId = "5B63C23B", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "800", Type = TableWidthUnitValues.Dxa };
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

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "459760A7", TextId = "77777777" };

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            Bold bold1 = new Bold();
            FontSize fontSize1 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "22" };

            runProperties1.Append(bold1);
            runProperties1.Append(fontSize1);
            runProperties1.Append(fontSizeComplexScript1);
            Text text1 = new Text();
            text1.Text = "ADDITIONAL COURSES";

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

            TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "009B2C1D", ParagraphId = "78EC2050", TextId = "77777777" };

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

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "1D1F29ED", TextId = "111BDF87" };

            Run run3 = new Run();

            RunProperties runProperties3 = new RunProperties();
            FontSize fontSize3 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "22" };

            runProperties3.Append(fontSize3);
            runProperties3.Append(fontSizeComplexScript3);
            Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text3.Text = "4 days /2018 ";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph2.Append(run3);

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

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "4F98306E", TextId = "77777777" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation1 = new Indentation() { Left = "144" };

            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(indentation1);

            Run run4 = new Run();

            RunProperties runProperties4 = new RunProperties();
            Bold bold2 = new Bold();
            FontSize fontSize4 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "22" };

            runProperties4.Append(bold2);
            runProperties4.Append(fontSize4);
            runProperties4.Append(fontSizeComplexScript4);
            Text text4 = new Text();
            text4.Text = "SUCCESFUL INVESTING THROUGH IPO (INITIAL PUBLIC OFFERINGS)";

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph3.Append(paragraphProperties1);
            paragraph3.Append(run4);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "1C86C2CE", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation2 = new Indentation() { Left = "144" };

            paragraphProperties2.Append(spacingBetweenLines2);
            paragraphProperties2.Append(indentation2);

            Run run5 = new Run();

            RunProperties runProperties5 = new RunProperties();
            FontSize fontSize5 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "22" };

            runProperties5.Append(fontSize5);
            runProperties5.Append(fontSizeComplexScript5);
            Text text5 = new Text();
            text5.Text = "Edward Dubinsky/";

            run5.Append(runProperties5);
            run5.Append(text5);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run6 = new Run();

            RunProperties runProperties6 = new RunProperties();
            FontSize fontSize6 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "22" };

            runProperties6.Append(fontSize6);
            runProperties6.Append(fontSizeComplexScript6);
            Text text6 = new Text();
            text6.Text = "Fintelect";

            run6.Append(runProperties6);
            run6.Append(text6);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph4.Append(paragraphProperties2);
            paragraph4.Append(run5);
            paragraph4.Append(proofError1);
            paragraph4.Append(run6);
            paragraph4.Append(proofError2);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);
            tableCell3.Append(paragraph4);

            tableRow2.Append(tableRowProperties1);
            tableRow2.Append(tableCell2);
            tableRow2.Append(tableCell3);

            TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "009B2C1D", ParagraphId = "5C05956C", TextId = "77777777" };

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

            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "15A0A835", TextId = "67A8D757" };

            Run run7 = new Run();

            RunProperties runProperties7 = new RunProperties();
            FontSize fontSize7 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "22" };

            runProperties7.Append(fontSize7);
            runProperties7.Append(fontSizeComplexScript7);
            Text text7 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text7.Text = "3 days /2018 ";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph5.Append(run7);

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

            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "4F61B891", TextId = "77777777" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation3 = new Indentation() { Left = "144" };

            paragraphProperties3.Append(spacingBetweenLines3);
            paragraphProperties3.Append(indentation3);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            Bold bold3 = new Bold();
            FontSize fontSize8 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "22" };

            runProperties8.Append(bold3);
            runProperties8.Append(fontSize8);
            runProperties8.Append(fontSizeComplexScript8);
            Text text8 = new Text();
            text8.Text = "SUCCESS STORY BY MULTIMILLIONAIR ROBET ALLEN";

            run8.Append(runProperties8);
            run8.Append(text8);

            paragraph6.Append(paragraphProperties3);
            paragraph6.Append(run8);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "72F27F2B", TextId = "77777777" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation4 = new Indentation() { Left = "144" };

            paragraphProperties4.Append(spacingBetweenLines4);
            paragraphProperties4.Append(indentation4);

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            FontSize fontSize9 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "22" };

            runProperties9.Append(fontSize9);
            runProperties9.Append(fontSizeComplexScript9);
            Text text9 = new Text();
            text9.Text = "Robert Allen";

            run9.Append(runProperties9);
            run9.Append(text9);

            paragraph7.Append(paragraphProperties4);
            paragraph7.Append(run9);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph6);
            tableCell5.Append(paragraph7);

            tableRow3.Append(tableRowProperties2);
            tableRow3.Append(tableCell4);
            tableRow3.Append(tableCell5);

            TableRow tableRow4 = new TableRow() { RsidTableRowAddition = "009B2C1D", ParagraphId = "03F6BF1E", TextId = "77777777" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            GridAfter gridAfter3 = new GridAfter() { Val = 1 };
            WidthAfterTableRow widthAfterTableRow3 = new WidthAfterTableRow() { Width = "360", Type = TableWidthUnitValues.Dxa };

            tableRowProperties3.Append(gridAfter3);
            tableRowProperties3.Append(widthAfterTableRow3);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "2550", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders6 = new TableCellBorders();
            TopBorder topBorder7 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder7 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder7 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders6.Append(topBorder7);
            tableCellBorders6.Append(leftBorder7);
            tableCellBorders6.Append(bottomBorder7);
            tableCellBorders6.Append(rightBorder7);

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(tableCellBorders6);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "5D42335D", TextId = "51893288" };

            Run run10 = new Run();

            RunProperties runProperties10 = new RunProperties();
            FontSize fontSize10 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "22" };

            runProperties10.Append(fontSize10);
            runProperties10.Append(fontSizeComplexScript10);
            Text text10 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text10.Text = "7 days /2017 ";

            run10.Append(runProperties10);
            run10.Append(text10);

            paragraph8.Append(run10);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph8);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "5700", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders7 = new TableCellBorders();
            TopBorder topBorder8 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder8 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)1U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder8 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders7.Append(topBorder8);
            tableCellBorders7.Append(leftBorder8);
            tableCellBorders7.Append(bottomBorder8);
            tableCellBorders7.Append(rightBorder8);

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(tableCellBorders7);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "4FB95C8A", TextId = "77777777" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation5 = new Indentation() { Left = "144" };

            paragraphProperties5.Append(spacingBetweenLines5);
            paragraphProperties5.Append(indentation5);

            Run run11 = new Run();

            RunProperties runProperties11 = new RunProperties();
            Bold bold4 = new Bold();
            FontSize fontSize11 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "22" };

            runProperties11.Append(bold4);
            runProperties11.Append(fontSize11);
            runProperties11.Append(fontSizeComplexScript11);
            Text text11 = new Text();
            text11.Text = "7 WEEKS OF GENIUS MINDSET";

            run11.Append(runProperties11);
            run11.Append(text11);

            paragraph9.Append(paragraphProperties5);
            paragraph9.Append(run11);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "151FE73D", TextId = "77777777" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation6 = new Indentation() { Left = "144" };

            paragraphProperties6.Append(spacingBetweenLines6);
            paragraphProperties6.Append(indentation6);
            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run12 = new Run();

            RunProperties runProperties12 = new RunProperties();
            FontSize fontSize12 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "22" };

            runProperties12.Append(fontSize12);
            runProperties12.Append(fontSizeComplexScript12);
            Text text12 = new Text();
            text12.Text = "Mikola";

            run12.Append(runProperties12);
            run12.Append(text12);
            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

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
            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run14 = new Run();

            RunProperties runProperties14 = new RunProperties();
            FontSize fontSize14 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "22" };

            runProperties14.Append(fontSize14);
            runProperties14.Append(fontSizeComplexScript14);
            Text text14 = new Text();
            text14.Text = "Latansky";

            run14.Append(runProperties14);
            run14.Append(text14);
            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph10.Append(paragraphProperties6);
            paragraph10.Append(proofError3);
            paragraph10.Append(run12);
            paragraph10.Append(proofError4);
            paragraph10.Append(run13);
            paragraph10.Append(proofError5);
            paragraph10.Append(run14);
            paragraph10.Append(proofError6);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph9);
            tableCell7.Append(paragraph10);

            tableRow4.Append(tableRowProperties3);
            tableRow4.Append(tableCell6);
            tableRow4.Append(tableCell7);

            TableRow tableRow5 = new TableRow() { RsidTableRowAddition = "009B2C1D", ParagraphId = "5B71406E", TextId = "77777777" };

            TableRowProperties tableRowProperties4 = new TableRowProperties();
            GridAfter gridAfter4 = new GridAfter() { Val = 1 };
            WidthAfterTableRow widthAfterTableRow4 = new WidthAfterTableRow() { Width = "360", Type = TableWidthUnitValues.Dxa };

            tableRowProperties4.Append(gridAfter4);
            tableRowProperties4.Append(widthAfterTableRow4);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "2550", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders8 = new TableCellBorders();
            TopBorder topBorder9 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder9 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder9 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder9 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders8.Append(topBorder9);
            tableCellBorders8.Append(leftBorder9);
            tableCellBorders8.Append(bottomBorder9);
            tableCellBorders8.Append(rightBorder9);

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(tableCellBorders8);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "2C0E9EF7", TextId = "73FE0629" };

            Run run15 = new Run();

            RunProperties runProperties15 = new RunProperties();
            FontSize fontSize15 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "22" };

            runProperties15.Append(fontSize15);
            runProperties15.Append(fontSizeComplexScript15);
            Text text15 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text15.Text = "5 days /2017 ";

            run15.Append(runProperties15);
            run15.Append(text15);

            paragraph11.Append(run15);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph11);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "5700", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders9 = new TableCellBorders();
            TopBorder topBorder10 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder10 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)1U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder10 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder10 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders9.Append(topBorder10);
            tableCellBorders9.Append(leftBorder10);
            tableCellBorders9.Append(bottomBorder10);
            tableCellBorders9.Append(rightBorder10);

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(tableCellBorders9);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "47F544FE", TextId = "77777777" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation7 = new Indentation() { Left = "144" };

            paragraphProperties7.Append(spacingBetweenLines7);
            paragraphProperties7.Append(indentation7);

            Run run16 = new Run();

            RunProperties runProperties16 = new RunProperties();
            Bold bold5 = new Bold();
            FontSize fontSize16 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "22" };

            runProperties16.Append(bold5);
            runProperties16.Append(fontSize16);
            runProperties16.Append(fontSizeComplexScript16);
            Text text16 = new Text();
            text16.Text = "MASTERPLAN ANALYSIS OF FINANCIAL MARKETS";

            run16.Append(runProperties16);
            run16.Append(text16);

            paragraph12.Append(paragraphProperties7);
            paragraph12.Append(run16);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "2764A741", TextId = "77777777" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation8 = new Indentation() { Left = "144" };

            paragraphProperties8.Append(spacingBetweenLines8);
            paragraphProperties8.Append(indentation8);

            Run run17 = new Run();

            RunProperties runProperties17 = new RunProperties();
            FontSize fontSize17 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "22" };

            runProperties17.Append(fontSize17);
            runProperties17.Append(fontSizeComplexScript17);
            Text text17 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text17.Text = "Davide ";

            run17.Append(runProperties17);
            run17.Append(text17);
            ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run18 = new Run();

            RunProperties runProperties18 = new RunProperties();
            FontSize fontSize18 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "22" };

            runProperties18.Append(fontSize18);
            runProperties18.Append(fontSizeComplexScript18);
            Text text18 = new Text();
            text18.Text = "Catanossi";

            run18.Append(runProperties18);
            run18.Append(text18);
            ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph13.Append(paragraphProperties8);
            paragraph13.Append(run17);
            paragraph13.Append(proofError7);
            paragraph13.Append(run18);
            paragraph13.Append(proofError8);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph12);
            tableCell9.Append(paragraph13);

            tableRow5.Append(tableRowProperties4);
            tableRow5.Append(tableCell8);
            tableRow5.Append(tableCell9);

            TableRow tableRow6 = new TableRow() { RsidTableRowAddition = "009B2C1D", ParagraphId = "0F118484", TextId = "77777777" };

            TableRowProperties tableRowProperties5 = new TableRowProperties();
            GridAfter gridAfter5 = new GridAfter() { Val = 1 };
            WidthAfterTableRow widthAfterTableRow5 = new WidthAfterTableRow() { Width = "360", Type = TableWidthUnitValues.Dxa };

            tableRowProperties5.Append(gridAfter5);
            tableRowProperties5.Append(widthAfterTableRow5);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "2550", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders10 = new TableCellBorders();
            TopBorder topBorder11 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder11 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder11 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder11 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders10.Append(topBorder11);
            tableCellBorders10.Append(leftBorder11);
            tableCellBorders10.Append(bottomBorder11);
            tableCellBorders10.Append(rightBorder11);

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(tableCellBorders10);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "22C1914E", TextId = "6503111B" };

            Run run19 = new Run();

            RunProperties runProperties19 = new RunProperties();
            FontSize fontSize19 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "22" };

            runProperties19.Append(fontSize19);
            runProperties19.Append(fontSizeComplexScript19);
            Text text19 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text19.Text = "1 day /2017 ";

            run19.Append(runProperties19);
            run19.Append(text19);

            paragraph14.Append(run19);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph14);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "5700", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders11 = new TableCellBorders();
            TopBorder topBorder12 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder12 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)1U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder12 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder12 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders11.Append(topBorder12);
            tableCellBorders11.Append(leftBorder12);
            tableCellBorders11.Append(bottomBorder12);
            tableCellBorders11.Append(rightBorder12);

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(tableCellBorders11);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "76F032E7", TextId = "77777777" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation9 = new Indentation() { Left = "144" };

            paragraphProperties9.Append(spacingBetweenLines9);
            paragraphProperties9.Append(indentation9);

            Run run20 = new Run();

            RunProperties runProperties20 = new RunProperties();
            Bold bold6 = new Bold();
            FontSize fontSize20 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "22" };

            runProperties20.Append(bold6);
            runProperties20.Append(fontSize20);
            runProperties20.Append(fontSizeComplexScript20);
            Text text20 = new Text();
            text20.Text = "REACHING PERSONAL MAXIMUM";

            run20.Append(runProperties20);
            run20.Append(text20);

            paragraph15.Append(paragraphProperties9);
            paragraph15.Append(run20);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "77DDB263", TextId = "77777777" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation10 = new Indentation() { Left = "144" };

            paragraphProperties10.Append(spacingBetweenLines10);
            paragraphProperties10.Append(indentation10);

            Run run21 = new Run();

            RunProperties runProperties21 = new RunProperties();
            FontSize fontSize21 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "22" };

            runProperties21.Append(fontSize21);
            runProperties21.Append(fontSizeComplexScript21);
            Text text21 = new Text();
            text21.Text = "Brian Tracy";

            run21.Append(runProperties21);
            run21.Append(text21);

            paragraph16.Append(paragraphProperties10);
            paragraph16.Append(run21);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph15);
            tableCell11.Append(paragraph16);

            tableRow6.Append(tableRowProperties5);
            tableRow6.Append(tableCell10);
            tableRow6.Append(tableCell11);

            TableRow tableRow7 = new TableRow() { RsidTableRowAddition = "009B2C1D", ParagraphId = "1ECC5A69", TextId = "77777777" };

            TableRowProperties tableRowProperties6 = new TableRowProperties();
            GridAfter gridAfter6 = new GridAfter() { Val = 1 };
            WidthAfterTableRow widthAfterTableRow6 = new WidthAfterTableRow() { Width = "360", Type = TableWidthUnitValues.Dxa };

            tableRowProperties6.Append(gridAfter6);
            tableRowProperties6.Append(widthAfterTableRow6);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "2550", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders12 = new TableCellBorders();
            TopBorder topBorder13 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder13 = new LeftBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder13 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder13 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders12.Append(topBorder13);
            tableCellBorders12.Append(leftBorder13);
            tableCellBorders12.Append(bottomBorder13);
            tableCellBorders12.Append(rightBorder13);

            tableCellProperties12.Append(tableCellWidth12);
            tableCellProperties12.Append(tableCellBorders12);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "2E157BDE", TextId = "26FF1DCC" };

            Run run22 = new Run();

            RunProperties runProperties22 = new RunProperties();
            FontSize fontSize22 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "22" };

            runProperties22.Append(fontSize22);
            runProperties22.Append(fontSizeComplexScript22);
            Text text22 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text22.Text = "1 day /2017 ";

            run22.Append(runProperties22);
            run22.Append(text22);

            paragraph17.Append(run22);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph17);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "5700", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders13 = new TableCellBorders();
            TopBorder topBorder14 = new TopBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder14 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)1U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder14 = new BottomBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder14 = new RightBorder() { Val = BorderValues.Single, Color = "FFFFFF", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders13.Append(topBorder14);
            tableCellBorders13.Append(leftBorder14);
            tableCellBorders13.Append(bottomBorder14);
            tableCellBorders13.Append(rightBorder14);

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(tableCellBorders13);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "093DCB2B", TextId = "77777777" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation11 = new Indentation() { Left = "144" };

            paragraphProperties11.Append(spacingBetweenLines11);
            paragraphProperties11.Append(indentation11);

            Run run23 = new Run();

            RunProperties runProperties23 = new RunProperties();
            Bold bold7 = new Bold();
            FontSize fontSize23 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "22" };

            runProperties23.Append(bold7);
            runProperties23.Append(fontSize23);
            runProperties23.Append(fontSizeComplexScript23);
            Text text23 = new Text();
            text23.Text = "ART OF THE SPEECH";

            run23.Append(runProperties23);
            run23.Append(text23);

            paragraph18.Append(paragraphProperties11);
            paragraph18.Append(run23);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphAddition = "009B2C1D", RsidRunAdditionDefault = "009E39C2", ParagraphId = "22A7EA18", TextId = "77777777" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { After = "100", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation12 = new Indentation() { Left = "144" };

            paragraphProperties12.Append(spacingBetweenLines12);
            paragraphProperties12.Append(indentation12);
            ProofError proofError9 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run24 = new Run();

            RunProperties runProperties24 = new RunProperties();
            FontSize fontSize24 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "22" };

            runProperties24.Append(fontSize24);
            runProperties24.Append(fontSizeComplexScript24);
            Text text24 = new Text();
            text24.Text = "Radislav";

            run24.Append(runProperties24);
            run24.Append(text24);
            ProofError proofError10 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run25 = new Run();

            RunProperties runProperties25 = new RunProperties();
            FontSize fontSize25 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "22" };

            runProperties25.Append(fontSize25);
            runProperties25.Append(fontSizeComplexScript25);
            Text text25 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text25.Text = " ";

            run25.Append(runProperties25);
            run25.Append(text25);
            ProofError proofError11 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run26 = new Run();

            RunProperties runProperties26 = new RunProperties();
            FontSize fontSize26 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "22" };

            runProperties26.Append(fontSize26);
            runProperties26.Append(fontSizeComplexScript26);
            Text text26 = new Text();
            text26.Text = "Gandapas";

            run26.Append(runProperties26);
            run26.Append(text26);
            ProofError proofError12 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph19.Append(paragraphProperties12);
            paragraph19.Append(proofError9);
            paragraph19.Append(run24);
            paragraph19.Append(proofError10);
            paragraph19.Append(run25);
            paragraph19.Append(proofError11);
            paragraph19.Append(run26);
            paragraph19.Append(proofError12);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph18);
            tableCell13.Append(paragraph19);

            tableRow7.Append(tableRowProperties6);
            tableRow7.Append(tableCell12);
            tableRow7.Append(tableCell13);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);
            table1.Append(tableRow5);
            table1.Append(tableRow6);
            table1.Append(tableRow7);
            return table1;
        }


    }
}
