using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using GridColumn = DocumentFormat.OpenXml.Wordprocessing.GridColumn;
using InsideHorizontalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder;
using InsideVerticalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using NonVisualGraphicFrameDrawingProperties = DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellBorders = DocumentFormat.OpenXml.Wordprocessing.TableCellBorders;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableGrid = DocumentFormat.OpenXml.Wordprocessing.TableGrid;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;

namespace WordDocumentGeneration
{
    public class WordDocumentManager
    {
        public void SaveDocument(GenerationData data, string filePath, string fileName)
        {
            using (var mem = new MemoryStream())
            {
                // Create Document
                using (var wordDocument =
                    WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
                {
                    // Add a main document part. 
                    var mainPart = wordDocument.AddMainDocumentPart();

                    var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    using (var stream = new FileStream(@"C:\Users\KRUROL\Downloads\Testa Amrop CV_eng - Copy\word\media\image1.jpeg", FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }

                    // Create the document structure and add some text.
                    mainPart.Document = FillDocumentFromData(data, mainPart.GetIdOfPart(imagePart));

                    AddHeaderFooter(mainPart, data);
                }
                
                mem.Position = 0;

                using (var file = new FileStream($"{filePath}\\{fileName}", FileMode.CreateNew, FileAccess.Write))
                {
                    mem.CopyTo(file);
                }
            }
        }

        public byte[] GetDocument(GenerationData data)
        {
            using (var mem = new MemoryStream())
            {
                // Create Document
                using (var wordDocument =
                    WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
                {
                    // Add a main document part. 
                    var mainPart = wordDocument.AddMainDocumentPart();

                    var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    using (var stream = new FileStream(@"C:\Users\KRUROL\Downloads\Testa Amrop CV_eng - Copy\word\media\image1.jpeg", FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }

                    // Create the document structure and add some text.
                    mainPart.Document = FillDocumentFromData(data, mainPart.GetIdOfPart(imagePart));

                    AddHeaderFooter(mainPart, data);

                }

                return mem.ToArray();
            }
        }

        private void AddHeaderFooter(MainDocumentPart mainPart, GenerationData data)
        {
            // Create a new header and footer part
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            var headerPartFirst = mainPart.AddNewPart<HeaderPart>();
            var footerPart = mainPart.AddNewPart<FooterPart>();

            // Get Id of the headerPart and footer parts
            var headerPartId = mainPart.GetIdOfPart(headerPart);
            var headerPartFirstId = mainPart.GetIdOfPart(headerPartFirst);
            var footerPartId = mainPart.GetIdOfPart(footerPart);
            
            GenerateHeaderPartContent(headerPart, data);
            GenerateHeaderFirstPartContent(headerPartFirst);
            GenerateFooterPartContent(footerPart, data);

            // Get SectionProperties and Replace HeaderReference and FooterRefernce with new Id
            var sections = mainPart.Document.Body.Elements<SectionProperties>();

            foreach (var section in sections)
            {
                section.RsidR = HexBinaryValue.FromString("009B2C1D");

                // Delete existing references to headers and footers
                section.RemoveAllChildren<HeaderReference>();
                section.RemoveAllChildren<FooterReference>();

                // Create the new header and footer reference node
                section.PrependChild(new HeaderReference { Id = headerPartId, Type = new EnumValue<HeaderFooterValues>{Value = HeaderFooterValues.Default}});
                section.PrependChild(new FooterReference { Id = footerPartId });
                section.PrependChild(new HeaderReference { Id = headerPartFirstId, Type = new EnumValue<HeaderFooterValues> { Value = HeaderFooterValues.Default } });


                section.PrependChild(new PageSize
                {
                    Width = UInt32Value.FromUInt32(11870),
                    Height = UInt32Value.FromUInt32(16787)
                });
                section.PrependChild(new PageMargin
                {
                    Top = Int32Value.FromInt32(1440),
                    Right = UInt32Value.FromUInt32(1440),
                    Bottom = Int32Value.FromInt32(1440),
                    Left = UInt32Value.FromUInt32(1440),
                    Header = UInt32Value.FromUInt32(720),
                    Footer = UInt32Value.FromUInt32(720),
                    Gutter = UInt32Value.FromUInt32(0)
                });
                section.PrependChild(new Columns
                {
                    Space = StringValue.FromString("720")
                });
                section.PrependChild(new TitlePage());
            }
        }

        private void GenerateHeaderFirstPartContent(HeaderPart headerPartFirst)
        {
            var header = new Header { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "w14 w15 w16se w16cid wp14" } };
            header.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            header.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            header.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            header.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            header.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            header.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            header.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            header.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            header.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            header.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            header.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            header.AppendChild(new Paragraph(new Run(new CarriageReturn()))
            {
                ParagraphId = HexBinaryValue.FromString("637CAE62"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("00F225EA"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009E39C2")
            });

            headerPartFirst.Header = header;
        }

        private void GenerateFooterPartContent(FooterPart footerPart, GenerationData data)
        {
            var footer = new Footer { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "w14 w15 w16se w16cid wp14" } };
            footer.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            footer.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            footer.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            footer.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            footer.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            footer.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            footer.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            footer.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            footer.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            footer.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            footer.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            footer.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            footer.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            footer.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            var paragraph = new Paragraph
            {
                ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties
                {
                    ParagraphStyleId = new ParagraphStyleId { Val = StringValue.FromString("leftRight") }
                },
                TextId = HexBinaryValue.FromString("77777777"),
                ParagraphId = HexBinaryValue.FromString("71F3416E"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009E39C2")
            };
            paragraph.AppendChild(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(data.TitleArea.Title)));
            paragraph.AppendChild(new Run(new TabChar()));
            paragraph.AppendChild(new Run(new FieldChar { FieldCharType = new EnumValue<FieldCharValues>{ Value = FieldCharValues.Begin }}));
            paragraph.AppendChild(new Run(new FieldCode("PAGE")));
            paragraph.AppendChild(new Run(new FieldChar { FieldCharType = new EnumValue<FieldCharValues> { Value = FieldCharValues.Separate } }));
            paragraph.AppendChild(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("2"))
            {
                RunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties
                {
                    NoProof = new NoProof()
                }
            });
            paragraph.AppendChild(new Run(new FieldChar { FieldCharType = new EnumValue<FieldCharValues> { Value = FieldCharValues.End } }));
            footer.AppendChild(paragraph);

            footerPart.Footer = footer;
        }

        private void GenerateHeaderPartContent(HeaderPart headerPart, GenerationData data)
        {
            var header = new Header { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "w14 w15 w16se w16cid wp14" } };
            header.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            header.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            header.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            header.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            header.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            header.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            header.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            header.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            header.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            header.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            header.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            var table = new Table();

            var tableProperties = new TableProperties
            {
                TableWidth = new TableWidth
                {
                    Width = StringValue.FromString("0"),
                    Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Auto }
                },
                TableIndentation = new TableIndentation
                {
                    Width = Int32Value.FromInt32(10),
                    Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
                },
                TableCellMarginDefault = new TableCellMarginDefault
                {
                    TableCellLeftMargin = new TableCellLeftMargin
                    {
                        Width = Int16Value.FromInt16(10),
                        Type = new EnumValue<TableWidthValues> { Value = TableWidthValues.Dxa}
                    },
                    TableCellRightMargin = new TableCellRightMargin
                    {
                        Width = Int16Value.FromInt16(10),
                        Type = new EnumValue<TableWidthValues> { Value = TableWidthValues.Dxa }
                    }
                },
                TableLook = new TableLook
                {
                    Val = HexBinaryValue.FromString("0000"),
                    FirstRow = OnOffValue.FromBoolean(false),
                    LastRow = OnOffValue.FromBoolean(false),
                    FirstColumn = OnOffValue.FromBoolean(false),
                    LastColumn = OnOffValue.FromBoolean(false),
                    NoHorizontalBand = OnOffValue.FromBoolean(false),
                    NoVerticalBand = OnOffValue.FromBoolean(false)
                }
            };
            table.AppendChild(tableProperties);

            var tableGrid = new TableGrid(new GridColumn { Width = StringValue.FromString("8980") });
            table.AppendChild(tableGrid);

            var tableRow = new TableRow
            {
                ParagraphId = HexBinaryValue.FromString("07A74B1D"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidTableRowAddition = HexBinaryValue.FromString("009B2C1D")
            };
            var tableCell = new TableCell
            {
                TableCellProperties = new TableCellProperties
                {
                    TableCellWidth = new TableCellWidth
                    {
                        Width = StringValue.FromString("9000"),
                        Type = new EnumValue<TableWidthUnitValues> {Value = TableWidthUnitValues.Dxa}
                    },
                    Shading = new Shading
                    {
                        Val = new EnumValue<ShadingPatternValues> {Value = ShadingPatternValues.Clear},
                        Color = StringValue.FromString("auto"),
                        Fill = StringValue.FromString("1C75BC")
                    }
                }
            };
            tableCell.AppendChild(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(data.TitleArea.Name))
            {
                RunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties
                {
                    Bold = new Bold(),
                    Caps = new Caps(),
                    Color = new Color
                    {
                        Val = StringValue.FromString("FFFFFF")
                    },
                    FontSize = new FontSize
                    {
                        Val = StringValue.FromString("21")
                    },
                    FontSizeComplexScript = new FontSizeComplexScript
                    {
                        Val = StringValue.FromString("21")
                    }
                }
            })
            {
                ParagraphId = HexBinaryValue.FromString("7EE08DA7"),
                TextId = HexBinaryValue.FromString("2A892E24"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("0007641E"),
                ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties
                {
                    SpacingBetweenLines = new SpacingBetweenLines
                    {
                        Before = StringValue.FromString("10"),
                        After = StringValue.FromString("10")
                    },
                    Justification = new Justification
                    {
                        Val = new EnumValue<JustificationValues> { Value = JustificationValues.Center }
                    }
                }
            });
            tableRow.AppendChild(tableCell);
            table.AppendChild(tableRow);

            header.AppendChild(table);

            headerPart.Header = header;
        }

        private static void AddMainLogo(OpenXmlElement elementToAppend, string relationshipId)
        {
            var element = new Drawing
            {
                Inline = new Inline
                {
                    DistanceFromTop = UInt32Value.FromUInt32(0),
                    DistanceFromBottom = UInt32Value.FromUInt32(0),
                    DistanceFromLeft = UInt32Value.FromUInt32(0),
                    DistanceFromRight = UInt32Value.FromUInt32(0),
                    AnchorId = HexBinaryValue.FromString("29B39BAA"),
                    EditId = HexBinaryValue.FromString("5481AAF3"),
                    Extent = new Extent
                    {
                        Cx = Int64Value.FromInt64(1257300),
                        Cy = Int64Value.FromInt64(1057275)
                    },
                    EffectExtent = new EffectExtent
                    {
                        LeftEdge = Int64Value.FromInt64(0),
                        TopEdge = Int64Value.FromInt64(0),
                        RightEdge = Int64Value.FromInt64(0),
                        BottomEdge = Int64Value.FromInt64(0)
                    },
                    DocProperties = new DocProperties
                    {
                        Id = UInt32Value.FromUInt32(1),
                        Name = "Picture 1"
                    },
                    NonVisualGraphicFrameDrawingProperties = new NonVisualGraphicFrameDrawingProperties
                    {
                        GraphicFrameLocks = new GraphicFrameLocks
                        {
                            NoChangeAspect = BooleanValue.FromBoolean(true)
                        }
                    },
                    Graphic = new Graphic
                    {
                        GraphicData = new GraphicData(new DocumentFormat.OpenXml.Drawing.Pictures.Picture
                        {
                            NonVisualPictureProperties = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties
                            {
                                NonVisualDrawingProperties = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                                {
                                    Id = UInt32Value.FromUInt32(0),
                                    Name = "Picture 1"
                                },
                                NonVisualPictureDrawingProperties = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties
                                {
                                    PictureLocks = new PictureLocks
                                    {
                                        NoChangeAspect = BooleanValue.FromBoolean(true),
                                        NoChangeArrowheads = BooleanValue.FromBoolean(true)
                                    }
                                }
                            },
                            BlipFill = new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(new BlipExtensionList(new BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }),
                                                    new Stretch(new FillRectangle()))
                            {
                                Blip = new Blip
                                {
                                    Embed = relationshipId,
                                    
                                },
                                SourceRectangle = new SourceRectangle(),
                                
                            },
                            ShapeProperties = new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(new PresetGeometry { Preset = new EnumValue<ShapeTypeValues> {Value = ShapeTypeValues.Rectangle}, AdjustValueList = new AdjustValueList()},
                                                                  new NoFill(),
                                                                  new DocumentFormat.OpenXml.Drawing.Outline(new NoFill()))
                            {
                                BlackWhiteMode = new EnumValue<BlackWhiteModeValues> { Value = BlackWhiteModeValues.Auto },
                                Transform2D = new Transform2D
                                {
                                    Offset = new Offset
                                    {
                                        X = Int64Value.FromInt64(0),
                                        Y = Int64Value.FromInt64(0)
                                    },
                                    Extents = new Extents
                                    {
                                        Cx = Int64Value.FromInt64(1257300),
                                        Cy = Int64Value.FromInt64(1057275)
                                    }
                                }
                            }
                        })
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }
                    }
                }
            };

            // Append the reference to body, the element should be in a Run.
            elementToAppend.AppendChild(new Paragraph(new Run(element) {RunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties
            {
                NoProof = new NoProof()
            }}));
        }

        private static Document FillDocumentFromData(GenerationData data, string pictureRelationshipId)
        {
            var document = new Document();

            var body = new Body();

            FillHeaderTable(body, pictureRelationshipId, data);

            document.AppendChild(body);

            return document;
        }

        private static void FillHeaderTable(OpenXmlElement body, string pictureRelationshipId, GenerationData data)
        {
            var table = new Table();

            FillHeaderTableProperties(table);

            FillHeaderTableGrid(table);

            FillFirstTableRow(table, pictureRelationshipId);
            FillSecondTableRow(table);
            FillThirdTableRow(table);
            FillFourthTableRow(table, data);
            FillFifthTableRow(table);
            FillSixthTableRow(table);

            table.AppendChild(new Paragraph
            {
                ParagraphId = HexBinaryValue.FromString("00DAFD9F"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009B2C1D")
            });

            body.AppendChild(table);
        }

        private static void FillSixthTableRow(OpenXmlElement table)
        {
            var tableRow = new TableRow
            {
                TextId = HexBinaryValue.FromString("77777777"),
                ParagraphId = HexBinaryValue.FromString("20373438"),
                RsidTableRowAddition = HexBinaryValue.FromString("009B2C1D")
            };
            var rowProperties = new TableRowProperties();
            rowProperties.AppendChild(new GridAfter
            {
                Val = Int32Value.FromInt32(2)
            });
            rowProperties.AppendChild(new WidthAfterTableRow
            {
                Width = StringValue.FromString("6375"),
                Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
            });
            tableRow.TableRowProperties = rowProperties;

            var tableCell = new TableCell
            {
                TableCellProperties = new TableCellProperties
                {
                    TableCellWidth = new TableCellWidth
                    {
                        Width = StringValue.FromString("800"),
                        Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
                    },
                    TableCellBorders = new TableCellBorders
                    {
                        TopBorder = new TopBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        LeftBorder = new LeftBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        BottomBorder = new BottomBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        RightBorder = new RightBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        }
                    }
                }
            };

            tableCell.AppendChild(new Paragraph
            {
                ParagraphId = HexBinaryValue.FromString("54B78A08"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009B2C1D")
            });
            tableRow.AppendChild(tableCell);
            table.AppendChild(tableRow);
        }

        private static void FillFifthTableRow(OpenXmlElement table)
        {
            var tableRow = new TableRow
            {
                TextId = HexBinaryValue.FromString("77777777"),
                ParagraphId = HexBinaryValue.FromString("485735CB"),
                RsidTableRowAddition = HexBinaryValue.FromString("009B2C1D")
            };
            var rowProperties = new TableRowProperties();
            rowProperties.AppendChild(new GridAfter
            {
                Val = Int32Value.FromInt32(2)
            });
            rowProperties.AppendChild(new WidthAfterTableRow
            {
                Width = StringValue.FromString("6375"),
                Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
            });
            tableRow.TableRowProperties = rowProperties;

            var tableCell = new TableCell
            {
                TableCellProperties = new TableCellProperties
                {
                    TableCellWidth = new TableCellWidth
                    {
                        Width = StringValue.FromString("800"),
                        Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
                    },
                    TableCellBorders = new TableCellBorders
                    {
                        TopBorder = new TopBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        LeftBorder = new LeftBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        BottomBorder = new BottomBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        RightBorder = new RightBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        }
                    }
                }
            };

            tableCell.AppendChild(new Paragraph
            {
                ParagraphId = HexBinaryValue.FromString("41648536"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009B2C1D")
            });
            tableRow.AppendChild(tableCell);
            table.AppendChild(tableRow);
        }

        private static void FillFourthTableRow(OpenXmlElement table, GenerationData data)
        {
            var tableRow = new TableRow
            {
                TextId = HexBinaryValue.FromString("77777777"),
                ParagraphId = HexBinaryValue.FromString("0E828CE9"),
                RsidTableRowAddition = HexBinaryValue.FromString("009B2C1D")
            };

            var tableCell1 = new TableCell
            {
                TableCellProperties = new TableCellProperties
                {
                    TableCellWidth = new TableCellWidth
                    {
                        Width = StringValue.FromString("800"),
                        Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
                    },
                    TableCellBorders = new TableCellBorders
                    {
                        TopBorder = new TopBorder
                        {
                            Val = new EnumValue<BorderValues> {  Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        LeftBorder = new LeftBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        BottomBorder = new BottomBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        RightBorder = new RightBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        }
                    }
                }
            };
            tableCell1.AppendChild(new Paragraph {
                ParagraphId = HexBinaryValue.FromString("7D33E7CB"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009B2C1D")
            });
            tableRow.AppendChild(tableCell1);

            var tableCell2 = new TableCell
            {
                TableCellProperties = new TableCellProperties
                {
                    TableCellWidth = new TableCellWidth
                    {
                        Width = StringValue.FromString("2550"),
                        Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
                    },
                    TableCellBorders = new TableCellBorders
                    {
                        TopBorder = new TopBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        LeftBorder = new LeftBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        BottomBorder = new BottomBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        RightBorder = new RightBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        }
                    }
                }
            };
            tableCell2.AppendChild(new Paragraph
            {
                ParagraphId = HexBinaryValue.FromString("4B505573"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009B2C1D")
            });
            tableRow.AppendChild(tableCell2);

            var tableCell3 = new TableCell
            {
                TableCellProperties = new TableCellProperties
                {
                    TableCellWidth = new TableCellWidth
                    {
                        Width = StringValue.FromString("3825"),
                        Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
                    },
                    TableCellBorders = new TableCellBorders
                    {
                        TopBorder = new TopBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        LeftBorder = new LeftBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        BottomBorder = new BottomBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        RightBorder = new RightBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        }
                    },
                    Shading = new Shading
                    {
                        Val = new EnumValue<ShadingPatternValues> { Value = ShadingPatternValues.Clear },
                        Color = StringValue.FromString("auto"),
                        Fill = StringValue.FromString("0069B4")
                    },
                    TableCellVerticalAlignment = new TableCellVerticalAlignment
                    {
                        Val = new EnumValue<TableVerticalAlignmentValues> { Value = TableVerticalAlignmentValues.Center }
                    }
                }
            };
            tableCell3.AppendChild(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(data.TitleArea.Title))
            {
                RunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties
                {
                    Color = new Color { Val = StringValue.FromString("FFFFFF") },
                    FontSize = new FontSize { Val = StringValue.FromString("18") },
                    FontSizeComplexScript = new FontSizeComplexScript { Val = StringValue.FromString("18") }
                }
            })
            {
                ParagraphId = HexBinaryValue.FromString("63EAFC9C"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009E39C2"),
                ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties
                {
                    SpacingBetweenLines = new SpacingBetweenLines
                    {
                        Before = StringValue.FromString("550"),
                        After = StringValue.FromString("800")
                    },
                    Indentation = new Indentation
                    {
                        Left = StringValue.FromString("432")
                    }
                }
            });
            tableCell3.AppendChild(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(data.TitleArea.Name))
            {
                RunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties
                {
                    Color = new Color { Val = StringValue.FromString("FFFFFF") },
                    FontSize = new FontSize { Val = StringValue.FromString("33") },
                    FontSizeComplexScript = new FontSizeComplexScript { Val = StringValue.FromString("33") }
                }
            })
            {
                ParagraphId = HexBinaryValue.FromString("48E9D2B9"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("0007641E"),
                ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties
                {
                    SpacingBetweenLines = new SpacingBetweenLines
                    {
                        After = StringValue.FromString("0")
                    },
                    Indentation = new Indentation
                    {
                        Left = StringValue.FromString("432")
                    }
                }
            });
            tableCell3.AppendChild(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(data.TitleArea.Date))
            {
                RunProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties
                {
                    Color = new Color { Val = StringValue.FromString("FFFFFF") },
                    FontSize = new FontSize { Val = StringValue.FromString("18") },
                    FontSizeComplexScript = new FontSizeComplexScript { Val = StringValue.FromString("18") }
                }
            })
            {
                ParagraphId = HexBinaryValue.FromString("4D45EC2E"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009E39C2"),
                ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties
                {
                    SpacingBetweenLines = new SpacingBetweenLines
                    {
                        Before = StringValue.FromString("800"),
                        After = StringValue.FromString("550"),
                        Line = StringValue.FromString("240"),
                        LineRule = new EnumValue<LineSpacingRuleValues> { Value = LineSpacingRuleValues.Auto }
                    },
                    Indentation = new Indentation
                    {
                        Left = StringValue.FromString("432")
                    }
                }
            });
            tableRow.AppendChild(tableCell3);

            table.AppendChild(tableRow);
        }

        private static void FillThirdTableRow(OpenXmlElement table)
        {
            var tableRow = new TableRow
            {
                TextId = HexBinaryValue.FromString("77777777"),
                ParagraphId = HexBinaryValue.FromString("13569F8F"),
                RsidTableRowAddition = HexBinaryValue.FromString("009B2C1D")
            };
            var rowProperties = new TableRowProperties();
            rowProperties.AppendChild(new GridAfter
            {
                Val = Int32Value.FromInt32(2)
            });
            rowProperties.AppendChild(new WidthAfterTableRow
            {
                Width = StringValue.FromString("6375"),
                Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
            });
            tableRow.TableRowProperties = rowProperties;

            var tableCell = new TableCell
            {
                TableCellProperties = new TableCellProperties
                {
                    TableCellWidth = new TableCellWidth
                    {
                        Width = StringValue.FromString("800"),
                        Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
                    },
                    TableCellBorders = new TableCellBorders
                    {
                        TopBorder = new TopBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        LeftBorder = new LeftBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        BottomBorder = new BottomBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        RightBorder = new RightBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        }
                    }
                }
            };

            tableCell.AppendChild(new Paragraph
            {
                ParagraphId = HexBinaryValue.FromString("595BC873"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009B2C1D")
            });

            tableRow.AppendChild(tableCell);

            table.AppendChild(tableRow);
        }

        private static void FillSecondTableRow(OpenXmlElement table)
        {
            var tableRow = new TableRow
            {
                TextId = HexBinaryValue.FromString("77777777"),
                ParagraphId = HexBinaryValue.FromString("3125C09D"),
                RsidTableRowAddition = HexBinaryValue.FromString("009B2C1D")
            };
            var rowProperties = new TableRowProperties();
            rowProperties.AppendChild(new GridAfter
            {
                Val = Int32Value.FromInt32(2)
            });
            rowProperties.AppendChild(new WidthAfterTableRow
            {
                Width = StringValue.FromString("6375"),
                Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
            });
            tableRow.TableRowProperties = rowProperties;

            var tableCell = new TableCell
            {
                TableCellProperties = new TableCellProperties
                {
                    TableCellWidth = new TableCellWidth
                    {
                        Width = StringValue.FromString("800"),
                        Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
                    },
                    TableCellBorders = new TableCellBorders
                    {
                        TopBorder = new TopBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        LeftBorder = new LeftBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        BottomBorder = new BottomBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        RightBorder = new RightBorder
                        {
                            Val = new EnumValue<BorderValues> { Value = BorderValues.Single },
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        }
                    }
                }
            };

            tableCell.AppendChild(new Paragraph
            {
                ParagraphId = HexBinaryValue.FromString("595BC873"),
                TextId = HexBinaryValue.FromString("77777777"),
                RsidParagraphAddition = HexBinaryValue.FromString("009B2C1D"),
                RsidRunAdditionDefault = HexBinaryValue.FromString("009B2C1D")
            });
            tableRow.AppendChild(tableCell);
            table.AppendChild(tableRow);
        }

        private static void FillFirstTableRow(OpenXmlElement table, string relationshipId)
        {
            var tableRow = new TableRow
            {
                TextId = HexBinaryValue.FromString("77777777"),
                ParagraphId = HexBinaryValue.FromString("080C4265"),
                RsidTableRowAddition = HexBinaryValue.FromString("009B2C1D")
            };
            var rowProperties = new TableRowProperties();
            rowProperties.AppendChild(new GridAfter
            {
                Val = Int32Value.FromInt32(2)
            });
            rowProperties.AppendChild(new WidthAfterTableRow
            {
                Width = StringValue.FromString("6375"),
                Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Dxa }
            });
            tableRow.TableRowProperties = rowProperties;

            var tableCell = new TableCell
            {
                TableCellProperties = new TableCellProperties
                {
                    TableCellWidth = new TableCellWidth
                    {
                        Width = StringValue.FromString("800"),
                        Type = new EnumValue<TableWidthUnitValues> {Value = TableWidthUnitValues.Dxa}
                    },
                    TableCellBorders = new TableCellBorders
                    {
                        TopBorder = new TopBorder
                        {
                            Val = new EnumValue<BorderValues> {Value = BorderValues.Single},
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        LeftBorder = new LeftBorder
                        {
                            Val = new EnumValue<BorderValues> {Value = BorderValues.Single},
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        BottomBorder = new BottomBorder
                        {
                            Val = new EnumValue<BorderValues> {Value = BorderValues.Single},
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        },
                        RightBorder = new RightBorder
                        {
                            Val = new EnumValue<BorderValues> {Value = BorderValues.Single},
                            Size = UInt32Value.FromUInt32(0),
                            Space = UInt32Value.FromUInt32(0),
                            Color = StringValue.FromString("FFFFFF")
                        }
                    }
                }
            };

            AddMainLogo(tableCell, relationshipId);

            tableRow.AppendChild(tableCell);

            table.AppendChild(tableRow);
        }

        private static void FillHeaderTableGrid(OpenXmlElement table)
        {
            var tableGrid = new TableGrid();
            tableGrid.AppendChild(new GridColumn
            {
                Width = StringValue.FromString("2000")
            });
            tableGrid.AppendChild(new GridColumn
            {
                Width = StringValue.FromString("2550")
            });
            tableGrid.AppendChild(new GridColumn
            {
                Width = StringValue.FromString("3825")
            });
            table.AppendChild(tableGrid);
        }

        private static void FillHeaderTableProperties(OpenXmlElement table)
        {
            var tableBorders = new TableBorders
            {
                TopBorder = new TopBorder
                {
                    Color = "000000",
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = new UInt32Value((uint)10),
                    Space = new UInt32Value((uint)0)
                },
                LeftBorder = new LeftBorder
                {
                    Color = "000000",
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = new UInt32Value((uint)10),
                    Space = new UInt32Value((uint)0)
                },
                BottomBorder = new BottomBorder
                {
                    Color = "000000",
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = new UInt32Value((uint)10),
                    Space = new UInt32Value((uint)0)
                },
                RightBorder = new RightBorder
                {
                    Color = "000000",
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = new UInt32Value((uint)10),
                    Space = new UInt32Value((uint)0)
                },
                InsideHorizontalBorder = new InsideHorizontalBorder
                {
                    Color = "000000",
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = new UInt32Value((uint)10),
                    Space = new UInt32Value((uint)0)
                },
                InsideVerticalBorder = new InsideVerticalBorder
                {
                    Color = "000000",
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = new UInt32Value((uint)10),
                    Space = new UInt32Value((uint)0)
                }
            };
            var tableCellMargins = new TableCellMarginDefault
            {
                TableCellLeftMargin = new TableCellLeftMargin
                {
                    Width = 10,
                    Type = new EnumValue<TableWidthValues> { Value = TableWidthValues.Dxa }
                },
                TableCellRightMargin = new TableCellRightMargin
                {
                    Width = 10,
                    Type = new EnumValue<TableWidthValues> { Value = TableWidthValues.Dxa }
                }
            };

            var tableProperties = new TableProperties
            {
                TableWidth = new TableWidth
                {
                    Width = "0",
                    Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Auto }
                },
                TableIndentation = new TableIndentation
                {
                    Width = new Int32Value(10),
                    Type = new EnumValue<TableWidthUnitValues> { Value = TableWidthUnitValues.Auto }
                },
                TableBorders = tableBorders,
                TableCellMarginDefault = tableCellMargins,
                TableLook = new TableLook
                {
                    Val = HexBinaryValue.FromString("0000"),
                    FirstRow = OnOffValue.FromBoolean(false),
                    LastRow = OnOffValue.FromBoolean(false),
                    FirstColumn = OnOffValue.FromBoolean(false),
                    LastColumn = OnOffValue.FromBoolean(false),
                    NoHorizontalBand = OnOffValue.FromBoolean(false),
                    NoVerticalBand = OnOffValue.FromBoolean(false)
                }
            };

            table.AppendChild(tableProperties);
        }
    }
}
