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

                }

                return mem.ToArray();
            }
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

            body.AppendChild(table);
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
