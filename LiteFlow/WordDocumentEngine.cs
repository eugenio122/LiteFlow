using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using LiteFlow.Models;

namespace LiteFlow
{
    public class WordDocumentEngine
    {
        private static string SanitizeForXml(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            text = text.Replace("\v", "\n");
            return Regex.Replace(text, @"[\x00-\x08\x0C\x0E-\x1F]", "");
        }

        public static void PrepareDocument(string templatePath, string outputPath, Dictionary<string, string> tags)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            if (!string.IsNullOrEmpty(templatePath) && File.Exists(templatePath))
            {
                File.Copy(templatePath, outputPath, true);

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
                {
                    var body = wordDoc.MainDocumentPart!.Document.Body;
                    foreach (var tag in tags)
                    {
                        string safeValue = SanitizeForXml(tag.Value);
                        foreach (var textNode in body!.Descendants<Text>().Where(t => t.Text.Contains(tag.Key)))
                        {
                            textNode.Text = textNode.Text.Replace(tag.Key, safeValue);
                        }
                    }
                    wordDoc.MainDocumentPart.Document.Save();
                }
            }
            else
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());
                    mainPart.Document.Save();
                }
            }
        }

        public static void AppendAllEvidence(List<EvidenceItem> items, string outputPath, LiteFlowProjectData projectConfig)
        {
            if (items == null || items.Count == 0) return;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart!;
                var body = mainPart.Document.Body!;

                int columns = 1;
                if (projectConfig.ReportLayout == LayoutMode.Compacto) columns = 2;
                else if (projectConfig.ReportLayout == LayoutMode.Mobile) columns = projectConfig.MobileColumns;

                if (columns <= 1)
                {
                    foreach (var item in items)
                    {
                        AppendSingleItem(mainPart, body, item, columns, projectConfig.ReportLayout);
                    }
                }
                else
                {
                    Table table = new Table();

                    int totalWidthTwips = 8800;
                    int cellWidthTwips = totalWidthTwips / columns;

                    TableProperties tblProp = new TableProperties(
                        new TableWidth() { Width = totalWidthTwips.ToString(), Type = TableWidthUnitValues.Dxa },
                        new TableLayout() { Type = TableLayoutValues.Fixed },
                        new TableBorders(
                            new TopBorder() { Val = BorderValues.None },
                            new BottomBorder() { Val = BorderValues.None },
                            new LeftBorder() { Val = BorderValues.None },
                            new RightBorder() { Val = BorderValues.None },
                            new InsideHorizontalBorder() { Val = BorderValues.None },
                            new InsideVerticalBorder() { Val = BorderValues.None }
                        ),
                        new TableCellMarginDefault(
                            new TopMargin() { Width = "115", Type = TableWidthUnitValues.Dxa },
                            new BottomMargin() { Width = "115", Type = TableWidthUnitValues.Dxa },
                            new LeftMargin() { Width = "115", Type = TableWidthUnitValues.Dxa },
                            new RightMargin() { Width = "115", Type = TableWidthUnitValues.Dxa }
                        )
                    );
                    table.AppendChild(tblProp);

                    TableGrid tableGrid = new TableGrid();
                    for (int i = 0; i < columns; i++) tableGrid.AppendChild(new GridColumn() { Width = cellWidthTwips.ToString() });
                    table.AppendChild(tableGrid);

                    for (int i = 0; i < items.Count; i += columns)
                    {
                        TableRow row = new TableRow();
                        row.AppendChild(new TableRowProperties(new CantSplit() { Val = OnOffOnlyValues.On }));

                        for (int j = 0; j < columns; j++)
                        {
                            TableCell cell = new TableCell();
                            cell.AppendChild(new TableCellProperties(
                                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = cellWidthTwips.ToString() },
                                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Top }
                            ));

                            if (i + j < items.Count)
                            {
                                AppendItemToCell(mainPart, cell, items[i + j], columns, projectConfig.ReportLayout);
                            }
                            else
                            {
                                cell.AppendChild(new Paragraph(new Run(new Text(""))));
                            }
                            row.AppendChild(cell);
                        }
                        table.AppendChild(row);
                    }
                    body.AppendChild(table);
                }
                mainPart.Document.Save();
            }
        }

        private static void AppendItemToCell(MainDocumentPart mainPart, OpenXmlElement parentElement, EvidenceItem item, int columns, LayoutMode layout)
        {
            var pNote = CreateNoteParagraph(item.Note, false);
            var pImg = CreateImageParagraph(mainPart, item, columns, layout, false);

            if (item.TextBelowImage)
            {
                if (pImg != null) parentElement.AppendChild(pImg);
                if (pNote != null) parentElement.AppendChild(pNote);
            }
            else
            {
                if (pNote != null) parentElement.AppendChild(pNote);
                if (pImg != null) parentElement.AppendChild(pImg);
            }

            parentElement.AppendChild(new Paragraph(new Run(new Text(""))));
        }

        private static void AppendSingleItem(MainDocumentPart mainPart, Body body, EvidenceItem item, int columns, LayoutMode layout)
        {
            bool noteNeedsKeepNext = !item.TextBelowImage;
            bool imgNeedsKeepNext = item.TextBelowImage;

            var pNote = CreateNoteParagraph(item.Note, noteNeedsKeepNext);
            var pImg = CreateImageParagraph(mainPart, item, columns, layout, imgNeedsKeepNext);

            if (item.TextBelowImage)
            {
                if (pImg != null) body.AppendChild(pImg);
                if (pNote != null) body.AppendChild(pNote);
            }
            else
            {
                if (pNote != null) body.AppendChild(pNote);
                if (pImg != null) body.AppendChild(pImg);
            }

            body.AppendChild(new Paragraph(new Run(new Text(""))));
        }

        private static Paragraph? CreateNoteParagraph(string note, bool keepNext)
        {
            if (string.IsNullOrWhiteSpace(note)) return null;

            string safeNote = SanitizeForXml(note);
            var textRun = new Run();
            textRun.AppendChild(new RunProperties(new Bold(), new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "404040" }, new FontSize() { Val = "20" }));

            string[] lines = safeNote.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                textRun.AppendChild(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
                if (i < lines.Length - 1) textRun.AppendChild(new Break());
            }

            var pProps = new ParagraphProperties();
            if (keepNext) pProps.AppendChild(new KeepNext() { Val = true });

            pProps.AppendChild(new SpacingBetweenLines() { Before = "100", After = "100" });
            pProps.AppendChild(new Justification() { Val = JustificationValues.Left });

            return new Paragraph(pProps, textRun);
        }

        // Carrega, grava no Word e descarta num instante. Nunca consome RAM extra!
        private static Paragraph? CreateImageParagraph(MainDocumentPart mainPart, EvidenceItem item, int columns, LayoutMode layout, bool keepNext)
        {
            Bitmap? bmpToUse = item.Image;
            bool disposeAfterUse = false;

            if (bmpToUse == null && !string.IsNullOrEmpty(item.DiskPath) && File.Exists(item.DiskPath))
            {
                using (var fs = new FileStream(item.DiskPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    bmpToUse = new Bitmap(Image.FromStream(fs));
                    disposeAfterUse = true;
                }
            }

            if (bmpToUse == null) return null;

            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (MemoryStream ms = new MemoryStream())
            {
                bmpToUse.Save(ms, ImageFormat.Png);
                ms.Position = 0;
                imagePart.FeedData(ms);
            }

            var element = CriarElementoImagem(mainPart.GetIdOfPart(imagePart), bmpToUse.Width, bmpToUse.Height, item.Note, columns, layout);

            if (disposeAfterUse) bmpToUse.Dispose();

            var pProps = new ParagraphProperties();
            if (keepNext) pProps.AppendChild(new KeepNext() { Val = true });
            pProps.AppendChild(new KeepLines() { Val = true });
            pProps.AppendChild(new Justification() { Val = JustificationValues.Center });

            return new Paragraph(pProps, new Run(element));
        }

        private static Drawing CriarElementoImagem(string relationshipId, int width, int height, string note, int columns, LayoutMode layout)
        {
            long baseWidthEMU = 5400000L;
            long maxWidthEMU = (baseWidthEMU / columns) - 150000L;
            long maxHeightEMU;

            if (layout == LayoutMode.Padrao)
            {
                bool isHorizontal = width > height;

                if (isHorizontal)
                {
                    maxHeightEMU = 5400000L;
                }
                else
                {
                    long alturaUtilFolha = 8200000L;
                    long alturaImagemPadrao = 7200000L;
                    long alturaImagemReduzidaPermitida = (long)(alturaImagemPadrao * 0.80);

                    int linhasTexto = 0;
                    if (!string.IsNullOrWhiteSpace(note))
                    {
                        var lines = note.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                        foreach (var line in lines) linhasTexto += 1 + (line.Length / 85);
                    }

                    long alturaDoTexto = (linhasTexto * 250000L) + 300000L;
                    long espacoDisponivelNaPagina = alturaUtilFolha - alturaDoTexto;

                    if (espacoDisponivelNaPagina >= alturaImagemReduzidaPermitida)
                        maxHeightEMU = Math.Min(alturaImagemPadrao, espacoDisponivelNaPagina);
                    else
                        maxHeightEMU = alturaImagemPadrao;
                }
            }
            else if (layout == LayoutMode.Compacto)
            {
                maxHeightEMU = 3400000L;
            }
            else
            {
                maxHeightEMU = 6500000L;
            }

            long originalWidthEMU = (long)(width * 9525);
            long originalHeightEMU = (long)(height * 9525);

            double ratioX = originalWidthEMU > maxWidthEMU ? (double)maxWidthEMU / originalWidthEMU : 1.0;
            double ratioY = originalHeightEMU > maxHeightEMU ? (double)maxHeightEMU / originalHeightEMU : 1.0;

            double finalRatio = Math.Min(ratioX, ratioY);
            long finalWidthEMU = (long)(originalWidthEMU * finalRatio);
            long finalHeightEMU = (long)(originalHeightEMU * finalRatio);

            string imageName = $"Evidencia_{Guid.NewGuid().ToString().Substring(0, 8)}";

            return new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = finalWidthEMU, Cy = finalHeightEMU },
                         new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                         new DW.DocProperties() { Id = (UInt32Value)1U, Name = imageName },
                         new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(new A.GraphicData(new PIC.Picture(
                             new PIC.NonVisualPictureProperties(new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = imageName + ".png" }, new PIC.NonVisualPictureDrawingProperties()),
                             new PIC.BlipFill(new A.Blip(new A.BlipExtensionList(new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" })) { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print }, new A.Stretch(new A.FillRectangle())),
                             new PIC.ShapeProperties(new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = finalWidthEMU, Cy = finalHeightEMU }), new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                         { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, EditId = "50D07946" });
        }
    }
}