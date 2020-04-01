using System;
using System.Collections.Generic;
using System.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXMLDemo.Models.OpenXMLUtilites
{
    public class WordDoc
    {
        public static void CreateWordprocessingDocument(List<string> HeaderData,List<List<string>> DataLines)
        {
            List<string> CsvData = new List<string>();
            

            string filepath = $@"{Constants.Locations.DesktopPath}\{Constants.FTP.DocFile}";
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                Paragraph newPara = new Paragraph(new Run(new Text("This Document is created by Kashish Jhaveri.")));

                body.Append(newPara);

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                FTP.DownloadFile(Constants.FTP.ImageUrl, Constants.Locations.DesktopPath + "\\" + Constants.Locations.ImageFile);
                using (FileStream stream = new FileStream(Constants.Locations.DesktopPath + "\\" + Constants.Locations.ImageFile, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart));

                for (int i = 0; i < DataLines.Count; i++)
                {
                    newPara = new Paragraph(new Run
                            (new Break() { Type = BreakValues.Page }));

                    body.Append(newPara);

                    CsvData = DataLines[i];
                    for (int j = 0; j < HeaderData.Count; j++)
                    {
                        Paragraph Para = new Paragraph(new Run(new Text($"\n {HeaderData[j]} : {CsvData[j]} \n")));
                        body.Append(Para);
                    }
                }

            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }
    }

}
