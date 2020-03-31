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
using WRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WBreak = DocumentFormat.OpenXml.Wordprocessing.Break;
using WText = DocumentFormat.OpenXml.Spreadsheet.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using EText = DocumentFormat.OpenXml.Spreadsheet.Text;
using System.Globalization;
using EFonts = DocumentFormat.OpenXml.Spreadsheet.Fonts;
using EFont = DocumentFormat.OpenXml.Spreadsheet.Font;
using EBorders = DocumentFormat.OpenXml.Spreadsheet.Borders;
using EBorder = DocumentFormat.OpenXml.Spreadsheet.Border;

namespace OpenXMLDemo.Models
{
    class OpenXMLUtilites
    {
        public class WordDoc
        {
            public static void CreateWordprocessingDocument(List<List<string>> DataLines)
            {
                List<string> CsvData = new List<string>();
                List<string> HeaderData = DataLines.First();
                DataLines.RemoveAt(0);

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

                    Paragraph newPara = new Paragraph(new WRun(new WText("This Document is created by Kashish Jhaveri.")));

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
                        newPara = new Paragraph(new WRun
                                (new WBreak() { Type = BreakValues.Page }));

                        body.Append(newPara);

                        CsvData = DataLines[i];
                        for (int j = 0; j < HeaderData.Count; j++)
                        {
                            Paragraph Para = new Paragraph(new WRun(new WText($"\n {HeaderData[j]} : {CsvData[j]} \n")));
                            body.Append(Para);
                        }
                    }

                }
            }

            private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
            {
                // Define the reference of the image.
                var element =
                     new DocumentFormat.OpenXml.Office.Drawing.Drawing(
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
                wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new WRun(element)));
            }
        }

        public class Excel
        {
            private static object shareStringPart;

            public static void CreateSpreadsheetWorkbook()
            {
                string filepath = $@"{Constants.Locations.DesktopPath}\{Constants.FTP.ExcelFile}";

                // Create a spreadsheet document by supplying the filepath.
                // By default, AutoSave = true, Editable = true, and Type = xlsx.
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
                {
                    // Add a WorkbookPart to the document.
                    WorkbookPart workbookpart = spreadSheet.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();

                    // Add a WorksheetPart to the WorkbookPart.
                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    // Add Sheets to the Workbook.
                    Sheets sheets = spreadSheet.WorkbookPart.Workbook.
                        AppendChild<Sheets>(new Sheets());

                    // Append a new worksheet and associate it with the workbook.
                    Sheet sheet = new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart.
                        GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    };
                    sheets.Append(sheet);

                    // Add minimal Stylesheet
                    var stylesPart = spreadSheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = new Stylesheet
                    {
                        Fonts = new EFonts(new EFont()),
                        Fills = new Fills(new Fill()),
                        Borders = new EBorders(new EBorder()),
                        CellStyleFormats = new CellStyleFormats(new CellFormat()),
                        CellFormats =
                            new CellFormats(
                                new CellFormat(),
                                new CellFormat
                                {
                                    NumberFormatId = 14,
                                    ApplyNumberFormat = true
                                })
                    };

                    workbookpart.Workbook.Save();
                    

                    Cell cell = InsertCellInWorksheet("A", 2, worksheetPart);

                    // Set the value of cell A1.
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    cell.CellValue = new CellValue("My Name is Kashish Jhaveri");


                    // Save the new worksheet.
                    worksheetPart.Worksheet.Save();
                }
            }

            //InsertText in Exsting excel file
            public static void InsertData(List<List<string>> DataLines)
            {
                uint rowIndex = 1;
                string filepath = $@"{Constants.Locations.DesktopPath}\{Constants.FTP.ExcelFile}";
                List<string> HeaderData = DataLines.First();
                HeaderData.Insert(0, "UniqueId");
                HeaderData.Add("BookAge");
                DataLines.RemoveAt(0);

                // Open the document for editing.
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filepath, true))
                {
                    //UniqueId(generate a Guid from your application using Guid.NewGuid() in C#)
                    // Insert a new worksheet.
                    WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);


                    char colname = 'A';
                    Cell cell;
                    foreach (var col in HeaderData)
                    {
                        cell = InsertCellInWorksheet(Convert.ToString(colname), rowIndex, worksheetPart);
                        cell.CellValue = new CellValue(col);
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        colname++;
                    }
                    foreach (List<string> CsvRow in DataLines)
                    {
                        if (rowIndex > 100)
                        {
                            break;
                        }
                        rowIndex++;
                        cell = InsertCellInWorksheet("A", rowIndex, worksheetPart);
                        cell.CellValue = new CellValue(Guid.NewGuid().ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);                        
                        
                        cell = InsertCellInWorksheet("B", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        cell.CellValue = new CellValue(CsvRow[0]);
                        cell = InsertCellInWorksheet("C", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        cell.CellValue = new CellValue(CsvRow[1]);
                        cell = InsertCellInWorksheet("D", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        cell.CellValue = new CellValue(CsvRow[2]);
                        cell = InsertCellInWorksheet("E", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cell.CellValue = new CellValue(CsvRow[3]);
                        cell = InsertCellInWorksheet("F", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cell.CellValue = new CellValue(CsvRow[4]);
                        cell = InsertCellInWorksheet("G", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        cell.CellValue = new CellValue(CsvRow[5]);
                        cell = InsertCellInWorksheet("H", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cell.CellValue = new CellValue(CsvRow[6]);
                        cell = InsertCellInWorksheet("I", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        cell.CellValue = new CellValue(CsvRow[7]);
                        cell = InsertCellInWorksheet("J", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cell.CellValue = new CellValue(CsvRow[8]);
                        cell = InsertCellInWorksheet("K", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cell.CellValue = new CellValue(CsvRow[9]);
                        cell = InsertCellInWorksheet("L", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                        cell.CellValue = new CellValue(Convert.ToDateTime(CsvRow[10]).ToOADate().ToString(CultureInfo.InvariantCulture));
                        cell = InsertCellInWorksheet("M", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        cell.CellValue = new CellValue(CsvRow[11]);

                        //Calculated Age of book
                        cell = InsertCellInWorksheet("N", rowIndex, worksheetPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cell.CellFormula = new CellFormula($"=INT((TODAY()-L{rowIndex})/365)");
                            
                    }
                    

                    // Save the new worksheet.
                    worksheetPart.Worksheet.Save();
                    spreadSheet.Close();
                }
            }

            //Insert New Worksheet
            private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
            {
                // Add a new worksheet part to the workbook.
                WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                newWorksheetPart.Worksheet.Save();

                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

                // Get a unique ID for the new sheet.
                uint sheetId = 1;

                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }



                string sheetName = "Sheet" + sheetId;

                // Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);
                workbookPart.Workbook.Save();

                return newWorksheetPart;
            }

            //Create new cell and insert it into excel file
            private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
            {
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                string cellReference = columnName + rowIndex;

                // If the worksheet does not contain a row with the specified row index, insert one.
                Row row;
                if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
                {
                    row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
                }
                else
                {
                    row = new Row() { RowIndex = rowIndex };
                    sheetData.Append(row);
                }

                // If there is not a cell with the specified column name, insert one.  
                if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
                {
                    return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
                }
                else
                {
                    // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                    Cell refCell = null;
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }

                    Cell newCell = new Cell() { CellReference = cellReference };
                    row.InsertBefore(newCell, refCell);

                    worksheet.Save();
                    return newCell;
                }
            }

        }

    }
}
