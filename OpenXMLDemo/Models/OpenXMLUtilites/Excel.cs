using System;
using System.Collections.Generic;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Globalization;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLDemo.Models.OpenXMLUtilites
{
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
                    Fonts = new Fonts(new Font()),
                    Fills = new Fills(new Fill()),
                    Borders = new Borders(new Border()),
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
        public static void InsertData(List<string> HeaderData, List<List<string>> DataLines)
        {
            uint rowIndex = 1;
            string filepath = $@"{Constants.Locations.DesktopPath}\{Constants.FTP.ExcelFile}";

            HeaderData.Insert(0, "UniqueId");
            HeaderData.Add("BookAge");

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
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
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
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
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
