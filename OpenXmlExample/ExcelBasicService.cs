using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OpenXmlExample
{
    public class ExcelBasicService
    {
        public void CreateExcelDoc(MemoryStream memoryStream)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Employees" };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                //List<Employee> employees = Employees.EmployeesList;

                DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData = worksheetPart.Worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

                // Constructing header
                Row row = new Row();

                row.Append(
                    ConstructCell("Id", CellValues.String),
                    ConstructCell("Name", CellValues.String),
                    ConstructCell("Birth Date", CellValues.String),
                    ConstructCell("Salary", CellValues.String));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Inserting each employee
                for (int i = 0; i < 5; i++)
                {
                    row = new Row();

                    row.Append(
                        ConstructCell(14.ToString(), CellValues.Number),
                        ConstructCell("Employee", CellValues.String),
                        ConstructCell(DateTime.Now.ToShortDateString(), CellValues.String),
                        ConstructCell((20000).ToString(), CellValues.Number));

                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
            }
        }

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }

        private static int CellReferenceToIndex(Cell cell)
        {
            int index = 0;
            string reference = cell.CellReference.ToString().ToUpper();
            foreach (char ch in reference)
            {
                if (Char.IsLetter(ch))
                {
                    int value = (int)ch - (int)'A';
                    index = (index == 0) ? value : ((index + 1) * 26) + value;
                }
                else
                {
                    return index;
                }
            }
            return index;
        }

        private void CheckForEmptyRow(WorkbookPart wbp, string sheetId)
        {
            Sheet sheet = wbp.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Id == sheetId);
            WorksheetPart wsp = (WorksheetPart)(wbp.GetPartById(sheetId));
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.SheetData> sheetData = wsp.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
            bool isRowEmpty = false;

            foreach (DocumentFormat.OpenXml.Spreadsheet.SheetData sd in sheetData)
            {
                IEnumerable<Row> row = sd.Elements<Row>(); // Get the row IEnumerator

                if (row == null)
                {
                    isRowEmpty = true;
                    break;
                }
            }

            if (isRowEmpty)
            {
                sheet.Remove();
                wbp.DeletePart(wsp);
            }
        }
    }
}
