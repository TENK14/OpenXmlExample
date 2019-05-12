using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OpenXmlExample
{
    public class ExcelService
    {

        public void CreateExcel(Stream stream)
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();


                #region Styles
                /**
                WorkbookStylesPart sp = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                // Step 1: Initial Stylesheet class.
                sp.Stylesheet = new Stylesheet();
                Stylesheet stylesheet = sp.Stylesheet;
                sp.Stylesheet.NumberingFormats = new NumberingFormats();
                //Step 2: Create a CellFormat Object and add this to stylesheet's CellFormats property.
                stylesheet.CellFormats = new CellFormats();
                stylesheet.CellFormats.Count = 2;
                // #.##% is also Excel style index 1
                NumberingFormat nf2decimal = new NumberingFormat();
                nf2decimal.NumberFormatId = UInt32Value.FromUInt32(3453);
                nf2decimal.FormatCode = StringValue.FromString("0.0%");
                sp.Stylesheet.NumberingFormats.Append(nf2decimal);
                var nformat4Decimal = new NumberingFormat
                {
                    NumberFormatId = UInt32Value.FromUInt32(999),// iExcelIndex++),
                    FormatCode = StringValue.FromString("#,##0.0000")
                };
                sp.Stylesheet.NumberingFormats.Append(nformat4Decimal);
                CellFormat cf1 = stylesheet.CellFormats.AppendChild(new CellFormat());
                cf1.FontId = 0;
                cf1.FillId = 0;
                cf1.BorderId = 0;
                cf1.FormatId = 0;
                //cf1.NumberFormatId = nf2decimal.NumberFormatId;
                cf1.ApplyNumberFormat = BooleanValue.FromBoolean(true);
                cf1.ApplyFont = true;
                //set cf1's property
                CellFormat cf2 = stylesheet.CellFormats.AppendChild(new CellFormat());
                cf2.FontId = 0;
                cf2.FillId = 0;
                cf2.BorderId = 0;
                cf2.FormatId = 0;
                //cf2.NumberFormatId = nf2decimal.NumberFormatId;
                cf2.ApplyNumberFormat = BooleanValue.FromBoolean(true);
                cf2.ApplyFont = true;
                //set cf2's property
                spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Save();
                /**/

                // Adding style
                WorkbookStylesPart stylePart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                //stylePart.Stylesheet = GenerateStylesheet();
                stylePart.Stylesheet = GenerateNumberingStylesheet();
                stylePart.Stylesheet.Save();
                #endregion Styles

                uint sheetId = 1;
                //spreadsheetDocument.WorkbookPart.Workbook.Sheets = new Sheets();
                //Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                WorksheetPart wsPart = workbookpart.AddNewPart<WorksheetPart>();
                wsPart.Worksheet = new Worksheet();
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(wsPart),
                    SheetId = sheetId++,
                    Name = "Details"
                };
                sheets.Append(sheet);

                //SheetData sheetData = new SheetData();
                //wsPart.Worksheet = new Worksheet(sheetData);
                SheetData sheetData = wsPart.Worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.SheetData());
                //MergeCells(wsPart.Worksheet, "A1", "L1");

                sheetData.AppendChild(new Row(new Cell()
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue("Details of Vikram"),
                    StyleIndex = 0
                },
                new Cell()
                {
                    DataType = CellValues.Number,
                    CellValue = new CellValue((23000).ToString()),
                    StyleIndex = 1//4//2
                },
                new Cell()
                {
                    DataType = CellValues.Number,
                    //CellValue = new CellValue((0.0219M).ToString()), // 2.19%
                    //CellValue = new CellValue((0.0219M).ToString().Replace('.', ',')), // 2.19%
                    //CellValue = new CellValue((0.0219M).ToString().Replace(',', '.')), // 2.19%
                    CellValue = new CellValue((1.1M / 100).ToString().Replace(',', '.')), // 2.19%
                    //CellValue = new CellValue((2).ToString()), // 2.19%
                    //CellValue = new CellValue("0.036"),
                    //StyleIndex = 3
                    //StyleIndex = 10
                    //StyleIndex = 0
                    //StyleIndex = 1 // ok
                    StyleIndex = 2
                    //StyleIndex = 4
                },
                new Cell()
                {
                    //DataType = CellValues.Date,
                    DataType = null, //CellValues.String,
                    //CellValue = new CellValue(DateTime.Now.ToShortDateString()),
                    //CellValue = new CellValue(new DateTime(2018, 4, 1)), // "43556"
                    //CellValue = new CellValue(new DateTime(2018, 4, 1).ToShortDateString()), // 43191, "43556"
                    CellValue = new CellValue(new DateTime(2018, 4, 1).ToOADate().ToString(CultureInfo.InvariantCulture)), // "43556"
                    StyleIndex = 3
                }));

                ////Next, add your cell, setting the StyleIndex to your new numbering format index:
                ////Where r = new Row() and row.LeaseStartDate = date in  ToOADate() format:
                //r.AppendChild(new Cell()
                //{
                //    CellReference = "D" + Convert.ToString(idx),
                //    CellValue = new CellValue() { Text = row.LeaseStartDate },
                //    StyleIndex = Convert.ToUInt32(sIndex)
                //});

                wsPart.Worksheet.Save();
                workbookpart.Workbook.Save();
            }
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }

                ),
                new Font(
                    new FontSize() { Val = 12 },
                    new Bold()

                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
                    { PatternType = PatternValues.Solid }) // Index 2 - header
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true }, // header
                    // Here is the CellFormat object you will need to create in order to have the 0.00% mask applied to your number. 
                    // In this case you want the predefined format number 10 or 0.00%:
                    new CellFormat()
                    {
                        NumberFormatId = (UInt32Value)10U,
                        FontId = (UInt32Value)2U,
                        FillId = (UInt32Value)0U,
                        BorderId = (UInt32Value)0U,
                        FormatId = (UInt32Value)0U,
                        ApplyNumberFormat = true
                    }
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }

        private Stylesheet GenerateNumberingStylesheet()
        {
            var stylesheet = new Stylesheet();
            // Create a numberingformat,

            stylesheet.NumberingFormats = new NumberingFormats();
            // #.##% is also Excel style index 1

            //uint iExcelIndex = 164;
            NumberingFormat nf2decimal = new NumberingFormat()
            {
                NumberFormatId = UInt32Value.FromUInt32(3453),
                FormatCode = StringValue.FromString("0.0%")
            };
            stylesheet.NumberingFormats.Append(nf2decimal);

            var nformat4Decimal = new NumberingFormat
            {
                //NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                NumberFormatId = UInt32Value.FromUInt32(3500),
                FormatCode = StringValue.FromString("#,##0.0000")
            };
            stylesheet.NumberingFormats.Append(nformat4Decimal);

            var dateFormat = new NumberingFormat()
            {
                NumberFormatId = (UInt32Value)4000,
                FormatCode = StringValue.FromString("dd.mm.yyyy")

            };
            stylesheet.NumberingFormats.Append(dateFormat);


            stylesheet.Fonts = new Fonts(
                new Font(),
                new Font(
                    new FontSize() { Val = 10 },
                    new Bold()
                )
            );
            // Create a cell format and apply the numbering format id

            var cellFormat = new CellFormat();
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.NumberFormatId = nf2decimal.NumberFormatId;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.ApplyFont = true;

            //append cell format for cells of header row
            stylesheet.CellFormats = new CellFormats();
            stylesheet.CellFormats.AppendChild<CellFormat>(cellFormat);

            cellFormat = new CellFormat();
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.NumberFormatId = nformat4Decimal.NumberFormatId;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.ApplyFont = true;
            //append cell format for cells of header row
            stylesheet.CellFormats.AppendChild<CellFormat>(cellFormat);

            // Percentage
            cellFormat = new CellFormat();
            cellFormat.FontId = 1;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.NumberFormatId = 10;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.ApplyFont = true;
            //append cell format for cells of header row
            stylesheet.CellFormats.AppendChild<CellFormat>(cellFormat);

            // Date
            cellFormat = new CellFormat();
            cellFormat.FontId = 1;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.NumberFormatId = dateFormat.NumberFormatId;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.ApplyFont = true;
            //append cell format for cells of header row
            stylesheet.CellFormats.AppendChild<CellFormat>(cellFormat);

            //update font count 
            stylesheet.CellFormats.Count = UInt32Value.FromUInt32((uint)stylesheet.CellFormats.ChildElements.Count);

            return stylesheet;
        }
    }
}
