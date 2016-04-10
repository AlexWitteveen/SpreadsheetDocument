using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using OpenXml = DocumentFormat.OpenXml;
using Package = DocumentFormat.OpenXml.Packaging;
using Excel = DocumentFormat.OpenXml.Spreadsheet;

namespace ALEX.Library.SpreadsheetDocument
{
    internal class ExcelReader
    {
        public static void Read(Spreadsheet document, string path)
        {
            document.Clear();

            var buffer = new MemoryStream();
            var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            stream.CopyTo(buffer);
            stream.Close();

            var importDocument = Package.SpreadsheetDocument.Open(buffer, false);
            var importWorkbook = new Excel.Workbook();
            importWorkbook.Load(importDocument.WorkbookPart);

            var sharedStringTable = ReadSharedStringTable(importDocument);
            var cellFormats = ReadStyleSheet(document, importDocument);

            foreach (Excel.Sheet importSheet in importWorkbook.Sheets)
            {
                ReadSheet(document, importDocument, importSheet, sharedStringTable, cellFormats);
            }
            importDocument.Close();
        }

        static Dictionary<string, string> ReadSharedStringTable(Package.SpreadsheetDocument importDocument)
        {
            var result = new Dictionary<string, string>();

            var importSharedStringTable = new Excel.SharedStringTable();
            if (importDocument.WorkbookPart.SharedStringTablePart != null)
                importSharedStringTable.Load(importDocument.WorkbookPart.SharedStringTablePart);

            int index = 0;
            foreach (var item in importSharedStringTable.Elements<Excel.SharedStringItem>())
            {
                result.Add(index.ToString(), item.Text != null ? item.Text.Text : null);
                index++;
            }
            return result;
        }

        static Dictionary<int, CellFormat> ReadStyleSheet(Spreadsheet document, Package.SpreadsheetDocument importDocument)
        {
            Dictionary<int, CellFormat> cellFormats = new Dictionary<int, CellFormat>();

            var importedStyleSheet = new Excel.Stylesheet();
            if (importDocument.WorkbookPart.WorkbookStylesPart != null)
                importedStyleSheet.Load(importDocument.WorkbookPart.WorkbookStylesPart);

            int index = 0;
            foreach (Excel.CellFormat item in importedStyleSheet.CellFormats)
            {
                if (item.ApplyFill != null && item.ApplyFill.Value && item.FillId != null)
                {
                    int fillID = (int)(uint)item.FillId.Value;
                    var fill = (Excel.Fill)importedStyleSheet.Fills.ElementAt(fillID);

                    if (fill.PatternFill.ForegroundColor != null)
                    {
                        string fillColor = fill.PatternFill.ForegroundColor.Rgb;
                        cellFormats.Add(index, document.CellFormats().CellFormat(fillColor));
                    }
                }
                index++;
            }
            return cellFormats;
        }

        static void ReadSheet(Spreadsheet document, Package.SpreadsheetDocument importDocument, Excel.Sheet importSheet,
            Dictionary<string, string> sharedStringTable, Dictionary<int, CellFormat> cellFormats)
        {
            Sheet sheet = document.Sheets.Sheet(importSheet.Name);
            var importWorksheet = new Excel.Worksheet();
            importWorksheet.Load((Package.WorksheetPart)importDocument.WorkbookPart.GetPartById(importSheet.Id));
            var sheetData = (Excel.SheetData)importWorksheet.Elements<Excel.SheetData>().First();

            foreach (var importColumn in sheetData.Elements<Excel.Column>())
                ReadColumn();

            foreach (var importRow in sheetData.Elements<Excel.Row>())
                ReadRow(sheet, importRow, sharedStringTable, cellFormats);
        }

        static void ReadColumn()
        {
        }

        static void ReadRow(Sheet sheet, Excel.Row importRow, Dictionary<string, string> sharedStringTable, Dictionary<int, CellFormat> cellFormats)
        {
            Row row = sheet.Row((int)(uint)importRow.RowIndex-1);
            foreach (var importCell in importRow.Elements<Excel.Cell>())
                ReadCell(row, importCell, sharedStringTable, cellFormats);
        }

        static void ReadCell(Row row, Excel.Cell importCell, Dictionary<string, string> sharedStringTable, Dictionary<int, CellFormat> cellFormats)
        {
            string cellIndex = importCell.CellReference;
            string columnIndex = cellIndex.Substring(0, cellIndex.IndexOfAny("0123456789".ToCharArray())).ToUpper();
            int index = 0;
            while (columnIndex.Length > 1)
            {
                index += ((byte)columnIndex[0]) - 64;
                index = index * 26;
                columnIndex = columnIndex.Substring(1);
            }
            index += ((byte)columnIndex[0]) - 65;
            Column column = row._rows._sheet.Column(index);

            Cell cell = row.Cell(column);

            if (importCell.CellValue != null)
            {
                if (importCell.DataType == null)
                    importCell.DataType = new DocumentFormat.OpenXml.EnumValue<Excel.CellValues>(Excel.CellValues.Number);
                switch (importCell.DataType.Value)
                {
                    case DocumentFormat.OpenXml.Spreadsheet.CellValues.String:
                        cell.StringValue = importCell.CellValue.Text;
                        break;
                    case DocumentFormat.OpenXml.Spreadsheet.CellValues.Number:
                        cell.NumberValue = Convert.ToDouble(importCell.CellValue.Text);
                        break;
                    case DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString:
                        cell.StringValue = sharedStringTable[importCell.CellValue.Text];
                        break;
                    case DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean:
                        cell.BooleanValue = importCell.CellValue.Text == "1" ? true : false;
                        break;
                    case DocumentFormat.OpenXml.Spreadsheet.CellValues.Error:
                        break;
                    default:
                        throw new Exception("Unsupported type");
                }

                if (importCell.StyleIndex!=null&& importCell.StyleIndex >= 0 && cellFormats.ContainsKey((int)(uint)importCell.StyleIndex))
                    cell.CellFormat = cellFormats[(int)(uint)importCell.StyleIndex];
            }
        }
    }
}
