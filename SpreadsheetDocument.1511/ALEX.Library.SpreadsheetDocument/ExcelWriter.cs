using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OpenXml = DocumentFormat.OpenXml;
using Package = DocumentFormat.OpenXml.Packaging;
using Excel = DocumentFormat.OpenXml.Spreadsheet;

namespace ALEX.Library.SpreadsheetDocument
{
    internal class ExcelWriter
    {
        public static void Write(Spreadsheet document, string path)
        {
            var exportedDocument = Package.SpreadsheetDocument.Create(path, OpenXml.SpreadsheetDocumentType.Workbook);
            var exportedWorkbookPart = exportedDocument.AddWorkbookPart();

            Dictionary<CellFormat, uint> cellFormatList = null;
            Excel.Stylesheet exportedStyleSheet = SaveStyleSheet(exportedWorkbookPart, ref cellFormatList, document);

            SaveDocument(exportedWorkbookPart, exportedStyleSheet, cellFormatList, document);
/*
            var exportedSharedStringTablePart = exportedWorkbookPart.AddNewPart<Package.SharedStringTablePart>();
            exportedSharedStringTablePart.SharedStringTable = SaveSharedStringTable();
            exportedSharedStringTablePart.SharedStringTable.Save();
*/

            exportedWorkbookPart.Workbook.Save();
            exportedDocument.Close();
        }

        static Excel.Stylesheet SaveStyleSheet(Package.WorkbookPart exportedWorkbookPart, ref Dictionary<CellFormat, uint> cellFormatList, Spreadsheet document)
        {
            var exportedStyleSheetPart = exportedWorkbookPart.AddNewPart<Package.WorkbookStylesPart>();
            cellFormatList = new Dictionary<CellFormat, uint>();
            Excel.Stylesheet exportedStyleSheet = new Excel.Stylesheet();
            exportedStyleSheetPart.Stylesheet = exportedStyleSheet;
            exportedStyleSheet.CellFormats = new Excel.CellFormats();
            exportedStyleSheet.CellFormats.Append(new Excel.CellFormat() { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 });
            exportedStyleSheet.Fills = new Excel.Fills();
            exportedStyleSheet.Fills.Append(new Excel.Fill(new Excel.PatternFill() 
            { 
                PatternType = new OpenXml.EnumValue<Excel.PatternValues>(Excel.PatternValues.None) 
            }));

            exportedStyleSheet.Fills.Append(new Excel.Fill(new Excel.PatternFill() 
            {
                PatternType = new OpenXml.EnumValue<Excel.PatternValues>(Excel.PatternValues.Gray125) 
            }));
            exportedStyleSheet.Fonts = new Excel.Fonts();
            exportedStyleSheet.Fonts.Append(new Excel.Font()
            {
                FontSize = new Excel.FontSize() { Val = 11 },
                Color = new Excel.Color() { Theme = 1 },
                FontName = new Excel.FontName() { Val = "Calibri" },
                FontFamilyNumbering = new Excel.FontFamilyNumbering() { Val = 2 },
                FontScheme = new Excel.FontScheme() { Val = new OpenXml.EnumValue<Excel.FontSchemeValues>(Excel.FontSchemeValues.Minor) }
            });
            exportedStyleSheet.Borders = new Excel.Borders();
            exportedStyleSheet.Borders.Append(new Excel.Border()
            {
                LeftBorder = new Excel.LeftBorder(),
                RightBorder = new Excel.RightBorder(),
                TopBorder = new Excel.TopBorder(),
                BottomBorder = new Excel.BottomBorder(),
                DiagonalBorder = new Excel.DiagonalBorder()
            });
            exportedStyleSheet.CellStyleFormats = new Excel.CellStyleFormats();
            exportedStyleSheet.CellStyleFormats.Append(new Excel.CellFormat() { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 });
            exportedStyleSheet.CellStyles = new Excel.CellStyles();
            exportedStyleSheet.CellStyles.Append(new Excel.CellStyle() { Name = "Normal", FormatId = 0, BuiltinId = 0 });
            exportedStyleSheet.DifferentialFormats = new Excel.DifferentialFormats();
            exportedStyleSheet.TableStyles = new Excel.TableStyles() { DefaultTableStyle = "TableStyleMedium9", DefaultPivotStyle = "PivotStyleLight16" };

            foreach (var cellFormat in document.CellFormats())
            {
                if (cellFormat.Count() > 0)
                {
                    cellFormatList.Add(cellFormat, (uint)exportedStyleSheet.CellFormats.ChildElements.Count());
                    exportedStyleSheet.CellFormats.Append(new Excel.CellFormat()
                    {
                        NumberFormatId = 0,
                        FontId = 0,
                        BorderId = 0,
                        FormatId = 0,
                        FillId = (uint)exportedStyleSheet.Fills.ChildElements.Count(),
                        ApplyFill = true
                    });
                    exportedStyleSheet.Fills.Append(new Excel.Fill(new Excel.PatternFill()
                    {
                        PatternType = new OpenXml.EnumValue<Excel.PatternValues>(Excel.PatternValues.Solid),
                        ForegroundColor = new Excel.ForegroundColor() { Rgb = new OpenXml.HexBinaryValue(cellFormat.FillColor) }
                    }));
                }
            }

            exportedStyleSheet.Save();
            return exportedStyleSheet;
        }

        static void SaveDocument(Package.WorkbookPart exportedWorkbookPart, Excel.Stylesheet styleSheet, Dictionary<CellFormat, uint> cellFormatList, Spreadsheet document)
        {
            var exportedWorkbook = new Excel.Workbook();
            exportedWorkbookPart.Workbook = exportedWorkbook;

            Excel.Sheets exportedSheets = new Excel.Sheets();
            exportedWorkbook.AppendChild(exportedSheets);

            uint sheetId = 1;
            foreach (var sheet in document.Sheets)
                SaveSheet(exportedWorkbookPart, styleSheet, cellFormatList, exportedSheets, sheet, sheetId++);
        }

/*
        static Excel.SharedStringTable SaveSharedStringTable()
        {
            var exportedSharedStringTable = new Excel.SharedStringTable();

            return exportedSharedStringTable;
        }
*/

        static void SaveSheet(Package.WorkbookPart exportedWorkbookPart, Excel.Stylesheet styleSheet, Dictionary<CellFormat, uint> cellFormatList, Excel.Sheets exportedSheets, Sheet sheet, uint sheetId)
        {
            var exportedWorksheetPart = exportedWorkbookPart.AddNewPart<Package.WorksheetPart>();
            string relId = exportedWorkbookPart.GetIdOfPart(exportedWorksheetPart);

            var exportedWorksheet = new Excel.Worksheet();
            exportedWorksheetPart.Worksheet = exportedWorksheet;

            var exportedColumns = new Excel.Columns();
            exportedWorksheet.Append(exportedColumns);

            var exportedSheetData = new Excel.SheetData();
            exportedWorksheet.Append(exportedSheetData);

            var exportedSheet = new Excel.Sheet() { Name = sheet.Name, Id = relId, SheetId = sheetId };
            if (sheet.Hidden) exportedSheet.State = Excel.SheetStateValues.Hidden;
            exportedSheets.Append(exportedSheet);

            foreach (var column in sheet.Columns.OrderBy(r=>r.Index))
                SaveColumn(exportedColumns, column);

            foreach (var row in sheet.Rows.OrderBy(r=>r.Index))
                SaveRow(exportedSheetData, styleSheet, cellFormatList, row);

            exportedWorksheetPart.Worksheet.Save();
        }

        static void SaveColumn(Excel.Columns exportedColumns, Column column)
        {
            var exportedColumn = new Excel.Column() { Min = ColumnIndexNum(column), Max = ColumnIndexNum(column) };
            if (column._hidden) exportedColumn.Hidden = true;
            exportedColumn.Width = 5.0;
            exportedColumns.Append(exportedColumn);
        }

        static void SaveRow(Excel.SheetData exportedSheetData, Excel.Stylesheet styleSheet, Dictionary<CellFormat, uint> cellFormatList, Row row)
        {
            Excel.Row exportedRow = new Excel.Row() { RowIndex = RowIndex(row), Hidden=row._hidden};
            if (row._hidden) exportedRow.Hidden = true;
            exportedSheetData.Append(exportedRow);

            foreach (var cell in row._cells.OrderBy(r=>r.Column._index))
                SaveCell(exportedRow, styleSheet, cellFormatList, cell);
        }

        static void SaveCell(Excel.Row exportedRow, Excel.Stylesheet styleSheet, Dictionary<CellFormat, uint> cellFormatList, Cell cell)
        {
            Excel.Cell exportedCell = new Excel.Cell() { CellReference= CellIndex(cell.Column, cell.Row)};
            exportedCell.DataType = new OpenXml.EnumValue<Excel.CellValues>(Excel.CellValues.String);
            switch (cell.Type)
            {
                case ExcelValueType.String:
                    exportedCell.CellValue = new Excel.CellValue(cell.StringValue);
                    exportedCell.DataType = new OpenXml.EnumValue<Excel.CellValues>(Excel.CellValues.String);
                    break;
                case ExcelValueType.Number:
                    exportedCell.CellValue = new Excel.CellValue(cell.NumberValue.ToString());
                    exportedCell.DataType = new OpenXml.EnumValue<Excel.CellValues>(Excel.CellValues.Number);
                    break;
                case ExcelValueType.Boolean:
                    exportedCell.CellValue = new Excel.CellValue((cell.BooleanValue ? 1 : 0).ToString());
                    exportedCell.DataType = new OpenXml.EnumValue<Excel.CellValues>(Excel.CellValues.Boolean);
                    break;
                case ExcelValueType.Null:
                    break;
                default:
                    throw new Exception("Unsupported type");
            }

            if (cell.CellFormat != null)
            {
                exportedCell.StyleIndex =  cellFormatList[cell.CellFormat];
            }

            exportedRow.Append(exportedCell);
        }
        static uint ColumnIndexNum(Column column)
        {
            return (uint)(column.Index + 1);
        }

        static string ColumnIndexAlpha(Column column)
        {
            int index = column.Index;
            string result = "";
            result = Convert.ToChar(index % 26 + 65) + result;
            index /= 26;
            while (index > 0)
            {
                result = Convert.ToChar(index % 26 + 64) + result;
                index /= 26;
            }
            return result;
        }

        static uint RowIndex(Row row)
        {
            return (uint)(row.Index + 1);
        }

        static string CellIndex(Column column, Row row)
        {
            return ColumnIndexAlpha(column) + RowIndex(row).ToString();
        }
    }
}
