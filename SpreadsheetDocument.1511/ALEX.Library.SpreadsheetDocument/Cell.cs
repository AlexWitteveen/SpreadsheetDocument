using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace ALEX.Library.SpreadsheetDocument
{
    public enum ExcelValueType { String, Number , Boolean, Null};

    public class Cell
    {
        Sheet _sheet = null;
        Column _column = null;
        Row _row = null;
        ExcelValueType _type = ExcelValueType.Null;
        string _stringValue = null;
        double _numberValue = double.NaN;
        bool _boolValue = false;
        CellFormat _cellFormat = null;

        internal Cell(Sheet sheet, Column column, Row row)
        {
            _sheet = sheet;
            _column = column;
            _row = row;
        }

        internal void Deleted()
        {
            if (_cellFormat != null)
                _cellFormat.RemoveCell(this);
            _cellFormat = null;
            _sheet = null;
            _column = null;
            _row = null;
        }

        public Sheet Sheet
        {
            get { return _sheet; }
        }

        public Column Column
        {
            get { return _column; }
        }

        public Row Row
        {
            get { return _row; }
        }

        public ExcelValueType Type
        {
            get { return _type; }
        }

        public string StringValue
        {
            get
            {
                switch (_type)
                {
                    case ExcelValueType.String:
                        return _stringValue;
                    case ExcelValueType.Null:
                        return null;
                    case ExcelValueType.Number:
                        return _numberValue.ToString();
                    case ExcelValueType.Boolean:
                        return _boolValue.ToString();
                    default:
                        throw new Exception("Unsupported type");
                }
            }
            set
            {
                _type = ExcelValueType.String;
                _stringValue = value;
                _numberValue = double.NaN;
                _boolValue = false;
            }
        }

        public double NumberValue
        {
            get
            {
                Debug.Assert(_type == ExcelValueType.Number);
                return _numberValue;
            }
            set
            {
                _type = ExcelValueType.Number;
                _numberValue = value;
                _stringValue = null;
                _boolValue = false;
            }
        }

        public bool BooleanValue
        {
            get
            {
                Debug.Assert(_type == ExcelValueType.Boolean);
                return _boolValue;
            }
            set
            {
                _type = ExcelValueType.Boolean;
                _boolValue = value;
                _numberValue = double.NaN;
                _stringValue = null;
            }
        }

        public CellFormat CellFormat 
        { 
            get 
            { 
                return _cellFormat; 
            } 
            set 
            {
                if (_cellFormat != null) _cellFormat.RemoveCell(this);
                _cellFormat = value;
                if (_cellFormat != null) _cellFormat.AddCell(this);
            } 
        }

        public void Clear()
        {
            _type = ExcelValueType.Null;
            _numberValue = double.NaN;
            _stringValue = null;
            _boolValue = false;
            _cellFormat = null;
        }
    }

    public class Cells : List<Cell>
    {
        Sheet _sheet = null;
        Column _column = null;
        Row _row = null;

        internal Cells(Sheet sheet, Column column, Row row)
        {
            _sheet = sheet;
            _column = column;
            _row = row;
        }

        internal void Deleted()
        {
            this.Clear();
            _sheet = null;
            _column = null;
            _row = null;
        }

/*
        public new void Clear()
        {
            foreach (var cell in this)
                cell.Deleted();
            base.Clear();
        }
*/
        internal void Clear(Cells cells)
        {
            for (int i=cells.Count-1;i>=0;i--)
            {
                var cell = cells[i];
                if (this.Contains(cell))
                {
                    cell.Deleted();
                    this.Remove(cell);
                }
            }
//            base.Clear();
        }

        public Cell Cell(Sheet sheet, Column column, Row row)
        {
            if (this.Count(r => r.Sheet == sheet && r.Column == column && r.Row == row) == 0)
            {
                Cell cell = new Cell(sheet, column, row);
                sheet.Cells.Add(cell);
                column._cells.Add(cell);
                row._cells.Add(cell);
            }
            return this.First(r => r.Sheet == sheet && r.Column == column && r.Row == row);
        }

        public Cell Cell(int columnIndex, int rowIndex)
        {
            Debug.Assert(_sheet != null);
            var column = _sheet.Columns.Column(columnIndex);
            var row = _sheet.Rows.Row(rowIndex);
            return Cell(column, row);
        }

        public Cell Cell(Column column, int rowIndex)
        {
            var row = column._columns._sheet.Rows.Row(rowIndex);
            return Cell(column, row);
        }

        public Cell Cell(int columnIndex, Row row)
        {
            var column = row._rows._sheet.Columns.Column(columnIndex);
            return Cell(column, row);
        }

        public Cell Cell(Column column, Row row)
        {
            return Cell(column._columns._sheet, column, row);
        }

        public Cell Cell(Column column)
        {
            Debug.Assert(_row != null);
            return Cell(_row._rows._sheet, column, _row);
        }

        public Cell Cell(Row row)
        {
            Debug.Assert(_column != null);
            return Cell(_column._columns._sheet, _column, row);
        }
    }
}
