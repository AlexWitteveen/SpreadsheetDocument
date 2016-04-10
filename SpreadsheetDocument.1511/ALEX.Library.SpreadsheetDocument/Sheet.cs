using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace ALEX.Library.SpreadsheetDocument
{
    public class Sheet
    {
        Sheets _sheets = null;
        string _name = null;
        internal Cells _cells = null;
        internal Columns _columns = null;
        internal Rows _rows = null;
        internal bool _hidden = false;

        internal Sheet(Sheets sheets, string name)
        {
            _sheets = sheets;
            _name = name;
            _cells = new Cells(this, null, null);
            _columns = new Columns(this);
            _rows = new Rows(this);
        }

        internal void Deleted()
        {
            _cells.Deleted();
            _columns.Deleted();
            _rows.Deleted();
            _sheets = null;
        }

        public void Clear()
        {
            Clear(_cells);
        }

        internal void Clear(Cells cells)
        {
            _cells.Clear(cells);
            _rows.Clear(cells);
            _columns.Clear(cells);
        }

        public string Name
        {
            get { return _name; }
        }

        public bool Hidden { get { return _hidden; } set { _hidden = value; } }

        public Column Column(int columnIndex)
        {
            return _columns.Column(columnIndex);
        }

        public Row Row(int rowIndex)
        {
            return _rows.Row(rowIndex);
        }

        public Cell Cell(int columnIndex, int rowIndex)
        {
            return _cells.Cell(columnIndex, rowIndex);
        }

        public Cell Cell(Column column, Row row)
        {
            return _cells.Cell(column, row);
        }

        internal Columns Columns
        {
            get { return _columns; }
        }

        internal Rows Rows
        {
            get { return _rows; }
        }

        internal Cells Cells
        {
            get { return _cells; }
        }
    }

    public class Sheets : List<Sheet>
    {
        Spreadsheet _document = null;

        internal Sheets(Spreadsheet document)
        {
            _document = document;
        }

        internal void Deleted()
        {
            foreach (var sheet in this)
                sheet.Deleted();
            Clear();
            _document = null;
        }

        public Sheet Sheet(string name)
        {
            if (this.Count(r => r.Name == name)==0)
                Add(new Sheet(this, name));
            return this.First(r => r.Name == name);
        }

        public Sheet Sheet(int index)
        {
            if (index < 0 || index >= this.Count)
                return null;
            return this[index];
        }
    }
}
