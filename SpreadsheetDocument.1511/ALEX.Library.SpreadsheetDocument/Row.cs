using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ALEX.Library.SpreadsheetDocument
{
    public class Row
    {
        internal Rows _rows = null;
        internal Cells _cells = null;
        internal int _index = -1;
        internal bool _hidden = false;

        internal Row(Rows rows, int index)
        {
            _rows = rows;
            _index = index;
            _cells = new Cells(null, null, this);
        }

        internal void Deleted()
        {
            _cells.Deleted();
            _rows = null;
        }

        public void Clear()
        {
            _rows._sheet.Clear(_cells);
        }

        internal void Clear(Cells cells)
        {
            _cells.Clear(cells);
        }

        public int Index
        {
            get { return _index; }
        }

        public bool Hidden { get { return _hidden; } set { _hidden = value; } }

        public Cell Cell(int columnIndex)
        {
            return _cells.Cell(columnIndex, this);
        }

        public Cell Cell(Column column)
        {
            return _cells.Cell(column, this);
        }

        public List<Cell> Cells
        {
            get
            {
                return _cells.OrderBy(r => r.Column.Index).ToList();
            }
        }

    }

    public class Rows : List<Row>
    {
        internal Sheet _sheet = null;

        internal Rows(Sheet sheet)
        {
            _sheet = sheet;
        }

        internal void Deleted()
        {
            foreach (var row in this)
                row.Deleted();
            this.Clear();
            _sheet = null;
        }

/*
        public new void Clear()
        {
            foreach (var row in this)
                row.Deleted();
            base.Clear();
        }
*/

        internal void Clear(Cells cells)
        {
            foreach (var row in this)
                row.Clear(cells);
        }

/*
        public Sheet Sheet
        {
            get { return _sheet; }
        }
*/
        public Row Row(int index)
        {
            if (this.Count(r => r.Index == index) == 0)
                Add(new Row(this, index));
            return this.First(r => r.Index == index);
        }

        public int MinRowIndex()
        {
            return this.Min(r => r.Index);
        }

        public int MaxRowIndex()
        {
            return this.Max(r => r.Index);
        }
    }
}
