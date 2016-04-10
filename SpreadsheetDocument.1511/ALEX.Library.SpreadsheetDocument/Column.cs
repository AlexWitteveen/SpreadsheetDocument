using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ALEX.Library.SpreadsheetDocument
{
    public class Column
    {
        internal Columns _columns = null;
        internal Cells _cells = null;
        internal int _index = -1;
        internal bool _hidden = false;

        internal Column(Columns columns, int index)
        {
            _columns = columns;
            _index = index;
            _cells = new Cells(null, this, null);
        }

        internal void Deleted()
        {
            _cells.Deleted();
            _columns = null;
        }

        public void Clear()
        {
            _columns._sheet.Clear(_cells);
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

        public Cell Cell(int rowIndex)
        {
            return _cells.Cell(this, rowIndex);
        }

        public Cell Cell(Row row)
        {
            return _cells.Cell(this, row);
        }

        public List<Cell> Cells
        {
            get
            {
                return _cells.OrderBy(r => r.Row.Index).ToList();
            }
        }
    }

    public class Columns : List<Column>
    {
        internal Sheet _sheet = null;

        internal Columns(Sheet sheet)
        {
            _sheet = sheet;
        }

        internal void Deleted()
        {
            this.Clear();
            _sheet = null;
        }

/*
        public new void Clear()
        {
            foreach (var column in this)
                column.Deleted();
            base.Clear();
        }
*/
        internal void Clear(Cells cells)
        {
            foreach (var column in this)
                column.Clear(cells);
        }

/*
        public Sheet Sheet
        {
            get { return _sheet; }
        }
*/
        public Column Column(int index)
        {
            if (this.Count(r => r.Index == index) == 0)
                Add(new Column(this, index));
            return this.First(r => r.Index == index);
        }

        internal int MinColumnIndex()
        {
            return this.Min(r => r.Index);
        }

        internal int MaxColumnIndex()
        {
            return this.Max(r => r.Index);
        }

    }
}
