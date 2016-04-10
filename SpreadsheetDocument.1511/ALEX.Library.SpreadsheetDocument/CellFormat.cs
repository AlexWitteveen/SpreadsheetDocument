using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ALEX.Library.SpreadsheetDocument
{
    public class CellFormat
    {
        string _fillColor = "FFFFFFF";
        CellFormats _cellFormats = null;
        List<Cell> _cells = new List<Cell>();

        internal CellFormat(CellFormats cellFormats, string fillColor)
        {
            _fillColor = fillColor;
            _cellFormats = cellFormats;
        }

        internal void Deleted()
        {
            _cells.Clear();
            _cellFormats = null;
        }

        public string FillColor
        {
            set { _fillColor = value; }
            get { return _fillColor; }
        }

        internal void AddCell(Cell cell)
        {
            if (!_cells.Contains(cell))
                _cells.Add(cell);
        }

        internal void RemoveCell(Cell cell)
        {
            if (_cells.Contains(cell))
                _cells.Remove(cell);
        }

        internal int Count()
        {
            return _cells.Count;
        }
    }

    public class CellFormats: List<CellFormat>
    {
        Spreadsheet _document = null;

        internal CellFormats(Spreadsheet document)
        {
            _document = document;
        }

        internal void Deleted()
        {
            foreach (var cellFormat in this)
                cellFormat.Deleted();
            _document = null;
        }

        public CellFormat CellFormat(string fillColor)
        {
            if (this.Count(r => r.FillColor == fillColor) == 0)
                this.Add(new CellFormat(this, fillColor));

            return this.First(r => r.FillColor == fillColor);
        }
    }
}
