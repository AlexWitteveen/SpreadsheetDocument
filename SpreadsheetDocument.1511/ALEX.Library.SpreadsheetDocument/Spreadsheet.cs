using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace ALEX.Library.SpreadsheetDocument
{
    public class Spreadsheet
    {
        Sheets _sheets = null;
        CellFormats _cellFormats = null;

        public Spreadsheet()
        {
            Init();
        }

        public void Clear()
        {
            Deleted();
            Init();
        }

        internal void Init()
        {
            _sheets = new Sheets(this);
            _cellFormats = new CellFormats(this);
        }

        internal void Deleted()
        {
            foreach (var sheet in _sheets)
                sheet.Deleted();
            _sheets = null;
            _cellFormats.Deleted();
            _cellFormats = null;
        }

        public bool ContainsSheet(string name)
        {
            name = name.ToLower();
            return _sheets.Count(r=>r.Name.ToLower() == name)>0;
        }

        public Sheet Sheet(string name)
        {
            return _sheets.Sheet(name);
        }

        public Sheet Sheet(int index)
        {
            return _sheets.Sheet(index);
        }

        public CellFormats CellFormats()
        {
            return _cellFormats;
        }

        public Sheets Sheets
        {
            get { return _sheets; }
        }

        public void Save(string path)
        {
            ExcelWriter.Write(this, path);
        }

        public void Load(string path)
        {
            ExcelReader.Read(this, path);
        }
    }
}
