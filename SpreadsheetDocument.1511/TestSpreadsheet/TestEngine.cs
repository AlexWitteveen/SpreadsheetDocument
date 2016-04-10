using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


using ALEX.Library.SpreadsheetDocument;

namespace TestSpreadsheet
{
    class TestEngine
    {
        public void Test1(MainWindow main)
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Excel 1.0.xlsx");

            var doc = new Spreadsheet();
            doc.Load(path);

            main.WriteLog($"Sheet count\t: {doc.Sheets.Count}");
        }
    }
}
