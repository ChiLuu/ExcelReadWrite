using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReadWrite
{
    public class Excel
    {
        _Application excel = new _Excel.Application();

        public string Path { get; set; }
        public Workbook Workbook { get; set; }
        public Worksheet Worksheet { get; set; }

        public Excel (){}

        public void InitializeExcel(string path)
        {
            this.Path = path;
            this.Workbook = excel.Workbooks.Open(path);
        }

    }
}
