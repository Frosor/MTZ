using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTZ
{
    public class ExcelData
    {
        public ExcelData(bool goOn, Application app, Workbook wb, Worksheet ws)
        {
            this.goOn = goOn;
            this.app = app;
            this.wb = wb;
            this.ws = ws;
        }
        public ExcelData()
        {

        }
        public bool goOn;
        public Application app;
        public Workbook wb;
        public Worksheet ws;
    }
}
