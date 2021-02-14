using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EvoExcelAddIn.GenericExtension
{
    public static class WorkbookExtensions
    {
        public static Worksheet GetWorksheetByName(this Workbook workbook, string name)
        {
            return workbook.Worksheets.OfType<Worksheet>().FirstOrDefault(ws => ws.Name == name);
        }
    }
}
