using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EvoExcelAddIn.GenericExtension;
using Microsoft.Office.Interop.Excel;

namespace EvoExcelAddIn.Core
{
    public class ExcelManager
    {
        public static Workbook _workBook;
        public static Worksheet _excelSheet;

        public static void ActivateExcel(string sheetName = null, bool reset = false)
        {
            var nativeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (nativeWorkbook != null)
            {
                _workBook = nativeWorkbook;
            }

            if (string.IsNullOrEmpty(sheetName))
            {
                _excelSheet = nativeWorkbook.ActiveSheet;
            }
            else
            {
                _excelSheet = nativeWorkbook.GetWorksheetByName(sheetName);

                if (_excelSheet == null)
                {
                    _excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)nativeWorkbook.Worksheets.Add();
                    _excelSheet.Name = sheetName;
                }
            }

            if (reset == true)
                _excelSheet.Cells.Clear();
        }

        public static bool SheetExists(string sheetName)
        {
            var nativeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (nativeWorkbook != null)
            {
                _workBook = nativeWorkbook;
            }

            return _workBook.GetWorksheetByName(sheetName) != null ? true : false;
        }

        public static int[] GetLastRowCol()
        {
            Range last = _excelSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            bool isMerged = (bool)last.MergeCells;
            if (isMerged)
            {
                last.UnMerge();
                last = _excelSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            }
            return new int[2] { last.Row, last.Column };
        }
    }
}
