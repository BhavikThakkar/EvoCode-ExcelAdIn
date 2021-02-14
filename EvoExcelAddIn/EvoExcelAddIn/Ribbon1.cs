using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Data;
using System.ComponentModel;
using EvoExcelAddIn.Core;
using Comments = EvoExcelAddIn.Model.Comments;
using System.Windows.Forms;
using EvoExcelAddIn.Forms;

namespace EvoExcelAddIn
{
    public partial class EvoTrial
    {
        Microsoft.Office.Interop.Excel.Workbook nativeWorkbook;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private async void btnGetdata_ClickAsync(object sender, RibbonControlEventArgs e)
        {
            var dataObj = await GetData();
            ExcelWriter<Comments> myExcel = new CommentWriter();
            myExcel.WriteDateToExcel("DATA", dataObj.ToList(), "A1", "C1");

            //nativeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //if (nativeWorkbook != null)
            //{
            //    Microsoft.Office.Tools.Excel.Workbook vstoWorkbook =
            //        Globals.Factory.GetVstoObject(nativeWorkbook);
            //}

            //Worksheet newWorksheet = nativeWorkbook.GetWorksheetByName("DATA");

            //if (newWorksheet != null)
            //{
            //    newWorksheet.Cells.Clear();
            //}
            //else
            //{
            //    newWorksheet = (Worksheet)nativeWorkbook.Worksheets.Add();
            //    newWorksheet.Name = "DATA";
            //}

            //var dataObj = await GetData();

            //System.Data.DataTable data = (System.Data.DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(dataObj), (typeof(System.Data.DataTable)));

            //string[,] importString = new string[data.Rows.Count, data.Columns.Count];
            ////populate the string[,] however you can
            //for (int r = 0; r < data.Rows.Count; r++)
            //{
            //    for (int c = 0; c < data.Columns.Count; c++)
            //    {
            //        importString[r, c] = data.Rows[r][c].ToString();
            //    }
            //}

            //Range oRange = newWorksheet.Range[newWorksheet.Cells[1, 1],
            //            newWorksheet.Cells[data.Rows.Count, data.Columns.Count]];
            //oRange.Value = importString;

        }

        public async Task<IList<Comments>> GetData()
        {
            using (var client = new HttpClient())
            {
                HttpResponseMessage response = client.GetAsync("https://jsonplaceholder.typicode.com/comments").Result;  // Blocking call!  
                if (response.IsSuccessStatusCode)
                {
                    // Get the response
                    var customerJsonString = await response.Content.ReadAsStringAsync();
                    Console.WriteLine("Your response data is: " + customerJsonString);

                    // Deserialise the data (include the Newtonsoft JSON Nuget package if you don't already have it)
                    IList<Comments> deserialized = JsonConvert.DeserializeObject<IList<Comments>>(custome‌​rJsonString);

                    return deserialized;
                }
            }

            return null;
        }

        private void btnLinkdata_Click(object sender, RibbonControlEventArgs e)
        {

            if (!ExcelManager.SheetExists("DATA"))
            {
                MessageBox.Show("Data Worksheet does not exists, try to get data first & try again.");
                return;
            }

            ExcelManager.ActivateExcel("DATA",false);
            var last = ExcelManager.GetLastRowCol();
            string lastRow = "A" + Convert.ToInt32(last[0]-1);

            ExcelManager.ActivateExcel(null,true);

            if (ExcelManager._excelSheet.Name == "DATA")
            {
                MessageBox.Show("You can not reference same cells for formulaes, please try with new worksheet.");
                return;
            }

            var _excelRange = ExcelManager._excelSheet.get_Range("A1", lastRow);
            _excelRange.Formula = "=DATA!A1";
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            About frmAbout = new About();
            frmAbout.ShowDialog();
        }
    }




}
