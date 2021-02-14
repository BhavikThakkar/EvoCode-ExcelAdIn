using EvoExcelAddIn.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace EvoExcelAddIn
{
    class CommentWriter : ExcelWriter<Comments>
    {
        private object[,] commentsData;
        private int rows,cols;

        public override object[] Headers
        {
            get
            {
                object[] headerName = {  "Id", "Name", "Email" };
                return headerName;
            }
        }

        public override object[,] ExcelData
        {
            get
            {
                return commentsData;
            }
        }

        public override int ColumnCount
        {
            get
            {
                return cols;
            }
        }

        public override int RowCount
        {
            get
            {
                return rows;
            }
        }

        public override void FillRowData(List<Comments> list)
        {

            System.Data.DataTable data = (System.Data.DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(list), (typeof(System.Data.DataTable)));

            rows = data.Rows.Count;
            cols = data.Columns.Count;
            string[,] importString = new string[data.Rows.Count, data.Columns.Count];

            for (int r = 0; r < data.Rows.Count; r++)
            {
                for (int c = 0; c < data.Columns.Count; c++)
                {
                    importString[r, c] = data.Rows[r][c].ToString();
                }
            }

            commentsData = importString;
        }
        
    }
}
