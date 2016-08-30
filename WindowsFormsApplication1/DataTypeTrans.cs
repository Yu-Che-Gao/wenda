using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    class DataTypeTrans
    {
        public static DataTable dvTodt(DataGridView dv)
        {
            DataTable dt = new DataTable();
            DataColumn dc;
            for (int i = 0; i < dv.Columns.Count; i++)
            {
                try
                {
                    dc = new DataColumn();
                    dc.ColumnName = dv.Columns[i].HeaderText.ToString();
                    dt.Columns.Add(dc);
                }
                catch (System.Data.DuplicateNameException ex)
                {
                    /*
                     * 問題:出現DuplicateNameException:出現重複的欄位名
                     * 解決:在輸出Excel欄位名多加空格
                     */
                    dc = new DataColumn();
                    dc.ColumnName = dv.Columns[i].HeaderText.ToString() + " ";
                    dt.Columns.Add(dc);
                }
            }
            for (int j = 0; j < dv.Rows.Count; j++)
            {
                DataRow dr = dt.NewRow();
                for (int x = 0; x < dv.Columns.Count; x++)
                {
                    dr[x] = dv.Rows[j].Cells[x].Value;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}
